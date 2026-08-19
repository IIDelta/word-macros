import os
import sys
import glob
import shutil
import zipfile
import tempfile

try:
    import win32com.client
    import ctypes
except ImportError:
    print("Error: pywin32 is not installed. Please run 'pip install -r requirements.txt'")
    sys.exit(1)

def inject_custom_ui(dotm_path, custom_ui_path):
    print("Injecting Custom Ribbon UI...")
    if not os.path.exists(custom_ui_path):
        print(f"Warning: Custom UI XML not found at {custom_ui_path}")
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(dotm_path, 'r') as zin:
            zin.extractall(tmpdir)
        
        custom_ui_dir = os.path.join(tmpdir, 'customUI')
        os.makedirs(custom_ui_dir, exist_ok=True)
        shutil.copy2(custom_ui_path, os.path.join(custom_ui_dir, 'customUI14.xml'))
        
        rels_path = os.path.join(tmpdir, '_rels', '.rels')
        if os.path.exists(rels_path):
            with open(rels_path, 'r', encoding='utf-8') as f:
                rels_content = f.read()
            if 'customUI14.xml' not in rels_content:
                rel_str = '<Relationship Id="customUIRelID" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
                rels_content = rels_content.replace('</Relationships>', f'  {rel_str}\n</Relationships>')
                with open(rels_path, 'w', encoding='utf-8') as f:
                    f.write(rels_content)
        
        with zipfile.ZipFile(dotm_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, _, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    zout.write(file_path, arcname)
    print("Custom Ribbon UI injected successfully.")

def build_dotm():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.join(base_dir, 'src', 'modules')
    dotm_path = os.path.join(base_dir, 'dist', 'MedicalWritingTools.dotm')
    custom_ui_path = os.path.join(base_dir, 'src', 'customUI', 'customUI.xml')

    print("Starting Microsoft Word...")
    try:
        word = win32com.client.Dispatch("Word.Application")
    except Exception as e1:
        print(f"Initial COM connection failed: {e1}")
        print("Attempting to recover by killing hung background Word processes and clearing COM cache...")
        import platform
        if platform.system() == "Windows":
            os.system("taskkill /F /IM WINWORD.EXE /T >nul 2>&1")
        import time
        time.sleep(2)
        
        # Clear win32com cache
        try:
            import tempfile
            gen_py_path = os.path.join(tempfile.gettempdir(), 'gen_py')
            if os.path.exists(gen_py_path):
                shutil.rmtree(gen_py_path)
        except:
            pass
            
        try:
            import pythoncom
            pythoncom.CoInitialize()
            word = win32com.client.dynamic.Dispatch("Word.Application")
        except Exception as e2:
            print(f"\n[FATAL ERROR] Failed to dispatch Word.Application even after recovery.\nError details: {e2}")
            sys.exit(1)
        
    word.Visible = False
    word.AutomationSecurity = 1 # msoAutomationSecurityLow (enables macros without prompts during automation)
    
    try:
        print(f"Creating new base template at {dotm_path}...")
        if os.path.exists(dotm_path):
            try:
                os.remove(dotm_path)
            except OSError:
                pass # If we can't remove it, let SaveAs2 try to overwrite or fail
        doc = word.Documents.Add()
        doc.SaveAs2(dotm_path, FileFormat=15) # 15 = wdFormatXMLTemplateMacroEnabled
            
        # Remove existing standard modules
        vb_comp = doc.VBProject.VBComponents
        for comp in list(vb_comp): # Convert to list to iterate safely while modifying
            if comp.Type == 1: # 1 = vbext_ct_StdModule
                print(f"Removing old module {comp.Name}")
                vb_comp.Remove(comp)
        
        # Import new modules
        for bas_file in glob.glob(os.path.join(src_dir, '*.bas')):
            print(f"Importing {os.path.basename(bas_file)}...")
            vb_comp.Import(bas_file)
            
        print("Saving document...")
        doc.Save()
        doc.Close()
        del vb_comp
        del doc
        print("Build successful!")
    except Exception as e:
        err_msg = str(e)
        print(f"Build failed: {err_msg}")
        if "VBProject" in err_msg or "Trust" in err_msg or "800a175d" in err_msg:
            msg = ("Trust access to the VBA project object model is required.\n\n"
                   "Please open Word, go to:\n"
                   "File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings\n"
                   "and check 'Trust access to the VBA project object model'.")
            print(msg)
            try:
                ctypes.windll.user32.MessageBoxW(0, msg, "Build Error: Macro Security", 0x10)
            except:
                pass
        sys.exit(1)
    finally:
        try:
            word.Quit()
            del word
            import gc
            gc.collect()
        except:
            pass

    # Now that Word is fully closed and locks are released, we can inject UI and deploy
    try:
        import time
        time.sleep(1.5) # Allow OS to fully release file locks

        # Inject Custom UI
        inject_custom_ui(dotm_path, custom_ui_path)
        
        # Deploy to Word STARTUP folder
        appdata = os.environ.get('APPDATA')
        if appdata:
            startup_dir = os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP')
            os.makedirs(startup_dir, exist_ok=True)
            dest_path = os.path.join(startup_dir, 'MedicalWritingTools.dotm')
            
            deployed = False
            while not deployed:
                try:
                    shutil.copy2(dotm_path, dest_path)
                    print(f"Auto-deployed successfully to: {dest_path}")
                    deployed = True
                except PermissionError as deploy_pe:
                    print(f"\n[Deploy Error] The destination file '{dest_path}' is locked.")
                    print("This happens when Microsoft Word or Outlook is running in the background.")
                    ans = input("Would you like to forcefully close all Word processes now to complete deployment? (y/n): ")
                    if ans.lower().strip() == 'y':
                        print("Forcefully closing WINWORD.EXE...")
                        os.system("taskkill /F /IM WINWORD.EXE /T >nul 2>&1")
                        import time
                        time.sleep(2) # Give Windows a moment to release file handles
                    else:
                        msg = ("Deploy failed: The destination file in your STARTUP folder is locked.\n\n"
                               "The local build in /dist/ was successful.\n\n"
                               "Opening the STARTUP folder for you now. You can manually copy 'dist\\MedicalWritingTools.dotm' there.")
                        print(f"\n{msg}\n\nManual installation path: {startup_dir}\nError details: {deploy_pe}")
                        try:
                            os.startfile(startup_dir)
                        except Exception as e:
                            print(f"Could not open file explorer: {e}")
                        
                        try:
                            ctypes.windll.user32.MessageBoxW(0, msg, "Deploy Error: STARTUP Folder Locked", 0x30)
                        except:
                            pass
                        break
            
    except PermissionError as pe:
        msg = ("Build failed: The local template in /dist/ is locked.\n\n"
               "Please fully close Microsoft Word (and ensure no instances are running in the background) before running the build.")
        print(f"Build failed: {msg}\nError details: {pe}")
        try:
            ctypes.windll.user32.MessageBoxW(0, msg, "Build Error: File Locked", 0x10)
        except:
            pass
    except Exception as e:
        print(f"Post-build step failed: {e}")

if __name__ == "__main__":
    build_dotm()
