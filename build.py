import os
import sys
import glob
import shutil
import tempfile
import time

try:
    import win32com.client
    import ctypes
except ImportError:
    print("Error: pywin32 is not installed. Please run 'pip install -r requirements.txt'")
    sys.exit(1)

def build_dotm():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.join(base_dir, 'src', 'modules')
    dotm_path = os.path.join(base_dir, 'dist', 'MedicalWritingTools.dotm')

    print("Connecting to Microsoft Word...")
    word = None
    try:
        # Try to connect to an existing Word instance, or start a new one
        word = win32com.client.Dispatch("Word.Application")
    except Exception as e1:
        print(f"Initial COM connection failed: {e1}")
        print("Attempting to recover by clearing COM cache...")
        try:
            gen_py_path = os.path.join(tempfile.gettempdir(), 'gen_py')
            if os.path.exists(gen_py_path):
                shutil.rmtree(gen_py_path)
            import pythoncom
            pythoncom.CoInitialize()
            word = win32com.client.dynamic.Dispatch("Word.Application")
        except Exception as e2:
            print(f"\n[FATAL ERROR] Failed to dispatch Word.Application.\nError details: {e2}")
            sys.exit(1)
        
    word.AutomationSecurity = 1 # msoAutomationSecurityLow
    
    # Check if Word was already running by counting documents
    # If it was 0, it might be a background instance. We don't hide it to avoid hiding user docs.
    
    # UNLOAD EXISTING ADD-IN TO RELEASE FILE LOCK
    appdata = os.environ.get('APPDATA')
    startup_dir = os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP') if appdata else ""
    dest_path = os.path.join(startup_dir, 'MedicalWritingTools.dotm') if startup_dir else ""
    
    if dest_path:
        try:
            for addin in word.AddIns:
                if addin.Name == "MedicalWritingTools.dotm":
                    print("Unloading existing add-in to release file lock...")
                    addin.Installed = False
        except Exception as e:
            print(f"Could not check/unload existing add-ins: {e}")

    try:
        print(f"Creating new base template at {dotm_path}...")
        if os.path.exists(dotm_path):
            try:
                os.remove(dotm_path)
            except OSError:
                pass
        
        # Disable screen updating so the user doesn't see flashing if Word is visible
        try:
            word.ScreenUpdating = False
        except:
            pass

        doc = word.Documents.Add(Visible=False)
        doc.SaveAs2(dotm_path, FileFormat=15)
            
        vb_comp = doc.VBProject.VBComponents
        for comp in list(vb_comp):
            if comp.Type == 1:
                print(f"Removing old module {comp.Name}")
                vb_comp.Remove(comp)
        
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
        sys.exit(1)
    finally:
        try:
            word.ScreenUpdating = True
        except:
            pass

    # Deploy
    if dest_path:
        os.makedirs(startup_dir, exist_ok=True)
        deployed = False
        attempts = 0
        
        while not deployed and attempts < 3:
            try:
                shutil.copy2(dotm_path, dest_path)
                print(f"Auto-deployed successfully to: {dest_path}")
                deployed = True
            except PermissionError:
                attempts += 1
                print(f"\n[Deploy] Destination locked. Attempting to forcefully kill WINWORD.EXE without asking (Attempt {attempts}/3)...")
                os.system("taskkill /F /IM WINWORD.EXE /T >nul 2>&1")
                time.sleep(2)
        
        if deployed:
            # Re-load the add-in
            try:
                # Dispatch again in case we killed it
                word = win32com.client.Dispatch("Word.Application")
                
                # Check if it's already in the AddIns collection
                addin_found = False
                for addin in word.AddIns:
                    if addin.Name == "MedicalWritingTools.dotm":
                        print("Re-enabling add-in seamlessly...")
                        addin.Installed = True
                        addin_found = True
                        break
                
                # If not found, add it
                if not addin_found:
                    print("Loading new add-in...")
                    word.AddIns.Add(FileName=dest_path, Install=True)
                    
            except Exception as e:
                print(f"Could not automatically re-enable add-in: {e}")
        else:
            print("Failed to deploy. Please copy manually.")

if __name__ == "__main__":
    build_dotm()
