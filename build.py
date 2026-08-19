import os
import sys
import glob
import shutil
import zipfile
import tempfile

try:
    import win32com.client
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
        shutil.copy2(custom_ui_path, os.path.join(custom_ui_dir, 'customUI.xml'))
        
        rels_path = os.path.join(tmpdir, '_rels', '.rels')
        if os.path.exists(rels_path):
            with open(rels_path, 'r', encoding='utf-8') as f:
                rels_content = f.read()
            if 'customUI.xml' not in rels_content:
                rel_str = '<Relationship Id="customUIRelID" Type="http://schemas.microsoft.com/office/2009/07/customui" Target="customUI/customUI.xml"/>'
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
    except Exception as e:
        print(f"Failed to dispatch Word.Application. Ensure you are running this on a Windows machine with Microsoft Word installed.\nError details: {e}")
        sys.exit(1)
        
    word.Visible = False
    
    try:
        if not os.path.exists(dotm_path) or os.path.getsize(dotm_path) == 0:
            print(f"Creating new base template at {dotm_path}...")
            if os.path.exists(dotm_path):
                os.remove(dotm_path)
            doc = word.Documents.Add()
            doc.SaveAs2(dotm_path, FileFormat=15) # 15 = wdFormatXMLTemplateMacroEnabled
        else:
            print(f"Opening {dotm_path}...")
            doc = word.Documents.Open(dotm_path)
            
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
        print("Build successful!")
        
        # Inject Custom UI
        inject_custom_ui(dotm_path, custom_ui_path)
        
        # Deploy to Word STARTUP folder
        appdata = os.environ.get('APPDATA')
        if appdata:
            startup_dir = os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP')
            os.makedirs(startup_dir, exist_ok=True)
            dest_path = os.path.join(startup_dir, 'MedicalWritingTools.dotm')
            try:
                shutil.copy2(dotm_path, dest_path)
                print(f"Auto-deployed successfully to: {dest_path}")
            except Exception as cp_err:
                print(f"Warning: Could not copy to STARTUP folder. Please ensure Word is closed. Error: {cp_err}")
    except Exception as e:
        print(f"Build failed: {e}")
        print("Note: Ensure 'Trust access to the VBA project object model' is enabled in Word -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings.")
    finally:
        word.Quit()

if __name__ == "__main__":
    build_dotm()
