import os
import sys
import glob
import shutil

try:
    import win32com.client
except ImportError:
    print("Error: pywin32 is not installed. Please run 'pip install -r requirements.txt'")
    sys.exit(1)

def build_dotm():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.join(base_dir, 'src', 'modules')
    dotm_path = os.path.join(base_dir, 'dist', 'IIDelta_MW_Tools.dotm')

    if not os.path.exists(dotm_path):
        print(f"Error: Base template not found at {dotm_path}")
        sys.exit(1)

    print("Starting Microsoft Word...")
    try:
        word = win32com.client.Dispatch("Word.Application")
    except Exception as e:
        print(f"Failed to dispatch Word.Application. Ensure you are running this on a Windows machine with Microsoft Word installed.\nError details: {e}")
        sys.exit(1)
        
    word.Visible = False
    
    try:
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
        
        # Deploy to Word STARTUP folder
        appdata = os.environ.get('APPDATA')
        if appdata:
            startup_dir = os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP')
            os.makedirs(startup_dir, exist_ok=True)
            dest_path = os.path.join(startup_dir, 'IIDelta_MW_Tools.dotm')
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
