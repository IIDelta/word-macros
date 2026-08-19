# IIDelta Medical Writing Tools

A centralized Word Add-in (`.dotm`) containing a suite of highly optimized VBA macros and a custom Ribbon tab designed to automate formatting, visual redlining, and document cleanup for regulatory submissions.

## 🚀 Features (V1)
- **Yellow Highlight Redline**: Converts tracked changes to yellow highlighted text.
- **Standard Redline**: Converts tracked changes to standard red font and strikethrough.
- **Remove Multiple Spaces**: Instantly cleans up irregular spacing.
- **Delete Hidden Text**: Permanently deletes hidden text.
- **Update Fields**: Refreshes all TOCs, tables, and document fields.
- **Reopen Comments**: Marks all resolved comments as active.

*Note: Large macros include a progress bar and ETA in the Word Status Bar and allow interacting with other documents during execution.*

## 📦 Installation for End Users

1. Close Microsoft Word completely.
2. Double-click the `install_MW_tools.bat` script located in the repository root.
3. Open Microsoft Word. The new **MW Tools** tab will be available on the Ribbon.

---

## 🛠️ Development Architecture

This repository separates raw VBA source code from the compiled Word Add-in for version control. We use Python via COM on Windows to compile and test the macros.

### Folder Structure

*   `/src/modules/`: Contains the raw, plain-text VBA modules (`.bas`). All logic edits happen here.
*   `/src/customUI/`: Contains the `customUI.xml` file defining the Ribbon buttons and callbacks.
*   `/dist/`: Contains the compiled `MedicalWritingTools.dotm` Global Template.
*   `/tests/`: Contains automated test scripts.

### Prerequisites for Building and Testing
You must run these tools on a **Windows machine** with Microsoft Word installed.
```cmd
pip install -r requirements.txt
```

### How to Compile Updates

We use an automated build script to inject the `.bas` files into the binary `.dotm` file.

```cmd
python build.py
```
*If you receive an error, ensure that **"Trust access to the VBA project object model"** is enabled in Word (File > Options > Trust Center > Trust Center Settings > Macro Settings).*

### How to Run Automated Tests

The testing suite will launch a hidden Word instance, apply the macros to mock data, and verify the output.

```cmd
python tests/test_macros.py
```