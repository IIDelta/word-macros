# IIDelta Medical Writing Tools

A centralized Word Add-in (`.dotm`) containing a suite of VBA macros and a custom Ribbon tab designed to automate formatting, visual redlining, and document cleanup for regulatory submissions.

## 📦 Installation for End Users

1. Close Microsoft Word completely.
2. Double-click the `install_MW_tools.bat` script located in the repository root.
3. Open Microsoft Word. The new **MW Tools** tab will be available on the Ribbon.

---

## 🛠️ Development Architecture

This repository separates raw VBA source code from the compiled Word Add-in for version control and AI-assisted (Antigravity) development.

### Folder Structure

*   `/src/modules/`: Contains the raw, plain-text VBA modules (`.bas`). All logic edits happen here.
*   `/src/customUI/`: Contains the `customUI.xml` file defining the Ribbon buttons and callbacks.
*   `/dist/`: Contains the compiled `IIDelta_MW_Tools.dotm` Global Template.

### The Component Map

1.  **`Mod_RibbonCallbacks.bas`**: The switchboard. Contains no logic. Only routes Ribbon clicks to target macros.
2.  **`Mod_Redlines.bas`**: Contains the complex algorithms for converting Tracked Changes into static visual formatting (Yellow Highlight and Standard Redline).
3.  **`Mod_Cleanup.bas`**: Houses document sterilization tools (Delete Hidden Text, Remove Multiple Spaces, Update Fields).
4.  **`Mod_Review.bas`**: Houses reviewer tools (Reopen Comments).

### How to Compile Updates

Because VBA cannot be compiled via the CLI, use this workflow:
1. Update the `.bas` code in `/src/modules/`.
2. Open `/dist/IIDelta_MW_Tools.dotm` in Word.
3. Open the VBA Editor (`Alt + F11`), remove the old module, and import the newly updated `.bas` file.
4. Save the `.dotm` file and commit to Git.