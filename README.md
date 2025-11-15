# Word Forward Macros

VBA macros and UserForms to manage a persistent "forward/copy list" in Microsoft Word.
The macros maintain a numeric-key -> text dataset and let you insert numbered forwarding lines into Word documents. Includes an advanced editor UI.

## Files

- `src/ModuleForwardList.bas` — main module (dataset, persistence, commands)
- `src/AppEventHandler.cls` — class module to handle Word application events (autosave)
- `src/UserForm1.txt` — code for selection form (paste into a UserForm)
- `src/UserForm2.txt` — code for additional options form (paste into a UserForm)
- `src/AdvancedEditorForm.txt` — advanced editor form code (paste into a UserForm)
- `README.md`, `LICENSE`, `.gitignore`

## Installation (developer)
1. Open Word → Alt+F11 to open VBA editor.
2. Insert → Module, paste `ModuleForwardList.bas` code.
3. Insert → Class Module, set `(Name)` to `AppEventHandler`, paste the `.cls` code.
4. Insert → UserForm, name and add controls then paste code from `UserForm1.txt`, `UserForm2.txt`, `AdvancedEditorForm.txt`.
5. Save the project (if storing as a global add-in, save in `Normal.dotm`).
6. Run `InitAppEventHandler` once (or restart Word).
7. Use `ShowAdvancedEditor` to manage dataset, and `ShowSelectionFormAndInsert` to insert lines.

## Persistence
The dataset is saved at: `%APPDATA%\ForwardList\WordItemsDataset.txt`

## License
MIT — see LICENSE file.
