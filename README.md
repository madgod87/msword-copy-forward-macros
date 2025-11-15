Word Forward Macros — README

A complete, ready-to-use Microsoft Word VBA macro system for managing a persistent, numbered “Copy forwarded to:” list.
Includes a persistent dataset saved at %APPDATA%\ForwardList\WordItemsDataset.txt, an advanced editor UI (add / edit / reorder / multi-delete / renumber), and the insertion macro that writes properly numbered lines into Word.

This README shows exactly how to import the code, how to build the UserForms (controls + properties), gives a visual demo layout you can follow, and a step-by-step user guide so others can set it up and use it without guesswork.

Table of contents

Overview

Repo files (what each file does)

Prerequisites

Quick install (summary)

Full install — step by step (modules, class, forms)

Module: ModuleForwardList.bas

Class: AppEventHandler.cls

UserForm1 (selection dialog)

UserForm2 (additional DM options)

AdvancedEditorForm (full editor) — layout + properties + ASCII demo

How to use (day-to-day)

Launch editor

Add / edit / delete / reorder items

Insert forwarding list into a document

Reset dataset / backup / restore

Special-case behaviors (ADM, Joint BDO, Gram Panchayat, “To … For Compliance.”)

Troubleshooting (common errors & fixes)

Contributing / tips for maintainers

License

1 — Overview

This project provides:

A persistent dataset of numeric-key -> text entries (no sub-keys like 1.1 — keys are plain integers).

An Advanced Editor UI to manage the dataset (inline edit, add at key, multi-delete, move up/down, renumber 1..N).

A selection form to choose multiple recipients and insert a correctly numbered forwarding list into the current Word document.

Autosave & autosync on Word open / close via an application event handler.

Data file path (default):

%APPDATA%\ForwardList\WordItemsDataset.txt

2 — Repo files (what to import)

Place the following files into src/ in your repo. For Word import, copy/paste or import .bas/.cls — for forms create empty forms and paste their code.

src/ModuleForwardList.bas — main module: persistence, dataset operations, insertion macro, helpers.

src/AppEventHandler.cls — class module (set class name to AppEventHandler) to hook Word events and autosave/load.

src/UserForm1.txt — code for the selection dialog (ListBox of items; OK/Cancel).

src/UserForm2.txt — code for multi-choice Additional DM options.

src/AdvancedEditorForm.txt — the advanced editor form code (create an empty UserForm named AdvancedEditorForm and paste code).

README.md — this file.

LICENSE — choose license (MIT recommended).

.gitignore

3 — Prerequisites

Windows + Microsoft Word with VBA (desktop version).

Basic ability to open the VBA editor (Alt+F11).

No external libraries required — code uses late-bound Scripting.Dictionary (no reference required).

4 — Quick install (summary)

Clone repo or copy src files locally.

In Word: Alt+F11 → Insert → Module → paste ModuleForwardList.bas.

Insert → Class Module → set (Name) to AppEventHandler → paste AppEventHandler.cls.

Insert → UserForm(s):

UserForm1 → add controls (see step 5).

UserForm2 → add controls (see step 5).

AdvancedEditorForm → create controls exactly as described below and paste code.

Save project (if in Normal.dotm save Normal.dotm).

Run InitAppEventHandler (or restart Word).

5 — Full install — step by step

Tip: You can import .bas and .cls files with File → Import File in VBA editor if you saved them from this repo. For forms, create the form manually and paste code into each form's code window.

5.1 Paste Module (ModuleForwardList)

Insert → Module

Paste the entire contents of ModuleForwardList.bas (this file contains the dataset logic, persistence, helper functions, and macros such as ShowSelectionFormAndInsert, ShowAdvancedEditor, AddDataItem, MoveDataItem, ResetDataset).

Save.

Important: The module declares Public itemsDict As Object and Public gAppHandler As AppEventHandler — keep these names unchanged.

5.2 Add Class Module (AppEventHandler)

Insert → Class Module

In the Properties window change (Name) to AppEventHandler

Paste the content of AppEventHandler.cls.

Save.

This class hooks Application events and autosaves the dataset on close.

5.3 Create UserForm1 (selection dialog)

Create Form:

Insert → UserForm → keep name UserForm1.

Add Controls:

ListBox1 — ListBox (Name: ListBox1)

OKButton — CommandButton (Name: OKButton, Caption: OK)

CancelButton — CommandButton (Name: CancelButton, Caption: Cancel)

Properties:

For ListBox1 you do not need to set ColumnCount here (code sets it before showing). Optionally set MultiSelect = fmMultiSelectMulti.

Paste Code:

Paste contents of UserForm1.txt into the form's code window.

5.4 Create UserForm2 (Additional DM options)

Create Form:

Insert → UserForm → name UserForm2.

Add Controls:

ListBox2 — ListBox (Name: ListBox2)

CommandButton3 — CommandButton (Name: CommandButton3, Caption: OK)

CommandButton4 — CommandButton (Name: CommandButton4, Caption: Cancel)

Properties:

ListBox2.MultiSelect = fmMultiSelectMulti (set in code too)

Paste Code:

Paste contents of UserForm2.txt.

5.5 Create AdvancedEditorForm (full editor)

This is the main UI for editing the dataset.

Create Form:

Insert → UserForm → in Properties set (Name) = AdvancedEditorForm, Caption = Advanced Editor.

Add Controls (exact names required):

ListBox1 — ListBox

(Name) = ListBox1

Left = 12, Top = 12, Width = 380, Height = 270 (approx)

MultiSelect = fmMultiSelectMulti

ColumnCount = 2

ColumnWidths = "320 pt;0 pt" (second column hidden stores numeric key)

txtInline — TextBox

(Name) = txtInline

Left = 12, Top = 288, Width = 380, Height ≈ 24

CommandButtons (all: Width ≈ 100, Height ≈ 28)

btnAdd Caption: Add Left: 404 Top: 12

btnMoveUp Caption: Move Up Left: 404 Top: 48

btnMoveDown Caption: Move Down Left: 404 Top: 84

btnDelete Caption: Delete Left: 404 Top: 120

btnEdit Caption: Apply Edit Left: 404 Top: 156

btnSaveOrder Caption: Save Order Left: 404 Top: 192

btnRefresh Caption: Refresh Left: 404 Top: 228

btnClose Caption: Close Left: 404 Top: 264

Optional: lblStatus — Label (Name: lblStatus)

Left = 12, Top = 320, Width = 380, Height = 20, WordWrap = True, Caption = ""

If you don't add this, the code still works (but we recommended earlier to include it).

Paste Code:

Paste the full AdvancedEditorForm code (the cleaned, robust version included in the repo). That code expects the exact control names above.

5.6 Save & initialize

Save the VBA project.

Run macro InitAppEventHandler (from the Macros list) once to initialize and create the data file if missing, or restart Word to run AutoExec.

6 — How to use (user guide)
6.1 Launching the Advanced Editor (manage dataset)

Macro: ShowAdvancedEditor

Purpose: view all items, add new items, edit value inline, delete multiple items, re-order items, renumber keys.

Basic flows

Add an item

Click Add.

You can enter a numeric key to insert the item at that key (existing items ≥ key will be shifted up by 1), or leave blank to append at the end.

Enter the display text when prompted.

Editor refreshes and saves automatically.

Inline edit value

Single-click an item in the list.

The item's value appears in txtInline.

Edit text in txtInline.

Click Apply Edit (btnEdit). The value saves to the dataset file.

Delete items (multi-select)

Select one or more items (hold Ctrl or Shift).

Click Delete. Confirm prompt. A timestamped backup file is created automatically.

Items removed and autosaved.

Move Up / Move Down

Select one or more items (they move as a block preserving relative order).

Click Move Up or Move Down.

Click Save Order to renumber keys 1..N according to current list order (recommended if you want compact sequential keys). The form also automatically rebuilds and saves after certain operations.

Refresh

Click Refresh to reload the list from memory (useful if you edited the file externally).

6.2 Insert forwarding list into a Word document

Macro: ShowSelectionFormAndInsert

Place the insertion cursor in Word at the start of a line (important).

Run macro ShowSelectionFormAndInsert.

A selection dialog (UserForm1) opens with key - value list.

Select the items you want (multi-select supported). Click OK.

For special items the macro may prompt:

Additional DM: will open UserForm2 to select subtypes (Gen, Dev, LR, ZP) — supports multi-select.

Joint BDO: prompts for number of joint BDOs. If you enter 1, prints 12) style; if >1 prints 12-15) range and increments numbering accordingly.

(All Gram Panchayat) or Shri/Smt: prompts for a count; if 1 it prints 12) not 12-12).

To … For Compliance.: prompts for count and prints each separately (one line per count).

The macro inserts header:

Copy forwarded to for information:


and then writes numbered entries starting 1) 2) etc, observing ranges and single-case formatting.

6.3 Backup and reset

Automatic backups: on delete operations a timestamped backup WordItemsDataset_backup_YYYY-MM-DD_HHMMSS.txt is created in %APPDATA%\ForwardList.

Reset to defaults

Run macro ResetDataset. This backs up the current dataset and recreates the default dataset file.

View current dataset

Run macro ListDatasetToImmediate to print the dataset (numeric order) to the Immediate Window (open by Ctrl+G in VBA editor).

7 — Special-case behaviors (details)

Single vs range numbering
Single counts print 12) (not 12-12)), ranges print 12-15) only when count > 1.

Additional District Magistrate
Selecting this item in the insertion dialog opens UserForm2 to select subtypes (Gen / Dev / LR / ZP). If you select multiple subtypes they are presented together and numbering increments per subtype selected.

Joint Block Development Officer
Prompts for a count. If count > 1, macro prints a range like 12-14) and increments the global sequence accordingly.

(All Gram Panchayat) and Shri/Smt
Prompts for a count. If 1, prints 12); if >1 prints 12-... range.

To ... For Compliance.
This new item prints each requested count as a separate line (no ranges).

8 — Troubleshooting (common errors & fixes)

Compile error: variable not defined — lblStatus
Fix: add a label named lblStatus to AdvancedEditorForm or remove code lines referring to it. Recommended: add lblStatus (Label) with properties shown earlier.

Compile error: user-defined type not defined (for AppEventHandler)
Make sure class module (Name) is set to AppEventHandler (not the default Class1). Then ensure Public gAppHandler As AppEventHandler is in the standard module.

Textbox (txtInline) not populating after selection
Ensure ListBox1.ColumnCount = 2 and ColumnWidths includes a hidden second column such as "320 pt;0 pt". The second column stores the numeric key.

Inserted numbering is wrong / items reorder unexpectedly
Use the editor Save Order to renumber keys sequentially after reordering. Avoid manual edits of the data file unless you know the format (key|value per line).

Where is the dataset file?
%APPDATA%\ForwardList\WordItemsDataset.txt — you can open it with a text editor. Each line is:

key|value (with '|' escaped as '||' in values)


I deleted the file accidentally
Look for backups in %APPDATA%\ForwardList\WordItemsDataset_backup_*.txt; copy a backup back to WordItemsDataset.txt.

Form controls not found error on compile
Confirm controls exist and names match exactly. The AdvancedEditorForm expects ListBox1, txtInline, btnAdd, btnEdit, btnDelete, btnMoveUp, btnMoveDown, btnSaveOrder, btnRefresh, btnClose. If lblStatus is referenced, add it or remove references.

9 — Contributing / maintainers guide

If you edit code, follow the modular layout: keep persistence helpers in ModuleForwardList.bas and UI logic in forms.

To add new special-case behaviors (e.g., new item types with unique prompts/numbering), modify ShowSelectionFormAndInsert in ModuleForwardList.bas — find the If InStr(valuePart, "...") blocks.

If you want to export the forms programmatically from Word, you can use VBA to write .frm/.bas text files, but manual paste/import is reliable across Office versions.

10 — License

This repo includes an MIT license template file. Replace <Your Name> and the year in LICENSE with your details.

Appendices
A — Example: data file format

Example lines inside %APPDATA%\ForwardList\WordItemsDataset.txt:

1|The District Magistrate, Nadia.
2|The District Magistrate and District Election Officer, Nadia.
23|To ........................... For Compliance.
24|Shri/Smt ………………………… For Compliance.


Note: | inside values is escaped as || in the file.

B — ASCII demo of AdvancedEditorForm layout (copy/paste into form designer)
+--------------------------------------------------------------+
| Advanced Editor                                              |
| +----------------------------------------------------------+ |
| | ListBox1 (key - value)                                   | |
| |  (multi-select, 2 columns, large area)                   | |
| |                                                          | |
| +----------------------------------------------------------+ |
| txtInline: [..............................................]  |
| lblStatus: (Loaded 25 items.)                                |
|                                                              |
|  [Add]    [Move Up]   [Move Down]   [Delete]   [Apply Edit]  |
|  [Save Order] [Refresh] [Close]                              |
+--------------------------------------------------------------+


Exact coordinates (recommended)

Form: Width=520, Height=420

ListBox1: Left=12, Top=12, Width=380, Height=270

txtInline: Left=12, Top=288, Width=380, Height=24

lblStatus: Left=12, Top=320, Width=380, Height=20

Buttons column start: Left=404, Top increments 12, 48, 84, 120... (Width=100, Height=28)