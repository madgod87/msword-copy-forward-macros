# 📄 Word Forwarding List Manager (VBA for Microsoft Word)

![Logo](logo.png)

A complete Microsoft Word **VBA automation system** for generating and managing a persistent, hierarchical **"Copy forwarded to:"** list.

Includes:

- 🔧 **Persistent dataset** stored in `%APPDATA%`
- 📝 **Advanced Editor UI** (Add / Edit / Delete / Move / Renumber)
- 📥 **Selection form** to insert forwarding lines into Word
- ♻️ **Automatic autosave & auto-load**
- 📌 **Special rules** for ADM, Joint BDO, Gram Panchayat, Compliance items

---

# ⭐ Features Overview

| Feature | Description |
|--------|-------------|
| Persistent Dataset | Stored in `%APPDATA%\ForwardList` |
| Advanced Editor | Full UI for inline edit, reorder, delete, add items |
| Word Insertion Macro | Inserts correctly numbered forwarding list |
| Autosave | Saves dataset automatically on Word close |
| Autosync | Reloads dataset on Word open |
| Backup System | Timestamped backups on delete/reset |
| Clean Architecture | Modules, class handler, and 3 UserForms |

---

# 📦 Repository Layout

```
word-forward-macros/
├─ src/
│  ├─ ModuleForwardList.bas
│  ├─ AppEventHandler.cls
│  ├─ UserForm1.txt
│  ├─ UserForm2.txt
│  └─ AdvancedEditorForm.txt
├─ README.md
├─ LICENSE
└─ .gitignore
```

---

# 🚀 Quick Installation

> For a **full setup guide**, see the section below.

1. Open Microsoft Word → **Alt + F11**
2. Insert → **Module** → paste `ModuleForwardList.bas`
3. Insert → **Class Module** → rename to `AppEventHandler` → paste `AppEventHandler.cls`
4. Insert → **UserForms**
   - `UserForm1` → paste code, add controls
   - `UserForm2` → paste code, add controls
   - `AdvancedEditorForm` → add all required controls and paste code
5. Save Word template or Normal.dotm
6. Run macro: **`InitAppEventHandler`**

---

# 📘 Full Installation Guide

## 1️⃣ Import Main Module
Insert → Module → paste content of `src/ModuleForwardList.bas`.

---

## 2️⃣ Add Application Event Handler
Insert → Class Module → Rename to `AppEventHandler` → Paste `src/AppEventHandler.cls`.

---

## 3️⃣ Create UserForm1 (Selection Form)

### Controls to Add

| Type | Name | Caption | Notes |
|------|------|---------|-------|
| ListBox | `ListBox1` | — | MultiSelect |
| CommandButton | `OKButton` | OK | Saves selection |
| CommandButton | `CancelButton` | Cancel | Clear + close |

Paste: `src/UserForm1.txt`.

---

## 4️⃣ Create UserForm2 (ADM Options)

### Controls to Add

| Type | Name | Caption |
|------|------|---------|
| ListBox | `ListBox2` | — |
| CommandButton | `CommandButton3` | OK |
| CommandButton | `CommandButton4` | Cancel |

Paste: `src/UserForm2.txt`.

---

## 5️⃣ Create AdvancedEditorForm (Main Editor)

### Controls Required

| Control | Name | Purpose |
|--------|------|----------|
| ListBox | `ListBox1` | Shows list of `key - value` |
| TextBox | `txtInline` | Inline editor |
| Label | `lblStatus` | Status messages |
| CommandButton | `btnAdd` | Add item |
| CommandButton | `btnEdit` | Apply edit |
| CommandButton | `btnDelete` | Multi-delete |
| CommandButton | `btnMoveUp` | Move selection up |
| CommandButton | `btnMoveDown` | Move selection down |
| CommandButton | `btnSaveOrder` | Renumber keys |
| CommandButton | `btnRefresh` | Reload dataset |
| CommandButton | `btnClose` | Close editor |

### Suggested Coordinates

```
Form size: Width=520, Height=420

ListBox1:  Left=12, Top=12, Width=380, Height=270
txtInline: Left=12, Top=288, Width=380, Height=24
lblStatus: Left=12, Top=320, Width=380, Height=20

Right column buttons:
btnAdd       Left=404 Top=12
btnMoveUp    Left=404 Top=48
btnMoveDown  Left=404 Top=84
btnDelete    Left=404 Top=120
btnEdit      Left=404 Top=156
btnSaveOrder Left=404 Top=192
btnRefresh   Left=404 Top=228
btnClose     Left=404 Top=264
```

### Layout Diagram

```
+--------------------------------------------------------------+
| Advanced Editor                                              |
| +----------------------------------------------------------+ |
| | ListBox1 (key - value)                                   | |
| +----------------------------------------------------------+ |
| txtInline: [..............................................]  |
| lblStatus: (Loaded X items.)                                 |
|                                                              |
|  [Add]  [Move Up]  [Move Down]  [Delete]  [Apply Edit]       |
|  [Save Order]  [Refresh]  [Close]                            |
+--------------------------------------------------------------+
```

---

# 📚 How to Use

## 🛠 Manage Dataset
Run:
```
ShowAdvancedEditor
```

You can:
- Add items  
- Edit inline  
- Delete  
- Multi-delete  
- Move Up/Down  
- Save Order  

---

## 📄 Insert Forwarding List
Run:
```
ShowSelectionFormAndInsert
```

Macro handles:
- ADM prompts  
- Joint BDO count  
- Gram Panchayat counts  
- “To … For Compliance” (individual lines)  

---

# 🔍 Special-case Behavior Summary

### 🟦 Additional District Magistrate
Options:
```
Gen, Dev, LR, ZP
```

### 🟨 Joint BDO
- If count = 1 → prints `12)`
- If count > 1 → prints range like `12–14)`

### 🟩 Gram Panchayat Groups
Same numbering rules.

### 🟧 “To … For Compliance”
Each entry printed individually:
```
6) To ...
7) To ...
8) To ...
```

---

# 💾 Dataset Persistence

Primary file:
```
%APPDATA%\ForwardList\WordItemsDataset.txt
```

Backups:
```
WordItemsDataset_backup_YYYY-MM-DD_HHMMSS.txt
```

Format:
```
key|value
```
(`|` escaped as `||` inside values)

---

# ❗ Troubleshooting

### 🚫 lblStatus not found
Add a Label named `lblStatus`.

### 🚫 Event handler not firing
Class name **must** be `AppEventHandler`.

### 🚫 Wrong numbering
After moving items, click **Save Order**.

### 🚫 Textbox doesn't fill
Ensure:
```
ListBox1.ColumnCount = 2
ListBox1.ColumnWidths = "320 pt;0 pt"
```

---

# 🤝 Contributing
Pull Requests welcome.

---

# 📄 License
MIT License — see LICENSE.

---

# 🖼 Screenshot placeholders

```
docs/images/editor.png
docs/images/selection-form.png
docs/images/insertion-demo.png
```

---

# ⭐ If you find this useful, star the repository!
