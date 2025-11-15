Attribute VB_Name = "ModuleForwardList"
Option Explicit

' -------------------------
' Module: ModuleForwardList (core dataset + persistence + UI launchers)
' -------------------------
' Configuration
Private Const DATA_FOLDER_NAME As String = "ForwardList"
Private Const DATA_FILE_NAME As String = "WordItemsDataset.txt" ' saved in %APPDATA%\ForwardList

' Public so AppEventHandler class can access it
Public itemsDict As Object ' late-bound Scripting.Dictionary

' Global handler instance (class must be named AppEventHandler)
Public gAppHandler As AppEventHandler

' Initialize dataset
Public Sub InitData()
    Dim fContent As String
    Set itemsDict = CreateObject("Scripting.Dictionary")
    fContent = LoadFromFile(GetDataFilePath())
    If Len(Trim(fContent)) > 0 Then
        DeserializeToDict fContent, itemsDict
    Else
        AddDefaultItems itemsDict
        SaveDictToFile itemsDict, GetDataFilePath()
    End If
End Sub

Public Sub AddDefaultItems(ByRef d As Object)
    If d Is Nothing Then Set d = CreateObject("Scripting.Dictionary")
    d.Add "1", "The District Magistrate, Nadia."
    d.Add "2", "The District Magistrate and District Election Officer, Nadia."
    d.Add "3", "The Returning Officer, 12th Parliamentary Constituency."
    d.Add "4", "The Returning Officer, 13th Parliamentary Constituency."
    d.Add "5", "The Additional District Magistrate (), Nadia."
    d.Add "6", "The Sub-Divisional Officer, Sadar, Nadia."
    d.Add "7", "The Sub-Divisional Officer, Sadar and Assistant Returning Officer, 12th Parliamentary Constituency, Nadia."
    d.Add "8", "The District Panchayat and Rural Development Officer, Nadia."
    d.Add "9", "The Block Development Officer, Krishnagar-I Development Block, Nadia."
    d.Add "10", "The Executive Officer, Krishnagar-I Panchayat Samity, Nadia."
    d.Add "11", "The Joint Block Development Officer, Krishnagar-I Dev. Block, Nadia."
    d.Add "12", "The Prodhan ………………………… (All Gram Panchayat)."
    d.Add "13", "The Upa-Prodhan  ………………………… (All Gram Panchayat)."
    d.Add "14", "The Member  ………………………… (All Gram Panchayat)."
    d.Add "15", "Executive Assistant/Executive Assistant in-charge ………………………… (All Gram Panchayat)."
    d.Add "16", "Secretary/Secretary in-charge ………………………… (All Gram Panchayat)."
    d.Add "17", "Nirman Sahayak/Nirman Sahayak in-charge ………………………… (All Gram Panchayat)."
    d.Add "18", "Sahayak ………………………… (All Gram Panchayat)."
    d.Add "19", "Gram Rozgar Sahayak/Assistant Gram Rozgar Sahayak ………………………… (All Gram Panchayat)."
    d.Add "20", "Skilled Technical Person ………………………… (All Gram Panchayat)."
    d.Add "21", "Villege Level Entrepreneur ………………………… (All Gram Panchayat)."
    d.Add "22", "Gram Panchayat Karmee ………………………… (All Gram Panchayat)."
    d.Add "23", "To ........................... For Compliance."
    d.Add "24", "Shri/Smt ………………………… For Compliance."
    d.Add "25", "Office File."
End Sub

' Add item (interactive)
Public Sub AddDataItem()
    If itemsDict Is Nothing Then InitData
    Dim keyStr As String, val As String, keyNum As Long
    keyStr = Trim(InputBox("Enter numeric key for new item (integer >=1):", "Add item - key"))
    If keyStr = vbNullString Then MsgBox "Cancelled.": Exit Sub
    If Not IsNumeric(keyStr) Then MsgBox "Key must be numeric integer.", vbExclamation: Exit Sub
    keyNum = CLng(keyStr)
    If keyNum < 1 Then MsgBox "Key must be >= 1.", vbExclamation: Exit Sub
    val = Trim(InputBox("Enter display text/value for key " & keyNum & ":", "Add item - value"))
    If val = vbNullString Then MsgBox "No value provided. Cancelled.": Exit Sub
    AddDataItemAtKey keyNum, val
    MsgBox "Added item at key " & keyNum & ".", vbInformation
End Sub

' Add at key helper (shifts keys >= key up by 1)
Public Sub AddDataItemAtKey(ByVal keyNum As Long, ByVal val As String)
    If itemsDict Is Nothing Then InitData
    Dim keysArr As Variant, i As Long, k As Long
    keysArr = itemsDict.Keys
    If itemsDict.Exists(CStr(keyNum)) Then
        Dim numericKeys() As Long
        ReDim numericKeys(0 To UBound(keysArr))
        For i = LBound(keysArr) To UBound(keysArr)
            numericKeys(i) = CLng(keysArr(i))
        Next i
        QuickSortLongArray numericKeys, LBound(numericKeys), UBound(numericKeys)
        For i = UBound(numericKeys) To LBound(numericKeys) Step -1
            If numericKeys(i) >= keyNum Then
                k = numericKeys(i)
                If itemsDict.Exists(CStr(k + 1)) Then itemsDict.Remove CStr(k + 1)
                itemsDict.Add CStr(k + 1), itemsDict(CStr(k))
                itemsDict.Remove CStr(k)
            End If
        Next i
        itemsDict.Add CStr(keyNum), val
    Else
        itemsDict.Add CStr(keyNum), val
    End If
    SaveDictToFile itemsDict, GetDataFilePath
End Sub

Public Sub RemoveDataItem()
    If itemsDict Is Nothing Then InitData
    Dim kstr As String
    kstr = Trim(InputBox("Enter the numeric key to remove:", "Remove data item"))
    If kstr = vbNullString Then Exit Sub
    If Not IsNumeric(kstr) Then MsgBox "Key must be numeric.", vbExclamation: Exit Sub
    If itemsDict.Exists(CStr(CLng(kstr))) Then
        BackupDatasetFileSilent
        itemsDict.Remove CStr(CLng(kstr))
        SaveDictToFile itemsDict, GetDataFilePath()
        MsgBox "Removed key " & CLng(kstr), vbInformation
    Else
        MsgBox "Key not found.", vbExclamation
    End If
End Sub

' Move (atomic)
Public Sub MoveDataItem()
    If itemsDict Is Nothing Then InitData
    Dim oldStr As String, newStr As String
    oldStr = Trim(InputBox("Enter the existing numeric key to move (e.g. 5):", "Move item - from"))
    If oldStr = vbNullString Then Exit Sub
    If Not IsNumeric(oldStr) Then MsgBox "Key must be numeric.", vbExclamation: Exit Sub
    newStr = Trim(InputBox("Enter the new numeric key for this item (e.g. 8):", "Move item - to"))
    If newStr = vbNullString Then Exit Sub
    If Not IsNumeric(newStr) Then MsgBox "Key must be numeric.", vbExclamation: Exit Sub
    Dim oldKey As Long, newKey As Long
    oldKey = CLng(oldStr)
    newKey = CLng(newStr)
    If oldKey < 1 Or newKey < 1 Then MsgBox "Keys must be >= 1.", vbExclamation: Exit Sub
    If Not itemsDict.Exists(CStr(oldKey)) Then MsgBox "Old key not found.", vbExclamation: Exit Sub
    If oldKey = newKey Then MsgBox "Old and new keys are identical; nothing to do.", vbInformation: Exit Sub

    Dim val As String
    val = itemsDict(CStr(oldKey))

    Dim keysArr As Variant, i As Long
    keysArr = itemsDict.Keys
    Dim numericKeys() As Long
    ReDim numericKeys(0 To UBound(keysArr))
    For i = LBound(keysArr) To UBound(keysArr)
        numericKeys(i) = CLng(keysArr(i))
    Next i
    QuickSortLongArray numericKeys, LBound(numericKeys), UBound(numericKeys)

    Dim newDict As Object
    Set newDict = CreateObject("Scripting.Dictionary")
    Dim k As Long
    For i = LBound(numericKeys) To UBound(numericKeys)
        k = numericKeys(i)
        If k = oldKey Then
        Else
            If oldKey < newKey Then
                If k > oldKey And k <= newKey Then
                    newDict.Add CStr(k - 1), itemsDict(CStr(k))
                Else
                    newDict.Add CStr(k), itemsDict(CStr(k))
                End If
            Else
                If k >= newKey And k < oldKey Then
                    newDict.Add CStr(k + 1), itemsDict(CStr(k))
                Else
                    newDict.Add CStr(k), itemsDict(CStr(k))
                End If
            End If
        End If
    Next i

    If newDict.Exists(CStr(newKey)) Then newDict.Remove CStr(newKey)
    newDict.Add CStr(newKey), val
    Set itemsDict = newDict
    SaveDictToFile itemsDict, GetDataFilePath
    MsgBox "Moved item from " & oldKey & " to " & newKey & ".", vbInformation
End Sub

' Show & insert
Public Sub ShowSelectionFormAndInsert()
    Dim seqNum As Long, keysArr As Variant, i As Long
    Dim outRange As Range
    On Error GoTo ErrHandler
    If Selection.Information(wdFirstCharacterColumnNumber) <> 1 Then
        MsgBox "Please set the cursor at the start of the line.", vbExclamation: Exit Sub
    End If
    If itemsDict Is Nothing Then InitData
    Set outRange = Selection.Range
    outRange.Collapse Direction:=wdCollapseStart
    outRange.InsertAfter "Copy forwarded to for information:" & vbCr
    outRange.Collapse Direction:=wdCollapseEnd

    Dim frm As New UserForm1
    With frm.ListBox1
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "300 pt;0 pt"
        On Error Resume Next: .MultiSelect = fmMultiSelectMulti: On Error GoTo 0
    End With

    keysArr = GetSortedNumericKeys(itemsDict)
    For i = LBound(keysArr) To UBound(keysArr)
        Dim k As String, v As String
        k = keysArr(i): v = itemsDict(k)
        frm.ListBox1.AddItem (k & " - " & v)
        frm.ListBox1.List(frm.ListBox1.ListCount - 1, 1) = k
    Next i

    frm.Show vbModal

    seqNum = 1
    For i = 0 To frm.ListBox1.ListCount - 1
        If frm.ListBox1.Selected(i) Then
            Dim itemKey As Long, valuePart As String
            itemKey = CLng(frm.ListBox1.List(i, 1))
            valuePart = itemsDict(CStr(itemKey))
            If InStr(valuePart, "The Additional District Magistrate") > 0 Then
                Dim addlDMOptions As Variant, j As Long, selCount As Long
                addlDMOptions = Array("Gen", "Dev", "LR", "ZP")
                Dim frmAddl As New UserForm2
                frmAddl.ListBox2.Clear
                For j = LBound(addlDMOptions) To UBound(addlDMOptions)
                    frmAddl.ListBox2.AddItem addlDMOptions(j)
                Next j
                On Error Resume Next: frmAddl.ListBox2.MultiSelect = fmMultiSelectMulti: On Error GoTo 0
                frmAddl.Show vbModal
                Dim addlSelections As String: addlSelections = "": selCount = 0
                For j = 0 To frmAddl.ListBox2.ListCount - 1
                    If frmAddl.ListBox2.Selected(j) Then
                        If addlSelections <> "" Then addlSelections = addlSelections & ", "
                        addlSelections = addlSelections & frmAddl.ListBox2.List(j)
                        selCount = selCount + 1
                    End If
                Next j
                If selCount > 0 Then
                    If selCount = 1 Then
                        outRange.InsertAfter seqNum & ")    The Additional District Magistrate (" & addlSelections & "), Nadia." & vbCr
                        seqNum = seqNum + 1
                    Else
                        outRange.InsertAfter seqNum & "-" & (seqNum + selCount - 1) & ")    The Additional District Magistrate (" & addlSelections & "), Nadia." & vbCr
                        seqNum = seqNum + selCount
                    End If
                    outRange.Collapse Direction:=wdCollapseEnd
                End If
            ElseIf InStr(valuePart, "Joint Block Development Officer") > 0 Then
                Dim answer As String, jointBDOCount As Long
                answer = InputBox("How many Joint BDOs you want to forward", "Input Required")
                If answer <> vbNullString And IsNumeric(answer) And Val(answer) >= 1 Then
                    jointBDOCount = CLng(Val(answer))
                    If jointBDOCount = 1 Then
                        outRange.InsertAfter seqNum & ")    " & valuePart & vbCr
                        seqNum = seqNum + 1
                    Else
                        outRange.InsertAfter seqNum & "-" & (seqNum + jointBDOCount - 1) & ")    " & valuePart & vbCr
                        seqNum = seqNum + jointBDOCount
                    End If
                    outRange.Collapse Direction:=wdCollapseEnd
                End If
            ElseIf InStr(valuePart, "To ........................... For Compliance.") > 0 Then
                Dim cntStr As String, cnt As Long, t As Long
                cntStr = InputBox("Enter a number for '" & valuePart & "' (each printed on its own line):", "Input Required")
                If cntStr <> vbNullString And IsNumeric(cntStr) And Val(cntStr) >= 1 Then
                    cnt = CLng(Val(cntStr))
                    For t = 1 To cnt
                        outRange.InsertAfter seqNum & ")    " & valuePart & vbCr
                        outRange.Collapse Direction:=wdCollapseEnd
                        seqNum = seqNum + 1
                    Next t
                End If
            ElseIf InStr(valuePart, "(All Gram Panchayat)") > 0 Or InStr(valuePart, "Shri/Smt") > 0 Then
                Dim userNumStr As String, userNum As Long
                userNumStr = InputBox("Enter a number for '" & valuePart & "'", "Input Required")
                If userNumStr <> vbNullString And IsNumeric(userNumStr) And Val(userNumStr) >= 1 Then
                    userNum = CLng(Val(userNumStr))
                    If userNum = 1 Then
                        outRange.InsertAfter seqNum & ")    " & valuePart & vbCr
                    Else
                        outRange.InsertAfter seqNum & "-" & (seqNum + userNum - 1) & ")    " & valuePart & vbCr
                    End If
                    outRange.Collapse Direction:=wdCollapseEnd
                    seqNum = seqNum + userNum
                End If
            Else
                outRange.InsertAfter seqNum & ")    " & valuePart & vbCr
                outRange.Collapse Direction:=wdCollapseEnd
                seqNum = seqNum + 1
            End If
        End If
    Next i
    MsgBox "Done. Inserted " & (seqNum - 1) & " entries.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' Persistence helpers
Public Sub SaveDictToFile(d As Object, filePath As String)
    Dim ts As Integer, k As Variant, s As String
    On Error GoTo ErrH
    EnsureDataFolderExists
    ts = FreeFile
    Open filePath For Output As #ts
    Dim sortedKeys As Variant, i As Long
    sortedKeys = GetSortedNumericKeys(d)
    For i = LBound(sortedKeys) To UBound(sortedKeys)
        k = sortedKeys(i)
        s = CStr(k) & "|" & Replace(CStr(d(k)), "|", "||")
        Print #ts, s
    Next i
    Close #ts
    Exit Sub
ErrH:
    MsgBox "Failed to save dataset to file: " & Err.Description, vbExclamation
    On Error Resume Next
    Close #ts
End Sub

Public Function LoadFromFile(filePath As String) As String
    Dim ts As Integer, line As String, out As String
    On Error GoTo NoFile
    ts = FreeFile
    Open filePath For Input As #ts
    out = ""
    Do While Not EOF(ts)
        Line Input #ts, line
        If Len(out) = 0 Then out = line Else out = out & vbCrLf & line
    Loop
    Close #ts
    LoadFromFile = out
    Exit Function
NoFile:
    LoadFromFile = ""
    On Error Resume Next
    Close #ts
End Function

Public Sub DeserializeToDict(s As String, ByRef d As Object)
    Dim lines As Variant, i As Long, kv As Variant, k As String
    If Len(Trim(s)) = 0 Then Exit Sub
    lines = Split(s, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        kv = Split(lines(i), "|")
        If UBound(kv) >= 1 Then
            k = kv(0)
            Dim j As Long, parts() As String, valStr As String
            ReDim parts(0 To UBound(kv) - 1)
            For j = 1 To UBound(kv)
                parts(j - 1) = kv(j)
            Next j
            valStr = Join(parts, "|")
            valStr = Replace(valStr, "||", "|")
            If Not d.Exists(k) Then d.Add k, valStr
        End If
    Next i
End Sub

Public Sub EnsureDataFolderExists()
    Dim folderPath As String
    folderPath = GetDataFolderPath()
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath
        On Error GoTo 0
    End If
End Sub

Public Function GetDataFolderPath() As String
    Dim appd As String
    appd = Environ("APPDATA")
    If Right(appd, 1) <> "\" Then appd = appd & "\"
    GetDataFolderPath = appd & DATA_FOLDER_NAME & "\"
End Function

Public Function GetDataFilePath() As String
    GetDataFilePath = GetDataFolderPath() & DATA_FILE_NAME
End Function

Public Function GetSortedNumericKeys(d As Object) As Variant
    Dim keysArr As Variant, numericKeys() As Long, i As Long
    If d Is Nothing Or d.Count = 0 Then GetSortedNumericKeys = Array(): Exit Function
    keysArr = d.Keys
    ReDim numericKeys(0 To UBound(keysArr))
    For i = LBound(keysArr) To UBound(keysArr)
        numericKeys(i) = CLng(keysArr(i))
    Next i
    QuickSortLongArray numericKeys, LBound(numericKeys), UBound(numericKeys)
    Dim out() As String
    ReDim out(0 To UBound(numericKeys))
    For i = LBound(numericKeys) To UBound(numericKeys)
        out(i) = CStr(numericKeys(i))
    Next i
    GetSortedNumericKeys = out
End Function

Public Sub QuickSortLongArray(arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim pivot As Long, i As Long, j As Long, temp As Long
    i = first: j = last: pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortLongArray arr, first, j
    If i < last Then QuickSortLongArray arr, i, last
End Sub

Public Sub QuickSortStringArray(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim pivot As String, i As Long, j As Long, temp As String
    i = first: j = last: pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortStringArray arr, first, j
    If i < last Then QuickSortStringArray arr, i, last
End Sub

Public Sub BackupDatasetFileSilent()
    Dim src As String, dst As String, ts As String
    src = GetDataFilePath()
    If Dir(src) = "" Then Exit Sub
    ts = Format(Now, "yyyy-mm-dd_HHNNSS")
    dst = GetDataFolderPath() & "WordItemsDataset_backup_" & ts & ".txt"
    On Error Resume Next
    FileCopy src, dst
    On Error GoTo 0
End Sub

Public Sub BackupDatasetFile()
    Dim src As String, dst As String, ts As String
    src = GetDataFilePath()
    If Dir(src) = "" Then
        MsgBox "No dataset file found to backup.", vbExclamation: Exit Sub
    End If
    ts = Format(Now, "yyyy-mm-dd_HHNNSS")
    dst = GetDataFolderPath() & "WordItemsDataset_backup_" & ts & ".txt"
    FileCopy src, dst
    MsgBox "Backup created: " & dst, vbInformation
End Sub

Public Sub ResetDataset()
    Dim f As String
    f = GetDataFilePath()
    EnsureDataFolderExists
    If Dir(f) <> "" Then
        If MsgBox("Backup current dataset then reset to defaults?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
        BackupDatasetFile
        On Error Resume Next: Kill f: On Error GoTo 0
    End If
    Set itemsDict = Nothing
    InitData
    MsgBox "Dataset reset to defaults and saved at:" & vbCrLf & GetDataFilePath(), vbInformation
End Sub

Public Sub ListDatasetToImmediate()
    If itemsDict Is Nothing Then InitData
    Dim keys As Variant, i As Long
    keys = GetSortedNumericKeys(itemsDict)
    Debug.Print "Dataset at: " & GetDataFilePath()
    For i = LBound(keys) To UBound(keys)
        Debug.Print keys(i) & " => " & itemsDict(CStr(keys(i)))
    Next i
    MsgBox "Dataset printed to Immediate Window (Ctrl+G).", vbInformation
End Sub

Public Function ModuleIsReady() As Boolean
    On Error Resume Next
    ModuleIsReady = Not (itemsDict Is Nothing)
End Function

Public Sub InitAppEventHandler()
    On Error Resume Next
    If gAppHandler Is Nothing Then
        Set gAppHandler = New AppEventHandler
        Set gAppHandler.App = Word.Application
    End If
    If itemsDict Is Nothing Then InitData
End Sub

Public Sub AutoExec()
    InitAppEventHandler
End Sub

Public Sub ShowAdvancedEditor()
    If itemsDict Is Nothing Then InitData
    AdvancedEditorForm.RefreshEditorList
    AdvancedEditorForm.Show vbModal
End Sub
