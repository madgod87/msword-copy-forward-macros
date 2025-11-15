' Class: AppEventHandler
Attribute VB_Name = "AppEventHandler"
Option Explicit

Public WithEvents App As Word.Application

Private Sub App_Quit()
    On Error Resume Next
    If Not ModuleIsReady Then Exit Sub
    SaveDictToFile itemsDict, GetDataFilePath
End Sub

Private Sub App_DocumentBeforeClose(ByVal Doc As Word.Document, Cancel As Boolean)
    On Error Resume Next
    If Not ModuleIsReady Then Exit Sub
    SaveDictToFile itemsDict, GetDataFilePath
End Sub

Private Sub App_DocumentOpen(ByVal Doc As Word.Document)
    On Error Resume Next
    If Not ModuleIsReady Then Exit Sub
    Dim f As String, content As String, newDict As Object
    f = GetDataFilePath()
    If Len(Dir(f)) > 0 Then
        content = LoadFromFile(f)
        If Len(Trim(content)) > 0 Then
            Set newDict = CreateObject("Scripting.Dictionary")
            DeserializeToDict content, newDict
            Set itemsDict = newDict
        End If
    End If
End Sub
