Attribute VB_Name = "modVBAExport"

Private Sub ExportVBAWithPicker()
    Dim exportPath As String

    exportPath = PickFolder("Select folder to export VBA code")
    If exportPath = "" Then Exit Sub

    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"

    ExportVBA exportPath
End Sub


Public Function PickFolder(Optional ByVal prompt As String = "Select a folder") As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = prompt
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function

Public Sub ExportVBA(ByVal path As String)
    Dim comp As Object
    
    If Right(path, 1) <> "\" Then path = path & "\"

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1 ' Standard Module
                comp.Export path & comp.Name & ".bas"
            Case 2 ' Class Module
                comp.Export path & comp.Name & ".cls"
        End Select
        Debug.Print "the path is "; path; "The file name is "; comp.Name
    Next comp
End Sub

Public Function CanAccessVBProject() As Boolean
    On Error Resume Next
    Dim x As Object
    Set x = ThisWorkbook.VBProject
    CanAccessVBProject = (Err.Number = 0)
    Err.Clear
End Function

