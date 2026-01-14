Attribute VB_Name = "modAudit"
Public Sub Data_LogImport(ByVal DataStoreName As String, _
                          ByVal Reason As String, _
                          ByVal RowsCount As Long, _
                          ByVal ImportedSheet As String)

    Dim ws As Worksheet
    Dim r As Long
    Dim VersionNumber As Long

    On Error Resume Next
    Set ws = Worksheets("ImportLog")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "ImportLog sheet is missing.", vbCritical
        Exit Sub
    End If

    VersionNumber = Audit_constant(5)
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1

    ws.Cells(r, 1).value = Now
    ws.Cells(r, 2).value = Application.UserName
    ws.Cells(r, 3).value = DataStoreName
    ws.Cells(r, 4).value = Reason
    ws.Cells(r, 5).value = RowsCount
    ws.Cells(r, 6).value = ImportedSheet
    ws.Cells(r, 7).value = VersionNumber
End Sub


Public Function ValidateImport( _
    ByVal DataStoreName As String, _
    ByVal ImportedSheet As String _
) As Boolean

    Dim lastRow As Long
    Dim response As VbMsgBoxResult
    Dim sheetExists As Boolean
    

    ValidateImport = False ' default safe
    UF_ControlCentre.UpdateStatus "Validating the request"
    lastRow = GetLastImportRow(DataStoreName, ImportedSheet)

    ' Never imported before
    If lastRow = 0 Then
        ValidateImport = True
        Exit Function
    End If

    sheetExists = WorksheetExists(ImportedSheet)

    ' Imported before but sheet no longer exists
    If Not sheetExists Then
        ValidateImport = True
        Exit Function
    End If

    ' Imported and still exists ? ask user
    response = MsgBox( _
        "This sheet has already been imported." & vbCrLf & _
        "Sheet: " & ImportedSheet & vbCrLf & _
        "Do you want to remove the existing data and re-import?", _
        vbYesNoCancel + vbExclamation, _
        "Re-import data?" _
    )

    Select Case response
        Case vbYes
            Call RemovePayRollData(ImportedSheet)
            ValidateImport = True

        Case vbNo, vbCancel
            ValidateImport = False
    End Select
End Function
Private Function RemovePayRollData(ByVal ImportID As String) As Long

    Dim sheetsToClean As Variant
    Dim ws As Worksheet
    Dim importCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim removedThisSheet As Long
    Dim removedTotal As Long

    If MsgBox("Remove imported data for Import ID:" & vbCrLf & ImportID & " ?", _
              vbYesNo + vbExclamation) = vbNo Then Exit Function

    sheetsToClean = Array( _
        "WeeklyHistory", _
        "AttendanceHistory", _
        "MonthlyHistory" _
    )

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    UF_ControlCentre.UpdateStatus "Removing data from the follwing " & Join(sheetsToClean, ", ")

    For Each ws In Worksheets
        If Not IsError(Application.Match(ws.Name, sheetsToClean, 0)) Then

            importCol = GetColumnByHeader(ws, "Import_Sheet")
            If importCol = 0 Then GoTo NextSheet

            removedThisSheet = 0
            lastRow = ws.Cells(ws.Rows.count, importCol).End(xlUp).row

            ' Bottom-up delete
            For r = lastRow To 2 Step -1
                If ws.Cells(r, importCol).value = ImportID Then
                    ws.Rows(r).Delete
                    removedThisSheet = removedThisSheet + 1
                End If
            Next r

            ' Log only if something removed
            If removedThisSheet > 0 Then
                Data_LogImport ws.Name, "Removed (asked the user) ", removedThisSheet, ImportID
                removedTotal = removedTotal + removedThisSheet
            End If

        End If
NextSheet:
    Next ws

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    UF_ControlCentre.UpdateStatus "Removed all data from " & ImportID

End Function



Public Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

Private Function GetLastImportRow( _
    ByVal DataStoreName As String, _
    ByVal ImportedSheet As String _
) As Long

    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim latestRow As Long
    Dim latestDate As Date

    Set ws = Worksheets("ImportLog")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    latestRow = 0
    latestDate = 0

    For i = 2 To lastRow
        If ws.Cells(i, 3).value = DataStoreName _
           And ws.Cells(i, 6).value = ImportedSheet Then

            If ws.Cells(i, 1).value > latestDate Then
                latestDate = ws.Cells(i, 1).value
                latestRow = i
            End If

        End If
    Next i

    GetLastImportRow = latestRow
End Function



