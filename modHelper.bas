Attribute VB_Name = "modHelper"
'new const
    Const COLUMN_OFFSET As Long = 4                ' this is the column where the template ends and the data starts
    Const HEADING_ROW As Long = 5                     ' The row where the heading titles are writen
    Const FIRST_DATA_ROW As Long = 8            ' This is the 1st employee row
    Const LAST_DATA_ROW As Long = 29             ' This is the last employee row
Public Type PayrollContext
        payYear As Long
        payMonth As Long
End Type

Public Function GetMonthNumberFromName(ByVal m As String) As Long

    Dim MonthNumber As Long
    
    Select Case m
        Case "January"
            MonthNumber = 1
        Case "February"
            MonthNumber = 2
        Case "March"
            MonthNumber = 3
        Case "April"
            MonthNumber = 4
        Case "May"
            MonthNumber = 5
        Case "June"
            MonthNumber = 6
        Case "July"
            MonthNumber = 7
        Case "August"
            MonthNumber = 8
        Case "September"
            MonthNumber = 9
        Case "October"
            MonthNumber = 10
        Case "November"
            MonthNumber = 11
        Case "December"
            MonthNumber = 12
    End Select

    GetMonthNumberFromName = MonthNumber
End Function
Public Function SheetTypeName(ByVal value As sheet_Type) As String
    Select Case value
        Case Booking: SheetTypeName = "Booking"
        Case payroll: SheetTypeName = "Payroll"
        Case Else:    SheetTypeName = "Unknown"
    End Select
End Function
Public Function BuildAttendanceKey(empID As Long, attDate As Date) As String
    BuildAttendanceKey = empID & "|" & Format(attDate, "yyyy-mm-dd")
End Function
Public Function GetWeekDays(startDate As Date, endDate As Date) As Variant
    Dim arr() As Date
    Dim i As Long
    Dim d As Date

    ReDim arr(1 To endDate - startDate + 1)

    i = 1
    For d = startDate To endDate
        arr(i) = d
        i = i + 1
    Next d

    GetWeekDays = arr
End Function
Public Sub DebugPrintSheetInfoStore()

    Dim store As Object
    Dim sheetType As Variant
    Dim info As Object
    Dim key As Variant

    Set store = GetSheetInfoStore()

    If store.count = 0 Then
        Debug.Print "SheetInfoStore is empty"
        Exit Sub
    End If

    Debug.Print "---- SheetInfoStore ----"

    For Each sheetType In store.Keys
        Debug.Print "SheetType:", SheetTypeName(sheetType)

        Set info = store(sheetType)

        For Each key In info.Keys
            Debug.Print "   ", key, "=", info(key)
        Next key
    Next sheetType

    Debug.Print "------------------------"

End Sub

Public Function IsBlankRow(ByVal rowNum As Long) As Boolean
    IsBlankRow = (rowNum = 10 Or rowNum = 20 Or rowNum = 26)
End Function

Public Function GetColumnByHeader(ws As Worksheet, headerName As String) As Long

    Dim hdr As range

    Set hdr = ws.Rows(1).Find( _
        What:=headerName, _
        LookAt:=xlWhole, _
        LookIn:=xlValues, _
        MatchCase:=False _
    )

    If hdr Is Nothing Then
        GetColumnByHeader = 0
    Else
        GetColumnByHeader = hdr.Column
    End If

End Function
Public Function GetCalculationRowMap() As Object

    Dim ws As Worksheet
    Set ws = Worksheets("Control")
    Static rowMap As Object
    If Not rowMap Is Nothing Then
        Set GetCalculationRowMap = rowMap
        Exit Function
    End If

    Set rowMap = CreateObject("Scripting.Dictionary")
    With ws
        rowMap("QE_LowerLimit") = .Cells(19, 7).value
        rowMap("QE_UpperLimit") = .Cells(20, 7).value
        rowMap("Pension_Employee_Rate") = .Cells(4, 7).value
        rowMap("Pension_Employer_Rate") = .Cells(6, 7).value
        rowMap("Pension_Owners_Rate") = .Cells(5, 7).value
        rowMap("Holiday_Rate") = .Cells(21, 7).value
        rowMap("Daily_Hours") = .Cells(22, 7).value
        rowMap("Carry_Over_Days") = .Cells(10, 7).value
        rowMap("Late_Days") = .Cells(11, 7).value
        rowMap("Sick_Days") = .Cells(12, 7).value
        rowMap("Absent_Days") = .Cells(13, 7).value
        rowMap("Employee_NI_Threshold") = .Cells(14, 7).value
        rowMap("Employee_NI_Rate_Category_A") = .Cells(16, 7).value
        rowMap("Employee_NI_Rate_Category_DEL") = .Cells(17, 7).value
        rowMap("Employer_NI_Threshold") = .Cells(15, 7).value
        rowMap("Employer_NI_Rate") = .Cells(18, 7).value
    End With
    Set GetCalculationRowMap = rowMap
    
    
End Function

Public Function GetPayrollContext(ByVal frm As UF_ControlCentre) As PayrollContext

    Dim ctx As PayrollContext
    
    ' Year
    ctx.payYear = CLng(frm.cboYear.value)

    ' Month (convert name to number)
    ctx.payMonth = modHelper.GetMonthNumberFromName(frm.cboMonth.value)

    GetPayrollContext = ctx    ' ? objects must be returned with Set ?

End Function

Function NI_Constant(row As Integer) As Double
    Dim ws As Worksheet
    Set ws = Worksheets("Control")
    ' we know the NI_values are in column d (4)
    NI_Constant = ws.Cells(row, 4).value
End Function
Function Audit_constant(row As Integer) As String
    Dim ws As Worksheet
    Set ws = Worksheets("Control")
    ' we know the NI_values are in column j (10)
    Audit_constant = ws.Cells(row, 10).value
End Function
Function Attendance_Status(row As Integer) As String
    Dim ws As Worksheet
    Set ws = Worksheets("AttendanceStatusConfig")
    ' we know the  Attendance_Status are in column A (1)
    Attendance_Status = ws.Cells(row, 1).value
End Function
Public Function GetEmployeeNode( _
    ByVal store As Object, _
    ByVal empID As Long) As Object

    If Not store.Exists(empID) Then
        Set store(empID) = CreateObject("Scripting.Dictionary")
    End If

    Set GetEmployeeNode = store(empID)

End Function


Public Sub EnsureYearInCombo(cbo As Object, ByVal yearValue As Long)

    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If CLng(cbo.List(i)) = yearValue Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

    ' Not found ? add it
    cbo.AddItem yearValue
    cbo.ListIndex = cbo.ListCount - 1

End Sub
Public Function ParseYearAndMonthFromSheetName( _
    ByVal sheetName As String) As YearMonthInfo

    Dim parts As Variant
    parts = Split(sheetName, "_")

    Dim result As New YearMonthInfo
    result.Year = CLng(parts(1))
    result.Month = GetMonthNumberFromName(parts(2))

    Set ParseYearAndMonthFromSheetName = result

End Function

Public Sub EnsureMonthInCombo(cbo As Object, ByVal MonthNumber As Long)

    Dim monthNameValue As String
    monthNameValue = MonthName(MonthNumber)

    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = monthNameValue Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

    cbo.AddItem monthNameValue
    cbo.ListIndex = cbo.ListCount - 1

End Sub

