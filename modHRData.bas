Attribute VB_Name = "modHRData"
    ' static column positions for the AttendanceHistory data store
    
    Const COL_EMPID As Long = 1
    Const COL_YEAR As Long = 2
    Const COL_MONTH As Long = 3
    Const COL_DATE As Long = 6
    Const COL_STATUS As Long = 7
Public Function LoadStatusConfig() As Object
    Dim ws As Worksheet
    Dim data As Variant
    Dim dict As Object
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("AttendanceStatusConfig")
    Set dict = CreateObject("Scripting.Dictionary")

    data = ws.range("A2:F" & ws.Cells(ws.Rows.count, 1).End(xlUp).row).value

    For i = 1 To UBound(data, 1)
        dict(data(i, 1)) = data(i, 2) & "|" & _
                           data(i, 3) & "|" & _
                           data(i, 4) & "|" & _
                           data(i, 5) & "|" & _
                           data(i, 6)
    Next i

    Set LoadStatusConfig = dict
End Function


Public Function LoadAttendanceHistory(ByVal CurrentYear As Long, _
                               ByVal CurrentMonth As Long) As Object

    Dim ws_DS As Worksheet
    Dim Attendancedata As Variant
    Dim store As Object
    Dim r As Long
    Dim key As String

    Set ws_DS = ThisWorkbook.Worksheets("AttendanceHistory")
    Set store = CreateObject("Scripting.Dictionary")

    Attendancedata = ws_DS.range("A2:H" & ws_DS.Cells(ws_DS.Rows.count, COL_EMPID).End(xlUp).row).value

    For r = 1 To UBound(Attendancedata, 1)

        If Attendancedata(r, COL_YEAR) = CurrentYear _
           And Attendancedata(r, COL_MONTH) = CurrentMonth Then

            key = BuildAttendanceKey(CLng(Attendancedata(r, COL_EMPID)), CDate(Attendancedata(r, COL_DATE)))
            store(key) = Attendancedata(r, COL_STATUS)

        End If
    Next r

    Set LoadAttendanceHistory = store

End Function
Public Sub WriteAttendanceRow( _
    ByVal wsHistA As Worksheet, _
    ByVal empID As Long, _
    ByVal weekIndex As Long, _
    ByVal workDate As Date, _
    ByVal attendanceCode As String, _
    ByVal sourceSheet As String)
    
    Dim lastDataRow As Long
    
    lastDataRow = wsHistA.Cells(wsHistA.Rows.count, 1).End(xlUp).row + 1

    If lastDataRow < 2 Then lastDataRow = 2
    
    wsHistA.Cells(lastDataRow, 1).Resize(1, 8).value = Array( _
        empID, _
        Year(workDate), _
        Month(workDate), _
        DatePart("ww", workDate, vbMonday, vbFirstFourDays), _
        weekIndex, _
        workDate, _
        attendanceCode, _
        sourceSheet _
    )

    

End Sub
Public Sub WriteEmployeeAttendance( _
    wsSource As Worksheet, _
    wsHistA As Worksheet, _
    empRow As Long, _
    empID As Long, _
    weekIndex As Long, _
    weekRange As Variant, _
    sourceSheetName As String)

    Dim dayOffset As Long
    Dim rawCode As String
    Dim normCode As String
    Dim workDate As Date
    Dim dayCol As Long

    For dayOffset = 0 To 6

        dayCol = ((weekIndex - 1) * (5 + 7)) + 8 + dayOffset

        rawCode = wsSource.Cells(empRow, dayCol).value
        normCode = modHRData.NormaliseAttendanceCode(rawCode)

        If normCode <> "NullString" Then
            workDate = weekRange(1) + dayOffset

            modHRData.WriteAttendanceRow _
                wsHistA, _
                empID, _
                weekIndex, _
                workDate, _
                normCode, _
                sourceSheetName

        End If

    Next dayOffset

End Sub
Public Function NormaliseAttendanceCode(rawCode As String) As String

    Select Case UCase(Trim(rawCode))
        Case "A": NormaliseAttendanceCode = Attendance_Status(7)
        Case "S": NormaliseAttendanceCode = Attendance_Status(4)
        Case "L": NormaliseAttendanceCode = Attendance_Status(3)
        Case "H": NormaliseAttendanceCode = Attendance_Status(2)
        Case "T": NormaliseAttendanceCode = Attendance_Status(5)
        Case "UPL": NormaliseAttendanceCode = Attendance_Status(6)
        

        Case Else
            NormaliseAttendanceCode = "NullString"
    End Select

End Function


Public Function EmployeeAttendanceData(ByVal filterYear As Long, _
    Optional ByVal filterMonth As Long = 0) As Object

    Dim wsA As Worksheet
    Dim dct As Object, emp As Object
    Dim r As Long, lastRow As Long
    Dim empID As Long

    

    Set wsA = Worksheets("AttendanceHistory")
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsA.Cells(wsA.Rows.count, "A").End(xlUp).row

    For r = 2 To lastRow

        empID = wsA.Cells(r, 1).value
        attDate = wsA.Cells(r, 6).value

        ' Year filter
        If Year(attDate) <> filterYear Then GoTo NextRow

        ' Month filter (if supplied)
        If filterMonth <> 0 Then
            If Month(attDate) <> filterMonth Then GoTo NextRow
        End If

        If Not dict.Exists(empID) Then
            Set emp = CreateObject("Scripting.Dictionary")
            
            emp("AbsentCount") = 0
            emp("LateCount") = 0
            emp("SickCount") = 0
            emp("HolidayCount") = 0
            emp("HolidayDays") = 0

            Set dict(empID) = emp
        End If

        Select Case wsA.Cells(r, 7).value
            Case "Absent":  emp("AbsentCount") = emp("AbsentCount") + 1
            Case "Late":    emp("LateCount") = emp("LateCount") + 1
            Case "Sick":    emp("SickCount") = emp("SickCount") + 1
            Case "Holiday": emp("HolidayCount") = emp("HolidayCount") + 1
        End Select

NextRow:
    Next r

    Set EmployeeAttendanceData = dict

End Function
Public Function ParsePayType(ByVal v As String) As PayTypeEnum
    Select Case UCase$(Trim$(v))
        Case "H": ParsePayType = PayHourly
        Case "S": ParsePayType = PaySalary
        Case Else: Err.Raise 5, , "Invalid PayType: " & v
    End Select
End Function

Public Function EmployeeLookup() As Object

    Dim ws As Worksheet
    Dim dict As Object, emp As Object
    Dim r As Long, lastRow As Long, empID As Long
    Dim rawPensionVal As String


    Set ws = Worksheets("Employee")
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    For r = 2 To lastRow
        empID = ws.Cells(r, 1).value
        rawPensionVal = ws.Cells(r, 14).value

        If Not dict.Exists(empID) Then
            Set emp = CreateObject("Scripting.Dictionary")
            emp("PayType") = ParsePayType(ws.Cells(r, 8).value)
            emp("DOB") = ws.Cells(r, 5).value
            emp("StartDate") = ws.Cells(r, 7).value
            emp("Salary") = ws.Cells(r, 10).value
            emp("Rate") = ws.Cells(r, 9).value
            emp("TaxCode") = ws.Cells(r, 13).value
            emp("Pension") = ws.Cells(r, 14).value
            emp("WorkerCategory") = modEmployeeRules.NormaliseWorkerCategory(ws.Cells(r, 14).value)
            emp("PensionStatus") = modEmployeeRules.NormalisePensionStatus(ws.Cells(r, 14).value)
            emp("NI_Catagory") = ws.Cells(r, 15).value
            emp("Apprentice") = ws.Cells(r, 16).value
            emp("AllowanceDays") = ws.Cells(r, 17).value

            Set dict(empID) = emp
        End If
    Next r

    Set EmployeeLookup = dict
End Function
Public Function EmployeeHoildayData(ByVal filterYear As Long, _
    Optional ByVal filterMonth As Long = 0) As Object

    Dim wsE As Worksheet, wsW As Worksheet, wsA As Worksheet
    Dim dict As Object, emp As Object, EmpLookup As Object
    Dim r As Long, lastRow As Long, empID As Long
    Dim DailyHours As Double, WK_Hours As Double, Taken_Days As Double
    Dim Accrued_Hours As Double, Allowance_Days As Double
    Dim Accrued_Days As Double, Allowance_Hours As Double, Taken_Hours As Double
    Dim Remaining_Hours As Double, Remaining_Days As Double
    Dim CarryOverHours As Double
    Dim CarryOverDays As Double
    Dim OpeningBalanceHours As Double, OpeningBalanceDays As Double
    Dim YearEndDate As Date, startDate As Date
    
    
    Set EmpLookup = EmployeeLookup()
    Set wsE = Worksheets("Employee")
    Set wsW = Worksheets("WeeklyHistory")
    Set wsA = Worksheets("AttendanceHistory")
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsA.Cells(wsA.Rows.count, "A").End(xlUp).row
    DailyHours = Calculation_constant(9)     ' e.g. 7.5


    For r = 2 To lastRow

        'look up information from employee table
        
        empID = wsA.Cells(r, 1).value
        If Not EmpLookup.Exists(empID) Then GoTo NextRow
        startDate = EmpLookup(empID)("StartDate")
        Allowance_Days = EmpLookup(empID)("AllowanceDays")
        
        
        'filter Year
        If wsA.Cells(r, 2).value <> filterYear Then GoTo NextRow

        OpeningBalanceHours = GetOpeningBalance("Hours", empID, filterYear)
        CarryOverHours = Calculation_constant(10) * DailyHours
        OpeningBalanceDays = GetOpeningBalance("Days", empID, filterYear)
        CarryOverDays = Calculation_constant(10)
    
    
        ' Worked hours
        WK_Hours = Application.WorksheetFunction.SumIfs( _
                        wsW.range("H:H"), _
                        wsW.range("A:A"), empID, _
                        wsW.range("B:B"), filterYear)
    
    
        ' Holiday taken (DAYS) from AttendanceHistory
        Taken_Days = Application.WorksheetFunction.CountIfs( _
                        wsA.range("A:A"), empID, _
                        wsA.range("B:B"), filterYear, _
                        wsA.range("G:G"), "Holiday")
    
    
    
        ' Accrued
        Accrued_Hours = WK_Hours * Calculation_constant(8) ' 0.1207
    
    
        Allowance_Hours = Allowance_Days * DailyHours
    
        If Accrued_Hours > Allowance_Hours Then
            Accrued_Hours = Allowance_Hours
        End If
    
        Accrued_Days = IIf(DailyHours > 0, Accrued_Hours / DailyHours, 0)
    
        ' Taken
        Taken_Hours = Taken_Days * DailyHours
    
        ' Remaining
        Remaining_Hours = (OpeningBalanceHours + Accrued_Hours) - Taken_Hours
        If Remaining_Hours = (Allowance_Hours + CarryOverHours) Then Remaining_Hours = 0
        
        Remaining_Days = (OpeningBalanceDays + Accrued_Days) - Taken_Days
        If Remaining_Days = (Allowance_Days + CarryOverDays) Then Remaining_Days = 0
    
        If Not dict.Exists(empID) Then
            Set emp = CreateObject("Scripting.Dictionary")

            emp("StartDate") = startDate
            emp("HolidayAllowanceInDays") = Allowance_Days
            emp("HolidayAllowanceInHours") = Allowance_Days * DailyHours
            emp("HolidaysAccruedInHours") = Accrued_Hours
            emp("HolidaysAccruedInDays") = Accrued_Days
            emp("HolidaysTakenInDays") = Taken_Days
            emp("HolidaysTakenInHours") = Taken_Hours
            emp("HolidaysRemainingInDays") = Remaining_Days
            emp("HolidaysRemainingInHours") = Remaining_Hours
            
            Set dict(empID) = emp
        End If
NextRow:
    Next r

    
    Set EmployeeHoildayData = dict

End Function

Private Function GetOpeningBalance(BalanceType As String, empID As Variant, Year As Long) As Double

    Dim ws As Worksheet
    Set ws = Worksheets("HolidayBalances")

    On Error Resume Next
    If BalanceType = "Hours" Then
        GetOpeningBalance = Application.WorksheetFunction.SumIfs( _
                            ws.range("C:C"), _
                            ws.range("A:A"), empID, _
                            ws.range("B:B"), Year)
    Else
        GetOpeningBalance = Application.WorksheetFunction.SumIfs( _
                            ws.range("D:D"), _
                            ws.range("A:A"), empID, _
                            ws.range("B:B"), Year)
    End If
    On Error GoTo 0

End Function
Public Function EmployeeHolidayBalance(ByVal filterYear As Long) As Object

    Dim dict As Object, emp As Object
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, empID As Long
    Dim yearFound As Boolean
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets("HolidayBalances")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    yearFound = False

    For r = 2 To lastRow

        If ws.Cells(r, "B").value = filterYear Then
            yearFound = True
            empID = ws.Cells(r, "A").value
             If Not dict.Exists(empID) Then
                Set emp = CreateObject("Scripting.Dictionary")
                emp("CarryOverHours") = ws.Cells(r, "C").value
                emp("CarryOverDays") = ws.Cells(r, "D").value
                
                dict(empID) = emp
            End If
        End If

    Next r

    ' Year not present in dataset
    If Not yearFound Then
        Set EmployeeHolidayBalance = Nothing
    Else
        Set EmployeeHolidayBalance = dict
    End If

End Function

