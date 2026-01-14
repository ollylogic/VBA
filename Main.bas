Attribute VB_Name = "Main"
Option Explicit

   
    
    'new const
    Const COLUMN_OFFSET As Long = 4                ' this is the column where the template ends and the data starts
    Const HEADING_ROW As Long = 5                     ' The row where the heading titles are writen
    
    Private firstDay As Date, lastDay As Date               ' 1st and last date of a range ( weeks or months)
    Private weekCount As Long, weekIndex As Long            ' 1st and last week of month
    Private currentCol As Long
    
    ' Program constant, normally pointer to rows on static sheets
    
    Const HEADING_START_ROW_ON_TEMPLATE As Long = 5        ' The heading offset row number
Const PAYROLLRAWDATA_WIDTH As Long = 5                 ' there are 5 cols created, so the app needs to know, to loop through weeks

    Const TEMPLATE_OFFSET As Long = 4                ' this is where the template ends
    Const FIRST_DAY_COL As Long = 9                        ' The data offset row number
    Const DAYS_IN_WEEK As Long = 7
    Const FIRST_EMP_ROW As Long = 8                        ' start reading data
    Const LAST_EMP_ROW As Long = 29                        ' stop loop through rows
    Const NODE_HOLIDAYS As String = "Holidays"             ' The Node label for the Holiday data set
    Const NODE_ATTENDANCE As String = "Attendance"         ' The Node label for the Attendance data set
    Const NODE_PAYROLL As String = "PayRoll"               ' The Node label for the PayRoll data set
    Const NODE_EMPID_PAYDATA As String = "empID_PayData"   ' The Node label for the empID_PayData data set
    Const NODE_HOLIDAYBALANCE As String = "HolidayBalance" ' The Node label for the HolidayBalance data set
    Const WEEKDAY_HEADER_ROW As Long = 7
    Const FIRST_EMP_ROW_FINANCIALS As Long = 11            ' This is the start row of a report
    Const LAST_EMP_ROW_FINANCIALS As Long = 32             ' This is the end rows of a report
    Const FIRST_EMP_ROW_TRACKER As Long = 11               ' This is the start row of a report
    Const LAST_EMP_ROW_TRACKER As Long = 32                ' This is the end rows of a report
    Const FIRST_COL As Long = 4
    


    








Public Sub system_Filter_combo_boxes(IsTheYearClosed As Boolean, ThisYear As String)


Dim ws As Worksheet, wsC As Worksheet
Dim r As Long, lastRow As Long, sheetYear As Long
Dim yearPart() As String, prefix As String

Dim Payroll_exclude As String, Attendance_exclude As String

    
Set wsC = Worksheets("Control")
 r = 4

    ' wipe the payroll combo boxs and then fill them excluding the closed years
    
    UF_ControlCentre.cbobookingsheets.Clear
    UF_ControlCentre.cboPayrollSheets.Clear
    
    For Each ws In ThisWorkbook.Worksheets

        prefix = Split(ws.Name, "_")(0)
        If prefix = "PayRoll" Or prefix = "Attendance" Then
            yearPart = Split(ws.Name, "_")
            ' Ensure we actually have a year
            If IsNumeric(yearPart(1)) = True Then
                sheetYear = CLng(yearPart(1))
                ' Show all payroll sheets in the control sheet
                wsC.Cells(r, 1).value = ws.Name
                r = r + 1

                ' Add to ComboBox only if allowed
                If sheetYear = CLng(UF_ControlCentre.cboYear.value) Then
                    If Not (IsTheYearClosed And sheetYear = ThisYear) Then
                        If prefix = "PayRoll" Then
                            UF_ControlCentre.cboPayrollSheets.AddItem ws.Name
                        Else
                            UF_ControlCentre.cbobookingsheets.AddItem ws.Name
                        End If
                        
                        
                    End If
                End If
            End If
        End If
    Next ws
End Sub


Public Function xxValidate_PayRollImport( _
    ByVal ImportID As String, _
    ByVal Sheet_Name As String _
) As Boolean

    Dim ws As Worksheet
    Dim importCol As Long
    Dim f As range

    Set ws = Worksheets(Sheet_Name)

    ' Try common import column headers
    importCol = GetColumnByHeader(ws, "Import_Sheet")
    If importCol = 0 Then importCol = GetColumnByHeader(ws, "ImportID")
    If importCol = 0 Then importCol = GetColumnByHeader(ws, "Import_ID")

    ' No import column found
    If importCol = 0 Then
        Validate_PayRollImport = False
        Exit Function
    End If

    ' Fast lookup
    Set f = ws.Columns(importCol).Find( _
                What:=ImportID, _
                LookAt:=xlWhole, _
                LookIn:=xlValues _
            )

    Validate_PayRollImport = Not f Is Nothing

End Function

Public Sub xxxCreateMonthlyDataInputSheet()

    Dim selectedYear As Long, Offset As Long
    Dim wsNew As Worksheet
    Dim sName As String, selectedMonth As String
    Dim rng As Variant
    Dim SelectedMonthNumber As Long ' we need the number  later
    Dim d As Date
    
    On Error GoTo Skip_and_end
    
    
    
    
    ' the selected year and month are now  shown in the form, we take varibles to use below
    
    selectedYear = UF_ControlCentre.cboYear
    selectedMonth = UF_ControlCentre.cboMonth
    SelectedMonthNumber = Hepler_GetMonthNumber(selectedMonth)   ' returns 1–12
    
    'CURRENT_SHEET_TYPE = This is selected based upon the button which called the procedure
    

    currentCol = START_COL_ON_TEMPLATE
    
    'This procedure uses a static template, Which is currently in the workbook coloured yellow,
    'this worksheet contains all the Employee formatted by the user ( can be changed)

    'Each has a different: name, Offset form the original template
    
    If CURRENT_SHEET_TYPE = stPayRoll Then
        sName = "PayRoll_" & selectedYear & "_" & selectedMonth
        Offset = PAYROLLRAWDATA_WIDTH + DAYS_IN_WEEK
    Else
        sName = "Attendance_" & selectedYear & "_" & selectedMonth
        Offset = DAYS_IN_WEEK
    End If
   
   
    Set wsNew = Helper_CreateTheRequiredSheet(sName)
   
   
    ' Checks if the sheet can be created
    If wsNew Is Nothing Then
        ' User declined or cancelled
        GoTo Skip_and_end
    End If

    'Get month range  from selectedyear and SelectedMonthNumber
        
    rng = GetDateRange(rtMonthly, selectedYear, SelectedMonthNumber)
    
    Format_WriteSheetTileAndPeriod wsNew, CDate(rng(1)), CDate(rng(2))
    
    weekCount = DateDiff("ww", CDate(rng(1)), CDate(rng(2)), vbMonday, vbFirstFourDays) + 1
    
    'This loop does all the work
    'We have given it the start and end weeks,
    'so it can use the ProcessTemplateWeek with the varibles set above to do the work ie
    'produce thats weeks information and then move the counter(currentCol) on  too do the next week and so on
    
    For weekIndex = 1 To weekCount
        ProcessTemplateWeek wsNew, selectedYear, SelectedMonthNumber, weekIndex, currentCol
        currentCol = currentCol + Offset
    Next weekIndex

    wsNew.Columns.AutoFit
    

    Exit Sub
Skip_and_end:
    
    MsgBox Err.Description
    
End Sub




Private Sub xxWritePeriodMothly(ws As Worksheet, firstDay As Date, lastDay As Date)

    Dim Title As String
    
    If CURRENT_SHEET_TYPE = stPayRoll Then
        Title = "Pay Period"
    Else
        Title = "Hoilday range"
        ws.range("A32").value = ""
        ws.range("B34:B38").value = ""
    End If
    
    ws.range("A2").value = Title
    
    ws.range("C2").value = _
        Format(firstDay, "d mmmm yyyy") & " – " & Format(lastDay, "d mmmm yyyy")

    With ws.Cells(2, 3)
        .Resize(1, 4).Merge
        .Font.Bold = True
        .Font.Size = 11
    End With

End Sub



Function xxxLoadAttendanceHistory(ByVal payYear As Long, ByVal payMonth As Long) As Object

    Dim ws As Worksheet
    Dim data As Variant
    Dim dict As Object
    Dim i As Long
    Dim key As String

    Set ws = ThisWorkbook.Worksheets("AttendanceHistory")
    Set dict = CreateObject("Scripting.Dictionary")

    data = ws.range("A2:H" & ws.Cells(ws.Rows.count, 1).End(xlUp).row).value

    For i = 1 To UBound(data, 1)
        If data(i, 2) = payYear And data(i, 3) = payMonth Then
            key = data(i, 1) & "|" & data(i, 6) ' EmpID|Date
            dict(key) = data(i, 7)              ' Status
        End If
    Next i

    Set LoadAttendanceHistory = dict
End Function


Private Sub xxProcessWeek(ws As Worksheet, payYear As Long, payMonth As Long, _
                        weekIndex As Long, startCol As Long)

    Dim weekRange As Variant
    Dim firstDataRow As Long, lastDataRow As Long
    Dim empRow As Long, row As Long, col As Long
    Dim holidayDict As Object
    
    firstDataRow = DATA_START_ROW_ON_TEMPLATE
    lastDataRow = LAST_ROW
    
    
    Set holidayDict = LoadAttendanceHistory()

    For empRow = firstDataRow To lastDataRow
        If IsBlankRow(empRow) Then GoTo NextEmp
        SetWeeklyEmployeeInput ws, empRow, startCol
NextEmp:
    Next empRow

    ApplyWeeklyBorders ws, startCol, lastDataRow
    
 

End Sub

Private Function SetTheTempleteOffset() As Long

If sheet_Type = payroll Then
        SetTheTempleteOffset = 5
    Else
        SetTheTempleteOffset = 0
    End If

End Function

Private Function xxHepler_GetMonthNumber(ByVal m As String) As Long

    Dim MonthNumber As Long
    
    Select Case m
        Case "Januay"
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

    Hepler_GetMonthNumber = MonthNumber
End Function




End Sub





Private Function xxDict_GetEmployeePayData() As Object

    Dim store As Object
    Dim ws As Worksheet
    Dim emp As Object, empNode As Object, att As Object
    Dim r As Long, lastRow As Long, empID As Long

    Set store = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets("Employee")

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        empID = ws.Cells(r, 1).value
         If Not store.Exists(empID) Then
            Set empNode = CreateObject("Scripting.Dictionary")
            Set att = CreateObject("Scripting.Dictionary")

            att("PayType") = ws.Cells(r, 8).value
            att("Salary") = ws.Cells(r, 10).value
            att("Rate") = ws.Cells(r, 9).value
            att("TaxCode") = ws.Cells(r, 13).value
            att("Pension") = ws.Cells(r, 14).value

            Set empNode(NODE_EMPID_PAYDATA) = att
            Set store(empID) = empNode
        End If
       Set att = store(empID)(NODE_EMPID_PAYDATA)
    Next r

    Set Dict_GetEmployeePayData = store

End Function







Function xxLoadEmployeePayrollData() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = Worksheets("Employee")

    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    ' EmpID ? payroll parameters (' pay type,Salary, Employee Tax code and Pension
    For r = 2 To lastRow
        dict(ws.Cells(r, 1).value) = Array( _
            ws.Cells(r, 8).value, _
            ws.Cells(r, 10).value, _
            ws.Cells(r, 13).value, _
            ws.Cells(r, 14).value _
        )
    Next r

    Set LoadEmployeePayrollData = dict
End Function


Private Function Dict_HoildayData(ByVal filterYear As Long, _
    Optional ByVal filterMonth As Long = 0) As Object

    Dim wsE As Worksheet, wsW As Worksheet, wsA As Worksheet
    Dim store As Object, empNode As Object, att As Object, Employee As Object
    Dim r As Long, lastRow As Long, empID As Long
    Dim DailyHours As Double, WK_Hours As Double, Taken_Days As Double
    Dim Accrued_Hours As Double, Allowance_Days As Double
    Dim Accrued_Days As Double, Allowance_Hours As Double, Taken_Hours As Double
    Dim Remaining_Hours As Double, Remaining_Days As Double
    Dim CarryOverHours As Double
    Dim CarryOverDays As Double
    Dim OpeningBalanceHours As Double, OpeningBalanceDays As Double
    Dim YearEndDate As Date, startDate As Date
    
    
    Set Employee = LoadEmployeeLookup()
    Set wsE = Worksheets("Employee")
    Set wsW = Worksheets("WeeklyHistory")
    Set wsA = Worksheets("AttendanceHistory")
    Set store = CreateObject("Scripting.Dictionary")

    lastRow = wsA.Cells(wsA.Rows.count, "A").End(xlUp).row
    DailyHours = Calculation_constant(9)     ' e.g. 7.5


    For r = 2 To lastRow

        'look up information from Employee table
        
        empID = wsA.Cells(r, 1).value
        If Not Employee.Exists(empID) Then GoTo NextRow
        startDate = Employee(empID)("StartDate")
        Allowance_Days = Employee(empID)("AllowanceDays")
        
        
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
    
        
    
        ' Get or create Employee node
        If Not store.Exists(empID) Then
            Set empNode = CreateObject("Scripting.Dictionary")
            Set att = CreateObject("Scripting.Dictionary")

            att("StartDate") = startDate
            att("HolidayAllowanceInDays") = Allowance_Days
            att("HolidayAllowanceInHours") = Allowance_Days * DailyHours
            att("HolidaysAccruedInHours") = Accrued_Hours
            att("HolidaysAccruedInDays") = Accrued_Days
            att("HolidaysTakenInDays") = Taken_Days
            att("HolidaysTakenInHours") = Taken_Hours
            att("HolidaysRemainingInDays") = Remaining_Days
            att("HolidaysRemainingInHours") = Remaining_Hours
            
            Set empNode(NODE_HOLIDAYS) = att
            Set store(empID) = empNode
        End If

        Set att = store(empID)(NODE_HOLIDAYS)

NextRow:
    Next r

    
    Set Dict_HoildayData = store

End Function


Private Function xxLoadEmployeeLookup() As Object

    Dim ws As Worksheet
    Dim dict As Object
    Dim r As Long, lastRow As Long
    Dim empID As Long
    Dim emp As Object

    Set ws = Worksheets("Employee")
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    For r = 2 To lastRow
        empID = ws.Cells(r, 1).value

        If Not dict.Exists(empID) Then
            Set emp = CreateObject("Scripting.Dictionary")
            emp("StartDate") = ws.Cells(r, 7).value
            emp("AllowanceDays") = ws.Cells(r, 15).value

            Set dict(empID) = emp
        End If
    Next r

    Set LoadEmployeeLookup = dict

End Function

Public Function Data_IsYearAlreadyClosed(Year As Long) As Boolean

    Dim found As Integer
    
    Data_IsYearAlreadyClosed = False
    
    found = Application.WorksheetFunction.CountIf( _
            Worksheets("HolidayBalances").range("B:B"), Year)
            
    If found > 0 Then
        Data_IsYearAlreadyClosed = True
        Exit Function
    End If
    
End Function
Public Function Data_CountImportedMonths(filterYear As Long) As Long

    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long

    Set ws = Worksheets("WeeklyHistory")
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row   ' Year column

    For r = 2 To lastRow
        If ws.Cells(r, "B").value = filterYear Then
            dict(ws.Cells(r, "C").value) = Empty   ' Month column
        End If
    Next r

    Data_CountImportedMonths = dict.count

End Function
Public Function Data_HasPostiveHolidayBalances(ByVal yearValue As Long) As Boolean
    

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = Worksheets("HolidayBalances")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    For i = 2 To lastRow
        If ws.Cells(i, "B").value = yearValue Then
            If ws.Cells(i, "c").value > 0 Then   ' F = Remaining Hours
                Data_HasPostiveHolidayBalances = True
                Exit Function
            End If
        End If
    Next i

    Data_HasPostiveHolidayBalances = False

End Function


Public Sub ProduceEmployeeReport(empID As Variant, filterYear As Long)

    
    Dim ws As Worksheet
    Dim stats As Object
    Set stats = BuildEmployeeStats(empID, filterYear)
    
    Set ws = Worksheets("Employee_Attendance")
    ' --- Fill report ---
    With ws
        ws.range("B12").value = empID
        ws.range("C12").value = UF_ControlCentre.cboEmployee.Column(1)
        ws.range("D12").value = stats("AbsentCount")
        ws.range("E12").value = stats("LateCount")
        ws.range("F12").value = stats("SickCount")
        ws.range("G12").value = stats("HolidayCount")
        ws.range("H12").value = stats("HolidayDates")
    End With

    ' --- Colour cells if thresholds reached ---
    Call Format_ColourCellIfThreshold(ws.range("E12"), stats("LateCount"), Calculation_constant(11))
    Call Format_ColourCellIfThreshold(ws.range("F12"), stats("SickCount"), Calculation_constant(12))
    Call Format_ColourCellIfThreshold(ws.range("D12"), stats("AbsentCount"), Calculation_constant(13))
    Call Format_ColourCellIfThreshold(ws.range("G12"), stats("HolidayCount"), Calculation_constant(14)) ' if needed


End Sub

Private Function GetEmployeeNode( _
    ByVal store As Object, _
    ByVal empID As Long) As Object

    If Not store.Exists(empID) Then
        Set store(empID) = CreateObject("Scripting.Dictionary")
    End If

    Set GetEmployeeNode = store(empID)

End Function

Public Function BuildEmployeeStats(ByVal empID As Long, _
    ByVal filterYear As Long) As Object

    Dim wsA As Worksheet
    Dim dict As Object
    Dim r As Long, lastRow As Long

    Set wsA = Worksheets("AttendanceHistory")
    Set dict = CreateObject("Scripting.Dictionary")

    dict("HolidayCount") = 0
    dict("HolidayDates") = ""
    dict("LateCount") = 0
    dict("SickCount") = 0
    dict("AbsentCount") = 0

    lastRow = wsA.Cells(wsA.Rows.count, "A").End(xlUp).row

    For r = 2 To lastRow
        If wsA.Cells(r, 1).value = empID _
           And wsA.Cells(r, 2).value = filterYear Then

            Select Case wsA.Cells(r, 7).value

                Case "Holiday"
                    dict("HolidayCount") = dict("HolidayCount") + 1
                    dict("HolidayDates") = dict("HolidayDates") & _
                        Format(wsA.Cells(r, 6).value, "dd/mm/yyyy") & ", "

                Case "Late"
                    dict("LateCount") = dict("LateCount") + 1

                Case "Sick"
                    dict("SickCount") = dict("SickCount") + 1

                Case "Absent"
                    dict("AbsentCount") = dict("AbsentCount") + 1

            End Select
        End If
    Next r

    If Len(dict("HolidayDates")) > 0 Then
        dict("HolidayDates") = Left(dict("HolidayDates"), _
                                    Len(dict("HolidayDates")) - 2)
    End If

    Set BuildEmployeeStats = dict

End Function
Public Function xxLoadAttendanceData(ByVal filterYear As Long, _
    Optional ByVal filterMonth As Long = 0) As Object

    Dim wsA As Worksheet
    Dim store As Object
    Dim r As Long, lastRow As Long
    Dim empID As Long
    Dim attDate As Date
    Dim empNode As Object, att As Object

    Set wsA = Worksheets("AttendanceHistory")
    Set store = CreateObject("Scripting.Dictionary")

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

        ' Get or create Employee node
        If Not store.Exists(empID) Then
            Set empNode = CreateObject("Scripting.Dictionary")
            Set att = CreateObject("Scripting.Dictionary")

            att("AbsentCount") = 0
            att("LateCount") = 0
            att("SickCount") = 0
            att("HolidayCount") = 0
            att("HolidayDays") = 0

            Set empNode(NODE_ATTENDANCE) = att
            Set store(empID) = empNode
        End If

        Set att = store(empID)(NODE_ATTENDANCE)

        Select Case wsA.Cells(r, 7).value
            Case "Absent":  att("AbsentCount") = att("AbsentCount") + 1
            Case "Late":    att("LateCount") = att("LateCount") + 1
            Case "Sick":    att("SickCount") = att("SickCount") + 1
            Case "Holiday": att("HolidayCount") = att("HolidayCount") + 1
        End Select

NextRow:
    Next r

    Set LoadAttendanceData = store

End Function


Public Sub WriteHolidayBalancesYearEnd(closeYear As Long)

    Dim wsHB As Worksheet
    Dim emp As Variant
    Dim arr As Variant
    Dim NextRow As Long
    Dim DailyHours As Double
    Dim CarryOverHours As Double
    Dim CarryOverDays As Double
    Dim Emp_data As Object
    
    Emp_data = LoadAttendanceData(closeYear)

    Set wsHB = Worksheets("HolidayBalances")

    NextRow = wsHB.Cells(wsHB.Rows.count, "A").End(xlUp).row + 1
    
    DailyHours = Calculation_constant(9)
    For Each emp In Emp_data.Keys

        arr = Emp_data(emp)

        ' arr mapping (from your function)
        ' arr(8)  = Remaining_Hours
        ' Carry-over = capped remaining


        CarryOverHours = WorksheetFunction.Min( _
                                arr(8), _
                                Calculation_constant(10) * DailyHours)

        CarryOverDays = IIf(DailyHours > 0, _
                            CarryOverHours / DailyHours, 0)

        wsHB.Cells(NextRow, 1).value = emp
        wsHB.Cells(NextRow, 2).value = closeYear
        wsHB.Cells(NextRow, 3).value = Round(CarryOverHours, 2)
        wsHB.Cells(NextRow, 4).value = Round(CarryOverDays, 2)
        wsHB.Cells(NextRow, 5).value = Now

        NextRow = NextRow + 1

    Next emp

    

End Sub
Public Function xxLoadMonthlyPayrollData( _
    ByVal filterYear As Long, _
    ByVal filterMonth As Long) As Object


    Dim store As Object, att As Object

    Set store = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = Worksheets("MonthlyHistory")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    Dim r As Long
    For r = 2 To lastRow

        If ws.Cells(r, "B").value = filterYear _
           And ws.Cells(r, "C").value = filterMonth Then

            Dim empID As Long
            empID = ws.Cells(r, "A").value

            Dim empNode As Object
            Set empNode = GetEmployeeNode(store, empID)

            Dim payroll As Object
            If Not empNode.Exists Then
                Set empNode = CreateObject("Scripting.Dictionary")
                Set att = CreateObject("Scripting.Dictionary")
            End If
            

            att("GrossWage") = ws.Cells(r, "D").value
            att("EmployeeTax") = ws.Cells(r, "E").value
            att("EmployeeNI") = ws.Cells(r, "F").value
            att("EmployerNI") = ws.Cells(r, "G").value
            att("TaxAllowance") = ws.Cells(r, "H").value
            att("EmployeePension") = ws.Cells(r, "I").value
            att("EmployerPension") = ws.Cells(r, "J").value
            att("TaxYear") = ws.Cells(r, "K").value
            
            Set empNode(NODE_PAYROLL) = att
            Set store(empID) = empNode
        End If
        Set att = store(empID)(NODE_PAYROLL)
    Next r

    Set LoadMonthlyPayrollData = store

End Function



 


Function MonthlySalaryProrated( _
    annualSalary As Double, _
    totalDaysWorked As Double, _
    expectedDays As Double _
) As Double

    If totalDaysWorked > expectedDays Then
        totalDaysWorked = expectedDays
    End If

    MonthlySalaryProrated = _
        Round((annualSalary / 12) * (totalDaysWorked / expectedDays), 2)

End Function
Public Sub Data_wrong_StoreAttendanceFromSheet( _
    ByVal wsInput As Worksheet, _
    ByVal wsHistA As Worksheet, _
    ByVal empID As Long, _
    ByVal empRow As Long, _
    ByVal weekIndex As Long, _
    ByVal WkRange As Variant, _
    ByRef lastHistRowA As Long _
)

    Dim dayOffset As Long
    Dim rawCode As String
    Dim normCode As String
    Dim workDate As Date
    Dim weekDayIndex As Long

    weekDayIndex = Weekday(WkRange(1), vbMonday) ' 1=Mon

    For dayOffset = 0 To DAYS_IN_WEEK - 1

        rawCode = wsInput.Cells( _
            empRow, _
            ((weekIndex - 1) * DAYS_IN_WEEK) + FIRST_DAY_COL + dayOffset _
        ).value

        normCode = NormaliseAttendanceCode(rawCode)

        If normCode <> "" Then

            workDate = DateAdd("d", dayOffset - (weekDayIndex - 1), WkRange(1))

            wsHistA.Cells(lastHistRowA, 1).Resize(1, 8).value = Array( _
                empID, _
                gPayYear, _
                gPayMonth, _
                DatePart("ww", WkRange(1), vbMonday, vbFirstFourDays), _
                weekIndex, _
                workDate, _
                normCode, _
                wsInput.Name _
            )

            lastHistRowA = lastHistRowA + 1

        End If
    Next dayOffset

End Sub
' Writes ONE attendance record

Public Sub Data_StoreAttendanceFromSheet(ByVal sheetName As String)

    Dim empID As Long, empRow As Long, weekIndex As Long
    Dim d As Date
    Dim startDate As Date, endDate As Date
    Dim code As String
    Dim wsHistA As Worksheet, ws As Worksheet
    Dim lastHistRowA As Long, lastRow As Long
    Dim rng As Variant
    Dim dayOffset As Long
    Dim rawCode As String
    Dim normCode As String
    Dim workDate As Date
    Dim DateParts() As String
    

    Set wsHistA = Worksheets("AttendanceHistory")
    Set ws = Worksheets(sheetName)
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastHistRowA = wsHistA.Cells(wsHistA.Rows.count, 1).End(xlUp).row + 1
    
    
    ' === PARSE YEAR / MONTH FROM SHEET NAME ===
    DateParts = Split(ws.Name, "_")
    EnsureYearInCombo UF_ControlCentre.cboYear, CLng(DateParts(1))
    EnsureMonthInCombo UF_ControlCentre.cboMonth, Hepler_GetMonthNumber(DateParts(2))
    
    ' === LOOP EmployeeS ===
    For empRow = FIRST_EMP_ROW To LAST_EMP_ROW
       
        If Not Helper_IsBlankRow(empRow) Then
            empID = ws.Cells(empRow, 1).value
            ' === LOOP WEEKS ===
            For weekIndex = 1 To 6   ' max possible weeks
                rng = GetDateRange(rtWeekly, CLng(DateParts(1)), Hepler_GetMonthNumber(DateParts(2)), weekIndex)
                ' --- SAFE WEEK VALIDATION ---
                If Not IsEmpty(rng) Then
                
                     For dayOffset = 0 To DAYS_IN_WEEK - 1
                        rawCode = ws.Cells( _
                            empRow, _
                            ((weekIndex - 1) * (PAYROLLRAWDATA_WIDTH + DAYS_IN_WEEK)) + FIRST_DAY_COL + dayOffset _
                            ).value

                        workDate = DateAdd("d", dayOffset - (weekIndex - 1), rng(1))
                        normCode = NormaliseAttendanceCode(rawCode)
                        If Not normCode = "NullString" Then WriteAttendanceRow wsHistA, empID, workDate, normCode, ws.Name, lastHistRowA
                        
                    Next dayOffset

                End If
                
            Next weekIndex
        End If
    Next empRow
    
    'audit import the lastHistRowsA -2 as they started on row 2 (headers) and its ready to enter a new row
    
    Data_LogImport wsHistA.Name, "Imported", lastHistRowA - 2, ws.Name


End Sub
Function GetWeekRange(ByVal yr As Long, ByVal mn As Long, ByVal wk As Long) As Variant
    Dim result(1 To 7) As Date
    Dim firstDayMonth As Date
    Dim firstMonday As Date
    Dim i As Long

    firstDayMonth = DateSerial(yr, mn, 1)

    ' Find first Monday ON or BEFORE the 1st of the month
    firstMonday = firstDayMonth
    Do While Weekday(firstMonday, vbMonday) <> 1
        firstMonday = firstMonday - 1
    Loop

    ' Build full week
    For i = 1 To 7
        result(i) = firstMonday + (wk - 1) * 7 + (i - 1)
    Next i

    GetWeekRange = result
End Function





Function SalaryWeeklyPay( _
    annualSalary As Double, _
    daysWorkedThisWeek As Double _
) As Double

    Const WORKING_DAYS_YEAR As Long = 260

    SalaryWeeklyPay = _
        Round((annualSalary / WORKING_DAYS_YEAR) * daysWorkedThisWeek, 2)

End Function
Function WorkingDaysInMonth(ByVal yr As Long, ByVal mn As Long) As Long
    Dim d As Date, lastDay As Date, count As Long

    d = DateSerial(yr, mn, 1)
    lastDay = DateSerial(yr, mn + 1, 0)

    Do While d <= lastDay
        If Weekday(d, vbMonday) <= 5 Then count = count + 1
        d = d + 1
    Loop

    WorkingDaysInMonth = count
End Function
Function GetMonthlyHours(ws As Worksheet, _
                           rowNum As Long, _
                           startCol As Long, _
                           weekCount As Long) As Double


    Const Offset As Long = 1      ' relative to week start
    

    Dim total As Double
    Dim w As Long
    Dim weekBlock As Variant

    For w = 0 To weekCount - 1

        ' Read one week block (12 columns) in one hit
        weekBlock = ws.Cells(rowNum, startCol + w * DAYS_IN_WEEK) _
                        .Resize(1, DAYS_IN_WEEK).value

        ' Pull only what you need
        If IsNumeric(weekBlock(1, Offset)) Then
            totalHours = totalHours + weekBlock(1, Offset)
            Debug.Print " week"; weekBlock(1, Offset)
        End If

    Next w

    GetMonthlyHours = total

End Function



Sub WriteWeekHeader(ws As Worksheet, current_Row As Long, current_Col As Long, Header As String)
    With ws
        .Cells(current_Row, current_Col).value = Header
        .range(.Cells(current_Row, current_Col), .Cells(current_Row, current_Col + 3)).Merge
        .Cells(current_Row, current_Col).HorizontalAlignment = xlCenter
        .Cells(current_Row, current_Col).Font.Bold = True
    End With
End Sub






   

Sub xxPopulateYearAndMonthListBoxesForPayRoll()

    Dim wsData As Worksheet
    Dim wsUI As Worksheet
    Dim dictYears As Object
    Dim dictMonths As Object
    Dim lastRow As Long
    Dim i As Long
    Dim arrYears() As Variant
    Dim arrMonths() As Variant

    Set wsData = Worksheets("MonthlyHistory")
    Set wsUI = ActiveSheet ' or specify the sheet with the listboxes
    Set dictYears = CreateObject("Scripting.Dictionary")
    Set dictMonths = CreateObject("Scripting.Dictionary")

    lastRow = wsData.Cells(wsData.Rows.count, "B").End(xlUp).row

    ' Collect distinct Years and Months
    For i = 2 To lastRow ' assumes headers in row 1
        If wsData.Cells(i, "B").value <> "" Then
            dictYears(wsData.Cells(i, "B").value) = 1
        End If
        If wsData.Cells(i, "C").value <> "" Then
            dictMonths(wsData.Cells(i, "C").value) = 1
        End If
    Next i

    ' Convert to arrays
    arrYears = dictYears.Keys
    arrMonths = dictMonths.Keys

    ' Assign
    
    range("o1").Resize(UBound(arrYears) + 1, 1).value = _
    Application.Transpose(arrYears)
    
    Dim outArr() As Variant
    Dim m As Long
    
    ReDim outArr(1 To UBound(arrMonths) - LBound(arrMonths) + 1, 1 To 1)
    
    For m = LBound(arrMonths) To UBound(arrMonths)
        outArr(m - LBound(arrMonths) + 1, 1) = MonthName(arrMonths(m))
    Next m
    
    range("P1").Resize(UBound(outArr), 1).value = outArr
    End Sub

Sub TrackEmployeeLeave()

    Dim wsData As Worksheet, wsSummary As Worksheet
    Dim lastRow As Long, summaryRow As Long
    Dim dict As Object
    Dim key As String
    Dim i As Long
    Dim keyv As Variant
    

    Set wsData = ThisWorkbook.Sheets("weeklyHistory")
    Set wsSummary = ThisWorkbook.Sheets("Holiday_Tracker")
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row

    ' Loop through weekly history
    For i = 2 To lastRow

        If wsData.Cells(i, 12).value = 1 Or wsData.Cells(i, 10).value > 0 Then 'AbsentFlag or HolidayDays

            key = wsData.Cells(i, 1).value & "|" & _
                  wsData.Cells(i, 2).value & "|" & _
                  wsData.Cells(i, 3).value

            If Not dict.Exists(key) Then
                dict.Add key, Array(0, 0)
            End If

            ' Add leave days
            dict(key)(0) = dict(key)(0) + wsData.Cells(i, 10).value

            ' Count absent weeks
            If wsData.Cells(i, 12).value = 1 Then
                dict(key)(1) = dict(key)(1) + 1
            End If

        End If
    Next i

    ' Output results
    summaryRow = 2
    Dim parts
    For Each keyv In dict.Keys
        parts = Split(key, "|")
        wsSummary.Cells(summaryRow, 1).value = parts(0)
        wsSummary.Cells(summaryRow, 2).value = parts(1)
        wsSummary.Cells(summaryRow, 3).value = parts(2)
        wsSummary.Cells(summaryRow, 4).value = dict(key)(0)
        wsSummary.Cells(summaryRow, 5).value = dict(key)(1)
        summaryRow = summaryRow + 1
    Next keyv

    MsgBox "Leave tracking completed!", vbInformation

End Sub


Private Function Helper_CreateTheRequiredSheet(ByVal sheetName As String) As Worksheet

Dim response As VbMsgBoxResult
Dim ws As Worksheet

    

RetryCreate:
    On Error GoTo ErrHandler
    Set ws = Helper_GetOrCreateSheet(sheetName, True, True)
    On Error GoTo 0
    
    'MsgBox "Sheet created successfully.", vbInformation
    Set Helper_CreateTheRequiredSheet = ws
    Exit Function

ErrHandler:
    response = MsgBox( _
        Err.Description & vbCrLf & vbCrLf & _
        "Do you want to delete the existing sheet and recreate it?", _
        vbYesNoCancel + vbExclamation, _
        "Sheet Already Exists")
        
    Select Case response
        Case vbYes
            Application.DisplayAlerts = False
            Worksheets(sheetName).Delete
            Application.DisplayAlerts = True
            Resume RetryCreate
        Case vbNo, vbCancel
            Set Helper_CreateTheRequiredSheet = Nothing
    End Select
    Err.Clear
    Exit Function
    
    
End Function

Private Function xxxHelper_GetOrCreateSheet( _
    ByVal sheetName As String, _
    ByVal useTemplate As Boolean, _
    Optional ByVal createOnly As Boolean = False _
) As Worksheet

    Dim ws As Worksheet
    Dim wsTemplate As Worksheet

    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    ' CREATE ONLY
    If createOnly Then
        If Not ws Is Nothing Then
            Err.Raise vbObjectError + 514, _
                "GetOrCreateSheet", _
                "Worksheet '" & sheetName & "' already exists."
        End If

        If useTemplate Then
            Set wsTemplate = Worksheets("Employee Template")
            wsTemplate.Copy After:=Worksheets(Worksheets.count)
            Set ws = ActiveSheet
            ws.Tab.Color = vbYellow
        Else
            Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))
        End If

        ws.Name = sheetName
        Set Helper_GetOrCreateSheet = ws
        Exit Function
    End If

    ' GET ONLY
    If ws Is Nothing Then
        Err.Raise vbObjectError + 515, _
            "GetOrCreateSheet", _
            "Worksheet '" & sheetName & "' does not exist."
    End If

    Set Helper_GetOrCreateSheet = ws
End Function


Sub Rpt_CreateHolidayBalances(Selected_year As String)

    Dim wsHist As Worksheet, wsRpt As Worksheet
    Dim Balances As Object
    Dim empID As Variant
    Dim r As Long

    
    Set wsHist = Worksheets("HolidayBalances")
    Set wsRpt = Worksheets("Holiday_Balances")
    
    Set Balances = EmployeeHolidayBalance(CDbl(Selected_year))
    
    If Balances Is Nothing Then
        MsgBox "No holiday balance data exists for " & Selected_year, vbInformation
        Exit Sub
    End If
  
    ' show the  report details
    With wsRpt
        
        .Cells(4, 3).value = Now()
        .Cells(5, 3).value = Selected_year
        .Cells(6, 3).value = Audit_constant(5)
        .Cells(7, 3).value = Application.UserName
    End With
    
       ' === LOOP EmployeeS ===
    For r = FIRST_EMP_ROW_FINANCIALS To LAST_EMP_ROW_FINANCIALS
        
        If IsNumeric(wsRpt.Cells(r, 1).value) Then
            empID = wsRpt.Cells(r, 1).value
            ' fill in results
            With wsRpt
                .Cells(r, FIRST_COL).value = Balances(empID)("CarryOverHours")
                .Cells(r, FIRST_COL + 1).value = Balances(empID)("CarryOverDays")
            End With
        End If
    Next r

    
   
End Sub

Sub Rpt_CreateAttendanceTracker(Selected_year As String)

    Dim wsRpt As Worksheet
    Dim Employee As Object
    Dim r As Long
    Dim empID As Long
    Dim FilteredBy As String
    
    
    If UF_ControlCentre.cboEmployee.value = "All Employees" Then
        MsgBox "Please select a Employee", vbOKOnly
        Exit Sub
    End If
    
    If UF_ControlCentre.OtnYear.value = True Then
        Set Employee = EmployeeAttendanceData(CDbl(Selected_year))
        FilteredBy = Selected_year
    Else
        Set Employee = EmployeeAttendanceData(CLng(Selected_year), _
        Month(DateValue("01 " & UF_ControlCentre.cboMonth.value & " " & Selected_year)))
        FilteredBy = Selected_year & " and " & UF_ControlCentre.cboMonth.value
    End If
    
    Set wsRpt = Worksheets("Attendance_Tracker")
    
    ' show the  report details
    With wsRpt
        .Cells(4, 3).value = Now()
        .Cells(5, 3).value = FilteredBy
        .Cells(6, 3).value = Audit_constant(5)
        .Cells(7, 3).value = Application.UserName
    End With
    

    ' === LOOP EmployeeS ===
    For r = FIRST_EMP_ROW_TRACKER To LAST_EMP_ROW_TRACKER
        
        If IsNumeric(wsRpt.Cells(r, 1).value) Then
            empID = wsRpt.Cells(r, 1).value
            ' fill in results
            If Employee.Exists(empID) Then
                With wsRpt
                    .Cells(r, FIRST_COL).value = Employee(empID)("AbsentCount")
                    .Cells(r, FIRST_COL + 1).value = Employee(empID)("LateCount")
                    .Cells(r, FIRST_COL + 2).value = Employee(empID)("sickCount")
                    .Cells(r, FIRST_COL + 3).value = Employee(empID)("HoildayCount")
                    .Cells(r, FIRST_COL + 4).value = Employee(empID)("HoildayDays")
                End With
            Else
                ' No attendance data ? default zeros
                With wsRpt
                    .Cells(r, FIRST_COL).Resize(1, 5).value = 0
                End With
            End If
        End If
    Next r
    
End Sub

Sub Rpr_CreateHoildayTracker(Selected_year As String)
    
    Dim wsHist As Worksheet, wsRpt As Worksheet
    Dim Employee As Object
    Dim empID As Variant
    Dim r As Long
    Dim FilteredBy As String
    

    
    Set wsHist = Worksheets("WeeklyHistory")
    Set wsRpt = Worksheets("Hoilday_Tracker")
    

    If UF_ControlCentre.OtnYear.value = True Then
        Set Employee = EmployeeHoildayData(CLng(Selected_year))
        FilteredBy = Selected_year
    Else
        Set Employee = EmployeeHoildayData(CLng(Selected_year), _
        Month(DateValue("01 " & UF_ControlCentre.cboMonth.value & " " & Selected_year)))
        FilteredBy = Selected_year & " and " & UF_ControlCentre.cboMonth.value
    End If
  
    ' show the  report details
    With wsRpt
        .Cells(4, 3).value = Now()
        .Cells(5, 3).value = FilteredBy
        .Cells(6, 3).value = Audit_constant(5)
        .Cells(7, 3).value = Application.UserName
    End With
    
       ' === LOOP EmployeeS ===
    For r = FIRST_EMP_ROW_TRACKER To LAST_EMP_ROW_TRACKER
        
        If IsNumeric(wsRpt.Cells(r, 1).value) Then
            empID = wsRpt.Cells(r, 1).value
            ' fill in results
            If Employee.Exists(empID) Then
                With wsRpt
                    .Cells(r, FIRST_COL).value = Employee(empID)("StartDate")
                    .Cells(r, FIRST_COL + 1).value = Employee(empID)("HolidayAllowanceInDays")
                    .Cells(r, FIRST_COL + 2).value = Round(Employee(empID)("HolidayAllowanceInHours"), 1) 'Days are rounded to 1
                    .Cells(r, FIRST_COL + 3).value = Round(Employee(empID)("HolidaysAccruedInHours"), 2) 'hours are rounded to 2
                    .Cells(r, FIRST_COL + 4).value = Round(Employee(empID)("HolidaysAccruedInDays"), 1) 'Days are rounded to 1
                    .Cells(r, FIRST_COL + 5).value = Round(Employee(empID)("HolidaysTakenInDays"), 1) 'Days are rounded to 1
                    .Cells(r, FIRST_COL + 6).value = Round(Employee(empID)("HolidaysTakenInHours"), 2) 'hours are rounded to 2
                    .Cells(r, FIRST_COL + 7).value = Round(Employee(empID)("HolidaysRemainingInDays"), 1) 'Days are rounded to 1
                    .Cells(r, FIRST_COL + 8).value = Round(Employee(empID)("HolidaysRemainingInHours"), 2) 'hours are rounded to 2
                End With
            Else
                ' No data ? default zeros
                With wsRpt
                    .Cells(r, FIRST_COL).Resize(1, 8).value = 0
                End With
            End If
        End If
    Next r
End Sub


