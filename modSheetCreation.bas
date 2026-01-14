Attribute VB_Name = "modSheetCreation"

Private Sub ProcessWeek( _
    ByVal wsTarget As Worksheet, _
    ByVal TargetYear As Long, _
    ByVal TargetMonthNumber As Long, _
    ByVal TargetWeekIndex As Long, _
    ByVal TargetCol As Long)

    Const HEADER_ROW As Long = 7
    Const DAYS_IN_WEEK As Long = 7
    Const PAYROLL_DAY_OFFSET As Long = 5

    Dim empRow As Long, empID As Long, day_index As Long, dayCol As Long, TargetRow As Long

    Dim key As String
    Dim attendanceDict As Object, statusRules As Object, sheetInfo As Object
    Dim Date_range As Variant

    Dim sheetType As sheet_Type
    Dim isPayroll As Boolean

    ' Cache context
    Set sheetInfo = CurrentSheetInfo()
    sheetType = GetCurrentSheetType()
    isPayroll = (sheetType = payroll)

    ' Load data once
    Set attendanceDict = modHRData.LoadAttendanceHistory(TargetYear, TargetMonthNumber)
    Set statusRules = modHRData.LoadStatusConfig()
    Date_range = modDateRanges.GetDateRange(rtWeekly, TargetYear, TargetMonthNumber, TargetWeekIndex)

    ' Write headings
    TargetRow = sheetInfo("Heading_Row")
    modSheetFormating.WriteSheetWeeklyHeading _
        wsTarget, Date_range, TargetRow, TargetCol, TargetYear, TargetMonthNumber

    ' Column headers
    TargetRow = sheetInfo("First_Data_Row") - 1

    If isPayroll Then
        modSheetFormating.WritePayrollHeaders wsTarget, TargetRow, TargetCol, Date_range
    Else
        modSheetFormating.WriteWeekDays wsTarget, TargetRow, TargetCol, Date_range
    End If

    TargetRow = TargetRow + 1

    ' === Main employee loop ===
    For empRow = TargetRow To sheetInfo("Last_Data_Row")

        If Not modHelper.IsBlankRow(empRow) Then

            empID = wsTarget.Cells(empRow, 1).value

            If isPayroll Then
                SetUpEmployeeInputWeek wsTarget, empRow, empID, TargetCol
            End If

            For day_index = 1 To DAYS_IN_WEEK

                dayCol = IIf(isPayroll, _
                             TargetCol + PAYROLL_DAY_OFFSET + day_index - 1, _
                             TargetCol + day_index - 1)

                key = modHelper.BuildAttendanceKey( _
                            empID, _
                            wsTarget.Cells(HEADER_ROW, dayCol).value)

                If attendanceDict.Exists(key) Then
                    modSheetFormating.ApplyAttendanceStatus _
                        wsTarget, empRow, dayCol, _
                        attendanceDict(key), TargetCol, statusRules
                End If

            Next day_index

        End If

    Next empRow

    ' Borders
    modSheetFormating.ApplyWeeklyBorders _
        wsTarget, TargetCol, sheetInfo("Last_Data_Row")

    ' Formula highlighting (same for all sheet types)
    modSheetFormating.ColourCellsWithformula _
        wsTarget.range( _
            wsTarget.Cells(sheetInfo("First_Data_Row"), TargetCol), _
            wsTarget.Cells(sheetInfo("Last_Data_Row"), TargetCol + sheetInfo("Offset")))

End Sub


Public Sub CreateTheNewSheet(ws As Worksheet)
 
 
    Dim Date_range As Variant
    Dim Start_Date As Date, End_Date As Date
    Dim MonthNumber As Long, weekCount As Long, TargetColumn As Long, weekIndex As Long
    
    Dim info As Object

    modControl.StoreSheetInformation  ' store by mode
    
    Call modHelper.DebugPrintSheetInfoStore
        
    MonthNumber = modHelper.GetMonthNumberFromName(UF_ControlCentre.cboMonth)
    'Get month range  from the form
        
    Date_range = modDateRanges.GetDateRange(rtMonthly, UF_ControlCentre.cboYear.value, MonthNumber)
    
    Start_Date = CDate(Date_range(1))
    End_Date = CDate(Date_range(2))
    
    TargetColumn = CurrentSheetInfo()("Starting_Column")
    
    modSheetFormating.WriteSheetTileAndPeriod ws, Start_Date, End_Date

    weekCount = DateDiff("ww", Start_Date, End_Date, vbMonday, vbFirstFourDays) + 1


    For weekIndex = 1 To weekCount
        ProcessWeek ws, UF_ControlCentre.cboYear, MonthNumber, weekIndex, TargetColumn
        TargetColumn = TargetColumn + CurrentSheetInfo()("Offset")
    Next weekIndex

    ws.Columns.AutoFit

End Sub

Private Sub SetUpEmployeeInputWeek(ByVal ws As Worksheet, ByVal empRow As Long, ByVal empID As Long, ByVal startCol As Long)

    Dim empData As Object
    Dim Emp_PayType As PayTypeEnum
    
    
    Set empData = EmployeeLookup
    
    If Not empData.Exists(empID) Then
        MsgBox "Employee ID not found: " & empID & "Stoping the program, please enter the Employee", vbCritical, "Payroll"
        Exit Sub
    End If
    
    Emp_PayType = empData(empID)("PayType")
    ' set the Formuala up
    
    If Emp_PayType = PaySalary Then
        ws.Cells(empRow, startCol).Formula = "=" & empData(empID)("Salary") & "/52/7*" & _
            ws.Cells(empRow, startCol + 1).Address(False, False)
        ' owners do not get hoilday pay
        ws.Cells(empRow, startCol + 4).Formula = "=" & _
            ws.Cells(empRow, startCol).Address(False, False)
    Else
        ws.Cells(empRow, startCol + 3).Formula = "=" & _
            ws.Cells(empRow, startCol + 2).Address(False, False) & _
            "*" & Round(modPayrollRules.AverageWeeklyPay(empID), 2)

        ws.Cells(empRow, startCol + 4).Formula = "=" & _
            ws.Cells(empRow, startCol).Address(False, False) & _
            "*" & empData(empID)("Rate") & "+" & ws.Cells(empRow, startCol + 3).Address(False, False)
    End If

    ws.Cells(empRow, startCol + 4).NumberFormat = "$#,##0.00"
End Sub



