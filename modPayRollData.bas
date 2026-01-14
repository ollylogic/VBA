Attribute VB_Name = "modPayRollData"
Public Enum CalcKey
    QE_LowerLimit
    QE_UpperLimit
    PensionEmployeeRate
    PensionEmployerRate
    HolidayRate
    DailyHours
    CarryOverDays
    LateDays
    SickDays
    AbsentDays
End Enum

Public Sub WriteWeeklyHistory( _
    ByRef payYear As Long, _
    ByRef payMonth As Long, _
    ByVal wsInput As Worksheet, _
    ByVal wsHistW As Worksheet, _
    ByVal empID As Long, _
    ByVal empRow As Long, _
    ByVal weekIndex As Long, _
    ByVal WkRange As Variant)

    Dim rate As Double, Hoilday As Double, lastDataRow As Long
    
    lastDataRow = wsHistW.Cells(wsHistW.Rows.count, 1).End(xlUp).row + 1

    If lastDataRow < 2 Then lastDataRow = 2
    
    rate = modDateRanges.GetWeeklyValue(wsInput, empRow, weekIndex, 1)
    Hoilday = modDateRanges.GetWeeklyValue(wsInput, empRow, weekIndex, 3)
    
    If Not (rate = 0 And Hoilday = 0) Then
        wsHistW.Cells(lastDataRow, 1).Resize(1, 11).value = Array( _
            empID, _
            payYear, _
            payMonth, _
            DatePart("ww", WkRange(1), vbMonday, vbFirstFourDays), _
            weekIndex, _
            WkRange(1), _
            WkRange(2), _
            rate, _
            modDateRanges.GetWeeklyValue(wsInput, empRow, weekIndex, 2), _
            Hoilday, _
            wsInput.Name _
        )
        
    End If
    
    

End Sub
Public Sub WriteMonthlyHistory( _
    ByRef wsMonthly_history As Worksheet, _
    ByVal empID As Long, _
    ByRef payYear As Long, _
    ByRef payMonth As Long, _
    ByVal GrossPay As Double, _
    ByVal employeeTax As Double, _
    ByVal employeeNI As Double, _
    ByVal employerNI As Double, _
    ByVal annualAllowance As Double, _
    ByVal employeePension As Double, _
    ByVal employerPension As Double, _
    ByVal sheetName As String)
    
    Dim lastDataRow As Long
    
    lastDataRow = wsMonthly_history.Cells(wsMonthly_history.Rows.count, 1).End(xlUp).row + 1
    
    If lastDataRow < 2 Then lastDataRow = 2

    wsMonthly_history.Cells(lastDataRow, 1).Resize(1, 12).value = Array( _
        empID, _
        payYear, _
        payMonth, _
        GrossPay, _
        employeeTax, _
        employeeNI, _
        employerNI, _
        annualAllowance, _
        employeePension, _
        employerPension, _
        Audit_constant(4), _
        sheetName _
        )
    
    

End Sub

Public Function GetEmployeePayrollInformation( _
    ByVal filterYear As Long, _
    ByVal filterMonth As Long) As Object


    Dim dict As Object, emp As Object
    Dim empID As Long, lastRow As Long, r As Long
    Dim ws As Worksheet
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = Worksheets("MonthlyHistory")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    
    For r = 2 To lastRow

        If ws.Cells(r, "B").value = filterYear _
           And ws.Cells(r, "C").value = filterMonth Then
            empID = ws.Cells(r, "A").value
            
            If Not dict.Exists(empID) Then
                Set emp = CreateObject("Scripting.Dictionary")

                emp("GrossWage") = ws.Cells(r, "D").value
                emp("EmployeeTax") = ws.Cells(r, "E").value
                emp("EmployeeNI") = ws.Cells(r, "F").value
                emp("EmployerNI") = ws.Cells(r, "G").value
                emp("TaxAllowance") = ws.Cells(r, "H").value
                emp("EmployeePension") = ws.Cells(r, "I").value
                emp("EmployerPension") = ws.Cells(r, "J").value
                emp("TaxYear") = ws.Cells(r, "K").value
            
                Set dict(empID) = emp
            End If
        End If
    Next r

    Set GetEmployeePayrollInformation = dict

End Function





Public Sub ImportPayrollData(sheetName As String)

    Dim ws As Worksheet
    Dim wsHistW As Worksheet, wsHistM As Worksheet, wsHistA As Worksheet
    Dim ctx As PayrollContext
    Dim employees As Object
    Dim empRow As Long

    ctx = modHelper.GetPayrollContext(UF_ControlCentre)

    Set ws = Worksheets(sheetName)
    Set wsHistW = Worksheets("WeeklyHistory")
    Set wsHistM = Worksheets("MonthlyHistory")
    Set wsHistA = Worksheets("AttendanceHistory")
    Set employees = EmployeeLookup()

    For empRow = 8 To 29
        If Not modHelper.IsBlankRow(empRow) Then
            ProcessEmployee ws, wsHistW, wsHistM, wsHistA, employees, empRow, ctx.payYear, ctx.payMonth
        End If
    Next empRow

    LogImportResults wsHistW, wsHistM, wsHistA, ws.Name

End Sub
Private Sub LogImportResults( _
    wsHistW As Worksheet, _
    wsHistM As Worksheet, _
    wsHistA As Worksheet, _
    sourceSheetName As String)

    
    LogOneHistory wsHistW, sourceSheetName
    LogOneHistory wsHistA, sourceSheetName
    LogOneHistory wsHistM, sourceSheetName

End Sub
Private Sub LogOneHistory( _
    wsHist As Worksheet, _
    sourceSheetName As String)

    Dim lastDataRow As Long
    Dim importedCount As Long

    lastDataRow = wsHist.Cells(wsHist.Rows.count, 1).End(xlUp).row

    ' Row 1 = headers ? data starts at row 2
    importedCount = Application.Max(0, lastDataRow - 1)

    modAudit.Data_LogImport _
        wsHist.Name, _
        "Imported", _
        importedCount, _
        sourceSheetName

End Sub

Private Sub ProcessEmployee( _
    ws As Worksheet, _
    wsHistW As Worksheet, _
    wsHistM As Worksheet, _
    wsHistA As Worksheet, _
    employees As Object, _
    empRow As Long, _
    payYear As Long, _
    payMonth As Long)

    Dim empID As Long
    empID = ws.Cells(empRow, 1).value

    ProcessEmployeeWeeks ws, wsHistW, wsHistA, empID, empRow, payYear, payMonth
    ProcessEmployeeMonthly ws, wsHistM, employees(empID), empID, empRow, payYear, payMonth

End Sub
Private Sub ProcessEmployeeWeeks( _
    ws As Worksheet, _
    wsHistW As Worksheet, _
    wsHistA As Worksheet, _
    empID As Long, _
    empRow As Long, _
    payYear As Long, _
    payMonth As Long)

    Dim weekIndex As Long
    Dim rng As Variant

    For weekIndex = 1 To 6
        rng = modDateRanges.GetDateRange(rtWeekly, payYear, payMonth, weekIndex)
        If IsEmpty(rng) Then Exit For

        WriteWeeklyHistory _
            payYear, payMonth, ws, wsHistW, empID, empRow, weekIndex, rng

       modHRData.WriteEmployeeAttendance ws, wsHistA, empID, empRow, weekIndex, rng, ws.Name
        
    Next weekIndex

End Sub

Private Sub ProcessEmployeeMonthly( _
    ws As Worksheet, _
    wsHistM As Worksheet, _
    emp As Object, _
    empID As Long, _
    empRow As Long, _
    payYear As Long, _
    payMonth As Long)

    
    Dim GrossPay As Double
    Dim employeeNI As Double, employerNI As Double
    Dim employeePension As Double, employerPension As Double
    Dim Allowance As Double
    Dim employeeTax As Double
    Dim workerCategory As String
    Dim pensionStatus As String

    workerCategory = emp("WorkerCategory")
    pensionStatus = emp("PensionStatus")

    GrossPay = modPayrollRules.GetGrossPay( _
                    emp("PayType"), _
                    modPayrollRules.GetMonthlyGrossPay(ws, empRow, 4), _
                    emp("Salary"))
                    
    employeeTax = modPayrollRules.GetEmployeeTax( _
                    emp("PayType"), _
                    emp("TaxCode"), _
                    GrossPay, _
                    emp("Salary"))
    
    
    If modEmployeeRules.PensionApplies(workerCategory, pensionStatus) Then
        Dim qualifyingEarnings As Double
        
        qualifyingEarnings = modPayrollRules.CalculateQualifyingEarnings( _
                            GrossPay)
                            
        If modEmployeeRules.EmployeeContributionApplies(workerCategory, pensionStatus) Then
            employeePension = modPayrollRules.CalculateEmployeePension( _
                        qualifyingEarnings)
        Else
            employeePension = 0
        End If

        If modEmployeeRules.EmployerContributionApplies(workerCategory, pensionStatus) Then
            employerPension = modPayrollRules.CalculateEmployerPension( _
                        emp("PayType"), _
                        qualifyingEarnings)
        Else
            employerPension = 0
        End If
    Else
        employeePension = 0
        employerPension = 0
    End If



    employeeNI = modPayrollRules.CalculateEmployeeNI(GrossPay, emp("NI_Catagory"))
    
    If EmployerNIExempt(emp("DOB"), emp("IsApprentice"), payYear, payMonth) Then
        employerNI = 0
    Else
        employerNI = modPayrollRules.CalculateEmployerNI(GrossPay)
    End If

    Allowance = modPayrollRules.GetAnnualTaxAllowance( _
                         emp("TaxCode"))

    WriteMonthlyHistory _
        wsHistM, empID, payYear, payMonth, GrossPay, employeeTax, employeeNI, employerNI, _
        Allowance, employeePension, employerPension, ws.Name

End Sub



