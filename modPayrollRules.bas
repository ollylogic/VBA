Attribute VB_Name = "modPayrollRules"
Option Explicit
Private Const PAYROLLRAWDATA_WIDTH As Long = 5                 ' there are 5 cols created, so the app needs to know, to loop through weeks
Private Const DAYS_IN_WEEK As Long = 7
Private Const WEEK_GROSS_OFFSET As Long = 5
Private Const WEEK_BLOCK_WIDTH As Long = PAYROLLRAWDATA_WIDTH + DAYS_IN_WEEK
Public Enum PayTypeEnum
    PayHourly = 1
    PaySalary = 2
End Enum

Public Enum NITypeEnum
    NIEmployee = 1
    NIEmployer = 2
End Enum


Private mCalcMap As Object

Private Function CalcMap() As Object
    If mCalcMap Is Nothing Then
        Set mCalcMap = modHelper.GetCalculationRowMap()
    End If
    Set CalcMap = mCalcMap
End Function
Public Function GetMonthlyTaxAllowance(ByVal TaxCode As String) As Double
    GetMonthlyTaxAllowance = GetAnnualTaxAllowance(TaxCode) / 12
End Function

Public Function GetAnnualTaxAllowance(ByVal TaxCode As String) As Double

    Dim numericCode As Long
    TaxCode = UCase$(Trim$(TaxCode))

    Select Case TaxCode
        Case "BR", "0T", "D0", "D1"
            GetAnnualTaxAllowance = 0
            Exit Function
    End Select

    numericCode = CLng(StripLetters(TaxCode))
    GetAnnualTaxAllowance = numericCode * 10

End Function


Private Function GetTaxRate(ByVal TaxCode As String) As Double

    Select Case UCase(Trim(TaxCode))
        Case "BR": GetTaxRate = 0.2
        Case "D0": GetTaxRate = 0.4
        Case "D1": GetTaxRate = 0.45
        Case "0T": GetTaxRate = 0.2   ' treated as basic unless banded logic added
        Case Else: GetTaxRate = 0   ' normal allowance-based calculation so
    End Select
 
End Function
Public Function GetEmployeeTax( _
    ByVal PayType As PayTypeEnum, _
    ByVal TaxCode As String, _
    ByVal GrossPay As Double, _
    ByVal Salary As Double) As Double

    Dim taxablePay As Double, rate As Double, TaxAllowance As Double
    

    TaxAllowance = GetAnnualTaxAllowance(TaxCode)
    rate = GetTaxRate(TaxCode)

    ' === Special tax codes (no allowance) ===
    If rate > 0 Then
        If PayType = PayHourly Then
            GetEmployeeTax = Round(GrossPay * rate, 2)
        Else
            GetEmployeeTax = Round((Salary / 12) * rate, 2)
        End If
        Exit Function
    End If

    ' === Normal allowance-based tax ===
    If PayType = PayHourly Then
        taxablePay = GrossPay - (TaxAllowance / 12)
    Else
        taxablePay = (Salary / 12) - (TaxAllowance / 12)
    End If

    If taxablePay > 0 Then
        GetEmployeeTax = Round(taxablePay * 0.2, 2)
    Else
        GetEmployeeTax = 0
    End If
    
    GetEmployeeTax = Round(GetEmployeeTax, 2)

End Function


Public Function GetMonthlyGrossPay( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal startCol As Long) As Double

    Dim total As Double
    Dim col As Long
    Dim grossCell As range

    col = startCol

    Do
        Set grossCell = ws.Cells(rowNum, col + 4)

        If IsEmpty(grossCell.value) Then Exit Do

        total = total + grossCell.value

        col = col + WEEK_BLOCK_WIDTH
    Loop

    GetMonthlyGrossPay = total
End Function



Public Function CalculateEmployeePension( _
    ByVal qualifyingEarnings As Currency) As Currency

    
    Dim employeeRate As Double
    
    employeeRate = CalcMap()("Pension_Employee_Rate")
    
    CalculateEmployeePension = _
        Round(qualifyingEarnings * employeeRate, 2)

End Function


Public Function CalculateEmployerPension( _
    ByVal PayType As PayTypeEnum, _
    ByVal qualifyingEarnings As Currency) As Currency

    Dim employerRate As Double
    
    
    Select Case PayType
        Case PayHourly
            employerRate = CalcMap()("Pension_Employer_Rate")
        Case PaySalary
            employerRate = CalcMap()("Pension_Owners_Rate")
    End Select
    
    CalculateEmployerPension = _
        Round(qualifyingEarnings * employerRate, 2)
End Function
Public Function EmployerNIExempt( _
    ByVal dateOfBirth As Date, _
    ByVal isApprentice As Boolean, _
    ByVal payYear As Long, _
    ByVal payMonth As Long) As Boolean
    
    
    Dim refDate As Date
    refDate = PayPeriodEndDate(payYear, payMonth)
    
    Dim age As Long
    age = AgeOnDate(dateOfBirth, refDate)
    
    If age < 21 Then
        EmployerNIExempt = True
    ElseIf age < 25 And isApprentice Then
        EmployerNIExempt = True
    Else
        EmployerNIExempt = False
    End If

End Function
Private Function PayPeriodEndDate( _
    ByVal payYear As Long, _
    ByVal payMonth As Long) As Date

    PayPeriodEndDate = DateSerial(payYear, payMonth + 1, 0)
End Function
Private Function AgeOnDate(ByVal dob As Date, ByVal refDate As Date) As Long
    AgeOnDate = DateDiff("yyyy", dob, refDate)
    If DateSerial(Year(refDate), Month(dob), Day(dob)) > refDate Then
        AgeOnDate = AgeOnDate - 1
    End If
End Function

Public Function CalculateEmployeeNI(ByVal GrossPay As Double, ByVal niCategory As String) As Double

    Dim threshold As Double
    Dim rate As Double

    threshold = CalcMap()("Employee_NI_Threshold")
    If UCase(Trim(niCategory)) = "A" Then
        rate = CDbl(CalcMap()("Employee_NI_Rate_Category_A"))
    Else
        rate = CDbl(CalcMap()("Employee_NI_Rate_Category_DEL"))
    End If
      
    If GrossPay > threshold Then
        CalculateEmployeeNI = (GrossPay - threshold) * rate
    Else
        CalculateEmployeeNI = 0
    End If

    CalculateEmployeeNI = Round(CalculateEmployeeNI, 2)
    
End Function


Public Function CalculateEmployerNI(GrossPay As Double) As Double

    Dim threshold As Double
    Dim rate As Double

    threshold = CalcMap()("Employer_NI_Threshold")
    rate = CalcMap()("Employer_NI_Rate")

  
    If GrossPay > threshold Then
        CalculateEmployerNI = (GrossPay - threshold) * rate
    Else
        CalculateEmployerNI = 0
    End If
    
    CalculateEmployerNI = Round(CalculateEmployerNI, 2)
End Function
Public Function AverageWeeklyPay(empID As Long) As Double
    Dim ws As Worksheet
    Set ws = Worksheets("WeeklyHistory")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    Dim totalHours As Double
    Dim workedHours As Double
    Dim paidWeeks As Integer
    Dim r As Long

    ' Loop backwards (most recent first) the logic here is we need the last 52 weeks of Worked weeks
    For r = lastRow To 2 Step -1

        If ws.Cells(r, 1).value = empID Then
        
            If Not (ws.Cells(r, 2).value = UF_ControlCentre.cboYear.value _
                And ws.Cells(r, 3).value = modHelper.GetMonthNumberFromName(UF_ControlCentre.cboMonth.value)) Then ' need to exclude this year month and week
                workedHours = ws.Cells(r, 8).value

                If workedHours > 0 Then
                    totalHours = totalHours + workedHours
                    paidWeeks = paidWeeks + 1
                End If
            End If

        End If

        If paidWeeks = 52 Then Exit For
    Next r

    If paidWeeks = 0 Then
        AverageWeeklyPay = 0
    Else
        AverageWeeklyPay = totalHours / paidWeeks
        
    End If
End Function

Public Function GetGrossPay(ByVal PayType As PayTypeEnum, _
                      ByVal GrossPay As Double, _
                      ByVal Salary As Double) As Double

     ' ===WORK OUT GROSSPAY ===
    GetGrossPay = 0
    
    Select Case PayType
        Case PayHourly
            GetGrossPay = GrossPay
    Case PaySalary
        GetGrossPay = Salary / 12
    End Select
    
    GetGrossPay = Round(GetGrossPay, 2)
End Function
Public Function CalculateQualifyingEarnings( _
    ByVal GrossPay As Currency) As Currency

    Dim cappedPay As Currency
    Dim QE_LowerLimit As Double
    Dim QE_UpperLimit As Double

    QE_LowerLimit = CalcMap()("QE_LowerLimit")
    QE_UpperLimit = CalcMap()("QE_UpperLimit")

    If GrossPay <= QE_LowerLimit Then
        CalculateQualifyingEarnings = 0
        Exit Function
    End If

    cappedPay = IIf(GrossPay > QE_UpperLimit, QE_UpperLimit, GrossPay)
    
    CalculateQualifyingEarnings = cappedPay - QE_LowerLimit

End Function

Private Function StripLetters(ByVal TaxCode As String) As Long
    Dim i As Long, result As String

    For i = 1 To Len(TaxCode)
        If mid$(TaxCode, i, 1) Like "#" Then
            result = result & mid$(TaxCode, i, 1)
        End If
    Next

    StripLetters = CLng(result)
End Function




Public Function IsPayrollYearLocked(ByVal payYear As Long) As Boolean

    Dim ws As Worksheet
    Set ws = Worksheets("PayRollYears")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, "A").value = payYear Then
            IsPayrollYearLocked = (ws.Cells(r, "B").value = True)
            Exit Function
        End If
    Next r

    ' Year not found = NOT locked
    IsPayrollYearLocked = False

End Function

Public Function DoseYearExits(ByVal payYear As Long) As Boolean

    Dim ws As Worksheet
    Set ws = Worksheets("MonthlyHistory")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, "B").value = payYear Then
            DoseYearExits = True
            Exit Function
        End If
    Next r

    ' Year not found = NOT locked
    DoseYearExits = False

End Function
Public Function DoseMonthExits(ByVal payMonth As String) As Boolean

    Dim ws As Worksheet
    Set ws = Worksheets("MonthlyHistory")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, "C").value = GetMonthNumberFromName(payMonth) Then
            DoseMonthExits = True
            Exit Function
        End If
    Next r

    ' Year not found = NOT locked
    DoseMonthExits = False

End Function
