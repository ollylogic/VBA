Attribute VB_Name = "modReports"
 Const FIRST_ROW As Long = 11
 Const LAST_ROW As Long = 32
 Const FIRST_COL As Long = 4

Public Sub CreateMonthlyFinancials(Selected_year As String, Selected_month As String)

    Dim wsHist As Worksheet, wsRpt As Worksheet
    Dim Financials As Object
    Dim empID As Variant
    Dim r As Long

    
    Set wsHist = Worksheets("MonthlyHistory")
    Set wsRpt = Worksheets("Monthly_Financials")
 
    Set Financials = modPayRollData.GetEmployeePayrollInformation(CDbl(Selected_year), _
    CDbl(Month(DateValue("01 " & Selected_month & " " & Selected_year))))
    
  
    ' show the  report details
    With wsRpt
        .Cells(2, 3).value = Audit_constant(4)
        .Cells(3, 3).value = Selected_year & " " & Selected_month
        .Cells(4, 3).value = Now()
        .Cells(5, 3).value = Audit_constant(5)
        .Cells(6, 3).value = Application.UserName
    End With
    
       ' === LOOP EmployeeS ===
    For r = FIRST_ROW To LAST_ROW
        
            empID = wsRpt.Cells(r, 1).value
        If IsNumeric(wsRpt.Cells(r, 1).value) Then
            ' fill in results
            With wsRpt
                .Cells(r, FIRST_COL).value = ZeroAsBlank(Financials(empID)("GrossWage"))
                .Cells(r, FIRST_COL + 1).value = ZeroAsBlank(Financials(empID)("EmployeeNI"))
                .Cells(r, FIRST_COL + 2).value = ZeroAsBlank(Financials(empID)("EmployerNI"))
                .Cells(r, FIRST_COL + 3).value = ZeroAsBlank(Financials(empID)("EmployeeTax"))
                .Cells(r, FIRST_COL + 4).value = ZeroAsBlank(Financials(empID)("TaxAllowance"))
                .Cells(r, FIRST_COL + 5).value = ZeroAsBlank(Financials(empID)("EmployeePension"))
                .Cells(r, FIRST_COL + 6).value = ZeroAsBlank(Financials(empID)("EmployerPension"))
            End With
        End If
    Next r
    
    wsRpt.Activate
    
     
End Sub


Private Function ZeroAsBlank(ByVal v As Double) As Variant
    If v = 0 Then
        ZeroAsBlank = ""
    Else
        ZeroAsBlank = v
    End If
End Function

