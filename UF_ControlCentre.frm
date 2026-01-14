VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ControlCentre 
   Caption         =   "Payroll & HR Control Centre"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900.001
   OleObjectBlob   =   "UF_ControlCentre.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ControlCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Private State As UIState
Private isUpdatingUI As Boolean


Private Sub RefreshState()

    Set State = New UIState

    Dim selectedYear As Long

    State.HasSelectedYear = _
        (Me.cboYear.ListIndex <> -1)

    If State.HasSelectedYear Then
        selectedYear = CLng(Me.cboYear.value)
        State.isYearLocked = IsPayrollYearLocked(selectedYear)
        State.YearExistsInData = DoseYearExits(selectedYear)
    Else
        State.isYearLocked = False
        State.YearExistsInData = False
    End If

    State.HasMonthlySheet = (Me.cboPayrollSheets.ListCount > 0)
    State.HasSelectedPayrollSheet = _
        (Me.cboPayrollSheets.ListIndex <> -1)

    State.MonthExistsInData = _
        (Me.cboMonth.ListIndex <> -1)

End Sub
Private Sub ReloadMonthlyCombos()
    Dim selectedYear As Long
    selectedYear = CLng(Me.cboYear.value)

    Dim isLocked As Boolean
    isLocked = IsPayrollYearLocked(selectedYear)

    Call LoadMonthlySheetCombos(selectedYear, isLocked)
End Sub


 Private Sub ApplyUIState()

'Debug.Print "HasSelectedYear:", (Me.cboYear.ListIndex <> -1), _
'            "Locked:", State.isYearLocked, _
'            "Selected:", State.HasSelectedPayrollSheet, _
'            "Index:", Me.cboPayrollSheets.ListIndex, _
'            "State.CanImport:", State.CanImport, _
'            "State.CanCreateMonth:", State.CanCreateMonth
'
            
    ApplyButton Me.btnImportPayrollData, _
                State.CanImport, _
                "Import payroll data from the selected sheet", _
                "Select a payroll sheet before importing"
            
    ApplyButton Me.btnCreatePayrollSheet, _
                State.CanCreateMonth, _
                "Create a new monthly payroll sheet", _
                "Year is locked – year end already completed"

    ApplyButton Me.btnRunMonthlyFinancials, _
                State.CanRunPayroll, _
                "Run payroll for the selected month", _
                "Selected year or month has no data"
                
    ApplyButton Me.btnCreateBookingSheet, _
                State.CanCreateMonth, _
                "Create a new monthly Booking sheet", _
                "Selected year or month has no data"
                
    ApplyComboBox Me.cboYear, _
              True, _
              "Select year", _
              "Year selection disabled"

    ApplyComboBox Me.cboYear, _
              True, _
              "Select year", _
              "Month selection disabled"
              
    ApplyComboBox Me.cboPayrollSheets, _
              State.HasMonthlySheet And Not State.isYearLocked, _
              "Select a payroll month", _
              "No payroll months available or year locked"

'
'    ApplyButton Me.cmdLockYear, _
'                State.CanDoYearEnd, _
'                "Lock the payroll year", _
'                "Year is already locked"
'
'    ApplyButton Me.cmdYearEndChecks, _
'                State.CanDoYearEnd, _
'                "Run year-end validation checks", _
'                "Year is already locked"
'
'    ApplyButton Me.cmdDoYearEnd, _
'                State.CanDoYearEnd, _
'                "Complete year end processing", _
'                "Year is already locked"

End Sub
Private Sub ApplyButton(btn As MSForms.CommandButton, _
                        enabled As Boolean, _
                        enabledTip As String, _
                        disabledTip As String)

    btn.enabled = enabled

    If enabled Then
        btn.ControlTipText = enabledTip
        btn.ForeColor = vbBlack
    Else
        btn.ControlTipText = disabledTip
        btn.ForeColor = RGB(160, 160, 160)
    End If

End Sub
Private Sub ApplyComboBox(cbo As MSForms.ComboBox, _
                          enabled As Boolean, _
                          enabledTip As String, _
                          disabledTip As String)

    cbo.enabled = enabled

    If enabled Then
        cbo.ControlTipText = enabledTip
        cbo.ForeColor = vbBlack
        cbo.BackColor = vbWhite
    Else
        cbo.ControlTipText = disabledTip
        cbo.ForeColor = RGB(160, 160, 160)
        cbo.BackColor = RGB(240, 240, 240)
    End If

End Sub

Private Sub UpdateUI()

    If isUpdatingUI Then Exit Sub
    isUpdatingUI = True

    On Error GoTo CleanExit

    RefreshState
    ApplyUIState

CleanExit:
    isUpdatingUI = False

End Sub
Private Sub ShowSheetWhenComplete(ByVal sheetName As String, Information As String)


    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    UpdateStatus Information
    ws.Activate
    
End Sub


Private Sub btnAttendanceTracker_Click()

    Call Rpt_CreateAttendanceTracker(UF_ControlCentre.cboYear.value)
    
    
    ShowSheetWhenComplete "Attendance_Tracker", "The Attendance_Tracker.....have been created"
    
End Sub



Private Sub btnClose_Click()

Unload Me
    
MsgBox "bye bye", vbInformation

End Sub
Private Function Check_PreFlight() As Boolean
    Check_PreFlight = (Data_CountImportedMonths(CLng(cboYear.value)) = 12)
End Function

Private Function Check_PayrollLocked() As Boolean
    Check_PayrollLocked = Data_IsYearAlreadyClosed(cboYear.value)
End Function

Private Function Check_HolidayPosition() As Boolean
    Check_HolidayPosition = Data_HasPostiveHolidayBalances(cboYear.value)
End Function


Private Function Check_Confirmation() As Boolean
    Check_Confirmation = (MsgBox("Confirm year end?", _
        vbYesNo + vbQuestion) = vbYes)
End Function


Private Sub btnCreatebookingSheet_Click()

    If Not State.CanCreateMonth Then Exit Sub
    
     SetCurrentSheetType Booking
     modControl.HandleSheet

End Sub




Private Sub btnCreatePayrollSheet_Click()

    If Not State.CanCreateMonth Then Exit Sub
    SetCurrentSheetType payroll
    modControl.HandleSheet

End Sub

Private Sub btnDoYearEnd_Click()

    If Not btnDoYearEnd.enabled Then
        MsgBox "Year End checks not passed.", vbExclamation
        Exit Sub
    End If

    If MsgBox("This will close the payroll year. Continue?", _
              vbYesNo + vbCritical) = vbNo Then Exit Sub

    RunYearEnd
    UpdateStatus "Year End completed successfully."

End Sub


Private Sub btnHoildayTracker_Click()

    Call Rpr_CreateHoildayTracker(UF_ControlCentre.cboYear.value)
    
    ShowSheetWhenComplete "Hoilday_Tracker", "Holiday Tracker.....have been created"
    
    
End Sub

Private Sub btnImportPayrollData_Click()

    If Not State.CanImport Then Exit Sub
     
    ' === PARSE YEAR / MONTH FROM SHEET NAME ===
    Dim ws As Worksheet
    Dim ym As YearMonthInfo
    
    Set ws = Worksheets(Me.cboPayrollSheets.value)
    
    Set ym = ParseYearAndMonthFromSheetName(ws.Name)

    EnsureYearInCombo UF_ControlCentre.cboYear, ym.Year
    EnsureMonthInCombo UF_ControlCentre.cboMonth, ym.Month
    
    
    If Not modAudit.ValidateImport("WeeklyHistory", Me.cboPayrollSheets.value) Then
        Exit Sub
    End If

    ' Safe to import
    UpdateStatus "Importing raw data to calculate the monthly Payroll..."
    
    modPayRollData.ImportPayrollData Me.cboPayrollSheets.value ' this splits the imported data in to 3 data stores so picked WeeklyHistory
     UpdateStatus "Monthly Payroll .....has been filled"
    
End Sub



Private Sub btnLockYear_Click()

    Dim yearEnd As Date
    
    yearEnd = Audit_constant(8)
    If MsgBox("This will close the year. Continue?", vbYesNo + vbCritical) = vbNo Then Exit Sub
         
    WriteHolidayBalancesYearEnd (Me.cboYear.value)
    
    
    
    UpdateStatus "Holiday balances have been written for year ending " & Me.cboYear.value
    
End Sub



Private Sub btnStartNewYear_Click()

    Dim User_information As String
    
    If AddTheNextYearToTheComboBox(UF_ControlCentre.cboYear.value) = True Then
    User_information = "The new year has been added and the form reset"
    Me.btnStartNewYear.enabled = False ' turn it off as you will not be setting up a new year aqain
    Else
    User_information = "The new year has not been added"
    End If
    
    UpdateStatus User_information
    
    
End Sub

Private Sub btnSaveCode_Click()
    
    If Not CanAccessVBProject Then
    MsgBox _
        "VBA export requires:" & vbCrLf & vbCrLf & _
        "File ? Options ? Trust Center ? Macro Settings" & vbCrLf & _
        "? Trust access to the VBA project object model", _
        vbCritical, "VBA Access Blocked"
    Exit Sub
    End If
    
    Dim exportPath As String
    exportPath = modControl.GetSavedExportPath

    If exportPath = "" Then
        exportPath = modVBAExport.PickFolder("Select folder to export VBA code")
        If exportPath = "" Then Exit Sub
    End If
    
    modVBAExport.ExportVBA exportPath
    modControl.SaveExportPath exportPath

    Me.UpdateStatus "The code has been saved down ready for storage"
    
End Sub

Private Sub btnYearEndChecks_Click()

    Dim allPassed As Boolean
    allPassed = True

    ' 1. Pre-flight validation
    If Check_PreFlight() Then
        SetCheckStatus lblYearEnd_step1, True
    Else
        SetCheckStatus lblYearEnd_step1, False
        allPassed = False
    End If

    ' 2. Payroll lock
    If Check_PayrollLocked() Then
        SetCheckStatus lblYearEnd_step2, True
    Else
        SetCheckStatus lblYearEnd_step2, False
        allPassed = False
    End If

    ' 3. Holiday position
    If Check_HolidayPosition() Then
        SetCheckStatus lblYearEnd_step3, True
    Else
        SetCheckStatus lblYearEnd_step3, False
        allPassed = False
    End If

    ' 5. Confirmation
    If Check_Confirmation() Then
        SetCheckStatus lblYearEnd_step4, True
    Else
        SetCheckStatus lblYearEnd_step4, False
        allPassed = False
    End If

    ' Enable Year End only if ALL passed
    btnDoYearEnd.enabled = allPassed

    If allPassed Then
        UpdateStatus "All checks passed. Ready to run Year End."
    Else
        UpdateStatus "Year End blocked. Fix failed checks."
    End If

End Sub
Private Sub SetCheckStatus(lbl As MSForms.Label, passed As Boolean)

    If passed Then
        lbl.Caption = "Pass"
        lbl.ForeColor = vbGreen
    Else
        lbl.Caption = "Failed"
        lbl.ForeColor = vbRed
    End If

End Sub
Private Sub ResetCheckStatus(lbl As MSForms.Label)
    lbl.Caption = "?"
    lbl.ForeColor = RGB(150, 150, 150)
End Sub




Private Sub btnRunMonthlyFinancials_Click()

    If Not State.CanRunPayroll Then Exit Sub
    
    Me.UpdateStatus "Creating the monthly financials....."
    modReports.CreateMonthlyFinancials Me.cboYear.value, Me.cboMonth.value
    Me.UpdateStatus "Monthly_Financials.....have been created"
End Sub


Private Sub cbobookingsheets_Change()

    If Me.cbobookingsheets.value <> "" Then
        Me.CmdImportbookings.enabled = True
    End If
    
    
End Sub

Private Sub cboPayrollSheets_Change()

  If isUpdatingUI Then Exit Sub
  UpdateUI
End Sub

Private Sub cboYear_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub CmdImportbookings_Click()

    Call Data_StoreAttendanceFromSheet(Me.cbobookingsheets.value)
    
    UpdateStatus "The Attendance form the booking sheet has been imported"
    
End Sub

Private Sub FraMonthly_Click()

End Sub

Private Sub OtnYear_Click()

End Sub

Private Sub UserForm_Initialize()

    ' Populate combos

    PopulateEmployeeCombo Me.cboEmployee
    PopulateYearCombo Me.cboYear
    PopulateMonthCombo Me.cboMonth

    ' Default page
    mpMain.value = 0
    UpdateStatus "Payroll Tasks"


    UpdateUI
    If Me.cboYear.ListCount > 0 Then ' Auto-select the first year on load:
        Me.cboYear.ListIndex = 0
    End If
End Sub
Private Sub EnsureYearInCombo(cbo As Object, ByVal yearValue As Long)

    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If CLng(cbo.List(i)) = yearValue Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

    cbo.AddItem yearValue
    cbo.ListIndex = cbo.ListCount - 1
    UpdateUI
    
End Sub
Private Sub cboYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim yr As Long

    If Trim(Me.cboYear.value) = "" Then Exit Sub

    If Not IsNumeric(Me.cboYear.value) Then
        MsgBox "Please enter a valid 4-digit year.", vbExclamation
        Cancel = True
        Exit Sub
    End If

    yr = CLng(Me.cboYear.value)

    If yr < 2000 Or yr > Year(Date) + 5 Then
        MsgBox "Year must be between 2000 and " & Year(Date) + 5, vbExclamation
        Cancel = True
        Exit Sub
    End If

    ' Add year if not already present
    EnsureYearInCombo Me.cboYear, yr

End Sub


' ===============================
' PAGE SWITCHING
' ===============================
Private Sub mpMain_Change()
    Select Case mpMain.value
        Case 0: UpdateStatus "Payroll Tasks"
        Case 1: UpdateStatus "HR Tasks"
    End Select
End Sub

Public Sub UpdateStatus(msg As String)
    lblstatus.Caption = msg
End Sub
Public Sub UpdateUIState()

    ' --- Create Month ---
    Me.cmdCreateMonth.enabled = Not isYearLocked

    ' --- Import / Re-import ---
    Me.cmdImport.enabled = _
        Not isYearLocked _
        And Me.cboPayrollSheets.ListCount > 0

    Me.cmdReImport.enabled = Me.cmdImport.enabled

    ' --- Run Payroll ---
    Me.cmdRunPayroll.enabled = _
        YearExistsInData _
        And MonthExistsInData

    ' --- Year End ---
    Me.cmdLockYear.enabled = Not isYearLocked
    Me.cmdYearEndChecks.enabled = Not isYearLocked
    Me.cmdDoYearEnd.enabled = Not isYearLocked

End Sub

Private Sub cboYear_Change()
    ReloadMonthlyCombos
    UpdateUI
End Sub

Private Sub cboMonth_Change()
     UpdateUI
End Sub





Private Sub btnYearEnd_Click()
    UpdateStatus "Year-End Closing..."

    If MsgBox("This will close the year and cannot be undone. Continue?", _
              vbYesNo + vbCritical) = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    WriteYearEndBalances
    LockYear
    AuditYearEnd
    RollForwardSettings

    Application.ScreenUpdating = True

    MsgBox "Year End completed successfully.", vbInformation
    
    UpdateStatus "Year-End Closed..."



End Sub
Public Sub LoadMonthlySheetCombos( _
        ByVal selectedYear As Long, _
        ByVal isYearLocked As Boolean)

    Dim ws As Worksheet
    Dim wsC As Worksheet
    Dim r As Long
    Dim sheetYear As Long
    Dim parts() As String
    Dim prefix As String

    Set wsC = Worksheets("Control")
    r = 4

    ' Reset combos
    UF_ControlCentre.cboPayrollSheets.Clear
    UF_ControlCentre.cbobookingsheets.Clear

    For Each ws In ThisWorkbook.Worksheets

        If InStr(ws.Name, "_") = 0 Then GoTo NextSheet

        parts = Split(ws.Name, "_")
        prefix = parts(0)

        If prefix <> "PayRoll" And prefix <> "Attendance" Then GoTo NextSheet
        If Not IsNumeric(parts(1)) Then GoTo NextSheet

        sheetYear = CLng(parts(1))

        ' Show all sheets in control sheet (debug/visibility)
        wsC.Cells(r, 1).value = ws.Name
        r = r + 1

        ' Add to combos only if allowed
        If sheetYear = selectedYear And Not isYearLocked Then
            If prefix = "PayRoll" Then
                UF_ControlCentre.cboPayrollSheets.AddItem ws.Name
            Else
                UF_ControlCentre.cbobookingsheets.AddItem ws.Name
            End If
        End If

NextSheet:
    Next ws

End Sub

Private Sub btnPayrollSummary_Click()
    MsgBox "Payroll Summary (stub)", vbInformation
End Sub

Private Sub btnHolidayBalances_Click()
    
    Call Rpt_CreateHolidayBalances(UF_ControlCentre.cboYear.value)
    
End Sub

Private Sub btnWarnings_Click()
    MsgBox "Absence Warnings (stub)", vbExclamation
End Sub



' ===============================
' REPORTS BUTTONS (STUBS)
' ===============================
Private Sub btnExportPayroll_Click()
    MsgBox "Export Payroll (stub)", vbInformation
End Sub

Private Sub btnExportHolidays_Click()
    MsgBox "Export Holidays (stub)", vbInformation
End Sub

Private Sub btnExportAttendance_Click()
    MsgBox "Export Attendance (stub)", vbInformation
End Sub

