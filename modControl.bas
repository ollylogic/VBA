Attribute VB_Name = "modControl"
Option Explicit
Const APP_VERISON As Double = 1.1                      ' The version of the code,used track changes
Const TEMPLATE_NAME As String = "Employee Template"    ' sheet name
Const NUMBER_OF_RAW_PAYROLL_DATA_COLUMNS As Long = 5   ' Amount of columns before the days of the week
Private sheetInfoStore As Object

' Enum  are user defind varibles that make the code easer to read
Public Enum rangeType
        rtMonthly = 1
        rtWeekly = 2
End Enum

Public Enum sheet_Type
        Booking = 1
        payroll = 2
End Enum
Private mCurrentSheetType As sheet_Type



 Public Sub SetCurrentSheetType(ByVal sheet_mode As sheet_Type)
    mCurrentSheetType = sheet_mode
End Sub
Public Function GetCurrentSheetType() As sheet_Type
    GetCurrentSheetType = mCurrentSheetType
End Function
Public Function GetSheetInfoStore() As Object
    If sheetInfoStore Is Nothing Then
        Set sheetInfoStore = CreateObject("Scripting.Dictionary")
    End If
    Set GetSheetInfoStore = sheetInfoStore
End Function
Public Function CurrentSheetInfo() As Object
    Set CurrentSheetInfo = GetSheetInfoStore(GetCurrentSheetType())
End Function
Public Sub StoreSheetInformation()

    Dim dict As Object
    Dim info As Object
    Dim sheetType As sheet_Type

    sheetType = GetCurrentSheetType()
    Set dict = GetSheetInfoStore()

    If Not dict.Exists(sheetType) Then
        Set info = CreateObject("Scripting.Dictionary")
        dict.Add sheetType, info          ' ? KEY FIX
    Else
        Set info = dict(sheetType)
    End If

    ' Common (static) values
    info("Starting_Column") = 4
    info("Heading_Row") = 5
    info("First_Data_Row") = 8
    info("Last_Data_Row") = 29

    Select Case sheetType

        Case payroll
            info("Sheet_Title") = "Monthly Time sheet"
            info("Sub_Title") = "Pay Period"
            info("Offset") = NUMBER_OF_RAW_PAYROLL_DATA_COLUMNS + 7

        Case Booking
            info("Sheet_Title") = "Monthly Holiday Booking sheet"
            info("Sub_Title") = "Booking Period"
            info("Offset") = 7

        Case Else
            Err.Raise vbObjectError + 300, , "Unknown SheetType"

    End Select

End Sub


Private Function GetRequiredDate() As Boolean

    Dim selectedYear As Long
    Dim selectedMonth As Long

    selectedYear = PromptForYear()
    If selectedYear = 0 Then Exit Function
    
    selectedMonth = PromptForMonth()
    If selectedMonth = 0 Then Exit Function
    
    EnsureYearInCombo UF_ControlCentre.cboYear, selectedYear
    EnsureMonthInCombo UF_ControlCentre.cboMonth, selectedMonth
    

    GetRequiredDate = True

End Function
Private Function PromptForMonth() As Long

Dim inputText As String
    Dim m As Long

    inputText = Trim(InputBox( _
        "Enter month number (1–12):", _
        "Month"))
        
    ' Cancel pressed
    If inputText = vbNullString Then Exit Function

     m = CLng(inputText)

    ' Range check
    If m < 1 Or m > 12 Then
        MsgBox "Month must be between 01 and 12.", vbExclamation
        Exit Function
    End If

    PromptForMonth = m

End Function
    


Public Sub PopulateEmployeeCombo(cbo As MSForms.ComboBox)
    Dim ws As Worksheet
    Dim data As Variant
    Dim listData() As Variant
    Dim lastRow As Long
    Dim i As Long

    Set ws = Worksheets("Employee")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    data = ws.range("A2:D" & lastRow).value

    ReDim listData(1 To UBound(data, 1) + 1, 1 To 2)

    listData(1, 1) = 0
    listData(1, 2) = "All Employees"

    For i = 1 To UBound(data, 1)
        listData(i + 1, 1) = data(i, 1)
        listData(i + 1, 2) = Trim(data(i, 3) & " " & data(i, 4))
    Next i

    With cbo
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "0 pt;150 pt"
        .List = listData
        .ListIndex = 0
    End With
End Sub
Public Sub PopulateYearCombo(cbo As MSForms.ComboBox)

    Dim ws As Worksheet
    Dim data As Variant
    Dim dict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim years() As Variant

    Set ws = Worksheets("MonthlyHistory")
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    data = ws.range("B2:B" & lastRow).value
    Set dict = CreateObject("Scripting.Dictionary")

    If IsArray(data) Then
    ' Standard logic for multiple records
        For i = 1 To UBound(data, 1)
            If IsNumeric(data(i, 1)) Then
                dict(CLng(data(i, 1))) = Empty
            End If
        Next i
        
        years = dict.Keys
        QuickSortLong years, LBound(years), UBound(years)
    Else
        years = Array(data) ' Creates a 1D array with 1 element
    End If


    With cbo
        .Clear
        .List = years
    End With
End Sub
Private Sub QuickSortLong(arr As Variant, ByVal first As Long, ByVal last As Long)

    Dim low As Long
    Dim high As Long
    Dim mid As Long
    Dim temp As Variant

    low = first
    high = last
    mid = arr((first + last) \ 2)

    Do While low <= high

        Do While arr(low) < mid
            low = low + 1
        Loop

        Do While arr(high) > mid
            high = high - 1
        Loop

        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp

            low = low + 1
            high = high - 1
        End If

    Loop

    If first < high Then QuickSortLong arr, first, high
    If low < last Then QuickSortLong arr, low, last

End Sub

Public Sub PopulateMonthCombo(cbo As MSForms.ComboBox)

    Dim ws As Worksheet
    Dim data As Variant
    Dim dict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim months(1 To 12) As String
    Dim result() As Variant
    Dim idx As Long

    Set ws = Worksheets("MonthlyHistory")
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    data = ws.range("C2:C" & lastRow).value
    Set dict = CreateObject("Scripting.Dictionary")

    
    If IsArray(data) Then
    ' Standard logic for multiple records
        For i = 1 To UBound(data, 1)
            If IsDate(data(i, 1)) Then
                dict(Month(data(i, 1))) = True
            ElseIf IsNumeric(data(i, 1)) Then
                dict(CLng(data(i, 1))) = True
            End If
        Next i
        
        ReDim result(1 To dict.count)
        idx = 1
        
        For i = 1 To 12
            If dict.Exists(i) Then
                result(idx) = MonthName(i)
                idx = idx + 1
            End If
        Next i
    Else
        result = Array(MonthName(data)) ' Creates a 1D array with 1 element
    End If


    With cbo
        .Clear
        .List = result
    End With
End Sub
Private Function PromptForYear() As Long
Dim inputText As String
    Dim yr As Long

    inputText = Trim(InputBox( _
        "Enter 4-digit year:", _
        "Year", _
        UF_ControlCentre.cboYear.value))
    
    ' Cancel pressed
    If inputText = vbNullString Then Exit Function

    If Len(inputText) <> 4 Or Not IsNumeric(inputText) Then ' condit to reject
        MsgBox "Please enter a valid 4-digit year.", vbExclamation
        Exit Function
    End If

    yr = CLng(inputText)

    If yr < 2000 Or yr > 2099 Then
        MsgBox "Year must be between 2000 and 2099.", vbExclamation
        Exit Function
    End If

    PromptForYear = yr

End Function

Public Sub HandleSheet()

    ' Ask user for Year / Month
    If Not GetRequiredDate() Then Exit Sub

    Dim sheetName As String, userInformation As String
    Dim ws As Worksheet, wsTemplate As Worksheet
    
    'Debug.Print "CurrentSheetType =", modHelper.SheetTypeName(GetCurrentSheetType())
    
    sheetName = BuildSheetName(GetCurrentSheetType())

    
    If Not modAudit.WorksheetExists(sheetName) Then
        Set wsTemplate = Worksheets(TEMPLATE_NAME)
        wsTemplate.Copy After:=Worksheets(Worksheets.count)
        Set ws = ActiveSheet
        ws.Name = sheetName
        ws.Tab.Color = vbYellow

        ' create new sheet
        modSheetCreation.CreateTheNewSheet ws
    Else
        ' show the one that is there
        Set ws = Worksheets(sheetName)
    End If


    '--- Apply layout based on type ---

    Dim sheetLabel As String
    sheetLabel = IIf(GetCurrentSheetType() = Booking, "Booking", "Payroll")

    If WorksheetExists(sheetName) Then
        userInformation = "The " & sheetLabel & _
                          " sheet for " & UF_ControlCentre.cboMonth & _
                          " has already been created"
    Else
        userInformation = "A " & sheetLabel & _
                          " sheet has been created for " & UF_ControlCentre.cboMonth
    End If


    ws.Activate

    UF_ControlCentre.UpdateStatus userInformation

End Sub


Private Function BuildSheetName(ByVal mode As sheet_Type) As String

    Dim yr As Long, mth As Long
    yr = CLng(UF_ControlCentre.cboYear.value)
    mth = CLng(GetMonthNumberFromName(UF_ControlCentre.cboMonth.value))

    Select Case mode
        Case payroll
            BuildSheetName = "PayRoll_" & UF_ControlCentre.cboYear & "_" & UF_ControlCentre.cboMonth
        Case Booking
            BuildSheetName = "Attendance_" & UF_ControlCentre.cboYear & "_" & UF_ControlCentre.cboMonth
    End Select

End Function



Public Sub PopulateCombo_box(ByVal cbo As Object, _
    Optional ByVal defaultText As String = vbNullString)

    ' working with the combo box passed in  populate it with the correct data from the data sets,
    ' if no  data is found a default is used. With cboEmployee add the default in as well
    
    Dim ws As Worksheet
    Dim data As Variant
    Dim listData As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object

    Select Case cbo.Name

       Case "cboEmployee"

            Set ws = Worksheets("Employee")
            lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
            If lastRow < 2 Then Exit Sub
        
            data = ws.range("A2:D" & lastRow).value
        
            ' +1 row for "All Employees"
            ReDim listData(1 To UBound(data, 1) + 1, 1 To 2)
        
            ' Default row
            listData(1, 1) = 0                ' EmployeeID = 0 means "All"
            listData(1, 2) = "All Employees"
        
            For i = 1 To UBound(data, 1)
                listData(i + 1, 1) = data(i, 1)
                listData(i + 1, 2) = Trim(data(i, 3) & " " & data(i, 4))
            Next i
        
            With cbo
                .Clear
                .ColumnCount = 2
                .ColumnWidths = "0 pt;150 pt"
                .List = listData
                .ListIndex = 0   ' select "All Employees"
            End With
        
            Exit Sub

       Case "cboYear"

            Set ws = Worksheets("MonthlyHistory")
            lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
        
            Set dict = CreateObject("Scripting.Dictionary")
        
            If lastRow >= 2 Then
                data = ws.range("B2:B" & lastRow).value
        
                For i = 1 To UBound(data, 1)
                    If IsNumeric(data(i, 1)) Then
                        dict(CLng(data(i, 1))) = Empty
                    End If
                Next i
            End If
        
            ' If no years found ? default to current year
            If dict.count = 0 Then
                dict(Year(Date)) = Empty
            End If
        
            listData = dict.Keys


        Case "cboMonth"

            Set ws = Worksheets("MonthlyHistory")
            lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row
            If lastRow < 2 Then Exit Sub
        
            data = ws.range("C2:C" & lastRow).value
        
            Set dict = CreateObject("Scripting.Dictionary")
        
            ' Collect unique months
            For i = 1 To UBound(data, 1)
                If IsDate(data(i, 1)) Or IsNumeric(data(i, 1)) Then
                    dict((data(i, 1))) = MonthName((data(i, 1)))
                End If
            Next i
        
            ' Output sorted Jan ? Dec
            ReDim listData(1 To dict.count)
        
            Dim idx As Long
            idx = 1
        
            For i = 1 To 12
                If dict.Exists(i) Then
                    listData(idx) = dict(i)
                    idx = idx + 1
                End If
            Next i

        
        Case Else
            Exit Sub
    End Select

    With cbo
        .Clear
        .List = listData
        If defaultText <> vbNullString Then
            .AddItem defaultText
            .value = defaultText
        End If
    End With

End Sub
Public Function GetSavedExportPath()

    Dim wsControl As Worksheet
    
    Set wsControl = Worksheets("Control")

    GetSavedExportPath = wsControl.Cells(7, 10).value
End Function
Public Function SaveExportPath(ByVal The_path As String)

    Dim wsControl As Worksheet
    
    Set wsControl = Worksheets("Control")

    wsControl.Cells(7, 10).value = The_path
End Function

Public Sub SetUpApp()
    
    Dim wsControl As Worksheet
    Dim r As Long
    Dim frm As UF_ControlCentre
    Set frm = UF_ControlCentre
    
    Set wsControl = Worksheets("Control")
    r = 4
    
    ' clear old data
    With wsControl
        Do Until .Cells(r, 1).value = ""
            .Cells(r, 1).value = ""
            r = r + 1
        Loop
    End With
   
    'get the code version out and show it in the control sheet and UF_ControlCentre' ( this is hard coded in application const)
    
    wsControl.Cells(5, 10).value = APP_VERISON
    

    With frm
        .StartUpPosition = 1 ' CenterOwner
        .lblstatus.Caption = "Let's go!"
        .lblAppVer.Caption = APP_VERISON
        .Show
    End With
     
End Sub



