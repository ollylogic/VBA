Attribute VB_Name = "modSheetFormating"
Public Sub WriteSheetTileAndPeriod(wsTarget As Worksheet, ByVal firstDay As Date, ByVal lastDay As Date)

    If GetCurrentSheetType() = Booking Then ' remove template information
        wsTarget.range("A32").value = ""
        wsTarget.range("B34:B38").value = ""
    End If

    With wsTarget
        .range("C1").value = CurrentSheetInfo()("Sheet_Title")
        .range("A2").value = CurrentSheetInfo()("Sub_Title")
        .range("C2").value = _
            Format(firstDay, "d mmmm yyyy") & " – " & Format(lastDay, "d mmmm yyyy")
    End With

    With wsTarget.Cells(2, 3)
        .Resize(1, 4).Merge
        .Font.Bold = False
        .Font.Size = 11
    End With
    
     With wsTarget.Cells(1, 3)
        .Resize(1, 4).Merge
        .Font.Bold = True
        .Font.Size = 14
    End With

End Sub
Public Sub WritePayrollHeaders( _
    wsTarget As Worksheet, _
    headerRow As Long, _
    TargetCol As Long, _
    Date_range As Variant)

    With wsTarget
        .Cells(headerRow, TargetCol).value = "Worked"
        .Cells(headerRow, TargetCol + 1).value = "Worked Days"
        .Cells(headerRow, TargetCol + 2).value = "Holiday Days"
        .Cells(headerRow, TargetCol + 3).value = "Holiday Pay"
        .Cells(headerRow, TargetCol + 4).value = "Gross Wage"

        modSheetFormating.WriteWeekDays _
            wsTarget, headerRow, TargetCol + 5, Date_range
    End With

End Sub
Public Sub WriteWeekDays(ws As Worksheet, TargetRow As Long, startCol As Long, weekRange As Variant)

    Dim days As Variant
    Dim weekDayIndex As Integer
    Dim i As Long
    Dim dayDate As Date
    Dim col As Long

    days = Array("Monday", "Tuesday", "Wednesday", _
                 "Thursday", "Friday", "Saturday", "Sunday")

    weekDayIndex = Weekday(weekRange(1), vbMonday) ' 1 = Mon

    For i = LBound(days) To UBound(days)

        col = startCol + SetTheTempleteOffset + i

        ' Calculate actual date for this weekday
        dayDate = DateAdd("d", i - (weekDayIndex - 1), weekRange(1))

        ' Store DATE, not text
        ws.Cells(TargetRow, col).value = dayDate

        ' Display only the day name
        ws.Cells(TargetRow, col).NumberFormat = "dddd"

        ' Grey out if outside month
        If Month(dayDate) <> Month(weekRange(1)) Then
            ws.Cells(TargetRow, col).Interior.Color = RGB(211, 211, 211)
        End If
    Next i

End Sub
Public Sub ApplyAttendanceStatus(ws As Worksheet, empRow As Long, dayCol As Long, _
                          status As String, startCol As Long, statusRules As Object)

    Dim parts As Variant
    Dim colourRGB() As String

    If Not statusRules.Exists(status) Then Exit Sub

    parts = Split(statusRules(status), "|")

    ' parts:
    ' 0 = Code
    ' 1 = Colour (R,G,B)
    ' 2 = CountsAsHoliday
    ' 3 = CountsAsWorked
    ' 4 = PayrollImpact

    ws.Cells(empRow, dayCol).value = parts(0)

    colourRGB = Split(parts(1), ",")
    ws.Cells(empRow, dayCol).Interior.Color = _
        RGB(colourRGB(0), colourRGB(1), colourRGB(2))
'    If UCase(parts(2)) = "TRUE" Then
'        ws.Cells(empRow, startCol + 2).value = _
'            ws.Cells(empRow, startCol + 2).value + 1
'    End If
'
'    If (parts(3)) = "TRUE" Then
'        ws.Cells(empRow, startCol + 1).value = _
'            ws.Cells(empRow, startCol + 1).value + 1
'    End If

End Sub
Public Sub ColourCellsWithformula(targetRange As range)


    On Error Resume Next
    targetRange.SpecialCells(xlCellTypeFormulas).Interior.Color = RGB(198, 224, 180)
    
    On Error GoTo 0

End Sub
Public Sub ColourCellIfThreshold(cell As range, value As Long, threshold As Long)

    If value >= threshold Then
        cell.Interior.Color = RGB(255, 199, 206)   ' light red
        cell.Font.Color = RGB(156, 0, 6)           ' dark red font
    Else
        cell.Interior.ColorIndex = xlNone          ' remove colour if below threshold
        cell.Font.Color = RGB(0, 0, 0)            ' default font
    End If

End Sub
Public Sub ApplyWeeklyBorders( _
    ByVal ws As Worksheet, _
    ByVal startCol As Long, _
    ByVal lastDataRow As Long)

    Dim borderRanges As Variant
    Dim i As Long, Offset As Long
    Dim rng As range

    ' Row blocks used in your template
    borderRanges = Array( _
        Array(5, 9), _
        Array(11, 19), _
        Array(21, 25), _
        Array(27, lastDataRow) _
    )
    Offset = CurrentSheetInfo()("Offset")
    
    
    For i = LBound(borderRanges) To UBound(borderRanges)
        Set rng = ws.range( _
            ws.Cells(borderRanges(i)(0), startCol), _
            ws.Cells(borderRanges(i)(1), (startCol + (Offset) - 1)) _
        )
       modSheetFormating.ApplyBorders rng
    Next i

End Sub
Public Sub ApplyBorders(rng As range)

    With rng
        ' Clear existing borders
        .Borders.LineStyle = xlNone

        ' Outside border (thick)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous

        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick

        ' Inside borders (thin)
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous

        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With

End Sub

Public Sub WriteSheetWeeklyHeading( _
    ByVal ws As Worksheet, _
    ByVal weekRange As Variant, _
    ByVal current_Row As Long, _
    ByVal current_Col As Long, _
    ByVal payYear As Long, _
    ByVal payMonth As Long)
    
    If modHelper.SheetTypeName(GetCurrentSheetType()) = "Payroll" Then
    
        With ws.Cells(current_Row, current_Col)
            .value = "Week " & weekIndex
            .Resize(1, (CurrentSheetInfo()("Offset"))).Merge
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        With ws.Cells(current_Row + 1, current_Col)
            .value = "Range :  " & weekRange(1) & "-" & weekRange(2)
            .Resize(1, 5).Merge
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        With ws.Cells(current_Row + 1, current_Col + 5)
            .value = "Days of the week"
            .Resize(1, 7).Merge
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    Else
    
        With ws.Cells(current_Row, current_Col)
            .value = "Range :  " & weekRange(1) & "-" & weekRange(2)
            .Resize(1, (7)).Merge
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        With ws.Cells(current_Row + 1, current_Col)
            .value = "Days of the week"
            .Resize(1, 7).Merge
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    End If
End Sub


