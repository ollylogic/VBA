Attribute VB_Name = "modDateRanges"
Public Function GetDateRange( _
    ByVal rangeType As rangeType, _
    ByVal yr As Long, _
    ByVal mn As Long, _
    Optional ByVal wk As Long = 0) As Variant

    Select Case rangeType
        Case rtMonthly
            GetDateRange = GetMonthlyRange(yr, mn)

        Case rtWeekly
            If wk <= 0 Then
                Err.Raise vbObjectError + 100, , "WeekIndex is required"
            End If
            GetDateRange = GetWeeklyRangeInMonth(yr, mn, wk)

        Case Else
            Set GetDateRange = Nothing
    End Select

End Function
Private Function GetMonthlyRange(ByVal yr As Long, ByVal mn As Long) As Variant

    Dim r(1 To 2) As Date

    r(1) = DateSerial(yr, mn, 1)
    r(2) = DateSerial(yr, mn + 1, 0)

    GetMonthlyRange = r

End Function
Public Function GetWeeklyValue(ws As Worksheet, empRow As Long, _
                                weekIndex As Long, valueOffset As Long) As Variant


    GetWeeklyValue = ws.Cells( _
        empRow, (4 - 1) + (weekIndex - 1) * (12) + valueOffset _
    ).value
End Function
Private Function GetWeeklyRangeInMonth( _
    ByVal yr As Long, _
    ByVal mn As Long, _
    ByVal wk As Long) As Variant

    Dim r(1 To 2) As Date
    Dim firstDayMonth As Date
    Dim lastDayMonth As Date
    Dim firstMonday As Date

    firstDayMonth = DateSerial(yr, mn, 1)
    lastDayMonth = DateSerial(yr, mn + 1, 0)

    ' First Monday on or before the month
    firstMonday = firstDayMonth
    Do While Weekday(firstMonday, vbMonday) <> 1
        firstMonday = firstMonday - 1
    Loop

    r(1) = firstMonday + (wk - 1) * 7
    r(2) = r(1) + 6

    ' Clip to month
    If r(1) < firstDayMonth Then r(1) = firstDayMonth
    If r(2) > lastDayMonth Then r(2) = lastDayMonth

    ' Outside month ? no range
    If r(1) > lastDayMonth Or r(2) < firstDayMonth Then
        GetWeeklyRangeInMonth = Empty
        Exit Function
    End If

    GetWeeklyRangeInMonth = r

End Function


