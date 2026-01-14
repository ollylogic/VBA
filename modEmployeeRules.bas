Attribute VB_Name = "modEmployeeRules"
Public Function NormaliseWorkerCategory(ByVal rawValue As Variant) As String

    Dim v As String
    v = UCase(Trim(CStr(rawValue)))

    Select Case v
        Case "YES-EJ", "POSTPONED", "NO-OPT OUT"
            NormaliseWorkerCategory = "JE"

        Case "NO-EW", "YES EW"
            NormaliseWorkerCategory = "EW"

    End Select

End Function

Public Function NormalisePensionStatus(ByVal rawValue As Variant) As String

    Dim v As String
    v = UCase(Trim(CStr(rawValue)))

    Select Case v
        Case "YES-EJ", "YES-EW"
            NormalisePensionStatus = "ENROLLED"
            
        Case "POSTPONED"
            NormalisePensionStatus = "POSTPONED"

        Case "NO-OPT OUT"
            NormalisePensionStatus = "OPTED_OUT"

        Case "NO-EW"
            NormalisePensionStatus = "NOT_JOINED"
        
         Case Else
            Err.Raise vbObjectError + 900, , _
                "Unknown Pension Status: " & rawValue

    End Select

End Function

Public Function PensionApplies( _
    ByVal workerCategory As String, _
    ByVal pensionStatus As String) As Boolean

    Select Case workerCategory

        Case "JE", "EJ"
            PensionApplies = (pensionStatus = "ENROLLED")

        Case "NEJ"
            PensionApplies = (pensionStatus = "ENROLLED")

        Case "EW"
            PensionApplies = False

    End Select

End Function
Public Function EmployerContributionApplies( _
    ByVal workerCategory As String, _
    ByVal pensionStatus As String) As Boolean

    If pensionStatus <> "ENROLLED" Then
        EmployerContributionApplies = False
        Exit Function
    End If

    Select Case workerCategory
        Case "JE", "NEJ"
            EmployerContributionApplies = True
        Case "EW"
            EmployerContributionApplies = False
        Case Else
            Err.Raise vbObjectError + 901, , _
                "Unknown WorkerCategory: " & workerCategory
    End Select

End Function

Public Function EmployeeContributionApplies( _
    ByVal workerCategory As String, _
    ByVal pensionStatus As String) As Boolean

    Select Case pensionStatus
        Case "ENROLLED"
            EmployeeContributionApplies = True
        Case Else
            EmployeeContributionApplies = False
    End Select

End Function



