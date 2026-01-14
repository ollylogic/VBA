VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- Core state ---
Public HasSelectedPayrollSheet As Boolean
Public isYearLocked As Boolean
Public HasMonthlySheet As Boolean
Public YearExistsInData As Boolean
Public MonthExistsInData As Boolean
Public HasSelectedYear  As Boolean



' --- Helper properties (intent-revealing) ---

Public Property Get CanCreateMonth() As Boolean
    CanCreateMonth = Not isYearLocked
End Property


Public Property Get CanRunPayroll() As Boolean
    CanRunPayroll = YearExistsInData And MonthExistsInData
End Property

Public Property Get CanDoYearEnd() As Boolean
    CanDoYearEnd = Not isYearLocked
End Property
Public Property Get CanImport() As Boolean
    CanImport = HasSelectedYear _
                And Not isYearLocked _
                And HasSelectedPayrollSheet
End Property

