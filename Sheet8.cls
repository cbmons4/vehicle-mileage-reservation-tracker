VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Application.EnableEvents = False
    FleetRefreshMonthlyMileageView False
SafeExit:
    Application.EnableEvents = True
End Sub
