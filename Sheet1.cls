VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target Is Nothing Then Exit Sub
    FleetDismissDashboardDateCalendar Target
SafeExit:
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Intersect(Target, Me.Range("D17,G17,B7:D7")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    If Not Intersect(Target, Me.Range("D17,G17")) Is Nothing Then
        FleetRefreshDashboard False
    ElseIf Not Intersect(Target, Me.Range("B7:D7")) Is Nothing Then
        FleetDashboardMileageVehicleChanged
    End If
SafeExit:
    Application.EnableEvents = True
End Sub

