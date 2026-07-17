VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Intersect(Target, Me.Range("C3,F3,C4")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    If Not Intersect(Target, Me.Range("C3,F3")) Is Nothing Then
        FleetWeeklyMonthYearChanged
    Else
        FleetGenerateWeeklyReservations True
    End If
SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target.CountLarge > 1 Then Exit Sub
    If Intersect(Target, Me.Range("A8:G13")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    FleetHandleWeeklyCalendarSelection Target
SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Application.EnableEvents = False
    FleetGenerateWeeklyReservations False
SafeExit:
    Application.EnableEvents = True
End Sub
