Attribute VB_Name = "modFleetUndoActions"
Option Explicit

Private Const EDIT_APP_TITLE As String = "Fleet Mileage Tool"
Private Const EDIT_MENU_SIZE As Long = 15
Private Const EDIT_SETUP_SHEET As String = "Setup"
Private Const MILEAGE_EDIT_CELL As String = "B13"
Private Const RESERVATION_EDIT_CELL As String = "B14"
Private Const MILEAGE_CANDIDATE_FIRST_ROW As Long = 20
Private Const RESERVATION_CANDIDATE_COLUMN As Long = 3
Private Const MILEAGE_EDIT_BUTTON As String = "btnUndoLastMileage"
Private Const RESERVATION_EDIT_BUTTON As String = "btnUndoLastCheckout"

Public Function FleetMileageEditActive() As Boolean
    FleetMileageEditActive = (Len(Trim$(CStr(EditSetupSheet().Range(MILEAGE_EDIT_CELL).value))) > 0)
End Function

Public Function FleetReservationEditActive() As Boolean
    FleetReservationEditActive = (Len(Trim$(CStr(EditSetupSheet().Range(RESERVATION_EDIT_CELL).value))) > 0)
End Function

Public Sub FleetEditMileage()
    On Error GoTo CleanFail

    Dim mileageTable As ListObject
    Dim popup As Object
    Dim rowIndex As Long
    Dim slot As Long
    Dim logId As Variant

    Set mileageTable = EditGetTable("tblMileageLog")
    EditSetupSheet().Range("B20:B34").ClearContents
    Set popup = EditCreatePopup("FleetRecentMileageMenu")

    For rowIndex = mileageTable.ListRows.Count To 1 Step -1
        logId = EditTableValue(mileageTable, rowIndex, "Log ID")
        If IsNumeric(logId) Then
            slot = slot + 1
            EditSetupSheet().Cells(MILEAGE_CANDIDATE_FIRST_ROW + slot - 1, 2).value = CLng(logId)
            EditAddPopupItem popup, CStr(slot) & ". " & EditMileageDescription(mileageTable, rowIndex), _
                "FleetSelectMileageEdit" & CStr(slot)
            If slot >= EDIT_MENU_SIZE Then Exit For
        End If
    Next rowIndex

    If slot = 0 Then
        popup.Delete
        MsgBox "No mileage entries are available to edit.", vbInformation, EDIT_APP_TITLE
        Exit Sub
    End If

    If Application.Visible Then popup.ShowPopup
    popup.Delete
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not popup Is Nothing Then popup.Delete
    MsgBox "Could not open recent mileage entries: " & Err.Description, vbExclamation, EDIT_APP_TITLE
End Sub

Public Sub FleetEditReservation()
    On Error GoTo CleanFail

    Dim reservationTable As ListObject
    Dim popup As Object
    Dim rowIndex As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim slot As Long
    Dim idsCsv As String

    Set reservationTable = EditGetTable("tblReservations")
    With EditSetupSheet().Range("C20:C34")
        .ClearContents
        .numberFormat = "@"
    End With
    Set popup = EditCreatePopup("FleetRecentReservationMenu")

    rowIndex = reservationTable.ListRows.Count
    Do While rowIndex >= 1 And slot < EDIT_MENU_SIZE
        If EditReservationRowCanBeEdited(reservationTable, rowIndex) Then
            EditReservationGroup reservationTable, rowIndex, firstRow, lastRow
            idsCsv = EditReservationIdsForRows(reservationTable, firstRow, lastRow)
            If idsCsv <> "" Then
                slot = slot + 1
                EditSetupSheet().Cells(MILEAGE_CANDIDATE_FIRST_ROW + slot - 1, RESERVATION_CANDIDATE_COLUMN).Value2 = CStr(idsCsv)
                EditAddPopupItem popup, CStr(slot) & ". " & EditReservationDescription(reservationTable, firstRow, lastRow), _
                    "FleetSelectReservationEdit" & CStr(slot)
            End If
            rowIndex = firstRow - 1
        Else
            rowIndex = rowIndex - 1
        End If
    Loop

    If slot = 0 Then
        popup.Delete
        MsgBox "No reservations are available to edit.", vbInformation, EDIT_APP_TITLE
        Exit Sub
    End If

    If Application.Visible Then popup.ShowPopup
    popup.Delete
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not popup Is Nothing Then popup.Delete
    MsgBox "Could not open recent reservations: " & Err.Description, vbExclamation, EDIT_APP_TITLE
End Sub

Public Sub FleetSelectMileageEdit1()
    EditLoadMileageCandidate 1
End Sub

Public Sub FleetSelectMileageEdit2()
    EditLoadMileageCandidate 2
End Sub

Public Sub FleetSelectMileageEdit3()
    EditLoadMileageCandidate 3
End Sub

Public Sub FleetSelectMileageEdit4()
    EditLoadMileageCandidate 4
End Sub

Public Sub FleetSelectMileageEdit5()
    EditLoadMileageCandidate 5
End Sub

Public Sub FleetSelectMileageEdit6()
    EditLoadMileageCandidate 6
End Sub

Public Sub FleetSelectMileageEdit7()
    EditLoadMileageCandidate 7
End Sub

Public Sub FleetSelectMileageEdit8()
    EditLoadMileageCandidate 8
End Sub

Public Sub FleetSelectMileageEdit9()
    EditLoadMileageCandidate 9
End Sub

Public Sub FleetSelectMileageEdit10()
    EditLoadMileageCandidate 10
End Sub

Public Sub FleetSelectMileageEdit11()
    EditLoadMileageCandidate 11
End Sub

Public Sub FleetSelectMileageEdit12()
    EditLoadMileageCandidate 12
End Sub

Public Sub FleetSelectMileageEdit13()
    EditLoadMileageCandidate 13
End Sub

Public Sub FleetSelectMileageEdit14()
    EditLoadMileageCandidate 14
End Sub

Public Sub FleetSelectMileageEdit15()
    EditLoadMileageCandidate 15
End Sub

Public Sub FleetSelectReservationEdit1()
    EditLoadReservationCandidate 1
End Sub

Public Sub FleetSelectReservationEdit2()
    EditLoadReservationCandidate 2
End Sub

Public Sub FleetSelectReservationEdit3()
    EditLoadReservationCandidate 3
End Sub

Public Sub FleetSelectReservationEdit4()
    EditLoadReservationCandidate 4
End Sub

Public Sub FleetSelectReservationEdit5()
    EditLoadReservationCandidate 5
End Sub

Public Sub FleetSelectReservationEdit6()
    EditLoadReservationCandidate 6
End Sub

Public Sub FleetSelectReservationEdit7()
    EditLoadReservationCandidate 7
End Sub

Public Sub FleetSelectReservationEdit8()
    EditLoadReservationCandidate 8
End Sub

Public Sub FleetSelectReservationEdit9()
    EditLoadReservationCandidate 9
End Sub

Public Sub FleetSelectReservationEdit10()
    EditLoadReservationCandidate 10
End Sub

Public Sub FleetSelectReservationEdit11()
    EditLoadReservationCandidate 11
End Sub

Public Sub FleetSelectReservationEdit12()
    EditLoadReservationCandidate 12
End Sub

Public Sub FleetSelectReservationEdit13()
    EditLoadReservationCandidate 13
End Sub

Public Sub FleetSelectReservationEdit14()
    EditLoadReservationCandidate 14
End Sub

Public Sub FleetSelectReservationEdit15()
    EditLoadReservationCandidate 15
End Sub

Public Sub FleetCommitMileageEdit(ByVal logDate As Date, ByVal vehicleId As String, _
    ByVal startOdometer As Long, ByVal endOdometer As Long, Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim mileageTable As ListObject
    Dim vehicleTable As ListObject
    Dim logIdText As String
    Dim rowIndex As Long
    Dim oldDate As Date
    Dim oldVehicleId As String
    Dim makeModel As String
    Dim editStage As String

    editStage = "start"
    EditDebugStatus "Mileage edit: " & editStage

    logIdText = Trim$(CStr(EditSetupSheet().Range(MILEAGE_EDIT_CELL).value))
    If logIdText = "" Then Err.Raise vbObjectError + 8201, EDIT_APP_TITLE, "No mileage entry is selected for editing."

    Set mileageTable = EditGetTable("tblMileageLog")
    Set vehicleTable = EditGetTable("tblVehicles")
    editStage = "locate row"
    EditDebugStatus "Mileage edit: " & editStage
    rowIndex = EditFindRowById(mileageTable, "Log ID", logIdText)
    If rowIndex = 0 Then Err.Raise vbObjectError + 8202, EDIT_APP_TITLE, "The selected mileage entry changed. Reopen Edit Mileage."

    oldDate = EditDateOnly(EditTableValue(mileageTable, rowIndex, "Date"))
    oldVehicleId = Trim$(CStr(EditTableValue(mileageTable, rowIndex, "Vehicle ID")))

    editStage = "write row"
    EditDebugStatus "Mileage edit: " & editStage
    EditSetTableValue mileageTable, rowIndex, "Date", EditDateOnly(logDate)
    EditSetTableValue mileageTable, rowIndex, "Vehicle ID", vehicleId
    If Not IsNumeric(EditTableValue(mileageTable, rowIndex, "Trip Count")) Then _
        EditSetTableValue mileageTable, rowIndex, "Trip Count", 1
    EditSetTableValue mileageTable, rowIndex, "Odometer Start", startOdometer
    EditSetTableValue mileageTable, rowIndex, "Odometer End", endOdometer
    EditSetTableValue mileageTable, rowIndex, "Miles Driven", endOdometer - startOdometer
    EditSetTableValue mileageTable, rowIndex, "Maintenance Flag", "No"

    editStage = "reset odometer"
    EditDebugStatus "Mileage edit: " & editStage
    EditResetVehicleOdometer vehicleTable, mileageTable, oldVehicleId
    If StrComp(oldVehicleId, vehicleId, vbTextCompare) <> 0 Then _
        EditResetVehicleOdometer vehicleTable, mileageTable, vehicleId

    editStage = "sync old month"
    EditDebugStatus "Mileage edit: " & editStage
    FleetSyncMonthlyMileageEntry oldDate, oldVehicleId
    If oldDate <> EditDateOnly(logDate) Or StrComp(oldVehicleId, vehicleId, vbTextCompare) <> 0 Then _
        FleetSyncMonthlyMileageEntry EditDateOnly(logDate), vehicleId
    editStage = "refresh mileage status"
    EditDebugStatus "Mileage edit: " & editStage
    FleetRefreshMileageLogVehicleStatus False

    makeModel = EditVehicleChoiceText(vehicleTable, vehicleId)
    editStage = "clear form"
    EditDebugStatus "Mileage edit: " & editStage
    FleetCancelMileageEdit False
    FleetClearMileageForm False
    editStage = "refresh dashboard"
    EditDebugStatus "Mileage edit: " & editStage
    FleetRefreshDashboard False
    EditDashboardMessage "Mileage updated: " & makeModel & ", " & Format(logDate, "m/d/yyyy") & ", " & _
        Format(startOdometer, "#,##0") & " to " & Format(endOdometer, "#,##0") & _
        " (" & Format(endOdometer - startOdometer, "#,##0") & " miles)."
    If showMessage Then MsgBox "Mileage entry updated.", vbInformation, EDIT_APP_TITLE
    Exit Sub

CleanFail:
    EditDebugStatus "Mileage edit failed at " & editStage & ": " & Err.Number & " - " & Err.Description
    If showMessage Then
        MsgBox "Could not update mileage: " & Err.Description, vbCritical, EDIT_APP_TITLE
    Else
        Err.Raise Err.Number, EDIT_APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetCommitReservationEdit(ByVal reserveStartDate As Date, ByVal reserveEndDate As Date, _
    ByVal vehicleId As String, ByVal driverName As String, ByVal startTime As Date, ByVal endTime As Date, _
    ByVal purpose As String, Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim reservationTable As ListObject
    Dim vehicleTable As ListObject
    Dim selectedIds As String
    Dim idItems As Variant
    Dim originalCount As Long
    Dim newCount As Long
    Dim sharedCount As Long
    Dim itemIndex As Long
    Dim rowIndex As Long
    Dim currentDate As Date
    Dim conflictText As String
    Dim notesText As String
    Dim lr As ListRow
    Dim makeModel As String

    selectedIds = Trim$(CStr(EditSetupSheet().Range(RESERVATION_EDIT_CELL).value))
    If selectedIds = "" Then Err.Raise vbObjectError + 8210, EDIT_APP_TITLE, "No reservation is selected for editing."

    Set reservationTable = EditGetTable("tblReservations")
    Set vehicleTable = EditGetTable("tblVehicles")
    idItems = Split(selectedIds, ",")
    originalCount = UBound(idItems) - LBound(idItems) + 1

    For itemIndex = LBound(idItems) To UBound(idItems)
        If EditFindRowById(reservationTable, "Reservation ID", Trim$(CStr(idItems(itemIndex)))) = 0 Then
            Err.Raise vbObjectError + 8211, EDIT_APP_TITLE, "The selected reservation changed. Reopen Edit Reservation."
        End If
    Next itemIndex

    currentDate = EditDateOnly(reserveStartDate)
    Do While currentDate <= EditDateOnly(reserveEndDate)
        conflictText = EditReservationConflictText(reservationTable, vehicleId, currentDate, selectedIds, startTime, endTime)
        If conflictText <> "" Then
            EditDashboardMessage "Vehicle is reserved: " & conflictText
            If showMessage Or Application.Visible Then MsgBox "Vehicle is reserved." & vbCrLf & vbCrLf & conflictText, vbExclamation, EDIT_APP_TITLE
            Exit Sub
        End If
        currentDate = DateAdd("d", 1, currentDate)
    Loop

    rowIndex = EditFindRowById(reservationTable, "Reservation ID", Trim$(CStr(idItems(LBound(idItems)))))
    notesText = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Notes")))
    newCount = DateDiff("d", EditDateOnly(reserveStartDate), EditDateOnly(reserveEndDate)) + 1
    sharedCount = originalCount
    If newCount < sharedCount Then sharedCount = newCount

    For itemIndex = 0 To sharedCount - 1
        rowIndex = EditFindRowById(reservationTable, "Reservation ID", Trim$(CStr(idItems(LBound(idItems) + itemIndex))))
        EditWriteReservationRow reservationTable, rowIndex, DateAdd("d", itemIndex, EditDateOnly(reserveStartDate)), _
            vehicleId, driverName, startTime, endTime, purpose, notesText
    Next itemIndex

    If newCount > originalCount Then
        For itemIndex = originalCount To newCount - 1
            Set lr = reservationTable.ListRows.Add
            EditSetListRowValue reservationTable, lr, "Reservation ID", EditNextNumber(reservationTable, "Reservation ID")
            EditWriteReservationListRow reservationTable, lr, DateAdd("d", itemIndex, EditDateOnly(reserveStartDate)), _
                vehicleId, driverName, startTime, endTime, purpose, notesText
        Next itemIndex
    ElseIf originalCount > newCount Then
        For itemIndex = originalCount - 1 To newCount Step -1
            rowIndex = EditFindRowById(reservationTable, "Reservation ID", Trim$(CStr(idItems(LBound(idItems) + itemIndex))))
            If rowIndex > 0 Then reservationTable.ListRows(rowIndex).Delete
        Next itemIndex
    End If

    MarkWeeklyReservationsDirty
    makeModel = EditVehicleChoiceText(vehicleTable, vehicleId)
    FleetCancelReservationEdit False
    FleetClearCheckoutForm False
    FleetRefreshDashboard False
    EditDashboardMessage "Reservation updated: " & driverName & ", " & makeModel & ", " & _
        EditDateRangeText(reserveStartDate, reserveEndDate) & " " & Format(startTime, "h:mm AM/PM") & _
        "-" & Format(endTime, "h:mm AM/PM") & "."
    If showMessage Then MsgBox "Reservation updated.", vbInformation, EDIT_APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Then
        MsgBox "Could not update reservation: " & Err.Description, vbCritical, EDIT_APP_TITLE
    Else
        Err.Raise Err.Number, EDIT_APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetDeleteSelectedMileage(Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim mileageTable As ListObject
    Dim vehicleTable As ListObject
    Dim logIdText As String
    Dim rowIndex As Long
    Dim logDate As Date
    Dim vehicleId As String
    Dim descriptionText As String
    Dim confirmation As VbMsgBoxResult

    logIdText = Trim$(CStr(EditSetupSheet().Range(MILEAGE_EDIT_CELL).value))
    If logIdText = "" Then Err.Raise vbObjectError + 8230, EDIT_APP_TITLE, "Select a mileage entry through Edit Mileage first."

    Set mileageTable = EditGetTable("tblMileageLog")
    Set vehicleTable = EditGetTable("tblVehicles")
    rowIndex = EditFindRowById(mileageTable, "Log ID", logIdText)
    If rowIndex = 0 Then Err.Raise vbObjectError + 8231, EDIT_APP_TITLE, "The selected mileage entry is no longer available."

    logDate = EditDateOnly(EditTableValue(mileageTable, rowIndex, "Date"))
    vehicleId = Trim$(CStr(EditTableValue(mileageTable, rowIndex, "Vehicle ID")))
    descriptionText = EditMileageDescription(mileageTable, rowIndex)
    If showMessage Then
        confirmation = MsgBox("Delete this mileage entry?" & vbCrLf & vbCrLf & descriptionText & vbCrLf & vbCrLf & _
            "This action cannot be undone.", vbYesNo + vbExclamation + vbDefaultButton2, EDIT_APP_TITLE)
        If confirmation <> vbYes Then Exit Sub
    End If

    mileageTable.ListRows(rowIndex).Delete
    EditResetVehicleOdometer vehicleTable, mileageTable, vehicleId
    FleetSyncMonthlyMileageEntry logDate, vehicleId
    FleetRefreshMileageLogVehicleStatus False
    FleetCancelMileageEdit False
    FleetClearMileageForm False
    FleetRefreshDashboard False
    EditDashboardMessage "Mileage entry deleted: " & descriptionText & "."
    If showMessage Then MsgBox "Mileage entry deleted.", vbInformation, EDIT_APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Then
        MsgBox "Could not delete mileage: " & Err.Description, vbCritical, EDIT_APP_TITLE
    Else
        Err.Raise Err.Number, EDIT_APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetDeleteSelectedReservation(Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim reservationTable As ListObject
    Dim selectedIds As String
    Dim descriptionText As String
    Dim rowIndex As Long
    Dim deletedCount As Long
    Dim confirmation As VbMsgBoxResult

    selectedIds = Trim$(CStr(EditSetupSheet().Range(RESERVATION_EDIT_CELL).value))
    If selectedIds = "" Then Err.Raise vbObjectError + 8240, EDIT_APP_TITLE, "Select a reservation through Edit Reservation first."

    Set reservationTable = EditGetTable("tblReservations")
    descriptionText = EditSelectedReservationDescription(reservationTable, selectedIds)
    If descriptionText = "" Then Err.Raise vbObjectError + 8241, EDIT_APP_TITLE, "The selected reservation is no longer available."
    If showMessage Then
        confirmation = MsgBox("Delete this reservation?" & vbCrLf & vbCrLf & descriptionText & vbCrLf & vbCrLf & _
            "All dates in this reservation will be removed. This action cannot be undone.", _
            vbYesNo + vbExclamation + vbDefaultButton2, EDIT_APP_TITLE)
        If confirmation <> vbYes Then Exit Sub
    End If

    For rowIndex = reservationTable.ListRows.Count To 1 Step -1
        If EditCsvContainsId(selectedIds, EditTableValue(reservationTable, rowIndex, "Reservation ID")) Then
            reservationTable.ListRows(rowIndex).Delete
            deletedCount = deletedCount + 1
        End If
    Next rowIndex
    If deletedCount = 0 Then Err.Raise vbObjectError + 8242, EDIT_APP_TITLE, "The selected reservation is no longer available."

    MarkWeeklyReservationsDirty
    FleetCancelReservationEdit False
    FleetClearCheckoutForm False
    FleetRefreshDashboard False
    EditDashboardMessage "Reservation deleted: " & descriptionText & "."
    If showMessage Then MsgBox "Reservation deleted.", vbInformation, EDIT_APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Then
        MsgBox "Could not delete reservation: " & Err.Description, vbCritical, EDIT_APP_TITLE
    Else
        Err.Raise Err.Number, EDIT_APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetCancelMileageEdit(Optional ByVal showMessage As Boolean = False)
    On Error Resume Next
    EditSetupSheet().Range(MILEAGE_EDIT_CELL).ClearContents
    EditSetButtonCaption "FleetSaveDashboardMileage", "SAVE MILEAGE"
    EditSetUtilityButton MILEAGE_EDIT_BUTTON, "EDIT MILEAGE", "FleetEditMileage", False
    If showMessage Then EditDashboardMessage "Mileage edit cancelled."
    On Error GoTo 0
End Sub

Public Sub FleetCancelReservationEdit(Optional ByVal showMessage As Boolean = False)
    On Error Resume Next
    EditSetupSheet().Range(RESERVATION_EDIT_CELL).ClearContents
    EditSetButtonCaption "FleetSaveDashboardCheckout", "SAVE RESERVATION"
    EditSetUtilityButton RESERVATION_EDIT_BUTTON, "EDIT RESERVATION", "FleetEditReservation", False
    If showMessage Then EditDashboardMessage "Reservation edit cancelled."
    On Error GoTo 0
End Sub

Public Sub FleetCancelAllEdits(Optional ByVal showMessage As Boolean = False)
    FleetCancelMileageEdit False
    FleetCancelReservationEdit False
    EditSetupSheet().Range("B20:C34").ClearContents
    If showMessage Then EditDashboardMessage "Edit mode cancelled."
End Sub

Private Sub EditLoadMileageCandidate(ByVal slot As Long)
    On Error GoTo CleanFail

    Dim mileageTable As ListObject
    Dim vehicleTable As ListObject
    Dim dashWs As Worksheet
    Dim logId As String
    Dim rowIndex As Long
    Dim vehicleId As String

    logId = Trim$(CStr(EditSetupSheet().Cells(MILEAGE_CANDIDATE_FIRST_ROW + slot - 1, 2).value))
    If logId = "" Then Exit Sub

    Set mileageTable = EditGetTable("tblMileageLog")
    Set vehicleTable = EditGetTable("tblVehicles")
    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    rowIndex = EditFindRowById(mileageTable, "Log ID", logId)
    If rowIndex = 0 Then Err.Raise vbObjectError + 8220, EDIT_APP_TITLE, "The selected mileage entry is no longer available."

    vehicleId = Trim$(CStr(EditTableValue(mileageTable, rowIndex, "Vehicle ID")))
    FleetCancelReservationEdit False
    EditSetupSheet().Range(MILEAGE_EDIT_CELL).value = logId
    EditEntrySet dashWs.Range("B6"), EditDateOnly(EditTableValue(mileageTable, rowIndex, "Date"))
    dashWs.Range("B6").numberFormat = "m/d/yyyy"
    EditEntrySet dashWs.Range("B7"), EditVehicleChoiceText(vehicleTable, vehicleId)
    EditEntrySet dashWs.Range("B8"), EditTableValue(mileageTable, rowIndex, "Odometer Start")
    EditEntrySet dashWs.Range("B9"), EditTableValue(mileageTable, rowIndex, "Odometer End")
    dashWs.Range("B8:B9").numberFormat = "#,##0"
    EditSetButtonCaption "FleetSaveDashboardMileage", "SAVE CHANGES"
    EditSetUtilityButton MILEAGE_EDIT_BUTTON, "DELETE MILEAGE", "FleetDeleteSelectedMileage", True
    EditDashboardMessage "Editing mileage entry " & logId & ". Select SAVE CHANGES to update it or DELETE MILEAGE to remove it. CLEAR cancels."
    If Application.Visible Then dashWs.Activate
    Exit Sub

CleanFail:
    MsgBox "Could not load mileage entry: " & Err.Description, vbExclamation, EDIT_APP_TITLE
End Sub

Private Sub EditLoadReservationCandidate(ByVal slot As Long)
    On Error GoTo CleanFail

    Dim reservationTable As ListObject
    Dim vehicleTable As ListObject
    Dim dashWs As Worksheet
    Dim idsCsv As String
    Dim rowIndex As Long
    Dim firstRow As Long
    Dim firstDate As Date
    Dim lastDate As Date
    Dim hasDate As Boolean
    Dim vehicleId As String
    Dim driverName As String
    Dim purpose As String
    Dim startValue As Variant
    Dim endValue As Variant

    idsCsv = Trim$(CStr(EditSetupSheet().Cells(MILEAGE_CANDIDATE_FIRST_ROW + slot - 1, RESERVATION_CANDIDATE_COLUMN).value))
    If idsCsv = "" Then Exit Sub

    Set reservationTable = EditGetTable("tblReservations")
    Set vehicleTable = EditGetTable("tblVehicles")
    Set dashWs = ThisWorkbook.Worksheets("Dashboard")

    For rowIndex = 1 To reservationTable.ListRows.Count
        If EditCsvContainsId(idsCsv, EditTableValue(reservationTable, rowIndex, "Reservation ID")) Then
            If firstRow = 0 Then
                firstRow = rowIndex
                vehicleId = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Vehicle ID")))
                driverName = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Driver Name")))
                purpose = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Purpose")))
                startValue = EditTableValue(reservationTable, rowIndex, "Start Time")
                endValue = EditTableValue(reservationTable, rowIndex, "End Time")
            End If
            EditExpandDateRange EditTableValue(reservationTable, rowIndex, "Date"), firstDate, lastDate, hasDate
        End If
    Next rowIndex

    If firstRow = 0 Or Not hasDate Then Err.Raise vbObjectError + 8221, EDIT_APP_TITLE, "The selected reservation is no longer available."

    FleetCancelMileageEdit False
    With EditSetupSheet().Range(RESERVATION_EDIT_CELL)
        .numberFormat = "@"
        .Value2 = CStr(idsCsv)
    End With
    If firstDate = lastDate Then
        EditEntrySet dashWs.Range("F6"), firstDate
    Else
        EditEntrySet dashWs.Range("F6"), EditDateRangeText(firstDate, lastDate)
    End If
    EditEntrySet dashWs.Range("F7"), EditVehicleChoiceText(vehicleTable, vehicleId)
    EditEntrySet dashWs.Range("F8"), driverName
    If EditIsTimeValue(startValue) Then EditEntrySet dashWs.Range("F9"), EditTimeAsDate(startValue) Else EditEntryClear dashWs.Range("F9")
    If EditIsTimeValue(endValue) Then EditEntrySet dashWs.Range("H9"), EditTimeAsDate(endValue) Else EditEntryClear dashWs.Range("H9")
    dashWs.Range("F9").numberFormat = "h:mm AM/PM"
    dashWs.Range("H9").numberFormat = "h:mm AM/PM"
    EditEntrySet dashWs.Range("F10"), purpose
    EditSetButtonCaption "FleetSaveDashboardCheckout", "SAVE CHANGES"
    EditSetUtilityButton RESERVATION_EDIT_BUTTON, "DELETE RESERVATION", "FleetDeleteSelectedReservation", True
    EditDashboardMessage "Editing reservation. Select SAVE CHANGES to update it or DELETE RESERVATION to remove it. CLEAR cancels."
    If Application.Visible Then dashWs.Activate
    Exit Sub

CleanFail:
    If Application.Visible Then
        MsgBox "Could not load reservation: " & Err.Description, vbExclamation, EDIT_APP_TITLE
    Else
        Err.Raise Err.Number, EDIT_APP_TITLE, Err.Description
    End If
End Sub

Private Function EditCreatePopup(ByVal popupName As String) As Object
    On Error Resume Next
    Application.CommandBars(popupName).Delete
    On Error GoTo 0
    Set EditCreatePopup = Application.CommandBars.Add(popupName, 5, False, True)
End Function

Private Sub EditAddPopupItem(ByVal popup As Object, ByVal caption As String, ByVal macroName As String)
    Dim item As Object
    Set item = popup.Controls.Add(1)
    item.caption = Replace(caption, "&", "&&")
    item.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Private Function EditMileageDescription(ByVal mileageTable As ListObject, ByVal rowIndex As Long) As String
    Dim vehicleTable As ListObject
    Dim vehicleId As String
    Dim logDate As Variant
    Dim startOdo As Variant
    Dim endOdo As Variant
    Dim miles As Variant

    Set vehicleTable = EditGetTable("tblVehicles")
    vehicleId = Trim$(CStr(EditTableValue(mileageTable, rowIndex, "Vehicle ID")))
    logDate = EditTableValue(mileageTable, rowIndex, "Date")
    startOdo = EditTableValue(mileageTable, rowIndex, "Odometer Start")
    endOdo = EditTableValue(mileageTable, rowIndex, "Odometer End")
    miles = EditTableValue(mileageTable, rowIndex, "Miles Driven")

    If IsDate(logDate) Then EditMileageDescription = Format(EditDateOnly(logDate), "m/d") & " | "
    EditMileageDescription = EditMileageDescription & EditVehicleChoiceText(vehicleTable, vehicleId)
    If IsNumeric(startOdo) And IsNumeric(endOdo) Then _
        EditMileageDescription = EditMileageDescription & " | " & Format(CLng(startOdo), "#,##0") & " to " & Format(CLng(endOdo), "#,##0")
    If IsNumeric(miles) Then EditMileageDescription = EditMileageDescription & " (" & Format(CDbl(miles), "#,##0") & " mi)"
End Function

Private Sub EditReservationGroup(ByVal reservationTable As ListObject, ByVal latestRow As Long, _
    ByRef firstRow As Long, ByRef lastRow As Long)
    Dim latestKey As String
    Dim rowIndex As Long
    Dim expectedId As Long
    Dim idValue As Variant
    Dim rowId As Variant
    Dim hasNumericId As Boolean

    firstRow = latestRow
    lastRow = latestRow
    latestKey = EditReservationKey(reservationTable, latestRow)
    idValue = EditTableValue(reservationTable, latestRow, "Reservation ID")
    hasNumericId = IsNumeric(idValue)
    If hasNumericId Then expectedId = CLng(idValue) - 1

    For rowIndex = latestRow - 1 To 1 Step -1
        If EditReservationKey(reservationTable, rowIndex) <> latestKey Then Exit For
        If hasNumericId Then
            rowId = EditTableValue(reservationTable, rowIndex, "Reservation ID")
            If Not IsNumeric(rowId) Then Exit For
            If CLng(rowId) <> expectedId Then Exit For
            expectedId = expectedId - 1
        End If
        firstRow = rowIndex
        If Not hasNumericId Then Exit For
    Next rowIndex
End Sub

Private Function EditReservationDescription(ByVal reservationTable As ListObject, ByVal firstRow As Long, ByVal lastRow As Long) As String
    Dim vehicleTable As ListObject
    Dim vehicleId As String
    Dim driverName As String
    Dim firstDate As Date
    Dim lastDate As Date
    Dim hasDate As Boolean
    Dim rowIndex As Long
    Dim startText As String
    Dim endText As String

    Set vehicleTable = EditGetTable("tblVehicles")
    vehicleId = Trim$(CStr(EditTableValue(reservationTable, lastRow, "Vehicle ID")))
    driverName = Trim$(CStr(EditTableValue(reservationTable, lastRow, "Driver Name")))
    startText = EditDisplayTime(EditTableValue(reservationTable, lastRow, "Start Time"))
    endText = EditDisplayTime(EditTableValue(reservationTable, lastRow, "End Time"))

    For rowIndex = firstRow To lastRow
        EditExpandDateRange EditTableValue(reservationTable, rowIndex, "Date"), firstDate, lastDate, hasDate
    Next rowIndex

    If hasDate Then EditReservationDescription = EditShortDateRangeText(firstDate, lastDate) & " | "
    EditReservationDescription = EditReservationDescription & EditVehicleChoiceText(vehicleTable, vehicleId)
    If driverName <> "" Then EditReservationDescription = EditReservationDescription & " | " & driverName
    If startText <> "" Or endText <> "" Then EditReservationDescription = EditReservationDescription & " | " & startText & "-" & endText
End Function

Private Function EditSelectedReservationDescription(ByVal reservationTable As ListObject, ByVal selectedIds As String) As String
    Dim vehicleTable As ListObject
    Dim rowIndex As Long
    Dim firstDate As Date
    Dim lastDate As Date
    Dim hasDate As Boolean
    Dim vehicleId As String
    Dim driverName As String
    Dim purposeText As String

    Set vehicleTable = EditGetTable("tblVehicles")
    For rowIndex = 1 To reservationTable.ListRows.Count
        If EditCsvContainsId(selectedIds, EditTableValue(reservationTable, rowIndex, "Reservation ID")) Then
            If vehicleId = "" Then
                vehicleId = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Vehicle ID")))
                driverName = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Driver Name")))
                purposeText = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Purpose")))
            End If
            EditExpandDateRange EditTableValue(reservationTable, rowIndex, "Date"), firstDate, lastDate, hasDate
        End If
    Next rowIndex

    If Not hasDate Then Exit Function
    EditSelectedReservationDescription = EditDateRangeText(firstDate, lastDate) & " | " & EditVehicleChoiceText(vehicleTable, vehicleId)
    If driverName <> "" Then EditSelectedReservationDescription = EditSelectedReservationDescription & " | " & driverName
    If purposeText <> "" Then EditSelectedReservationDescription = EditSelectedReservationDescription & " | " & purposeText
End Function

Private Function EditReservationIdsForRows(ByVal reservationTable As ListObject, ByVal firstRow As Long, ByVal lastRow As Long) As String
    Dim rowIndex As Long
    Dim idValue As Variant
    Dim result As String

    For rowIndex = firstRow To lastRow
        idValue = EditTableValue(reservationTable, rowIndex, "Reservation ID")
        If IsNumeric(idValue) Then
            If result <> "" Then result = result & ","
            result = result & CStr(CLng(idValue))
        End If
    Next rowIndex
    EditReservationIdsForRows = result
End Function

Private Function EditReservationKey(ByVal reservationTable As ListObject, ByVal rowIndex As Long) As String
    Dim statusText As String

    statusText = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Status")))
    If EditIsShopStatus(statusText) Or StrComp(statusText, "Cancelled", vbTextCompare) = 0 Then Exit Function

    EditReservationKey = UCase$(Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Vehicle ID")))) & "|" & _
        UCase$(Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Driver Name")))) & "|" & _
        UCase$(EditDisplayTime(EditTableValue(reservationTable, rowIndex, "Start Time"))) & "|" & _
        UCase$(EditDisplayTime(EditTableValue(reservationTable, rowIndex, "End Time"))) & "|" & _
        UCase$(Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Purpose")))) & "|" & _
        UCase$(statusText) & "|" & UCase$(Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Notes"))))
End Function

Private Function EditReservationRowCanBeEdited(ByVal reservationTable As ListObject, ByVal rowIndex As Long) As Boolean
    Dim statusText As String
    statusText = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Status")))
    EditReservationRowCanBeEdited = (Not EditIsShopStatus(statusText) And StrComp(statusText, "Cancelled", vbTextCompare) <> 0)
End Function

Private Function EditReservationConflictText(ByVal reservationTable As ListObject, ByVal vehicleId As String, _
    ByVal reserveDate As Date, ByVal excludedIds As String, ByVal requestedStart As Date, _
    ByVal requestedEnd As Date) As String
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim statusText As String
    Dim driverText As String
    Dim purposeText As String
    Dim startText As String
    Dim endText As String
    Dim existingStartValue As Variant
    Dim existingEndValue As Variant
    Dim existingStart As Date
    Dim existingEnd As Date
    Dim hasConflict As Boolean

    If reservationTable.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To reservationTable.ListRows.Count
        If Not EditCsvContainsId(excludedIds, EditTableValue(reservationTable, rowIndex, "Reservation ID")) Then
            If StrComp(Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
                rowDate = EditTableValue(reservationTable, rowIndex, "Date")
                If IsDate(rowDate) Then
                    If EditDateOnly(rowDate) = EditDateOnly(reserveDate) Then
                        statusText = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Status")))
                        If StrComp(statusText, "Cancelled", vbTextCompare) <> 0 Then
                            hasConflict = False
                            If EditIsShopStatus(statusText) Then
                                hasConflict = True
                            Else
                                existingStartValue = EditTableValue(reservationTable, rowIndex, "Start Time")
                                existingEndValue = EditTableValue(reservationTable, rowIndex, "End Time")
                                If EditIsTimeValue(existingStartValue) And EditIsTimeValue(existingEndValue) Then
                                    existingStart = EditTimeAsDate(existingStartValue)
                                    existingEnd = EditTimeAsDate(existingEndValue)
                                    If CDbl(existingEnd) > CDbl(existingStart) Then
                                        hasConflict = (CDbl(requestedStart) < CDbl(existingEnd) And _
                                            CDbl(requestedEnd) > CDbl(existingStart))
                                    Else
                                        hasConflict = True
                                    End If
                                Else
                                    hasConflict = True
                                End If
                            End If

                            If hasConflict Then
                                EditReservationConflictText = Format(EditDateOnly(rowDate), "m/d/yyyy")
                                If EditIsShopStatus(statusText) Then
                                    EditReservationConflictText = EditReservationConflictText & " - " & statusText
                                Else
                                driverText = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Driver Name")))
                                purposeText = Trim$(CStr(EditTableValue(reservationTable, rowIndex, "Purpose")))
                                startText = EditDisplayTime(EditTableValue(reservationTable, rowIndex, "Start Time"))
                                endText = EditDisplayTime(EditTableValue(reservationTable, rowIndex, "End Time"))
                                If driverText <> "" Then EditReservationConflictText = EditReservationConflictText & " by " & driverText
                                If startText <> "" Or endText <> "" Then EditReservationConflictText = EditReservationConflictText & " (" & startText & " - " & endText & ")"
                                If purposeText <> "" Then EditReservationConflictText = EditReservationConflictText & " - " & purposeText
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next rowIndex
End Function

Private Sub EditWriteReservationRow(ByVal reservationTable As ListObject, ByVal rowIndex As Long, _
    ByVal reserveDate As Date, ByVal vehicleId As String, ByVal driverName As String, _
    ByVal startTime As Date, ByVal endTime As Date, ByVal purpose As String, ByVal notesText As String)
    EditSetTableValue reservationTable, rowIndex, "Date", EditDateOnly(reserveDate)
    EditSetTableValue reservationTable, rowIndex, "Vehicle ID", vehicleId
    EditSetTableValue reservationTable, rowIndex, "Driver Name", driverName
    EditSetTableValue reservationTable, rowIndex, "Start Time", timeValue(startTime)
    EditSetTableValue reservationTable, rowIndex, "End Time", timeValue(endTime)
    EditSetTableValue reservationTable, rowIndex, "Purpose", purpose
    EditSetTableValue reservationTable, rowIndex, "Status", "Reserved"
    EditSetTableValue reservationTable, rowIndex, "Notes", notesText
End Sub

Private Sub EditWriteReservationListRow(ByVal reservationTable As ListObject, ByVal lr As ListRow, _
    ByVal reserveDate As Date, ByVal vehicleId As String, ByVal driverName As String, _
    ByVal startTime As Date, ByVal endTime As Date, ByVal purpose As String, ByVal notesText As String)
    EditSetListRowValue reservationTable, lr, "Date", EditDateOnly(reserveDate)
    EditSetListRowValue reservationTable, lr, "Vehicle ID", vehicleId
    EditSetListRowValue reservationTable, lr, "Driver Name", driverName
    EditSetListRowValue reservationTable, lr, "Start Time", timeValue(startTime)
    EditSetListRowValue reservationTable, lr, "End Time", timeValue(endTime)
    EditSetListRowValue reservationTable, lr, "Purpose", purpose
    EditSetListRowValue reservationTable, lr, "Status", "Reserved"
    EditSetListRowValue reservationTable, lr, "Notes", notesText
End Sub

Private Sub EditResetVehicleOdometer(ByVal vehicleTable As ListObject, ByVal mileageTable As ListObject, ByVal vehicleId As String)
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim rowId As Variant
    Dim endOdometer As Variant
    Dim latestDate As Date
    Dim latestId As Long
    Dim latestOdometer As Long
    Dim found As Boolean
    Dim vehicleRow As Long

    If vehicleId = "" Or mileageTable.DataBodyRange Is Nothing Then Exit Sub
    For rowIndex = 1 To mileageTable.ListRows.Count
        If StrComp(Trim$(CStr(EditTableValue(mileageTable, rowIndex, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = EditTableValue(mileageTable, rowIndex, "Date")
            rowId = EditTableValue(mileageTable, rowIndex, "Log ID")
            endOdometer = EditTableValue(mileageTable, rowIndex, "Odometer End")
            If IsDate(rowDate) And IsNumeric(endOdometer) Then
                If Not found Or EditDateOnly(rowDate) > latestDate Or _
                    (EditDateOnly(rowDate) = latestDate And IsNumeric(rowId) And CLng(rowId) > latestId) Then
                    latestDate = EditDateOnly(rowDate)
                    If IsNumeric(rowId) Then latestId = CLng(rowId)
                    latestOdometer = CLng(endOdometer)
                    found = True
                End If
            End If
        End If
    Next rowIndex

    If found Then
        vehicleRow = EditFindRowById(vehicleTable, "Vehicle ID", vehicleId)
        If vehicleRow > 0 Then EditSetTableValue vehicleTable, vehicleRow, "Odometer", latestOdometer
    End If
End Sub

Private Function EditVehicleChoiceText(ByVal vehicleTable As ListObject, ByVal vehicleId As String) As String
    Dim rowIndex As Long
    Dim makeModel As String

    rowIndex = EditFindRowById(vehicleTable, "Vehicle ID", vehicleId)
    If rowIndex > 0 Then makeModel = Trim$(CStr(EditTableValue(vehicleTable, rowIndex, "Make/Model")))
    If makeModel <> "" Then
        EditVehicleChoiceText = vehicleId & " - " & makeModel
    Else
        EditVehicleChoiceText = vehicleId
    End If
End Function

Private Function EditGetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set EditGetTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not EditGetTable Is Nothing Then Exit Function
    Next ws
    Err.Raise vbObjectError + 8290, EDIT_APP_TITLE, "Missing table: " & tableName
End Function

Private Function EditSetupSheet() As Worksheet
    Set EditSetupSheet = ThisWorkbook.Worksheets(EDIT_SETUP_SHEET)
End Function

Private Function EditTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    EditTableValue = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).index).value
End Function

Private Sub EditSetTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal value As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).index).value = value
End Sub

Private Sub EditSetListRowValue(ByVal lo As ListObject, ByVal lr As ListRow, ByVal columnName As String, ByVal value As Variant)
    lr.Range.Cells(1, lo.ListColumns(columnName).index).value = value
End Sub

Private Function EditFindRowById(ByVal lo As ListObject, ByVal idColumn As String, ByVal idValue As String) As Long
    Dim rowIndex As Long
    Dim currentValue As Variant

    If lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        currentValue = EditTableValue(lo, rowIndex, idColumn)
        If IsNumeric(currentValue) And IsNumeric(idValue) Then
            If CDbl(currentValue) = CDbl(idValue) Then
                EditFindRowById = rowIndex
                Exit Function
            End If
        ElseIf StrComp(Trim$(CStr(currentValue)), Trim$(idValue), vbTextCompare) = 0 Then
            EditFindRowById = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function EditNextNumber(ByVal lo As ListObject, ByVal idColumn As String) As Long
    On Error GoTo Fallback
    If lo.DataBodyRange Is Nothing Then
        EditNextNumber = 1
    Else
        EditNextNumber = CLng(Application.WorksheetFunction.Max(lo.ListColumns(idColumn).DataBodyRange)) + 1
    End If
    Exit Function

Fallback:
    EditNextNumber = lo.ListRows.Count + 1
End Function

Private Function EditCsvContainsId(ByVal idsCsv As String, ByVal idValue As Variant) As Boolean
    Dim normalizedId As String
    If IsNumeric(idValue) Then normalizedId = CStr(CLng(idValue)) Else normalizedId = Trim$(CStr(idValue))
    EditCsvContainsId = (InStr(1, "," & idsCsv & ",", "," & normalizedId & ",", vbTextCompare) > 0)
End Function

Private Function EditIsShopStatus(ByVal statusText As String) As Boolean
    statusText = UCase$(Trim$(statusText))
    EditIsShopStatus = (statusText = "MAINTENANCE" Or statusText = "SHOP" Or statusText = "IN SHOP" Or Left$(statusText, 6) = "SHOP -")
End Function

Private Function EditIsTimeValue(ByVal value As Variant) As Boolean
    If IsDate(value) Then
        EditIsTimeValue = True
    ElseIf IsNumeric(value) Then
        EditIsTimeValue = (CDbl(value) >= 0 And CDbl(value) < 1)
    ElseIf Trim$(CStr(value)) <> "" Then
        EditIsTimeValue = IsDate(Trim$(CStr(value)))
    End If
End Function

Private Function EditTimeAsDate(ByVal value As Variant) As Date
    If IsDate(value) Then
        EditTimeAsDate = timeValue(CDate(value))
    ElseIf IsNumeric(value) Then
        EditTimeAsDate = TimeSerial(0, 0, 0) + CDbl(value)
    Else
        EditTimeAsDate = timeValue(CDate(Trim$(CStr(value))))
    End If
End Function

Private Function EditDisplayTime(ByVal value As Variant) As String
    If EditIsTimeValue(value) Then EditDisplayTime = Format(EditTimeAsDate(value), "h:mm AM/PM")
End Function

Private Function EditDateOnly(ByVal value As Variant) As Date
    EditDateOnly = DateSerial(Year(CDate(value)), Month(CDate(value)), Day(CDate(value)))
End Function

Private Sub EditExpandDateRange(ByVal value As Variant, ByRef firstDate As Date, ByRef lastDate As Date, ByRef hasDate As Boolean)
    Dim candidate As Date
    If Not IsDate(value) Then Exit Sub
    candidate = EditDateOnly(value)
    If Not hasDate Then
        firstDate = candidate
        lastDate = candidate
        hasDate = True
    ElseIf candidate < firstDate Then
        firstDate = candidate
    ElseIf candidate > lastDate Then
        lastDate = candidate
    End If
End Sub

Private Function EditDateRangeText(ByVal startDate As Date, ByVal endDate As Date) As String
    If EditDateOnly(startDate) = EditDateOnly(endDate) Then
        EditDateRangeText = Format(startDate, "m/d/yyyy")
    Else
        EditDateRangeText = Format(startDate, "m/d/yyyy") & " - " & Format(endDate, "m/d/yyyy")
    End If
End Function

Private Function EditShortDateRangeText(ByVal startDate As Date, ByVal endDate As Date) As String
    If EditDateOnly(startDate) = EditDateOnly(endDate) Then
        EditShortDateRangeText = Format(startDate, "m/d")
    Else
        EditShortDateRangeText = Format(startDate, "m/d") & "-" & Format(endDate, "m/d")
    End If
End Function

Private Function EditEntryCell(ByVal targetCell As Range) As Range
    If targetCell.MergeCells Then
        Set EditEntryCell = targetCell.MergeArea.Cells(1, 1)
    Else
        Set EditEntryCell = targetCell
    End If
End Function

Private Sub EditEntrySet(ByVal targetCell As Range, ByVal value As Variant)
    EditEntryCell(targetCell).value = value
End Sub

Private Sub EditEntryClear(ByVal targetCell As Range)
    EditEntryCell(targetCell).ClearContents
End Sub

Private Sub EditSetButtonCaption(ByVal macroName As String, ByVal caption As String)
    Dim shp As Shape
    Dim actionText As String

    On Error Resume Next
    For Each shp In ThisWorkbook.Worksheets("Dashboard").Shapes
        actionText = CStr(shp.OnAction)
        If InStr(1, actionText, macroName, vbTextCompare) > 0 Then
            shp.TextFrame.Characters.Text = caption
            Exit For
        End If
    Next shp
    On Error GoTo 0
End Sub

Private Sub EditSetUtilityButton(ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, ByVal destructive As Boolean)
    Dim shp As Shape

    On Error Resume Next
    Set shp = ThisWorkbook.Worksheets("Dashboard").Shapes(shapeName)
    If shp Is Nothing Then Exit Sub
    shp.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame2.TextRange.Text = caption
    If destructive Then
        shp.Fill.ForeColor.RGB = RGB(177, 26, 34)
        shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        shp.Fill.ForeColor.RGB = RGB(238, 240, 242)
        shp.TextFrame.Characters.Font.Color = RGB(0, 59, 92)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 59, 92)
    End If
    On Error GoTo 0
End Sub

Private Sub EditDashboardMessage(ByVal messageText As String)
    With ThisWorkbook.Worksheets("Dashboard").Range("A29")
        If .MergeCells Then
            .MergeArea.Cells(1, 1).value = messageText
        Else
            .value = messageText
        End If
    End With
End Sub

Private Sub EditDebugStatus(ByVal statusText As String)
    On Error Resume Next
    EditSetupSheet().Range("B15").value = statusText
    On Error GoTo 0
End Sub
