Attribute VB_Name = "modFleetUndoActions"
Option Explicit

Private Const UNDO_APP_TITLE As String = "Fleet Mileage Tool"

Public Sub FleetUndoPing()
    ThisWorkbook.Worksheets("Dashboard").Range("A29").value = "Undo module ready."
End Sub

Public Sub FleetUndoLastMileage(Optional ByVal confirmUndo As Boolean = True)
    On Error GoTo CleanFail

    Dim mileageTable As ListObject
    Dim vehicleTable As ListObject
    Dim rowIndex As Long
    Dim logDate As Date
    Dim vehicleId As String
    Dim startOdo As Variant
    Dim description As String

    Set mileageTable = UndoGetTable("tblMileageLog")
    Set vehicleTable = UndoGetTable("tblVehicles")

    rowIndex = UndoLatestNumericRow(mileageTable, "Log ID")
    If rowIndex = 0 Then
        MsgBox "No mileage entries are available to undo.", vbInformation, UNDO_APP_TITLE
        Exit Sub
    End If

    logDate = UndoDateOnly(UndoTableValue(mileageTable, rowIndex, "Date"))
    vehicleId = Trim$(CStr(UndoTableValue(mileageTable, rowIndex, "Vehicle ID")))
    startOdo = UndoTableValue(mileageTable, rowIndex, "Odometer Start")
    description = UndoMileageDescription(mileageTable, vehicleTable, rowIndex)

    If confirmUndo Then
        If MsgBox("Undo the most recent mileage entry?" & vbCrLf & vbCrLf & description, _
            vbQuestion + vbYesNo, UNDO_APP_TITLE) = vbNo Then Exit Sub
    End If

    mileageTable.ListRows(rowIndex).Delete
    UndoResetVehicleOdometer vehicleTable, mileageTable, vehicleId, startOdo
    UndoRunWorkbookMacroWithTwoArguments "FleetSyncMonthlyMileageEntry", logDate, vehicleId
    UndoRunWorkbookMacro "FleetRefreshDashboard", False
    UndoDashboardMessage "Undid mileage entry: " & description
    If confirmUndo Then MsgBox "Mileage entry undone.", vbInformation, UNDO_APP_TITLE
    Exit Sub

CleanFail:
    If confirmUndo Then
        MsgBox "Could not undo mileage entry: " & Err.description, vbCritical, UNDO_APP_TITLE
    Else
        Err.Raise Err.Number, UNDO_APP_TITLE, Err.description
    End If
End Sub

Public Sub FleetUndoLastCheckout(Optional ByVal confirmUndo As Boolean = True)
    On Error GoTo CleanFail

    Dim reservationTable As ListObject
    Dim vehicleTable As ListObject
    Dim latestRow As Long
    Dim firstDeleteRow As Long
    Dim lastDeleteRow As Long
    Dim deleteCount As Long
    Dim rowIndex As Long
    Dim description As String

    Set reservationTable = UndoGetTable("tblReservations")
    Set vehicleTable = UndoGetTable("tblVehicles")

    latestRow = UndoLatestReservationRow(reservationTable)
    If latestRow = 0 Then
        MsgBox "No checkout reservations are available to undo.", vbInformation, UNDO_APP_TITLE
        Exit Sub
    End If

    UndoReservationGroup reservationTable, latestRow, firstDeleteRow, lastDeleteRow, deleteCount
    If deleteCount = 0 Then
        MsgBox "No checkout reservations are available to undo.", vbInformation, UNDO_APP_TITLE
        Exit Sub
    End If

    description = UndoReservationDescription(reservationTable, vehicleTable, firstDeleteRow, lastDeleteRow)
    If confirmUndo Then
        If MsgBox("Undo the most recent checkout reservation?" & vbCrLf & vbCrLf & description, _
            vbQuestion + vbYesNo, UNDO_APP_TITLE) = vbNo Then Exit Sub
    End If

    For rowIndex = lastDeleteRow To firstDeleteRow Step -1
        reservationTable.ListRows(rowIndex).Delete
    Next rowIndex

    UndoInvalidateWeeklyCache
    UndoRunWorkbookMacro "FleetRefreshDashboard", False
    UndoDashboardMessage "Undid checkout reservation: " & description
    If confirmUndo Then MsgBox "Checkout reservation undone.", vbInformation, UNDO_APP_TITLE
    Exit Sub

CleanFail:
    If confirmUndo Then
        MsgBox "Could not undo checkout reservation: " & Err.description, vbCritical, UNDO_APP_TITLE
    Else
        Err.Raise Err.Number, UNDO_APP_TITLE, Err.description
    End If
End Sub

Private Function UndoGetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set UndoGetTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not UndoGetTable Is Nothing Then Exit Function
    Next ws

    Err.Raise vbObjectError + 8100, UNDO_APP_TITLE, "Missing table: " & tableName
End Function

Private Function UndoTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    UndoTableValue = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).index).value
End Function

Private Function UndoLatestNumericRow(ByVal lo As ListObject, ByVal idColumn As String) As Long
    Dim rowIndex As Long
    Dim idValue As Variant
    Dim bestId As Long
    Dim hasNumericId As Boolean

    If lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        idValue = UndoTableValue(lo, rowIndex, idColumn)
        If IsNumeric(idValue) Then
            If Not hasNumericId Or CLng(idValue) > bestId Then
                bestId = CLng(idValue)
                UndoLatestNumericRow = rowIndex
                hasNumericId = True
            End If
        End If
    Next rowIndex

    If UndoLatestNumericRow = 0 Then UndoLatestNumericRow = lo.ListRows.Count
End Function

Private Function UndoLatestReservationRow(ByVal reservationTable As ListObject) As Long
    Dim rowIndex As Long
    Dim idValue As Variant
    Dim bestId As Long
    Dim hasNumericId As Boolean
    Dim fallbackRow As Long
    Dim statusText As String

    If reservationTable.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To reservationTable.ListRows.Count
        statusText = Trim$(CStr(UndoTableValue(reservationTable, rowIndex, "Status")))
        If Not UndoIsShopStatus(statusText) Then
            fallbackRow = rowIndex
            idValue = UndoTableValue(reservationTable, rowIndex, "Reservation ID")
            If IsNumeric(idValue) Then
                If Not hasNumericId Or CLng(idValue) > bestId Then
                    bestId = CLng(idValue)
                    UndoLatestReservationRow = rowIndex
                    hasNumericId = True
                End If
            End If
        End If
    Next rowIndex

    If UndoLatestReservationRow = 0 Then UndoLatestReservationRow = fallbackRow
End Function

Private Sub UndoReservationGroup(ByVal reservationTable As ListObject, ByVal latestRow As Long, _
    ByRef firstDeleteRow As Long, ByRef lastDeleteRow As Long, ByRef deleteCount As Long)
    Dim latestKey As String
    Dim rowIndex As Long
    Dim idValue As Variant
    Dim expectedId As Long
    Dim hasNumericId As Boolean

    latestKey = UndoReservationKey(reservationTable, latestRow)
    idValue = UndoTableValue(reservationTable, latestRow, "Reservation ID")
    hasNumericId = IsNumeric(idValue)
    If hasNumericId Then expectedId = CLng(idValue)

    firstDeleteRow = latestRow
    lastDeleteRow = latestRow
    For rowIndex = latestRow To 1 Step -1
        If UndoReservationKey(reservationTable, rowIndex) <> latestKey Then
            If deleteCount > 0 Then Exit For
        ElseIf hasNumericId Then
            idValue = UndoTableValue(reservationTable, rowIndex, "Reservation ID")
            If Not IsNumeric(idValue) Then Exit For
            If CLng(idValue) <> expectedId Then Exit For
            firstDeleteRow = rowIndex
            deleteCount = deleteCount + 1
            expectedId = expectedId - 1
        Else
            firstDeleteRow = rowIndex
            deleteCount = 1
            Exit For
        End If
    Next rowIndex
End Sub

Private Function UndoMileageDescription(ByVal mileageTable As ListObject, ByVal vehicleTable As ListObject, ByVal rowIndex As Long) As String
    Dim vehicleId As String
    Dim logDate As Variant
    Dim startOdo As Variant
    Dim endOdo As Variant
    Dim miles As Variant

    vehicleId = Trim$(CStr(UndoTableValue(mileageTable, rowIndex, "Vehicle ID")))
    logDate = UndoTableValue(mileageTable, rowIndex, "Date")
    startOdo = UndoTableValue(mileageTable, rowIndex, "Odometer Start")
    endOdo = UndoTableValue(mileageTable, rowIndex, "Odometer End")
    miles = UndoTableValue(mileageTable, rowIndex, "Miles Driven")

    UndoMileageDescription = UndoVehicleChoiceText(vehicleTable, vehicleId)
    If IsDate(logDate) Then UndoMileageDescription = UndoMileageDescription & ", " & Format(UndoDateOnly(logDate), "m/d/yyyy")
    If IsNumeric(startOdo) And IsNumeric(endOdo) Then
        UndoMileageDescription = UndoMileageDescription & ", " & Format(CLng(startOdo), "#,##0") & " to " & Format(CLng(endOdo), "#,##0")
    End If
    If IsNumeric(miles) Then UndoMileageDescription = UndoMileageDescription & " (" & Format(CDbl(miles), "#,##0") & " miles)"
End Function

Private Function UndoReservationDescription(ByVal reservationTable As ListObject, ByVal vehicleTable As ListObject, _
    ByVal firstDeleteRow As Long, ByVal lastDeleteRow As Long) As String
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim firstDate As Date
    Dim lastDate As Date
    Dim hasDate As Boolean
    Dim vehicleId As String
    Dim driverName As String
    Dim purposeText As String
    Dim startText As String
    Dim endText As String
    Dim dateCount As Long
    Dim pluralText As String

    vehicleId = Trim$(CStr(UndoTableValue(reservationTable, lastDeleteRow, "Vehicle ID")))
    driverName = Trim$(CStr(UndoTableValue(reservationTable, lastDeleteRow, "Driver Name")))
    purposeText = Trim$(CStr(UndoTableValue(reservationTable, lastDeleteRow, "Purpose")))
    startText = UndoDisplayTime(UndoTableValue(reservationTable, lastDeleteRow, "Start Time"))
    endText = UndoDisplayTime(UndoTableValue(reservationTable, lastDeleteRow, "End Time"))

    For rowIndex = firstDeleteRow To lastDeleteRow
        rowDate = UndoTableValue(reservationTable, rowIndex, "Date")
        If IsDate(rowDate) Then
            If Not hasDate Then
                firstDate = UndoDateOnly(rowDate)
                lastDate = UndoDateOnly(rowDate)
                hasDate = True
            ElseIf UndoDateOnly(rowDate) < firstDate Then
                firstDate = UndoDateOnly(rowDate)
            ElseIf UndoDateOnly(rowDate) > lastDate Then
                lastDate = UndoDateOnly(rowDate)
            End If
        End If
        dateCount = dateCount + 1
    Next rowIndex

    If dateCount = 1 Then pluralText = " date" Else pluralText = " dates"
    UndoReservationDescription = UndoVehicleChoiceText(vehicleTable, vehicleId)
    If hasDate Then UndoReservationDescription = UndoReservationDescription & ", " & UndoDateRangeText(firstDate, lastDate)
    If driverName <> "" Then UndoReservationDescription = UndoReservationDescription & ", " & driverName
    If startText <> "" Or endText <> "" Then UndoReservationDescription = UndoReservationDescription & ", " & startText & "-" & endText
    If purposeText <> "" Then UndoReservationDescription = UndoReservationDescription & ", " & purposeText
    UndoReservationDescription = UndoReservationDescription & " (" & CStr(dateCount) & pluralText & ")"
End Function

Private Function UndoReservationKey(ByVal reservationTable As ListObject, ByVal rowIndex As Long) As String
    Dim statusText As String

    statusText = Trim$(CStr(UndoTableValue(reservationTable, rowIndex, "Status")))
    If UndoIsShopStatus(statusText) Then Exit Function

    UndoReservationKey = UCase$(Trim$(CStr(UndoTableValue(reservationTable, rowIndex, "Vehicle ID")))) & "|" & _
        UCase$(Trim$(CStr(UndoTableValue(reservationTable, rowIndex, "Driver Name")))) & "|" & _
        UCase$(UndoDisplayTime(UndoTableValue(reservationTable, rowIndex, "Start Time"))) & "|" & _
        UCase$(UndoDisplayTime(UndoTableValue(reservationTable, rowIndex, "End Time"))) & "|" & _
        UCase$(Trim$(CStr(UndoTableValue(reservationTable, rowIndex, "Purpose")))) & "|" & _
        UCase$(statusText)
End Function

Private Function UndoVehicleChoiceText(ByVal vehicleTable As ListObject, ByVal vehicleId As String) As String
    Dim rowIndex As Long
    Dim makeModel As String

    If Not vehicleTable.DataBodyRange Is Nothing Then
        For rowIndex = 1 To vehicleTable.ListRows.Count
            If StrComp(Trim$(CStr(UndoTableValue(vehicleTable, rowIndex, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
                makeModel = Trim$(CStr(UndoTableValue(vehicleTable, rowIndex, "Make/Model")))
                Exit For
            End If
        Next rowIndex
    End If

    If makeModel <> "" Then
        UndoVehicleChoiceText = vehicleId & " - " & makeModel
    Else
        UndoVehicleChoiceText = vehicleId
    End If
End Function

Private Sub UndoResetVehicleOdometer(ByVal vehicleTable As ListObject, ByVal mileageTable As ListObject, _
    ByVal vehicleId As String, ByVal fallbackOdometer As Variant)
    Dim remainingOdometer As Variant

    remainingOdometer = UndoLatestOdometer(mileageTable, vehicleId)
    If IsNumeric(remainingOdometer) Then
        UndoSetVehicleOdometer vehicleTable, vehicleId, CLng(remainingOdometer)
    ElseIf IsNumeric(fallbackOdometer) Then
        UndoSetVehicleOdometer vehicleTable, vehicleId, CLng(fallbackOdometer)
    End If
End Sub

Private Function UndoLatestOdometer(ByVal mileageTable As ListObject, ByVal vehicleId As String) As Variant
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim endOdo As Variant
    Dim latestDate As Date
    Dim bestOdo As Long
    Dim hasOdo As Boolean

    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To mileageTable.ListRows.Count
        If StrComp(Trim$(CStr(UndoTableValue(mileageTable, rowIndex, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = UndoTableValue(mileageTable, rowIndex, "Date")
            endOdo = UndoTableValue(mileageTable, rowIndex, "Odometer End")
            If IsDate(rowDate) And IsNumeric(endOdo) Then
                If Not hasOdo Then
                    latestDate = UndoDateOnly(rowDate)
                    bestOdo = CLng(endOdo)
                    hasOdo = True
                ElseIf UndoDateOnly(rowDate) > latestDate Then
                    latestDate = UndoDateOnly(rowDate)
                    bestOdo = CLng(endOdo)
                ElseIf UndoDateOnly(rowDate) = latestDate And CLng(endOdo) > bestOdo Then
                    bestOdo = CLng(endOdo)
                End If
            End If
        End If
    Next rowIndex

    If hasOdo Then UndoLatestOdometer = bestOdo
End Function

Private Sub UndoSetVehicleOdometer(ByVal vehicleTable As ListObject, ByVal vehicleId As String, ByVal odometer As Long)
    Dim rowIndex As Long

    If vehicleTable.DataBodyRange Is Nothing Then Exit Sub
    For rowIndex = 1 To vehicleTable.ListRows.Count
        If StrComp(Trim$(CStr(UndoTableValue(vehicleTable, rowIndex, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            vehicleTable.DataBodyRange.Cells(rowIndex, vehicleTable.ListColumns("Odometer").index).value = odometer
            Exit Sub
        End If
    Next rowIndex
End Sub

Private Function UndoIsShopStatus(ByVal statusText As String) As Boolean
    statusText = UCase$(Trim$(statusText))
    UndoIsShopStatus = (statusText = "MAINTENANCE" Or statusText = "SHOP" Or statusText = "IN SHOP" Or Left$(statusText, 6) = "SHOP -")
End Function

Private Function UndoDateOnly(ByVal value As Variant) As Date
    If IsDate(value) Then
        UndoDateOnly = DateSerial(Year(CDate(value)), Month(CDate(value)), Day(CDate(value)))
    Else
        UndoDateOnly = Date
    End If
End Function

Private Function UndoDateRangeText(ByVal startDate As Date, ByVal endDate As Date) As String
    If UndoDateOnly(startDate) = UndoDateOnly(endDate) Then
        UndoDateRangeText = Format(startDate, "m/d/yyyy")
    Else
        UndoDateRangeText = Format(startDate, "m/d/yyyy") & " - " & Format(endDate, "m/d/yyyy")
    End If
End Function

Private Function UndoDisplayTime(ByVal timeValue As Variant) As String
    If IsDate(timeValue) Then
        UndoDisplayTime = Format(CDate(timeValue), "h:mm AM/PM")
    ElseIf IsNumeric(timeValue) Then
        UndoDisplayTime = Format(TimeSerial(0, 0, 0) + CDbl(timeValue), "h:mm AM/PM")
    Else
        UndoDisplayTime = Trim$(CStr(timeValue))
    End If
End Function

Private Sub UndoDashboardMessage(ByVal messageText As String)
    With ThisWorkbook.Worksheets("Dashboard").Range("A29")
        If .MergeCells Then
            .MergeArea.Cells(1, 1).value = messageText
        Else
            .value = messageText
        End If
    End With
End Sub

Private Sub UndoInvalidateWeeklyCache()
    On Error Resume Next
    ThisWorkbook.Worksheets("Weekly Reservations").Range("L3").value = "DIRTY"
    On Error GoTo 0
End Sub

Private Sub UndoRunWorkbookMacro(ByVal macroName As String, Optional ByVal argumentValue As Variant)
    On Error Resume Next
    If IsMissing(argumentValue) Then
        Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
    Else
        Application.Run "'" & ThisWorkbook.Name & "'!" & macroName, argumentValue
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub UndoRunWorkbookMacroWithTwoArguments(ByVal macroName As String, ByVal firstArgument As Variant, ByVal secondArgument As Variant)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName, firstArgument, secondArgument
    Err.Clear
    On Error GoTo 0
End Sub
