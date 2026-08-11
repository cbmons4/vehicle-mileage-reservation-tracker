Attribute VB_Name = "modFleetMileageTool"
Option Explicit

Private Const APP_TITLE As String = "Fleet Mileage Tool"
Private Const PORTFOLIO_REPOSITORY_URL As String = "https://github.com/cbmons4/vehicle-mileage-reservation-tracker"

Public Sub FleetAddVehicle()
    On Error GoTo CleanFail

    Dim lo As ListObject
    Dim lr As ListRow
    Dim vehicleId As String
    Dim makeModel As String
    Dim plate As String
    Dim status As String
    Dim branch As String
    Dim serviceDate As Variant
    Dim odometer As Variant
    Dim notes As String

    Set lo = GetTable("tblVehicles")

    vehicleId = PromptRequired("Vehicle ID")
    If vehicleId = "" Then Exit Sub
    If Not FindCell(lo, "Vehicle ID", vehicleId) Is Nothing Then
        MsgBox "Vehicle ID " & vehicleId & " already exists.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    makeModel = PromptRequired("Make / Model")
    If makeModel = "" Then Exit Sub
    plate = PromptRequired("License Plate")
    If plate = "" Then Exit Sub
    status = PromptChoice("Status", "Active,Retired,Maintenance,Out of Service", "Active")
    If status = "" Then Exit Sub
    branch = PromptRequired("Branch", "Demo Fleet")
    If branch = "" Then Exit Sub
    serviceDate = PromptDate("Last Service Date", Format(Date, "m/d/yyyy"))
    If IsEmpty(serviceDate) Then Exit Sub
    odometer = PromptNumber("Current Odometer", "0")
    If IsEmpty(odometer) Then Exit Sub
    notes = InputBox("Notes (optional)", APP_TITLE)

    Set lr = lo.ListRows.Add
    SetRowValue lo, lr, "Vehicle ID", vehicleId
    SetRowValue lo, lr, "Make/Model", makeModel
    SetRowValue lo, lr, "License Plate", plate
    SetRowValue lo, lr, "Status", status
    SetRowValue lo, lr, "Branch", branch
    SetRowValue lo, lr, "Last Service Date", CDate(serviceDate)
    SetRowValue lo, lr, "Odometer", CLng(odometer)
    SetRowValue lo, lr, "Notes", notes

    FleetRefreshDashboard False
    MsgBox "Vehicle added.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    MsgBox "Could not add vehicle: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub FleetAddDriver()
    On Error GoTo CleanFail

    Dim lo As ListObject
    Dim vehicleTable As ListObject
    Dim lr As ListRow
    Dim driverName As String
    Dim employeeId As String
    Dim contact As String
    Dim licenseExpiry As Variant
    Dim assignedVehicle As String
    Dim notes As String

    Set lo = GetTable("tblDrivers")
    Set vehicleTable = GetTable("tblVehicles")

    driverName = PromptRequired("Driver Name")
    If driverName = "" Then Exit Sub
    employeeId = PromptRequired("Employee ID")
    If employeeId = "" Then Exit Sub
    contact = InputBox("Contact (email or phone)", APP_TITLE)
    licenseExpiry = PromptDate("License Expiry", Format(DateAdd("yyyy", 1, Date), "m/d/yyyy"))
    If IsEmpty(licenseExpiry) Then Exit Sub
    assignedVehicle = InputBox("Assigned Vehicle ID (optional)" & vbCrLf & vbCrLf & ListValues(vehicleTable, "Vehicle ID", "Status", "Active"), APP_TITLE)
    assignedVehicle = Trim$(assignedVehicle)
    If assignedVehicle <> "" Then
        If FindCell(vehicleTable, "Vehicle ID", assignedVehicle) Is Nothing Then
            If MsgBox("Vehicle ID was not found. Add the driver anyway?", vbQuestion + vbYesNo, APP_TITLE) = vbNo Then Exit Sub
        End If
    End If
    notes = InputBox("Notes (optional)", APP_TITLE)

    Set lr = lo.ListRows.Add
    SetRowValue lo, lr, "Driver Name", driverName
    SetRowValue lo, lr, "Employee ID", employeeId
    SetRowValue lo, lr, "Contact", contact
    SetRowValue lo, lr, "License Expiry", CDate(licenseExpiry)
    SetRowValue lo, lr, "Assigned Vehicle", assignedVehicle
    SetRowValue lo, lr, "Notes", notes

    FleetRefreshDashboard False
    MsgBox "Driver added.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    MsgBox "Could not add driver: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub FleetReserveVehicle()
    FleetSaveDashboardCheckout
End Sub

Public Sub FleetLogMileage()
    FleetSaveDashboardMileage
End Sub

Public Sub FleetOpenMonthlyMileage()
    On Error GoTo CleanFail
    FleetRefreshMonthlyMileageView False
    Exit Sub

CleanFail:
    MsgBox "Could not open Monthly Mileage: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetOpenWeeklyReservations()
    On Error GoTo CleanFail
    FleetShowWeeklyReservations
    Exit Sub

CleanFail:
    MsgBox "Could not open Weekly Reservations: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetSaveDashboardMileage(Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim dashWs As Worksheet
    Dim lo As ListObject
    Dim vehicleTable As ListObject
    Dim lr As ListRow
    Dim logStartDate As Date
    Dim logEndDate As Date
    Dim vehicleId As String
    Dim startOdo As Variant
    Dim endOdo As Variant
    Dim makeModel As String

    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    Set lo = GetTable("tblMileageLog")
    Set vehicleTable = GetTable("tblVehicles")

    If Not TryDashboardDateRange(EntryValue(dashWs.Range("B6")), logStartDate, logEndDate) Then
        MsgBox "Enter date of start of trip.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B6")
        Exit Sub
    End If
    If logStartDate <> logEndDate Then
        MsgBox "Mileage can only be logged for one trip date. Use a single date for mileage entry.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B6")
        Exit Sub
    End If

    vehicleId = VehicleIdFromChoice(CStr(EntryValue(dashWs.Range("B7"))))
    If vehicleId = "" Or FindCell(vehicleTable, "Vehicle ID", vehicleId) Is Nothing Then
        MsgBox "Choose the vehicle from the drop-down list.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B7")
        Exit Sub
    End If

    startOdo = EntryValue(dashWs.Range("B8"))
    If Not IsNumeric(startOdo) Then
        MsgBox "Enter odometer at start of trip.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B8")
        Exit Sub
    End If

    endOdo = EntryValue(dashWs.Range("B9"))
    If Not IsNumeric(endOdo) Then
        MsgBox "Enter odometer at end of trip.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B9")
        Exit Sub
    End If

    If CDbl(endOdo) < CDbl(startOdo) Then
        MsgBox "Ending odometer cannot be lower than starting odometer.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B9")
        Exit Sub
    End If

    If FleetMileageEditActive() Then
        FleetCommitMileageEdit logStartDate, vehicleId, CLng(startOdo), CLng(endOdo), showMessage
        Exit Sub
    End If

    Set lr = lo.ListRows.Add
    SetRowValue lo, lr, "Log ID", NextNumber(lo, "Log ID")
    SetRowValue lo, lr, "Date", logStartDate
    SetRowValue lo, lr, "Vehicle ID", vehicleId
    SetRowValue lo, lr, "Trip Count", 1
    SetRowValue lo, lr, "Odometer Start", CLng(startOdo)
    SetRowValue lo, lr, "Odometer End", CLng(endOdo)
    SetRowValue lo, lr, "Miles Driven", CLng(endOdo) - CLng(startOdo)
    SetRowValue lo, lr, "Maintenance Flag", "No"
    SetRowValue lo, lr, "Notes", ""

    SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Odometer", CLng(endOdo)
    makeModel = CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Make/Model"))
    FleetRefreshMileageLogVehicleStatus False

    UpdateMonthlyMileageEntry logStartDate, vehicleId
    FleetClearMileageForm False
    FleetRefreshDashboard False
    dashWs.Range("A29").value = "Mileage logged: " & makeModel & " (" & vehicleId & "), " & _
        Format(logStartDate, "m/d/yyyy") & ", " & Format(CLng(startOdo), "#,##0") & " to " & _
        Format(CLng(endOdo), "#,##0") & " (" & Format(CLng(endOdo) - CLng(startOdo), "#,##0") & " miles)."
    If showMessage Then MsgBox "Mileage logged.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Or Application.Visible Then
        MsgBox "Could not log mileage: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetClearMileageForm(Optional ByVal showMessage As Boolean = True)
    Dim dashWs As Worksheet
    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    FleetCancelMileageEdit False
    dashWs.Range("B6").numberFormat = "m/d/yyyy"
    EntrySetValue dashWs.Range("B6"), Date
    EntryClear dashWs.Range("B7")
    EntryClear dashWs.Range("B8")
    EntryClear dashWs.Range("B9")
    EntryClear dashWs.Range("A29")
    If showMessage Then MsgBox "Mileage entry cleared.", vbInformation, APP_TITLE
End Sub

Public Sub FleetSaveDashboardCheckout(Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim dashWs As Worksheet
    Dim lo As ListObject
    Dim vehicleTable As ListObject
    Dim lr As ListRow
    Dim reserveStartDate As Date
    Dim reserveEndDate As Date
    Dim reserveCurrentDate As Date
    Dim vehicleId As String
    Dim driverName As String
    Dim startValue As Variant
    Dim endValue As Variant
    Dim startTime As Date
    Dim endTime As Date
    Dim purpose As String
    Dim status As String
    Dim vehicleStatus As String
    Dim makeModel As String
    Dim conflictText As String
    Dim savedCount As Long
    Dim nextReservationId As Long

    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    Set lo = GetTable("tblReservations")
    Set vehicleTable = GetTable("tblVehicles")
    EnsureVehicleSupportColumns vehicleTable

    SetDebugStatus "Checkout: date"
    If Not TryDashboardDateRange(EntryValue(dashWs.Range("F6")), reserveStartDate, reserveEndDate) Then
        DashboardValidationFailure showMessage, "Enter the checkout date."
        EntrySelect dashWs.Range("F6")
        Exit Sub
    End If

    SetDebugStatus "Checkout: vehicle"
    vehicleId = VehicleIdFromChoice(CStr(EntryValue(dashWs.Range("F7"))))
    If vehicleId = "" Or FindCell(vehicleTable, "Vehicle ID", vehicleId) Is Nothing Then
        DashboardValidationFailure showMessage, "Choose the vehicle from the drop-down list."
        EntrySelect dashWs.Range("F7")
        Exit Sub
    End If

    vehicleStatus = Trim$(CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Status")))
    If StrComp(vehicleStatus, "Maintenance", vbTextCompare) = 0 Then
        DashboardValidationFailure showMessage, "This vehicle is marked " & ShopStatusText(vehicleTable, vehicleId) & ". Use the Shop Status panel to make it available again."
        EntrySelect dashWs.Range("F7")
        Exit Sub
    End If

    driverName = Trim$(CStr(EntryValue(dashWs.Range("F8"))))
    If driverName = "" Then
        DashboardValidationFailure showMessage, "Enter staff name."
        EntrySelect dashWs.Range("F8")
        Exit Sub
    End If

    startValue = EntryValue(dashWs.Range("F9"))
    If Not IsTimeEntry(startValue) Then
        DashboardValidationFailure showMessage, "Choose a start time."
        EntrySelect dashWs.Range("F9")
        Exit Sub
    End If
    startTime = TimeEntryAsDate(startValue)

    endValue = EntryValue(dashWs.Range("H9"))
    If Not IsTimeEntry(endValue) Then
        DashboardValidationFailure showMessage, "Choose an end time."
        EntrySelect dashWs.Range("H9")
        Exit Sub
    End If
    endTime = TimeEntryAsDate(endValue)

    If CDbl(endTime) <= CDbl(startTime) Then
        DashboardValidationFailure showMessage, "End time must be later than start time."
        EntrySelect dashWs.Range("H9")
        Exit Sub
    End If

    purpose = Trim$(CStr(EntryValue(dashWs.Range("F10"))))
    If purpose = "" Then
        DashboardValidationFailure showMessage, "Enter route, project, city, or highway info."
        EntrySelect dashWs.Range("F10")
        Exit Sub
    End If

    status = "Reserved"

    If FleetReservationEditActive() Then
        FleetCommitReservationEdit reserveStartDate, reserveEndDate, vehicleId, driverName, startTime, endTime, purpose, showMessage
        Exit Sub
    End If

    SetDebugStatus "Checkout: conflicts"
    reserveCurrentDate = reserveStartDate
    Do While reserveCurrentDate <= reserveEndDate
        conflictText = ReservationConflictText(lo, vehicleId, reserveCurrentDate, startTime, endTime)
        If conflictText <> "" Then
            EntrySetValue dashWs.Range("A29"), "Vehicle is reserved: " & conflictText
            If showMessage Then MsgBox "Vehicle is reserved." & vbCrLf & vbCrLf & conflictText, vbExclamation, APP_TITLE
            Exit Sub
        End If
        reserveCurrentDate = DateAdd("d", 1, reserveCurrentDate)
    Loop

    SetDebugStatus "Checkout: add rows"
    nextReservationId = NextNumber(lo, "Reservation ID")
    reserveCurrentDate = reserveStartDate
    Do While reserveCurrentDate <= reserveEndDate
        Set lr = lo.ListRows.Add
        SetRowValue lo, lr, "Reservation ID", nextReservationId
        SetRowValue lo, lr, "Date", reserveCurrentDate
        SetRowValue lo, lr, "Vehicle ID", vehicleId
        SetRowValue lo, lr, "Driver Name", driverName
        SetRowValue lo, lr, "Start Time", Format(startTime, "h:mm AM/PM")
        SetRowValue lo, lr, "End Time", Format(endTime, "h:mm AM/PM")
        SetRowValue lo, lr, "Purpose", purpose
        SetRowValue lo, lr, "Status", status
        SetRowValue lo, lr, "Notes", ""
        savedCount = savedCount + 1
        nextReservationId = nextReservationId + 1
        reserveCurrentDate = DateAdd("d", 1, reserveCurrentDate)
    Loop
    MarkWeeklyReservationsDirty

    makeModel = CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Make/Model"))
    SetDebugStatus "Checkout: clear"
    FleetClearCheckoutForm False
    SetDebugStatus "Checkout: refresh dashboard"
    FleetRefreshDashboard False
    dashWs.Range("A29").value = "Checkout saved: " & driverName & ", " & makeModel & " (" & vehicleId & "), " & _
        DashboardDateRangeText(reserveStartDate, reserveEndDate) & " " & Format(startTime, "h:mm AM/PM") & "-" & _
        Format(endTime, "h:mm AM/PM") & " [" & status & "], " & CStr(savedCount) & " date(s)."
    If showMessage Then MsgBox "Checkout saved.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Or Application.Visible Then
        MsgBox "Could not save checkout: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetClearCheckoutForm(Optional ByVal showMessage As Boolean = True)
    Dim dashWs As Worksheet
    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    FleetCancelReservationEdit False
    dashWs.Range("F6").numberFormat = "m/d/yyyy"
    EntrySetValue dashWs.Range("F6"), Date
    EntryClear dashWs.Range("F7")
    EntryClear dashWs.Range("F8")
    EntrySetValue dashWs.Range("F9"), TimeSerial(8, 0, 0)
    EntrySetValue dashWs.Range("H9"), TimeSerial(17, 0, 0)
    EntryClear dashWs.Range("F10")
    EntryClear dashWs.Range("A29")
    If showMessage Then MsgBox "Checkout entry cleared.", vbInformation, APP_TITLE
End Sub

Public Sub FleetClearDashboardFields()
    Dim dashWs As Worksheet

    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    FleetClearMileageForm False
    FleetClearCheckoutForm False
    EntryClear dashWs.Range("B13")
    EntrySetValue dashWs.Range("B14"), "Shop - Central Garage"
    EntryClear dashWs.Range("A29")
    ClearDashboardDateCalendar dashWs
    MsgBox "Dashboard fields cleared.", vbInformation, APP_TITLE
End Sub

Public Sub FleetOpenMileageDateCalendar()
    OpenDashboardDateCalendar "B6", False
End Sub

Public Sub FleetOpenCheckoutDateCalendar()
    OpenDashboardDateCalendar "F6", True
End Sub

Public Sub FleetDashboardCalendarPreviousMonth()
    MoveDashboardDateCalendarMonth -1
End Sub

Public Sub FleetDashboardCalendarNextMonth()
    MoveDashboardDateCalendarMonth 1
End Sub

Public Sub FleetDashboardCalendarToggleRange()
    On Error GoTo CleanFail
    Dim previousScreenUpdating As Boolean
    Dim previousEvents As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ToggleDashboardCalendarRange
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not change calendar mode: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetDashboardCalendarPickDate()
    On Error GoTo CleanFail
    Dim previousScreenUpdating As Boolean
    Dim previousEvents As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    SelectDashboardCalendarDate CStr(Application.Caller)
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not select date: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetDashboardCalendarPickDateName(ByVal shapeName As String)
    On Error GoTo CleanFail
    Dim previousScreenUpdating As Boolean
    Dim previousEvents As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    SelectDashboardCalendarDate shapeName
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not select date: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetMarkVehicleInShop(Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim dashWs As Worksheet
    Dim vehicleTable As ListObject
    Dim vehicleId As String
    Dim shopLocation As String
    Dim shopStartValue As Variant
    Dim shopStart As Date

    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    Set vehicleTable = GetTable("tblVehicles")
    EnsureVehicleSupportColumns vehicleTable

    vehicleId = VehicleIdFromChoice(CStr(EntryValue(dashWs.Range("B13"))))
    If vehicleId = "" Or FindCell(vehicleTable, "Vehicle ID", vehicleId) Is Nothing Then
        MsgBox "Choose the vehicle from the Shop Status vehicle list.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B13")
        Exit Sub
    End If

    shopLocation = Trim$(CStr(EntryValue(dashWs.Range("B14"))))
    If Not IsAllowedShopLocation(shopLocation) Then
        MsgBox "Choose Shop - Central Garage, Shop - North Garage, or Shop - South Garage.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B14")
        Exit Sub
    End If

    SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Status", "Maintenance"
    SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Shop Location", Replace(shopLocation, "Shop - ", "", 1, 1, vbTextCompare)
    If Not IsDate(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Shop Start Date")) Then
        SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Shop Start Date", Date
    End If
    shopStartValue = GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Shop Start Date")
    If IsDate(shopStartValue) Then
        shopStart = DateOnly(shopStartValue)
    Else
        shopStart = Date
    End If
    AddShopReservationHistory vehicleId, shopStart, Date, shopLocation
    AppendVehicleNote vehicleTable, vehicleId, shopLocation & " as of " & Format(Date, "m/d/yyyy")
    MarkWeeklyReservationsDirty
    FleetRefreshMileageLogVehicleStatus False
    RefreshMonthlyMileageShopRange Date, MonthlyMileageProjectionEndDate()

    FleetRefreshDashboard False
    dashWs.Range("A29").value = "Shop status saved: " & VehicleChoiceTextForId(vehicleTable, vehicleId) & " is now " & shopLocation & "."
    If showMessage Then MsgBox "Vehicle marked " & shopLocation & ".", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Or Application.Visible Then
        MsgBox "Could not update shop status: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetMarkVehicleAvailable(Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim dashWs As Worksheet
    Dim vehicleTable As ListObject
    Dim vehicleId As String
    Dim shopStart As Variant
    Dim shopLocation As String
    Dim vehicleStatus As String
    Dim historyEnd As Date
    Dim projectionEnd As Date

    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    Set vehicleTable = GetTable("tblVehicles")
    EnsureVehicleSupportColumns vehicleTable

    vehicleId = VehicleIdFromChoice(CStr(EntryValue(dashWs.Range("B13"))))
    If vehicleId = "" Or FindCell(vehicleTable, "Vehicle ID", vehicleId) Is Nothing Then
        MsgBox "Choose the vehicle from the Shop Status vehicle list.", vbExclamation, APP_TITLE
        EntrySelect dashWs.Range("B13")
        Exit Sub
    End If

    vehicleStatus = Trim$(CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Status")))
    If StrComp(vehicleStatus, "Maintenance", vbTextCompare) <> 0 Then
        If showMessage Then MsgBox "This vehicle is already available. No shop history was changed.", vbInformation, APP_TITLE
        Exit Sub
    End If

    shopStart = GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Shop Start Date")
    shopLocation = ShopStatusText(vehicleTable, vehicleId)
    historyEnd = DateAdd("d", -1, Date)
    If IsDate(shopStart) Then
        If historyEnd >= DateOnly(shopStart) Then
            AddShopReservationHistory vehicleId, DateOnly(shopStart), historyEnd, shopLocation
            If Not ShopReservationHistoryComplete(vehicleId, DateOnly(shopStart), historyEnd) Then
                Err.Raise vbObjectError + 181, APP_TITLE, "Past shop history could not be saved. The vehicle remains marked in shop."
            End If
        End If
    End If

    DeleteShopReservationHistoryFromDate vehicleId, Date

    SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Status", "Active"
    SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Shop Location", ""
    SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Shop Start Date", ""
    AppendVehicleNote vehicleTable, vehicleId, "Returned from shop " & Format(Date, "m/d/yyyy")
    MarkWeeklyReservationsDirty
    FleetRefreshMileageLogVehicleStatus False
    projectionEnd = MonthlyMileageProjectionEndDate()
    RefreshMonthlyMileageShopRange Date, projectionEnd

    FleetRefreshDashboard False
    dashWs.Range("A29").value = "Shop status saved: " & VehicleChoiceTextForId(vehicleTable, vehicleId) & " is available."
    If showMessage Then MsgBox "Vehicle marked available. Past shop entries were retained.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Or Application.Visible Then
        MsgBox "Could not update shop status: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetDashboardMileageVehicleChanged()
    On Error GoTo CleanFail

    Dim dashWs As Worksheet
    Dim vehicleTable As ListObject
    Dim vehicleId As String
    Dim odometer As Variant

    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    Set vehicleTable = GetTable("tblVehicles")
    vehicleId = VehicleIdFromChoice(CStr(EntryValue(dashWs.Range("B7"))))
    If vehicleId = "" Then Exit Sub

    odometer = GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Odometer")
    If IsNumeric(odometer) Then
        EntrySetValue dashWs.Range("B8"), CLng(odometer)
        EntryClear dashWs.Range("B9")
    End If
    Exit Sub

CleanFail:
End Sub

Public Sub FleetRefreshDashboard(Optional ByVal showMessage As Boolean = True)
    On Error GoTo CleanFail

    Dim previousEvents As Boolean

    previousEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    UpdateDashboardData
    ThisWorkbook.Worksheets("Dashboard").Calculate
    FleetRefreshDashboardReviewHeader ThisWorkbook.Worksheets("Dashboard")

    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Range("A28").value = "Last refreshed: " & Format(Now, "m/d/yyyy h:mm AM/PM")
    ThisWorkbook.Worksheets("Dashboard").Activate
    On Error GoTo CleanFail

    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = True
    If showMessage Then MsgBox "Dashboard refreshed.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = True
    If showMessage Then
        MsgBox "Could not refresh dashboard: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetSyncMonthlyMileageEntry(ByVal logDate As Date, ByVal vehicleId As String)
    UpdateMonthlyMileageEntry DateOnly(logDate), Trim$(vehicleId)
End Sub

Public Sub FleetRefreshReports(Optional ByVal activateReports As Boolean = False)
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    FleetGenerateMonthlyMileage False
    FleetGenerateWeeklyReservations False
    Application.ScreenUpdating = True

    If activateReports Then ThisWorkbook.Worksheets("Monthly Mileage").Activate
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    If activateReports Then
        MsgBox "Could not refresh reports: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetOpenWorkbook()
    On Error GoTo CleanExit

    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim weeklyWs As Worksheet

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    FleetCancelAllEdits False
    Set weeklyWs = EnsureWorksheet("Weekly Reservations")
    weeklyWs.Range("C4").value = Date
    ApplyWorkbookMetadata
    ApplyWorkbookNavigationTabs False
    ApplyDashboardNavigationButtonMacros
    FleetRefreshMileageLogVehicleStatus False
    FleetRefreshMonthlyMileageShopStatus False
    FleetRefreshDashboard False
    ApplyDashboardNavigationButtonMacros
    ApplyWorkbookNavigationTabs True
    ApplySoftwareWindowChrome ThisWorkbook.Worksheets("Dashboard")

CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
End Sub

Private Sub ApplyWorkbookMetadata()
    On Error Resume Next

    Dim documentTitle As String

    documentTitle = "Vehicle Mileage and Reservation Tracker"
    If Len(documentTitle) = 0 Then documentTitle = "Vehicle Mileage and Reservation Tracker"
    ThisWorkbook.BuiltinDocumentProperties("Title").value = documentTitle
    ThisWorkbook.BuiltinDocumentProperties("Subject").value = "Excel VBA vehicle mileage and reservation workflow demonstration"
    ThisWorkbook.BuiltinDocumentProperties("Author").value = "Charles Monsanto"
    ThisWorkbook.BuiltinDocumentProperties("Manager").value = ""
    ThisWorkbook.BuiltinDocumentProperties("Company").value = ""
    ThisWorkbook.BuiltinDocumentProperties("Category").value = "Excel VBA Portfolio"
    ThisWorkbook.BuiltinDocumentProperties("Keywords").value = "Excel; VBA; fleet; mileage; reservations; reporting; process improvement"
    ThisWorkbook.BuiltinDocumentProperties("Content Status").value = "Portfolio Edition"
    ThisWorkbook.BuiltinDocumentProperties("Comments").value = _
        "Designed and built by Charles Monsanto as an Excel VBA portfolio project. " & _
        "All vehicles, staff, identifiers, mileage, reservations, and service locations are fictional. " & _
        "Source and documentation: " & PORTFOLIO_REPOSITORY_URL

    If WorkbookIsCloudHosted() Then
        ThisWorkbook.AutoSaveOn = True
    End If

    On Error GoTo 0
End Sub

Private Function WorkbookIsCloudHosted() As Boolean
    Dim locationText As String

    locationText = LCase$(ThisWorkbook.FullName & " " & ThisWorkbook.Path)
    WorkbookIsCloudHosted = (InStr(1, locationText, "https://", vbTextCompare) > 0 Or _
        InStr(1, locationText, "onedrive", vbTextCompare) > 0)
End Function

Private Sub ApplyWorkbookNavigationTabs(Optional ByVal activateDashboard As Boolean = False)
    On Error Resume Next

    Dim ws As Worksheet

    ThisWorkbook.Worksheets("Dashboard").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("Monthly Mileage").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("Weekly Reservations").Visible = xlSheetVisible

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case "Dashboard"
                ws.Tab.Color = RGB(0, 83, 143)
            Case "Monthly Mileage"
                ws.Tab.Color = RGB(255, 181, 17)
            Case "Weekly Reservations"
                ws.Tab.Color = RGB(0, 98, 155)
            Case Else
                ws.Visible = xlSheetHidden
        End Select
    Next ws

    If activateDashboard Then ApplySoftwareWindowChrome ThisWorkbook.Worksheets("Dashboard")

    On Error GoTo 0
End Sub

Private Sub ApplySoftwareWindowChrome(Optional ByVal targetSheet As Worksheet = Nothing)
    On Error Resume Next

    Dim targetWindow As Window
    Dim previousScreenUpdating As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If Not targetSheet Is Nothing Then
        If Not (Application.ActiveSheet Is targetSheet) Then targetSheet.Activate
    End If
    Set targetWindow = ActiveWindow
    If Not targetWindow Is Nothing Then
        If targetWindow.DisplayWorkbookTabs Then targetWindow.DisplayWorkbookTabs = False
        If targetWindow.DisplayGridlines Then targetWindow.DisplayGridlines = False
        If targetWindow.DisplayHeadings Then targetWindow.DisplayHeadings = False
    End If
    Application.ScreenUpdating = previousScreenUpdating

    On Error GoTo 0
End Sub

Private Sub ApplyDashboardNavigationButtonMacros()
    On Error Resume Next

    With ThisWorkbook.Worksheets("Dashboard")
        StyleDashboardNavigationShape .Shapes("btnMONTHLYMILEAGEA11D11"), _
            "GO TO MILEAGE TRACKER TAB", "FleetShowMonthlyMileage"
        StyleDashboardNavigationShape .Shapes("btnWEEKLYCHECKOUTBOARDE12H15"), _
            "GO TO RESERVATIONS TAB", "FleetShowWeeklyReservations"
    End With
    EnsureDashboardUtilityButtons
    AlignDashboardStaticShapes

    On Error GoTo 0
End Sub

Private Sub StyleDashboardNavigationShape(ByVal shp As Shape, ByVal caption As String, ByVal macroName As String)
    shp.Visible = msoTrue
    shp.OnAction = macroName
    shp.Fill.Visible = msoTrue
    shp.Fill.Transparency = 0
    shp.Fill.ForeColor.RGB = RGB(0, 34, 68)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(230, 233, 236)
    shp.Line.Weight = 1
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Name = "Aptos"
    shp.TextFrame.Characters.Font.Size = 11
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.HorizontalAlignment = xlCenter
    shp.TextFrame.VerticalAlignment = xlCenter
    shp.Placement = xlMove
    shp.ZOrder msoBringToFront
End Sub

Private Sub EnsureDashboardUtilityButtons()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Rows(30).RowHeight = 26
    AddOrUpdateDashboardUtilityButton ws, "btnCLEARFIELDSA16H16", "A30:B30", "CLEAR FIELDS", "FleetClearDashboardFields", RGB(238, 240, 242)
    AddOrUpdateDashboardUtilityButton ws, "btnUndoLastMileage", "C30:E30", "EDIT MILEAGE", "FleetEditMileage", RGB(238, 240, 242)
    AddOrUpdateDashboardUtilityButton ws, "btnUndoLastCheckout", "F30:H30", "EDIT RESERVATION", "FleetEditReservation", RGB(238, 240, 242)
End Sub

Private Sub AddOrUpdateDashboardUtilityButton(ByVal ws As Worksheet, ByVal shapeName As String, _
    ByVal address As String, ByVal caption As String, ByVal macroName As String, ByVal fillColor As Long)
    Dim rng As Range
    Dim shp As Shape

    Set rng = ws.Range(address)
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
        shp.Name = shapeName
    End If

    shp.Visible = msoTrue
    shp.Left = rng.Left + 2
    shp.Top = rng.Top + 3
    shp.Width = rng.Width - 4
    shp.Height = rng.Height - 6
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.Transparency = 0
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(215, 220, 225)
    shp.Line.Weight = 1
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Name = "Aptos"
    shp.TextFrame.Characters.Font.Color = RGB(0, 34, 68)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.HorizontalAlignment = xlCenter
    shp.TextFrame.VerticalAlignment = xlCenter
    shp.OnAction = macroName
    shp.Placement = xlMove
    shp.ZOrder msoBringToFront
End Sub

Public Sub FleetGenerateMonthlyMileage(Optional ByVal activateSheet As Boolean = True)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim vehicleTable As ListObject
    Dim mileageTable As ListObject
    Dim reservationTable As ListObject
    Dim minDate As Date
    Dim maxDate As Date
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim hasDates As Boolean
    Dim startRow As Long

    FleetMaterializeActiveShopHistory False
    Set ws = EnsureWorksheet("Monthly Mileage")
    Set vehicleTable = GetTable("tblVehicles")
    Set mileageTable = GetTable("tblMileageLog")
    Set reservationTable = GetTable("tblReservations")

    SetDebugStatus "Monthly: clear"
    ClearReportSheet ws
    SetDebugStatus "Monthly: date range"
    GetDateRange mileageTable, "Date", minDate, maxDate, hasDates
    If Not hasDates Then
        minDate = Date
        maxDate = Date
    End If

    SetDebugStatus "Monthly: drawing"
    monthStart = DateSerial(Year(minDate), Month(minDate), 1)
    startRow = 1
    Do While monthStart <= DateSerial(Year(maxDate), Month(maxDate), 1)
        SetDebugStatus "Monthly: " & Format(monthStart, "mmm yyyy")
        monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)
        DrawMonthlyMileageBlock ws, startRow, monthStart, monthEnd, vehicleTable, mileageTable, reservationTable
        startRow = startRow + Day(monthEnd) + 5
        monthStart = DateAdd("m", 1, monthStart)
    Loop

    SetDebugStatus "Monthly: widths"
    ws.Columns("A:J").ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 15.5
    ws.Columns(5).ColumnWidth = 15.5
    ws.Columns(7).ColumnWidth = 15.5
    ws.Columns(9).ColumnWidth = 15.5
    SetDebugStatus "Monthly: navigation"
    If activateSheet Or Application.Visible Then
        EnsureMonthlyMileageNavigation ws
    End If
    If activateSheet Then
        FreezeMonthlyMileageNavigation ws
        ApplySoftwareWindowChrome ws
    End If
    SetDebugStatus "Monthly: done"
    Exit Sub

CleanFail:
    If activateSheet Then
        MsgBox "Could not build Monthly Mileage: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetGenerateWeeklyReservations(Optional ByVal activateSheet As Boolean = True)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim vehicleTable As ListObject
    Dim reservationTable As ListObject
    Dim selectedDate As Date
    Dim weekStart As Date
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    FleetMaterializeActiveShopHistory False
    Set ws = EnsureWorksheet("Weekly Reservations")
    Set vehicleTable = GetTable("tblVehicles")
    Set reservationTable = GetTable("tblReservations")
    EnsureVehicleSupportColumns vehicleTable
    selectedDate = GetSelectedReservationDate(ws)
    weekStart = StartOfWeek(selectedDate)

    ClearWeeklyReservationsSheet ws
    ws.Rows("1:40").Hidden = False
    DrawWeeklyReservationControls ws, selectedDate, weekStart
    EnsureWeeklyDashboardNavigationButton ws
    DrawWeeklyReservationBlock ws, 15, weekStart, vehicleTable, reservationTable
    StoreWeeklyReservationsCache ws, weekStart, WeeklyReservationSignature(reservationTable, vehicleTable, weekStart)

    ws.Range("A:A").ColumnWidth = 10
    ws.Range("B:B").ColumnWidth = 14
    ws.Range("C:C").ColumnWidth = 18
    ws.Range("D:E").ColumnWidth = 12
    ws.Range("F:J").ColumnWidth = 14
    ProtectWeeklyReservationsSheet ws
    If activateSheet Then
        ApplySoftwareWindowChrome ws
    End If
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Sub

CleanFail:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    If activateSheet Then
        MsgBox "Could not build Weekly Reservations: " & Err.Description, vbCritical, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Private Function OpenWeeklyReservationsFromCache() As Boolean
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim selectedDate As Date
    Dim weekStart As Date
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set ws = EnsureWorksheet("Weekly Reservations")
    selectedDate = GetSelectedReservationDate(ws)
    weekStart = StartOfWeek(selectedDate)

    OpenWeeklyReservationsFromCache = WeeklyReservationsCacheValid(ws, weekStart)

    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Function

CleanFail:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    OpenWeeklyReservationsFromCache = False
End Function

Private Function WeeklyReservationsCacheValid(ByVal ws As Worksheet, ByVal weekStart As Date) As Boolean
    Dim selectedDate As Date
    Dim expectedWeekCaption As String

    If UCase$(Trim$(CStr(ws.Range("A1").value))) <> "WEEKLY RESERVATIONS" Then Exit Function
    If CLng(Val(ws.Range("L1").value)) <> CLng(weekStart) Then Exit Function
    If UCase$(Trim$(CStr(ws.Range("L3").value))) <> "READY" Then Exit Function
    If Not IsDate(ws.Range("C4").value) Then Exit Function

    selectedDate = DateOnly(CDate(ws.Range("C4").value))
    If StrComp(Trim$(CStr(ws.Range("C3").value)), Format$(selectedDate, "mmmm"), vbTextCompare) <> 0 Then Exit Function
    If CLng(Val(ws.Range("F3").value)) <> Year(selectedDate) Then Exit Function
    expectedWeekCaption = Format$(weekStart, "mmmm d") & " - " & _
        Format$(DateAdd("d", 6, weekStart), "mmmm d, yyyy")
    If StrComp(Trim$(CStr(ws.Range("C5").value)), expectedWeekCaption, vbTextCompare) <> 0 Then Exit Function
    WeeklyReservationsCacheValid = True
End Function

Private Function WeeklyReservationsShellReady(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    WeeklyReservationsShellReady = _
        (UCase$(Trim$(CStr(ws.Range("A1").value))) = "WEEKLY RESERVATIONS" And _
        ws.Range("A15:J25").Rows.Count = 11)
    On Error GoTo 0
End Function

Private Sub StoreWeeklyReservationsCache(ByVal ws As Worksheet, ByVal weekStart As Date, ByVal signature As String)
    On Error Resume Next
    ws.Range("L1").value = CLng(weekStart)
    ws.Range("L2").value = signature
    ws.Range("L3").value = "READY"
    ws.Columns("K:M").Hidden = True
    On Error GoTo 0
End Sub

Public Sub MarkWeeklyReservationsDirty()
    On Error Resume Next
    With ThisWorkbook.Worksheets("Weekly Reservations")
        .Range("L3").value = "DIRTY"
        .Columns("K:M").Hidden = True
    End With
    On Error GoTo 0
End Sub

Private Function WeeklyReservationSignature(ByVal reservationTable As ListObject, ByVal vehicleTable As ListObject, ByVal weekStart As Date) As String
    Dim reservationText As String
    Dim vehicleText As String
    Dim vehicleId As String
    Dim vehicleIndex As Long
    Dim vehicleValues As Variant
    Dim vehicleIdIndex As Long
    Dim vehicleStatusIndex As Long
    Dim shopLocationIndex As Long
    Dim shopStartDateIndex As Long
    Dim reservationCount As Long
    Dim lastReservationId As Variant

    vehicleText = "V" & CStr(vehicleTable.ListRows.Count)
    If Not vehicleTable.DataBodyRange Is Nothing Then
        vehicleValues = vehicleTable.DataBodyRange.Value2
        vehicleIdIndex = vehicleTable.ListColumns("Vehicle ID").index
        vehicleStatusIndex = vehicleTable.ListColumns("Status").index
        shopLocationIndex = vehicleTable.ListColumns("Shop Location").index
        shopStartDateIndex = vehicleTable.ListColumns("Shop Start Date").index
        For vehicleIndex = 1 To UBound(vehicleValues, 1)
            vehicleId = CStr(vehicleValues(vehicleIndex, vehicleIdIndex))
            vehicleText = vehicleText & "|" & SignatureText(vehicleId) & ":" & _
                SignatureText(vehicleValues(vehicleIndex, vehicleStatusIndex)) & ":" & _
                SignatureText(vehicleValues(vehicleIndex, shopLocationIndex)) & ":" & _
                SignatureText(vehicleValues(vehicleIndex, shopStartDateIndex))
        Next vehicleIndex
    End If

    reservationText = "R0"
    If Not reservationTable.DataBodyRange Is Nothing Then
        reservationCount = reservationTable.ListRows.Count
        lastReservationId = reservationTable.DataBodyRange.Cells( _
            reservationCount, reservationTable.ListColumns("Reservation ID").index).Value2
        reservationText = "R" & CStr(reservationCount) & ":" & SignatureText(lastReservationId)
    End If

    WeeklyReservationSignature = CStr(CLng(weekStart)) & "|" & vehicleText & "|" & reservationText
End Function

Private Function SignatureText(ByVal value As Variant) As String
    If IsDate(value) Then
        SignatureText = Format$(CDate(value), "yyyymmdd hhnnss")
    Else
        SignatureText = Replace(Trim$(CStr(value)), "|", "/")
    End If
End Function

Public Sub FleetPreviousReservationWeek()
    MoveReservationWeek -7
End Sub

Public Sub FleetThisReservationWeek()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet("Weekly Reservations")
    ws.Range("C4").value = Date
    FleetGenerateWeeklyReservations True
End Sub

Public Sub FleetNextReservationWeek()
    MoveReservationWeek 7
End Sub

Public Sub FleetPreviousReservationMonth()
    MoveReservationMonth -1
End Sub

Public Sub FleetNextReservationMonth()
    MoveReservationMonth 1
End Sub

Public Sub FleetToggleReservationCalendar()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim previousScreenUpdating As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = EnsureWorksheet("Weekly Reservations")
    If ws.Rows("6:13").Hidden Then
        DrawReservationCalendarPopup ws, GetSelectedReservationDate(ws)
    Else
        ws.Rows("6:13").Hidden = True
    End If
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not open calendar: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetCalendarPreviousMonth()
    MoveReservationCalendarMonth -1
End Sub

Public Sub FleetCalendarNextMonth()
    MoveReservationCalendarMonth 1
End Sub

Public Sub FleetWeeklyMonthYearChanged()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim selectedDate As Date
    Dim targetMonth As Long
    Dim targetYear As Long
    Dim targetDay As Long
    Dim maxDay As Long

    Set ws = EnsureWorksheet("Weekly Reservations")
    selectedDate = GetSelectedReservationDate(ws)
    targetMonth = MonthNumberFromName(CStr(ws.Range("C3").value))
    If targetMonth = 0 Then targetMonth = Month(selectedDate)

    If IsNumeric(ws.Range("F3").value) Then
        targetYear = CLng(ws.Range("F3").value)
    Else
        targetYear = Year(selectedDate)
    End If

    If targetYear < 2020 Or targetYear > 2035 Then targetYear = Year(selectedDate)
    maxDay = Day(DateSerial(targetYear, targetMonth + 1, 0))
    targetDay = Day(selectedDate)
    If targetDay > maxDay Then targetDay = maxDay

    RefreshWeeklyReservationsForDate ws, DateSerial(targetYear, targetMonth, targetDay)
    Exit Sub

CleanFail:
    MsgBox "Could not update the selected week: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetHandleWeeklyCalendarSelection(ByVal targetCell As Range)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim selectedDate As Date
    Dim calendarMonth As Date
    Dim value As Variant

    Set ws = targetCell.Worksheet
    If Intersect(targetCell, ws.Range("A6:G13")) Is Nothing Then Exit Sub

    If targetCell.address(False, False) = "A6" Then
        calendarMonth = CalendarMonthFromHeader(ws)
        DrawReservationCalendarPopup ws, DateAdd("m", -1, calendarMonth)
        Exit Sub
    End If
    If targetCell.address(False, False) = "G6" Then
        calendarMonth = CalendarMonthFromHeader(ws)
        DrawReservationCalendarPopup ws, DateAdd("m", 1, calendarMonth)
        Exit Sub
    End If

    value = targetCell.value
    If Not IsDate(value) Then Exit Sub

    selectedDate = DateOnly(CDate(value))
    RefreshWeeklyReservationsForDate ws, selectedDate

CleanExit:
    Exit Sub

CleanFail:
    MsgBox "Could not select calendar date: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub FleetShowDashboard()
    FleetDismissDashboardChoicePopup
    ApplySoftwareWindowChrome ThisWorkbook.Worksheets("Dashboard")
End Sub

Public Sub FleetShowVehicles()
    ThisWorkbook.Worksheets("Vehicles").Activate
End Sub

Public Sub FleetShowDrivers()
    With ThisWorkbook.Worksheets("Drivers")
        .Visible = xlSheetVisible
        .Activate
    End With
End Sub

Public Sub FleetShowReservations()
    ThisWorkbook.Worksheets("Reservations").Activate
End Sub

Public Sub FleetShowMileageLog()
    FleetRefreshMileageLogVehicleStatus False
    ThisWorkbook.Worksheets("Mileage Log").Activate
End Sub

Public Sub FleetShowImportSummary()
    ThisWorkbook.Worksheets("Import Summary").Activate
End Sub

Public Sub FleetShowMonthlyMileage()
    On Error GoTo CleanFail

    FleetDismissDashboardChoicePopup
    FleetRefreshMonthlyMileageView False
    Exit Sub

CleanFail:
    If Application.Visible Then
        MsgBox "Could not open Monthly Mileage: " & Err.Description, vbExclamation, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetRefreshMonthlyMileageView(Optional ByVal showMessage As Boolean = False)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim targetMonth As Date
    Dim previousScreenUpdating As Boolean
    Dim previousEvents As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    FleetRefreshMonthlyMileageShopStatus False
    Set ws = EnsureWorksheet("Monthly Mileage")
    targetMonth = ExistingMonthlyMileageMonthOrFallback(ws, DateSerial(Year(Date), Month(Date), 1))
    EnsureMonthlyMileageNavigation ws
    If Application.Visible Then
        FreezeMonthlyMileageNavigation ws
        SelectMonthlyMileageMonth ws, targetMonth
        ApplySoftwareWindowChrome ws
    End If
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    If showMessage Then MsgBox "Monthly Mileage refreshed.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    If showMessage Or Application.Visible Then
        MsgBox "Could not refresh Monthly Mileage: " & Err.Description, vbExclamation, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Public Sub FleetShowMonthlyMileageCurrent()
    FleetShowMonthlyMileage
End Sub

Public Sub FleetShowWeeklyReservations()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim selectedDate As Date
    Dim weekStart As Date
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    FleetDismissDashboardChoicePopup
    FleetMaterializeActiveShopHistory False
    Set ws = ThisWorkbook.Worksheets("Weekly Reservations")
    selectedDate = GetSelectedReservationDate(ws)
    weekStart = StartOfWeek(selectedDate)
    If Not WeeklyReservationsCacheValid(ws, weekStart) Then
        If WeeklyReservationsShellReady(ws) Then
            RefreshWeeklyReservationsGrid ws, selectedDate
        Else
            FleetGenerateWeeklyReservations False
        End If
    End If
    ApplySoftwareWindowChrome ws
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    If Application.Visible Then
        MsgBox "Could not open Weekly Reservations: " & errorDescription, vbExclamation, APP_TITLE
    Else
        Err.Raise errorNumber, APP_TITLE, errorDescription
    End If
End Sub

Public Sub FleetHideSetup()
    On Error Resume Next
    ThisWorkbook.Worksheets("Setup").Visible = xlSheetHidden
    ApplySoftwareWindowChrome ThisWorkbook.Worksheets("Dashboard")
End Sub

Private Sub MoveReservationWeek(ByVal daysToMove As Long)
    Dim ws As Worksheet
    Dim selectedDate As Date

    Set ws = EnsureWorksheet("Weekly Reservations")
    selectedDate = GetSelectedReservationDate(ws)
    RefreshWeeklyReservationsForDate ws, DateAdd("d", daysToMove, selectedDate)
End Sub

Private Sub MoveReservationMonth(ByVal monthsToMove As Long)
    Dim ws As Worksheet
    Dim selectedDate As Date

    Set ws = EnsureWorksheet("Weekly Reservations")
    selectedDate = GetSelectedReservationDate(ws)
    RefreshWeeklyReservationsForDate ws, DateAdd("m", monthsToMove, selectedDate)
End Sub

Private Sub MoveReservationCalendarMonth(ByVal monthsToMove As Long)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim calendarMonth As Date
    Dim previousScreenUpdating As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = EnsureWorksheet("Weekly Reservations")
    calendarMonth = CalendarMonthFromHeader(ws)
    DrawReservationCalendarPopup ws, DateAdd("m", monthsToMove, calendarMonth)
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not move the calendar month: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub OpenDashboardDateCalendar(ByVal targetAddress As String, Optional ByVal allowRange As Boolean = False)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim seedDate As Variant
    Dim seedStartDate As Date
    Dim seedEndDate As Date
    Dim stateWs As Worksheet
    Dim previousScreenUpdating As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    seedDate = EntryValue(ws.Range(targetAddress))
    If TryDashboardDateRange(seedDate, seedStartDate, seedEndDate) Then
        seedDate = seedStartDate
    Else
        seedDate = Date
    End If

    CreateOrUpdateName "DashboardCalendarTarget", ws.Range(targetAddress)
    Set stateWs = DashboardCalendarStateSheet()
    stateWs.Range("P1").value = targetAddress
    stateWs.Range("P2").value = IIf(allowRange, "TRUE", "FALSE")
    stateWs.Range("P3").value = "Single"
    stateWs.Range("P4").ClearContents
    DrawDashboardDateCalendar ws, CDate(seedDate), allowRange
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not open calendar: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub MoveDashboardDateCalendarMonth(ByVal monthsToMove As Long)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim calendarMonth As Date
    Dim previousScreenUpdating As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    calendarMonth = DashboardCalendarMonthFromHeader(ws)
    DrawDashboardDateCalendar ws, DateAdd("m", monthsToMove, calendarMonth), DashboardCalendarAllowsRange()
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not move the calendar month: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub ToggleDashboardCalendarRange()
    Dim ws As Worksheet
    Dim stateWs As Worksheet
    Dim calendarMonth As Date

    If Not DashboardCalendarAllowsRange() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set stateWs = DashboardCalendarStateSheet()
    If DashboardCalendarRangeMode() Then
        stateWs.Range("P3").value = "Single"
    Else
        stateWs.Range("P3").value = "Range"
    End If
    stateWs.Range("P4").ClearContents
    calendarMonth = DashboardCalendarMonthFromHeader(ws)
    DrawDashboardDateCalendar ws, calendarMonth, True
End Sub

Private Sub DrawDashboardDateCalendar(ByVal ws As Worksheet, ByVal calendarSeedDate As Date, Optional ByVal allowRange As Boolean = False)
    Dim calendarMonth As Date
    Dim selectedDate As Date
    Dim selectedStartDate As Date
    Dim selectedEndDate As Date
    Dim hasSelection As Boolean
    Dim pendingStartDate As Date
    Dim hasPendingStart As Boolean
    Dim firstGridDate As Date
    Dim currentDate As Date
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim dayOffset As Long
    Dim dayNames As Variant
    Dim targetRange As Range
    Dim panelArea As Range
    Dim leftPos As Double
    Dim topPos As Double
    Dim calendarWidth As Double
    Dim calendarHeight As Double
    Dim headerHeight As Double
    Dim dayHeaderHeight As Double
    Dim dayHeight As Double
    Dim dayWidth As Double
    Dim navWidth As Double
    Dim modeHeight As Double
    Dim gridTop As Double
    Dim fillColor As Long
    Dim fontColor As Long
    Dim dayShape As Shape

    On Error Resume Next
    Set targetRange = ThisWorkbook.names("DashboardCalendarTarget").RefersToRange
    On Error GoTo 0
    If targetRange Is Nothing Then Exit Sub

    selectedDate = Date
    hasSelection = TryDashboardDateRange(EntryValue(targetRange), selectedStartDate, selectedEndDate)
    If hasSelection Then selectedDate = selectedStartDate
    If allowRange And DashboardCalendarRangeMode() Then
        If IsDate(DashboardCalendarStateSheet().Range("P4").value) Then
            pendingStartDate = DateOnly(DashboardCalendarStateSheet().Range("P4").value)
            hasPendingStart = True
            selectedDate = pendingStartDate
        End If
    End If
    calendarMonth = DateSerial(Year(calendarSeedDate), Month(calendarSeedDate), 1)
    firstGridDate = DateAdd("d", 1 - Weekday(calendarMonth, vbSunday), calendarMonth)
    dayNames = Array("SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT")

    ClearDashboardDateCalendar ws

    If targetRange.Column < 5 Then
        Set panelArea = ws.Range("A7:D14")
    Else
        Set panelArea = ws.Range("E7:H14")
    End If

    leftPos = panelArea.Left + 4
    topPos = panelArea.Top + 2
    calendarWidth = panelArea.Width - 8
    headerHeight = 24
    dayHeaderHeight = 18
    dayHeight = 20
    If allowRange Then modeHeight = 18
    calendarHeight = headerHeight + modeHeight + dayHeaderHeight + (6 * dayHeight) + 8
    dayWidth = calendarWidth / 7
    navWidth = 28
    gridTop = topPos + headerHeight + modeHeight

    AddDashboardCalendarShape ws, "dashCalBack", "", leftPos - 2, topPos - 2, calendarWidth + 4, calendarHeight + 4, RGB(255, 255, 255), RGB(0, 0, 0), "", ""
    AddDashboardCalendarShape ws, "dashCalPrev", "<", leftPos, topPos, navWidth, headerHeight, RGB(0, 83, 143), RGB(255, 255, 255), "FleetDashboardCalendarPreviousMonth", ""
    Set dayShape = AddDashboardCalendarShape(ws, "dashCalHeader", Format(calendarMonth, "mmmm yyyy"), leftPos + navWidth, topPos, calendarWidth - (2 * navWidth), headerHeight, RGB(0, 83, 143), RGB(255, 255, 255), "", CStr(CLng(calendarMonth)))
    dayShape.TextFrame.Characters.Font.Size = 10
    AddDashboardCalendarShape ws, "dashCalNext", ">", leftPos + calendarWidth - navWidth, topPos, navWidth, headerHeight, RGB(0, 83, 143), RGB(255, 255, 255), "FleetDashboardCalendarNextMonth", ""
    If allowRange Then
        Set dayShape = AddDashboardCalendarShape(ws, "dashCalRangeMode", IIf(DashboardCalendarRangeMode(), "RANGE ON", "RANGE OFF"), _
            leftPos, topPos + headerHeight, calendarWidth, modeHeight, IIf(DashboardCalendarRangeMode(), RGB(255, 181, 17), RGB(232, 244, 252)), _
            RGB(0, 83, 143), "FleetDashboardCalendarToggleRange", "")
        dayShape.TextFrame.Characters.Font.Size = 8
    End If

    For colIndex = 0 To 6
        AddDashboardCalendarShape ws, "dashCalDow" & CStr(colIndex + 1), CStr(dayNames(colIndex)), leftPos + (colIndex * dayWidth), gridTop, dayWidth, dayHeaderHeight, RGB(232, 244, 252), RGB(0, 83, 143), "", ""
    Next colIndex

    dayOffset = 0
    For rowIndex = 0 To 5
        For colIndex = 0 To 6
            currentDate = DateAdd("d", dayOffset, firstGridDate)
            fillColor = RGB(255, 255, 255)
            fontColor = RGB(0, 0, 0)
            If Month(currentDate) <> Month(calendarMonth) Then
                fillColor = RGB(242, 242, 242)
                fontColor = RGB(128, 128, 128)
            End If
            If hasSelection And DateOnly(selectedStartDate) <> DateOnly(selectedEndDate) _
                And DateOnly(currentDate) >= DateOnly(selectedStartDate) And DateOnly(currentDate) <= DateOnly(selectedEndDate) Then
                fillColor = RGB(255, 242, 204)
                fontColor = RGB(0, 0, 0)
            End If
            If hasSelection And (DateOnly(currentDate) = DateOnly(selectedStartDate) Or DateOnly(currentDate) = DateOnly(selectedEndDate)) Then
                fillColor = RGB(0, 83, 143)
                fontColor = RGB(255, 255, 255)
            ElseIf hasPendingStart And DateOnly(currentDate) = DateOnly(pendingStartDate) Then
                fillColor = RGB(0, 83, 143)
                fontColor = RGB(255, 255, 255)
            ElseIf DateOnly(currentDate) = Date Then
                fillColor = RGB(255, 230, 153)
                fontColor = RGB(0, 0, 0)
            End If

            Set dayShape = AddDashboardCalendarShape(ws, "dashCalDay" & Format(currentDate, "yyyymmdd"), CStr(Day(currentDate)), _
                leftPos + (colIndex * dayWidth), gridTop + dayHeaderHeight + (rowIndex * dayHeight), _
                dayWidth, dayHeight, fillColor, fontColor, "FleetDashboardCalendarPickDate", CStr(CLng(currentDate)))
            If DateOnly(currentDate) = Date Then
                dayShape.Line.ForeColor.RGB = RGB(255, 121, 0)
                dayShape.Line.Weight = 2
            End If
            dayOffset = dayOffset + 1
        Next colIndex
    Next rowIndex
End Sub

Public Sub FleetDismissDashboardDateCalendar(ByVal targetCell As Range)
    On Error Resume Next
    FleetDismissDashboardChoicePopup
    On Error GoTo SafeExit
    If targetCell Is Nothing Then Exit Sub
    If StrComp(targetCell.Worksheet.Name, "Dashboard", vbTextCompare) <> 0 Then Exit Sub

    Dim ws As Worksheet
    Dim calendarBack As Shape
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    Set calendarBack = ws.Shapes("dashCalBack")
    On Error GoTo SafeExit
    If calendarBack Is Nothing Then Exit Sub
    ClearDashboardDateCalendar ws
SafeExit:
End Sub

Private Sub ClearDashboardDateCalendar(ByVal ws As Worksheet)
    DeleteShapesByPrefix ws, "dashCal"
    On Error Resume Next
    ws.Range("J5:P12").UnMerge
    ws.Range("J5:P12").Clear
    On Error GoTo 0
End Sub

Private Function DashboardCalendarMonthFromHeader(ByVal ws As Worksheet) As Date
    Dim serialText As String

    On Error Resume Next
    serialText = CStr(ws.Shapes("dashCalHeader").AlternativeText)
    If Err.Number = 0 And Len(serialText) > 0 Then
        DashboardCalendarMonthFromHeader = CDate(CLng(serialText))
    Else
        DashboardCalendarMonthFromHeader = DateSerial(Year(Date), Month(Date), 1)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SelectDashboardCalendarDate(ByVal shapeName As String)
    Dim ws As Worksheet
    Dim selectedDate As Date
    Dim targetRange As Range
    Dim stateWs As Worksheet
    Dim startDate As Date
    Dim endDate As Date

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    selectedDate = DateOnly(CDate(CLng(ws.Shapes(shapeName).AlternativeText)))
    Set targetRange = ThisWorkbook.names("DashboardCalendarTarget").RefersToRange

    If DashboardCalendarAllowsRange() And DashboardCalendarRangeMode() Then
        Set stateWs = DashboardCalendarStateSheet()
        If Not IsDate(stateWs.Range("P4").value) Then
            stateWs.Range("P4").value = selectedDate
            targetRange.numberFormat = "m/d/yyyy"
            EntrySetValue targetRange, selectedDate
            DrawDashboardDateCalendar ws, selectedDate, True
            Exit Sub
        End If

        startDate = DateOnly(stateWs.Range("P4").value)
        endDate = selectedDate
        If endDate < startDate Then SwapDates startDate, endDate
        targetRange.numberFormat = "@"
        EntrySetValue targetRange, DashboardDateRangeText(startDate, endDate)
        stateWs.Range("P4").ClearContents
        stateWs.Range("P3").value = "Single"
        ClearDashboardDateCalendar ws
        EntrySelect targetRange
        Exit Sub
    End If

    targetRange.numberFormat = "m/d/yyyy"
    EntrySetValue targetRange, selectedDate
    ClearDashboardDateCalendar ws
    EntrySelect targetRange
End Sub

Private Function AddDashboardCalendarShape(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
    ByVal leftPos As Double, ByVal topPos As Double, ByVal shapeWidth As Double, ByVal shapeHeight As Double, _
    ByVal fillColor As Long, ByVal fontColor As Long, ByVal macroName As String, ByVal altText As String) As Shape

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, shapeWidth, shapeHeight)
    shp.Name = shapeName
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(180, 180, 180)
    shp.Line.Weight = 1
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Color = fontColor
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 8
    shp.TextFrame.HorizontalAlignment = xlCenter
    shp.TextFrame.VerticalAlignment = xlCenter
    shp.Placement = xlMove
    If macroName <> "" Then shp.OnAction = macroName
    If altText <> "" Then shp.AlternativeText = altText
    Set AddDashboardCalendarShape = shp
End Function

Private Function GetSelectedReservationDate(ByVal ws As Worksheet) As Date
    Dim value As Variant

    value = ws.Range("C4").value
    If IsDate(value) Then
        GetSelectedReservationDate = DateOnly(CDate(value))
        Exit Function
    End If

    value = ws.Range("B2").value
    If IsDate(value) Then
        GetSelectedReservationDate = DateOnly(CDate(value))
    Else
        GetSelectedReservationDate = Date
    End If
End Function

Private Function GetSelectedReservationWeekStart(ByVal ws As Worksheet) As Date
    Dim value As Variant

    value = ws.Range("B2").value
    If IsDate(value) Then
        GetSelectedReservationWeekStart = StartOfWeek(CDate(value))
    Else
        GetSelectedReservationWeekStart = StartOfWeek(Date)
    End If
End Function

Private Function GetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set GetTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not GetTable Is Nothing Then Exit Function
    Next ws

    Err.Raise vbObjectError + 7100, APP_TITLE, "Missing table: " & tableName
End Function

Private Function EntryCell(ByVal targetCell As Range) As Range
    If targetCell.MergeCells Then
        Set EntryCell = targetCell.MergeArea.Cells(1, 1)
    Else
        Set EntryCell = targetCell.Cells(1, 1)
    End If
End Function

Private Function EntryValue(ByVal targetCell As Range) As Variant
    EntryValue = EntryCell(targetCell).value
End Function

Private Sub EntrySetValue(ByVal targetCell As Range, ByVal value As Variant)
    EntryCell(targetCell).value = value
End Sub

Private Sub EntryClear(ByVal targetCell As Range)
    EntryCell(targetCell).value = vbNullString
End Sub

Private Sub EntrySelect(ByVal targetCell As Range)
    EntryCell(targetCell).Select
End Sub

Private Sub DashboardValidationFailure(ByVal showMessage As Boolean, ByVal messageText As String)
    If showMessage Or Application.Visible Then
        MsgBox messageText, vbExclamation, APP_TITLE
    Else
        Err.Raise vbObjectError + 7201, APP_TITLE, messageText
    End If
End Sub

Private Sub SetDebugStatus(ByVal statusText As String)
    On Error Resume Next
    ThisWorkbook.Worksheets("Setup").Range("B12").value = statusText
    On Error GoTo 0
End Sub

Private Function IsTimeEntry(ByVal value As Variant) As Boolean
    If IsDate(value) Then
        IsTimeEntry = True
    ElseIf IsNumeric(value) Then
        IsTimeEntry = (CDbl(value) >= 0 And CDbl(value) < 1)
    End If
End Function

Private Function TimeEntryAsDate(ByVal value As Variant) As Date
    If IsDate(value) Then
        TimeEntryAsDate = timeValue(CDate(value))
    Else
        TimeEntryAsDate = TimeSerial(0, 0, 0) + CDbl(value)
    End If
End Function

Private Sub SetRowValue(ByVal lo As ListObject, ByVal lr As ListRow, ByVal columnName As String, ByVal value As Variant)
    lr.Range.Cells(1, lo.ListColumns(columnName).index).value = value
End Sub

Private Function FindCell(ByVal lo As ListObject, ByVal columnName As String, ByVal key As String) As Range
    Dim cell As Range

    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each cell In lo.ListColumns(columnName).DataBodyRange.Cells
        If StrComp(Trim$(CStr(cell.value)), Trim$(key), vbTextCompare) = 0 Then
            Set FindCell = cell
            Exit Function
        End If
    Next cell
End Function

Private Function GetRelatedValue(ByVal lo As ListObject, ByVal keyColumn As String, ByVal key As String, ByVal returnColumn As String) As Variant
    Dim keyCell As Range
    Dim rowIndex As Long

    Set keyCell = FindCell(lo, keyColumn, key)
    If keyCell Is Nothing Then
        GetRelatedValue = ""
    Else
        rowIndex = keyCell.row - lo.DataBodyRange.row + 1
        GetRelatedValue = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(returnColumn).index).value
    End If
End Function

Private Sub SetRelatedValue(ByVal lo As ListObject, ByVal keyColumn As String, ByVal key As String, ByVal targetColumn As String, ByVal value As Variant)
    Dim keyCell As Range
    Dim rowIndex As Long

    Set keyCell = FindCell(lo, keyColumn, key)
    If keyCell Is Nothing Then Exit Sub
    rowIndex = keyCell.row - lo.DataBodyRange.row + 1
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(targetColumn).index).value = value
End Sub

Private Sub AppendVehicleNote(ByVal lo As ListObject, ByVal vehicleId As String, ByVal noteText As String)
    Dim currentNotes As String

    currentNotes = CStr(GetRelatedValue(lo, "Vehicle ID", vehicleId, "Notes"))
    If Len(Trim$(currentNotes)) > 0 Then
        currentNotes = currentNotes & "; " & noteText
    Else
        currentNotes = noteText
    End If
    SetRelatedValue lo, "Vehicle ID", vehicleId, "Notes", currentNotes
End Sub

Private Function NextNumber(ByVal lo As ListObject, ByVal idColumn As String) As Long
    On Error GoTo Fallback
    If lo.DataBodyRange Is Nothing Then
        NextNumber = 1
    Else
        NextNumber = CLng(Application.WorksheetFunction.Max(lo.ListColumns(idColumn).DataBodyRange)) + 1
    End If
    Exit Function

Fallback:
    NextNumber = lo.ListRows.Count + 1
End Function

Private Function PromptRequired(ByVal prompt As String, Optional ByVal defaultValue As String = "") As String
    PromptRequired = Trim$(InputBox(prompt & vbCrLf & vbCrLf & "Leave blank to cancel.", APP_TITLE, defaultValue))
End Function

Private Function PromptDate(ByVal prompt As String, ByVal defaultValue As String) As Variant
    Dim value As String

    Do
        value = PromptRequired(prompt, defaultValue)
        If value = "" Then Exit Function
        If IsDate(value) Then
            PromptDate = CDate(value)
            Exit Function
        End If
        MsgBox "Enter a valid date.", vbExclamation, APP_TITLE
    Loop
End Function

Private Function PromptNumber(ByVal prompt As String, ByVal defaultValue As String) As Variant
    Dim value As String

    Do
        value = PromptRequired(prompt, defaultValue)
        If value = "" Then Exit Function
        If IsNumeric(value) Then
            PromptNumber = CDbl(value)
            Exit Function
        End If
        MsgBox "Enter a valid number.", vbExclamation, APP_TITLE
    Loop
End Function

Private Function PromptChoice(ByVal prompt As String, ByVal choices As String, ByVal defaultValue As String) As String
    Dim value As String
    Dim item As Variant

    Do
        value = PromptRequired(prompt & vbCrLf & "Choices: " & choices, defaultValue)
        If value = "" Then Exit Function
        For Each item In Split(choices, ",")
            If StrComp(value, Trim$(CStr(item)), vbTextCompare) = 0 Then
                PromptChoice = Trim$(CStr(item))
                Exit Function
            End If
        Next item
        MsgBox "Choose one of: " & choices, vbExclamation, APP_TITLE
    Loop
End Function

Private Function ListValues(ByVal lo As ListObject, ByVal valueColumn As String, Optional ByVal filterColumn As String = "", Optional ByVal filterValue As String = "") As String
    Dim i As Long
    Dim valueText As String
    Dim filterText As String
    Dim result As String
    Dim shown As Long

    If lo.DataBodyRange Is Nothing Then
        ListValues = "No records yet."
        Exit Function
    End If

    For i = 1 To lo.ListRows.Count
        If filterColumn <> "" Then
            filterText = Trim$(CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(filterColumn).index).value))
            If StrComp(filterText, filterValue, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        valueText = Trim$(CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(valueColumn).index).value))
        If valueText <> "" Then
            result = result & valueText & vbCrLf
            shown = shown + 1
        End If
        If shown >= 25 Then
            result = result & "...more in the sheet"
            Exit For
        End If

NextRow:
    Next i

    If result = "" Then result = "No matching records."
    ListValues = result
End Function

Private Sub UpdateDashboardData()
    Dim dashWs As Worksheet
    Dim dataWs As Worksheet
    Dim vehicleTable As ListObject
    Dim mileageTable As ListObject
    Dim reservationTable As ListObject
    Dim fiscalYearList As String
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim monthList As String
    Dim firstFyStart As Date
    Dim lastFyStart As Date
    Dim hasFiscalDates As Boolean
    Dim fiscalYearStart As Date
    Dim fiscalYearEnd As Date
    Dim vehicleIndex As Long
    Dim vehicleId As String
    Dim statusText As String
    Dim rowNumber As Long
    Dim monthMiles As Double
    Dim cumulativeMiles As Double
    Dim scheduledHours As Double
    Dim tripDays As Long
    Dim tripCount As Long
    Dim lastOdometer As Variant
    Dim totalMiles As Double
    Dim totalCumulativeMiles As Double
    Dim totalScheduledHours As Double
    Dim totalTripDays As Long
    Dim totalTrips As Long

    Set dashWs = ThisWorkbook.Worksheets("Dashboard")
    Set dataWs = EnsureWorksheet("DashboardData")
    Set vehicleTable = GetTable("tblVehicles")
    Set mileageTable = GetTable("tblMileageLog")
    Set reservationTable = GetTable("tblReservations")

    EnsureVehicleSupportColumns vehicleTable

    fiscalYearList = BuildFiscalYearValidationList(mileageTable, reservationTable, firstFyStart, lastFyStart, hasFiscalDates)
    fiscalYearStart = DashboardSelectedFiscalYearStart(dashWs, firstFyStart, lastFyStart, hasFiscalDates)
    fiscalYearEnd = DateAdd("d", -1, DateAdd("yyyy", 1, fiscalYearStart))
    monthList = BuildFiscalYearMonthValidationList(fiscalYearStart)
    monthStart = DashboardSelectedMonthStartForFiscalYear(dashWs, fiscalYearStart)
    monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)

    dashWs.Range("D17").numberFormat = "@"
    dashWs.Range("G17").numberFormat = "mmmm yyyy"
    dashWs.Range("D17").value = fiscalYearLabel(fiscalYearStart)
    dashWs.Range("G17").value = monthStart
    dashWs.Range("C17").ClearContents
    If Not dashWs.Range("H17").MergeCells Then dashWs.Range("H17").ClearContents
    ApplyFiscalYearValidation dashWs.Range("D17"), fiscalYearList
    ApplyMonthValidation dashWs.Range("G17"), monthList
    UpdateDashboardChoiceLists dataWs, vehicleTable, fiscalYearList, monthList
    ApplyDashboardEntryValidation dashWs

    ClearDashboardSummary dashWs
    WriteSummaryHeaders dashWs.Range("A19:H19")
    dataWs.Range("A:H").Clear
    WriteSummaryHeaders dataWs.Range("A1:H1")

    For vehicleIndex = 1 To vehicleTable.ListRows.Count
        vehicleId = CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))
        rowNumber = 19 + vehicleIndex
        monthMiles = MonthMilesForVehicle(mileageTable, monthStart, monthEnd, vehicleId)
        cumulativeMiles = FiscalYearMilesForVehicle(mileageTable, vehicleId, fiscalYearStart, monthStart)
        scheduledHours = ScheduledHoursForVehicleMonth(reservationTable, vehicleId, monthStart, monthEnd)
        tripDays = TripDaysForVehicleMonth(mileageTable, vehicleId, monthStart, monthEnd)
        tripCount = MonthTripsForVehicle(mileageTable, monthStart, monthEnd, vehicleId)
        lastOdometer = LastOdometerForVehicleThroughDate(mileageTable, vehicleId, monthEnd)
        statusText = CheckoutStatusForVehicle(vehicleTable, reservationTable, vehicleId, Date)

        WriteVehicleMonthlyRow dashWs, rowNumber, _
            CStr(TableValue(vehicleTable, vehicleIndex, "Make/Model")), _
            lastOdometer, statusText, monthMiles, tripDays, tripCount, scheduledHours, cumulativeMiles
        WriteVehicleMonthlyRow dataWs, vehicleIndex + 1, _
            CStr(TableValue(vehicleTable, vehicleIndex, "Make/Model")), _
            lastOdometer, statusText, monthMiles, tripDays, tripCount, scheduledHours, cumulativeMiles

        StyleCheckoutStatus dashWs.Cells(rowNumber, 3), statusText
        StyleMonthlyMiles dashWs.Cells(rowNumber, 4), monthMiles
        totalMiles = totalMiles + monthMiles
        totalTripDays = totalTripDays + tripDays
        totalTrips = totalTrips + tripCount
        totalScheduledHours = totalScheduledHours + scheduledHours
        totalCumulativeMiles = totalCumulativeMiles + cumulativeMiles
    Next vehicleIndex

    WriteVehicleMonthlyRow dashWs, 24, "TOTAL", "", "", totalMiles, totalTripDays, totalTrips, totalScheduledHours, totalCumulativeMiles
    With dashWs.Range("A24:H24")
        .Interior.Color = RGB(248, 248, 248)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
    End With

    dashWs.Range("A26").value = "Monthly mileage target: 500 miles per vehicle"
    dashWs.Range("A27").value = "Fiscal year: " & Format(fiscalYearStart, "m/d/yyyy") & " through " & Format(fiscalYearEnd, "m/d/yyyy") & "; FYTD miles are summed from fiscal year start through the selected month."
End Sub

Private Function BuildMonthValidationList(ByVal mileageTable As ListObject, ByRef firstMonth As Date, ByRef lastMonth As Date, ByRef hasMonths As Boolean) As String
    Dim minDate As Date
    Dim maxDate As Date
    Dim currentMonth As Date
    Dim result As String

    GetDateRange mileageTable, "Date", minDate, maxDate, hasMonths
    If Not hasMonths Then
        firstMonth = DateSerial(Year(Date), Month(Date), 1)
        lastMonth = firstMonth
    Else
        firstMonth = DateSerial(Year(minDate), Month(minDate), 1)
        lastMonth = DateSerial(Year(maxDate), Month(maxDate), 1)
    End If

    currentMonth = firstMonth
    Do While currentMonth <= lastMonth
        If result <> "" Then result = result & ","
        result = result & Format(currentMonth, "mmmm yyyy")
        currentMonth = DateAdd("m", 1, currentMonth)
    Loop
    BuildMonthValidationList = result
End Function

Private Function BuildFiscalYearMonthValidationList(ByVal fiscalYearStart As Date) As String
    Dim currentMonth As Date
    Dim fiscalYearEnd As Date
    Dim result As String

    currentMonth = fiscalYearStart
    fiscalYearEnd = DateAdd("d", -1, DateAdd("yyyy", 1, fiscalYearStart))
    Do While currentMonth <= fiscalYearEnd
        If result <> "" Then result = result & ","
        result = result & Format(currentMonth, "mmmm yyyy")
        currentMonth = DateAdd("m", 1, currentMonth)
    Loop
    BuildFiscalYearMonthValidationList = result
End Function

Private Function DashboardSelectedMonthStartForFiscalYear(ByVal dashWs As Worksheet, ByVal fiscalYearStart As Date) As Date
    Dim candidate As Date
    Dim fiscalYearEnd As Date

    fiscalYearEnd = DateAdd("d", -1, DateAdd("yyyy", 1, fiscalYearStart))
    If TryMonthStartFromValue(dashWs.Range("G17").value, candidate) Then
        If candidate >= fiscalYearStart And candidate <= fiscalYearEnd Then
            DashboardSelectedMonthStartForFiscalYear = candidate
            Exit Function
        End If
    End If

    DashboardSelectedMonthStartForFiscalYear = DateSerial(Year(fiscalYearEnd), Month(fiscalYearEnd), 1)
End Function

Private Function DashboardSelectedMonthStart(ByVal dashWs As Worksheet, ByVal firstMonth As Date, ByVal lastMonth As Date, ByVal hasMonths As Boolean) As Date
    Dim candidate As Date
    Dim currentMonth As Date

    If TryMonthStartFromValue(dashWs.Range("G17").value, candidate) Then
        If Not hasMonths Or (candidate >= firstMonth And candidate <= lastMonth) Then
            DashboardSelectedMonthStart = candidate
            Exit Function
        End If
    End If

    currentMonth = DateSerial(Year(Date), Month(Date), 1)
    If hasMonths Then
        If currentMonth < firstMonth Or currentMonth > lastMonth Then currentMonth = lastMonth
    End If
    DashboardSelectedMonthStart = currentMonth
End Function

Private Function TryMonthStartFromValue(ByVal value As Variant, ByRef monthStart As Date) As Boolean
    If IsDate(value) Then
        monthStart = DateSerial(Year(CDate(value)), Month(CDate(value)), 1)
        TryMonthStartFromValue = True
        Exit Function
    End If

    If IsNumeric(value) Then
        If CDbl(value) > 20000 And CDbl(value) < 60000 Then
            monthStart = DateSerial(Year(CDate(CDbl(value))), Month(CDate(CDbl(value))), 1)
            TryMonthStartFromValue = True
            Exit Function
        End If
    End If

    TryMonthStartFromValue = TryMonthStartFromLabel(Trim$(CStr(value)), monthStart)
End Function

Private Function TryMonthStartFromLabel(ByVal labelText As String, ByRef monthStart As Date) As Boolean
    Dim parsedDate As Date

    If labelText = "" Then Exit Function
    If TryMonthStartFromMonthYearText(labelText, monthStart) Then
        TryMonthStartFromLabel = True
        Exit Function
    End If

    If IsDate(labelText) Then
        parsedDate = CDate(labelText)
        monthStart = DateSerial(Year(parsedDate), Month(parsedDate), 1)
        TryMonthStartFromLabel = True
        Exit Function
    End If

    On Error Resume Next
    parsedDate = CDate("1 " & labelText)
    If Err.Number = 0 Then
        monthStart = DateSerial(Year(parsedDate), Month(parsedDate), 1)
        TryMonthStartFromLabel = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function TryMonthStartFromMonthYearText(ByVal labelText As String, ByRef monthStart As Date) As Boolean
    Dim cleaned As String
    Dim parts As Variant
    Dim monthNumber As Long
    Dim yearNumber As Long

    cleaned = Trim$(Replace(labelText, Chr$(160), " "))
    cleaned = Replace(cleaned, "-", " ")
    Do While InStr(cleaned, "  ") > 0
        cleaned = Replace(cleaned, "  ", " ")
    Loop
    If cleaned = "" Then Exit Function

    parts = Split(cleaned, " ")
    If UBound(parts) < 1 Then Exit Function

    monthNumber = MonthNumberFromName(CStr(parts(0)))
    If monthNumber = 0 Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    yearNumber = CLng(parts(1))
    If yearNumber < 100 Then yearNumber = 2000 + yearNumber
    If yearNumber < 1900 Or yearNumber > 9999 Then Exit Function

    monthStart = DateSerial(yearNumber, monthNumber, 1)
    TryMonthStartFromMonthYearText = True
End Function

Private Sub ApplyMonthValidation(ByVal targetCell As Range, ByVal listText As String)
    On Error Resume Next
    targetCell.Validation.Delete
    On Error GoTo 0

    If Len(listText) > 0 Then
        targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listText
        targetCell.Validation.IgnoreBlank = False
        targetCell.Validation.InCellDropdown = True
    End If
End Sub

Private Sub UpdateDashboardChoiceLists(ByVal dataWs As Worksheet, ByVal vehicleTable As ListObject, ByVal fiscalYearList As String, ByVal monthList As String)
    Dim vehicleIndex As Long
    Dim rowNumber As Long
    Dim currentTime As Date

    dataWs.Range("J:O").ClearContents
    dataWs.Range("J1").value = "Vehicle Choices"
    dataWs.Range("K1").value = "Fiscal Years"
    dataWs.Range("L1").value = "Months"
    dataWs.Range("M1").value = "Times"
    dataWs.Range("N1").value = "Shop Locations"
    dataWs.Range("O1").value = "Checkout Statuses"

    For vehicleIndex = 1 To vehicleTable.ListRows.Count
        dataWs.Cells(vehicleIndex + 1, "J").value = VehicleChoiceTextForRow(vehicleTable, vehicleIndex)
    Next vehicleIndex

    WriteCsvListToColumn dataWs, "K", fiscalYearList
    WriteMonthListToColumn dataWs, "L", monthList

    rowNumber = 2
    currentTime = TimeSerial(6, 0, 0)
    Do While currentTime <= TimeSerial(18, 0, 0)
        dataWs.Cells(rowNumber, "M").value = currentTime
        dataWs.Cells(rowNumber, "M").numberFormat = "h:mm AM/PM"
        currentTime = DateAdd("n", 15, currentTime)
        rowNumber = rowNumber + 1
    Loop

    dataWs.Range("N2").value = "Shop - Central Garage"
    dataWs.Range("N3").value = "Shop - North Garage"
    dataWs.Range("N4").value = "Shop - South Garage"
    dataWs.Range("O2").value = "Reserved"

    CreateOrUpdateName "VehicleChoiceList", dataWs.Range("J2:J" & CStr(Application.Max(2, vehicleTable.ListRows.Count + 1)))
    CreateOrUpdateName "FiscalYearChoiceList", NonEmptyListRange(dataWs, "K")
    CreateOrUpdateName "MonthChoiceList", NonEmptyListRange(dataWs, "L")
    CreateOrUpdateName "TimeChoiceList", dataWs.Range("M2:M50")
    CreateOrUpdateName "ShopLocationChoiceList", dataWs.Range("N2:N4")
    CreateOrUpdateName "CheckoutStatusChoiceList", dataWs.Range("O2:O2")
End Sub

Private Sub WriteMonthListToColumn(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal csvList As String)
    Dim item As Variant
    Dim rowNumber As Long
    Dim monthStart As Date

    rowNumber = 2
    For Each item In Split(csvList, ",")
        If TryMonthStartFromLabel(Trim$(CStr(item)), monthStart) Then
            ws.Cells(rowNumber, columnLetter).value = monthStart
            ws.Cells(rowNumber, columnLetter).numberFormat = "mmmm yyyy"
            rowNumber = rowNumber + 1
        End If
    Next item
End Sub

Private Sub ApplyDashboardEntryValidation(ByVal dashWs As Worksheet)
    ApplyNamedValidation dashWs.Range("D17:E17"), "=FiscalYearChoiceList"
    ApplyNamedValidation dashWs.Range("G17"), "=MonthChoiceList"
    ApplyNamedValidation dashWs.Range("B7:D7"), "=VehicleChoiceList"
    ApplyNamedValidation dashWs.Range("F7:H7"), "=VehicleChoiceList"
    ApplyNamedValidation dashWs.Range("B13:D13"), "=VehicleChoiceList"
    ApplyNamedValidation dashWs.Range("B14:D14"), "=ShopLocationChoiceList"
    ApplyNamedValidation dashWs.Range("F9"), "=TimeChoiceList"
    ApplyNamedValidation dashWs.Range("H9"), "=TimeChoiceList"
End Sub

Private Sub ApplyNamedValidation(ByVal targetRange As Range, ByVal formulaText As String)
    On Error Resume Next
    targetRange.Validation.Delete
    targetRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formulaText
    targetRange.Validation.IgnoreBlank = False
    targetRange.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

Private Sub WriteCsvListToColumn(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal csvList As String)
    Dim item As Variant
    Dim rowNumber As Long

    rowNumber = 2
    For Each item In Split(csvList, ",")
        If Trim$(CStr(item)) <> "" Then
            ws.Cells(rowNumber, columnLetter).value = Trim$(CStr(item))
            rowNumber = rowNumber + 1
        End If
    Next item
End Sub

Private Function NonEmptyListRange(ByVal ws As Worksheet, ByVal columnLetter As String) As Range
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).row
    If lastRow < 2 Then lastRow = 2
    Set NonEmptyListRange = ws.Range(columnLetter & "2:" & columnLetter & CStr(lastRow))
End Function

Private Sub CreateOrUpdateName(ByVal rangeName As String, ByVal targetRange As Range)
    On Error Resume Next
    ThisWorkbook.names(rangeName).Delete
    Err.Clear
    On Error GoTo 0
    ThisWorkbook.names.Add Name:=rangeName, RefersTo:="='" & targetRange.Worksheet.Name & "'!" & targetRange.address
End Sub

Public Sub FleetRefreshMileageLogVehicleStatus(Optional ByVal showMessage As Boolean = False)
    On Error GoTo CleanFail

    Dim mileageTable As ListObject
    Dim vehicleTable As ListObject
    Dim reservationTable As ListObject
    Dim statusColumn As ListColumn
    Dim statusRange As Range
    Dim shopFormat As FormatCondition
    Dim shopStatusByKey As Object
    Dim mileageValues As Variant
    Dim reservationValues As Variant
    Dim existingValues As Variant
    Dim statusValues() As Variant
    Dim vehicleId As String
    Dim statusText As String
    Dim rowDate As Variant
    Dim historyKey As String
    Dim rowIndex As Long
    Dim vehicleIdIndex As Long
    Dim logDateIndex As Long
    Dim reservationDateIndex As Long
    Dim reservationVehicleIndex As Long
    Dim reservationStatusIndex As Long
    Dim reservationPurposeIndex As Long
    Dim reservationDriverIndex As Long
    Dim needsWrite As Boolean
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set mileageTable = GetTable("tblMileageLog")
    Set vehicleTable = GetTable("tblVehicles")
    Set reservationTable = GetTable("tblReservations")
    EnsureVehicleSupportColumns vehicleTable
    FleetMaterializeActiveShopHistory False

    If TableHasColumn(mileageTable, "Vehicle Status on Date") Then
        Set statusColumn = mileageTable.ListColumns("Vehicle Status on Date")
    ElseIf TableHasColumn(mileageTable, "Current Vehicle Status") Then
        Set statusColumn = mileageTable.ListColumns("Current Vehicle Status")
        statusColumn.Name = "Vehicle Status on Date"
        needsWrite = True
    Else
        Set statusColumn = mileageTable.ListColumns.Add
        statusColumn.Name = "Vehicle Status on Date"
        needsWrite = True
    End If

    statusColumn.Range.EntireColumn.ColumnWidth = 24
    statusColumn.Range.HorizontalAlignment = xlLeft
    statusColumn.Range.WrapText = False

    If Not mileageTable.DataBodyRange Is Nothing Then
        Set shopStatusByKey = CreateObject("Scripting.Dictionary")
        shopStatusByKey.CompareMode = vbTextCompare
        If Not reservationTable.DataBodyRange Is Nothing Then
            reservationValues = reservationTable.DataBodyRange.Value2
            reservationDateIndex = reservationTable.ListColumns("Date").index
            reservationVehicleIndex = reservationTable.ListColumns("Vehicle ID").index
            reservationStatusIndex = reservationTable.ListColumns("Status").index
            reservationPurposeIndex = reservationTable.ListColumns("Purpose").index
            reservationDriverIndex = reservationTable.ListColumns("Driver Name").index
            For rowIndex = 1 To UBound(reservationValues, 1)
                rowDate = reservationValues(rowIndex, reservationDateIndex)
                statusText = Trim$(CStr(reservationValues(rowIndex, reservationStatusIndex)))
                If (IsDate(rowDate) Or IsNumeric(rowDate)) And IsShopReservationStatus(statusText) Then
                    vehicleId = Trim$(CStr(reservationValues(rowIndex, reservationVehicleIndex)))
                    historyKey = MonthlyShopStatusKey(DateOnly(rowDate), vehicleId)
                    shopStatusByKey(historyKey) = ShopReservationLabel(statusText, _
                        reservationValues(rowIndex, reservationPurposeIndex), reservationValues(rowIndex, reservationDriverIndex))
                End If
            Next rowIndex
        End If

        mileageValues = mileageTable.DataBodyRange.Value2
        existingValues = statusColumn.DataBodyRange.Value2
        vehicleIdIndex = mileageTable.ListColumns("Vehicle ID").index
        logDateIndex = mileageTable.ListColumns("Date").index
        ReDim statusValues(1 To mileageTable.ListRows.Count, 1 To 1)
        For rowIndex = 1 To mileageTable.ListRows.Count
            vehicleId = Trim$(CStr(mileageValues(rowIndex, vehicleIdIndex)))
            rowDate = mileageValues(rowIndex, logDateIndex)
            If IsDate(rowDate) Or IsNumeric(rowDate) Then
                historyKey = MonthlyShopStatusKey(DateOnly(rowDate), vehicleId)
            Else
                historyKey = vbNullString
            End If
            If historyKey <> vbNullString And shopStatusByKey.Exists(historyKey) Then
                statusValues(rowIndex, 1) = shopStatusByKey(historyKey)
            ElseIf (IsDate(rowDate) Or IsNumeric(rowDate)) And VehicleInShopOnDate(vehicleTable, vehicleId, DateOnly(rowDate)) Then
                statusValues(rowIndex, 1) = ShopStatusText(vehicleTable, vehicleId)
            Else
                statusValues(rowIndex, 1) = "Available"
            End If
            If Not needsWrite Then
                If CStr(existingValues(rowIndex, 1)) <> CStr(statusValues(rowIndex, 1)) Then needsWrite = True
            End If
        Next rowIndex

        Set statusRange = statusColumn.DataBodyRange
        If needsWrite Then statusRange.Value2 = statusValues
        statusRange.Font.Color = RGB(0, 0, 0)
        statusRange.Font.Bold = False
        If statusRange.FormatConditions.Count = 0 Then
            Set shopFormat = statusRange.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=LEFT(" & statusRange.Cells(1, 1).address(False, False) & ",6)=""Shop -""")
            shopFormat.Interior.Color = RGB(255, 199, 206)
            shopFormat.Font.Color = RGB(156, 0, 6)
            shopFormat.Font.Bold = True
        End If
    End If

    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    If showMessage Then MsgBox "Mileage Log vehicle status refreshed.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    If showMessage Then
        MsgBox "Could not refresh Mileage Log vehicle status: " & Err.Description, vbExclamation, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Private Sub EnsureVehicleSupportColumns(ByVal vehicleTable As ListObject)
    If Not TableHasColumn(vehicleTable, "Shop Location") Then
        vehicleTable.ListColumns.Add.Name = "Shop Location"
    End If
    If Not TableHasColumn(vehicleTable, "Shop Start Date") Then
        vehicleTable.ListColumns.Add.Name = "Shop Start Date"
    End If
End Sub

Private Function VehicleChoiceTextForRow(ByVal vehicleTable As ListObject, ByVal rowIndex As Long) As String
    VehicleChoiceTextForRow = CStr(TableValue(vehicleTable, rowIndex, "Vehicle ID")) & " - " & CStr(TableValue(vehicleTable, rowIndex, "Make/Model"))
End Function

Private Function VehicleChoiceTextForId(ByVal vehicleTable As ListObject, ByVal vehicleId As String) As String
    Dim makeModel As String

    makeModel = CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Make/Model"))
    If makeModel <> "" Then
        VehicleChoiceTextForId = vehicleId & " - " & makeModel
    Else
        VehicleChoiceTextForId = vehicleId
    End If
End Function

Private Function VehicleIdFromChoice(ByVal choiceText As String) As String
    Dim position As Long

    choiceText = Trim$(choiceText)
    position = InStr(1, choiceText, " - ", vbTextCompare)
    If position > 1 Then
        VehicleIdFromChoice = Trim$(Left$(choiceText, position - 1))
    Else
        VehicleIdFromChoice = choiceText
    End If
End Function

Private Function IsAllowedCheckoutStatus(ByVal statusText As String) As Boolean
    Select Case UCase$(Trim$(statusText))
        Case "RESERVED", "SCHEDULED", "CHECKED OUT", "RETURNED", "CANCELLED"
            IsAllowedCheckoutStatus = True
    End Select
End Function

Private Function ReservationConflictText(ByVal reservationTable As ListObject, ByVal vehicleId As String, _
    ByVal reserveDate As Date, ByVal requestedStart As Date, ByVal requestedEnd As Date) As String
    Dim i As Long
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
    For i = 1 To reservationTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(reservationTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = TableValue(reservationTable, i, "Date")
            If IsDate(rowDate) Then
                If DateOnly(rowDate) = DateOnly(reserveDate) Then
                    statusText = Trim$(CStr(TableValue(reservationTable, i, "Status")))
                    If StrComp(statusText, "Cancelled", vbTextCompare) <> 0 Then
                        hasConflict = False
                        If IsShopReservationStatus(statusText) Then
                            hasConflict = True
                        Else
                            existingStartValue = TableValue(reservationTable, i, "Start Time")
                            existingEndValue = TableValue(reservationTable, i, "End Time")
                            If IsTimeEntry(existingStartValue) And IsTimeEntry(existingEndValue) Then
                                existingStart = TimeEntryAsDate(existingStartValue)
                                existingEnd = TimeEntryAsDate(existingEndValue)
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
                            ReservationConflictText = Format(DateOnly(rowDate), "m/d/yyyy")
                            If IsShopReservationStatus(statusText) Then
                                ReservationConflictText = ReservationConflictText & " - " & ShopReservationLabel(statusText, _
                                    TableValue(reservationTable, i, "Purpose"), TableValue(reservationTable, i, "Driver Name"))
                            Else
                                driverText = Trim$(CStr(TableValue(reservationTable, i, "Driver Name")))
                                purposeText = Trim$(CStr(TableValue(reservationTable, i, "Purpose")))
                                startText = ReservationDisplayTime(TableValue(reservationTable, i, "Start Time"))
                                endText = ReservationDisplayTime(TableValue(reservationTable, i, "End Time"))
                                If driverText <> "" Then ReservationConflictText = ReservationConflictText & " by " & driverText
                                If startText <> "" Or endText <> "" Then ReservationConflictText = ReservationConflictText & " (" & startText & " - " & endText & ")"
                                If purposeText <> "" Then ReservationConflictText = ReservationConflictText & " - " & purposeText
                            End If
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Function IsShopReservationStatus(ByVal statusText As String) As Boolean
    statusText = UCase$(Trim$(statusText))
    If statusText = "MAINTENANCE" Or statusText = "SHOP" Or statusText = "IN SHOP" Then
        IsShopReservationStatus = True
    ElseIf Left$(statusText, 6) = "SHOP -" Then
        IsShopReservationStatus = True
    End If
End Function

Private Function ShopReservationLabel(ByVal statusText As Variant, ByVal purposeText As Variant, ByVal driverText As Variant) As String
    Dim result As String

    result = Trim$(CStr(statusText))
    If result = "" Or Not IsShopReservationStatus(result) Then result = Trim$(CStr(purposeText))
    If result = "" Then result = Trim$(CStr(driverText))
    If result = "" Then result = "Shop"
    If UCase$(result) = "MAINTENANCE" Or UCase$(result) = "SHOP" Or UCase$(result) = "IN SHOP" Then result = "Shop"
    ShopReservationLabel = result
End Function

Private Function ReservationDisplayTime(ByVal timeValue As Variant) As String
    Dim textValue As String

    If IsError(timeValue) Then Exit Function
    If IsDate(timeValue) Or IsNumeric(timeValue) Then
        ReservationDisplayTime = Format(TimeEntryAsDate(timeValue), "h:mm AM/PM")
    Else
        textValue = Trim$(CStr(timeValue))
        If IsDate(textValue) Then
            ReservationDisplayTime = Format(CDate(textValue), "h:mm AM/PM")
        Else
            ReservationDisplayTime = textValue
        End If
    End If
End Function

Private Function IsAllowedShopLocation(ByVal shopLocation As String) As Boolean
    Select Case UCase$(Trim$(shopLocation))
        Case "SHOP - CENTRAL GARAGE", "SHOP - NORTH GARAGE", "SHOP - SOUTH GARAGE"
            IsAllowedShopLocation = True
    End Select
End Function

Private Function ShopStatusText(ByVal vehicleTable As ListObject, ByVal vehicleId As String) As String
    Dim shopLocation As String

    EnsureVehicleSupportColumns vehicleTable
    shopLocation = Trim$(CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Shop Location")))
    If shopLocation = "" Then
        ShopStatusText = "Shop"
    ElseIf InStr(1, shopLocation, "Shop - ", vbTextCompare) = 1 Then
        ShopStatusText = shopLocation
    Else
        ShopStatusText = "Shop - " & shopLocation
    End If
End Function

Private Function AddShopReservationHistory(ByVal vehicleId As String, ByVal startDate As Date, ByVal endDate As Date, ByVal shopLocation As String) As Long
    Dim reservationTable As ListObject
    Dim lr As ListRow
    Dim currentDate As Date
    Dim existingKeys As Object
    Dim reservationValues As Variant
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim rowVehicleId As String
    Dim rowStatus As String
    Dim dateIndex As Long
    Dim vehicleIdIndex As Long
    Dim statusIndex As Long
    Dim historyKey As String
    Dim nextReservationId As Long

    If endDate < startDate Then Exit Function
    Set reservationTable = GetTable("tblReservations")
    Set existingKeys = CreateObject("Scripting.Dictionary")
    existingKeys.CompareMode = vbTextCompare

    If Not reservationTable.DataBodyRange Is Nothing Then
        reservationValues = reservationTable.DataBodyRange.Value2
        dateIndex = reservationTable.ListColumns("Date").index
        vehicleIdIndex = reservationTable.ListColumns("Vehicle ID").index
        statusIndex = reservationTable.ListColumns("Status").index
        For rowIndex = 1 To UBound(reservationValues, 1)
            rowVehicleId = Trim$(CStr(reservationValues(rowIndex, vehicleIdIndex)))
            If StrComp(rowVehicleId, vehicleId, vbTextCompare) = 0 Then
                rowDate = reservationValues(rowIndex, dateIndex)
                rowStatus = Trim$(CStr(reservationValues(rowIndex, statusIndex)))
                If (IsDate(rowDate) Or IsNumeric(rowDate)) And IsShopReservationStatus(rowStatus) Then
                    existingKeys(MonthlyShopStatusKey(DateOnly(rowDate), vehicleId)) = True
                End If
            End If
        Next rowIndex
    End If

    nextReservationId = NextNumber(reservationTable, "Reservation ID")
    currentDate = startDate
    Do While currentDate <= endDate
        historyKey = MonthlyShopStatusKey(currentDate, vehicleId)
        If Not existingKeys.Exists(historyKey) Then
            Set lr = reservationTable.ListRows.Add
            SetRowValue reservationTable, lr, "Reservation ID", nextReservationId
            SetRowValue reservationTable, lr, "Date", currentDate
            SetRowValue reservationTable, lr, "Vehicle ID", vehicleId
            SetRowValue reservationTable, lr, "Driver Name", ""
            SetRowValue reservationTable, lr, "Start Time", ""
            SetRowValue reservationTable, lr, "End Time", ""
            SetRowValue reservationTable, lr, "Purpose", ""
            SetRowValue reservationTable, lr, "Status", shopLocation
            SetRowValue reservationTable, lr, "Notes", "Shop status history"
            existingKeys(historyKey) = True
            nextReservationId = nextReservationId + 1
            AddShopReservationHistory = AddShopReservationHistory + 1
        End If
        currentDate = DateAdd("d", 1, currentDate)
    Loop
End Function

Private Function ShopReservationHistoryComplete(ByVal vehicleId As String, ByVal startDate As Date, ByVal endDate As Date) As Boolean
    Dim reservationTable As ListObject
    Dim existingKeys As Object
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim statusText As String
    Dim rowVehicleId As String
    Dim currentDate As Date

    If endDate < startDate Then
        ShopReservationHistoryComplete = True
        Exit Function
    End If

    Set reservationTable = GetTable("tblReservations")
    If reservationTable.DataBodyRange Is Nothing Then Exit Function
    Set existingKeys = CreateObject("Scripting.Dictionary")
    existingKeys.CompareMode = vbTextCompare
    For rowIndex = 1 To reservationTable.ListRows.Count
        rowVehicleId = Trim$(CStr(TableValue(reservationTable, rowIndex, "Vehicle ID")))
        If StrComp(rowVehicleId, vehicleId, vbTextCompare) = 0 Then
            rowDate = TableValue(reservationTable, rowIndex, "Date")
            statusText = Trim$(CStr(TableValue(reservationTable, rowIndex, "Status")))
            If (IsDate(rowDate) Or IsNumeric(rowDate)) And IsShopReservationStatus(statusText) Then
                existingKeys(MonthlyShopStatusKey(DateOnly(rowDate), vehicleId)) = True
            End If
        End If
    Next rowIndex

    currentDate = DateOnly(startDate)
    Do While currentDate <= DateOnly(endDate)
        If Not existingKeys.Exists(MonthlyShopStatusKey(currentDate, vehicleId)) Then Exit Function
        currentDate = DateAdd("d", 1, currentDate)
    Loop
    ShopReservationHistoryComplete = True
End Function

Private Function DeleteShopReservationHistoryFromDate(ByVal vehicleId As String, ByVal firstDate As Date) As Long
    Dim reservationTable As ListObject
    Dim reservationValues As Variant
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim rowVehicleId As String
    Dim statusText As String
    Dim dateIndex As Long
    Dim vehicleIdIndex As Long
    Dim statusIndex As Long

    Set reservationTable = GetTable("tblReservations")
    If reservationTable.DataBodyRange Is Nothing Then Exit Function

    reservationValues = reservationTable.DataBodyRange.Value2
    dateIndex = reservationTable.ListColumns("Date").index
    vehicleIdIndex = reservationTable.ListColumns("Vehicle ID").index
    statusIndex = reservationTable.ListColumns("Status").index
    For rowIndex = UBound(reservationValues, 1) To 1 Step -1
        rowVehicleId = Trim$(CStr(reservationValues(rowIndex, vehicleIdIndex)))
        If StrComp(rowVehicleId, vehicleId, vbTextCompare) = 0 Then
            rowDate = reservationValues(rowIndex, dateIndex)
            statusText = Trim$(CStr(reservationValues(rowIndex, statusIndex)))
            If (IsDate(rowDate) Or IsNumeric(rowDate)) And IsShopReservationStatus(statusText) Then
                If DateOnly(rowDate) >= DateOnly(firstDate) Then
                    reservationTable.ListRows(rowIndex).Delete
                    DeleteShopReservationHistoryFromDate = DeleteShopReservationHistoryFromDate + 1
                End If
            End If
        End If
    Next rowIndex
End Function

Public Function FleetMaterializeActiveShopHistory(Optional ByVal showMessage As Boolean = False) As Long
    On Error GoTo CleanFail

    Dim vehicleTable As ListObject
    Dim vehicleIndex As Long
    Dim vehicleId As String
    Dim vehicleStatus As String
    Dim shopStartValue As Variant
    Dim shopStart As Date
    Dim shopLocation As String

    Set vehicleTable = GetTable("tblVehicles")
    EnsureVehicleSupportColumns vehicleTable
    For vehicleIndex = 1 To vehicleTable.ListRows.Count
        vehicleStatus = Trim$(CStr(TableValue(vehicleTable, vehicleIndex, "Status")))
        If StrComp(vehicleStatus, "Maintenance", vbTextCompare) = 0 Then
            vehicleId = Trim$(CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID")))
            shopStartValue = TableValue(vehicleTable, vehicleIndex, "Shop Start Date")
            If IsDate(shopStartValue) Then
                shopStart = DateOnly(shopStartValue)
            Else
                shopStart = Date
                SetRelatedValue vehicleTable, "Vehicle ID", vehicleId, "Shop Start Date", shopStart
            End If
            shopLocation = ShopStatusText(vehicleTable, vehicleId)
            FleetMaterializeActiveShopHistory = FleetMaterializeActiveShopHistory + _
                AddShopReservationHistory(vehicleId, shopStart, Date, shopLocation)
        End If
    Next vehicleIndex

    If FleetMaterializeActiveShopHistory > 0 Then MarkWeeklyReservationsDirty
    If showMessage Then
        MsgBox CStr(FleetMaterializeActiveShopHistory) & " new shop-day entries saved.", vbInformation, APP_TITLE
    End If
    Exit Function

CleanFail:
    If showMessage Then
        MsgBox "Could not save active shop history: " & Err.Description, vbExclamation, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Function

Public Function FleetRecordShopHistoryPeriod(ByVal vehicleId As Variant, ByVal startDateValue As Variant, _
    ByVal endDateValue As Variant, ByVal shopLocation As Variant) As String

    On Error GoTo CleanFail
    Dim stage As String
    Dim startDate As Date
    Dim endDate As Date
    Dim addedCount As Long

    stage = "validate dates"
    If Not IsDate(startDateValue) Or Not IsDate(endDateValue) Then Err.Raise vbObjectError + 182, APP_TITLE, "Valid shop dates are required."
    startDate = DateOnly(CDate(startDateValue))
    endDate = DateOnly(CDate(endDateValue))
    If endDate < startDate Then Err.Raise vbObjectError + 183, APP_TITLE, "Shop end date cannot be before the start date."

    stage = "write history"
    addedCount = AddShopReservationHistory(Trim$(CStr(vehicleId)), startDate, endDate, Trim$(CStr(shopLocation)))
    stage = "verify history"
    If Not ShopReservationHistoryComplete(Trim$(CStr(vehicleId)), startDate, endDate) Then
        Err.Raise vbObjectError + 184, APP_TITLE, "Shop history could not be verified after adding " & CStr(addedCount) & " day(s)."
    End If
    MarkWeeklyReservationsDirty
    FleetRecordShopHistoryPeriod = "OK|" & CStr(addedCount)
    Exit Function

CleanFail:
    FleetRecordShopHistoryPeriod = "ERROR|" & stage & "|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function VehicleInShopOnDate(ByVal vehicleTable As ListObject, ByVal vehicleId As String, ByVal targetDate As Date) As Boolean
    Dim vehicleStatus As String
    Dim shopStart As Variant

    EnsureVehicleSupportColumns vehicleTable
    vehicleStatus = Trim$(CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Status")))
    If StrComp(vehicleStatus, "Maintenance", vbTextCompare) <> 0 Then Exit Function

    shopStart = GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Shop Start Date")
    If IsDate(shopStart) Then
        VehicleInShopOnDate = (DateOnly(targetDate) >= DateOnly(shopStart))
    Else
        VehicleInShopOnDate = True
    End If
End Function

Private Sub StyleWeeklyShopCell(ByVal targetCell As Range)
    targetCell.Interior.Color = RGB(242, 220, 219)
    targetCell.Font.Color = RGB(255, 0, 0)
    targetCell.Font.Bold = True
    targetCell.Font.Italic = True
    targetCell.HorizontalAlignment = xlCenter
    targetCell.VerticalAlignment = xlCenter
End Sub

Private Sub ClearDashboardSummary(ByVal dashWs As Worksheet)
    dashWs.Range("A19:H24").ClearContents
    dashWs.Range("A20:H24").Interior.Color = RGB(255, 255, 255)
    dashWs.Range("A26:H28").ClearContents
    dashWs.Range("A26:H28").Interior.Color = RGB(255, 255, 255)
End Sub

Private Sub WriteSummaryHeaders(ByVal headerRange As Range)
    headerRange.value = Array("Vehicle", "Last Mileage", "Checkout Status (This Week)", "Month Miles", "Trip Days", "Trips", "Scheduled Hours", "FYTD Miles")
    headerRange.Interior.Color = RGB(0, 83, 143)
    headerRange.Font.Color = RGB(255, 255, 255)
    headerRange.Font.Bold = True
    headerRange.HorizontalAlignment = xlCenter
    headerRange.WrapText = True
    ApplyBlackGrid headerRange
End Sub

Private Sub WriteVehicleMonthlyRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal vehicleName As String, ByVal lastOdometer As Variant, ByVal checkoutStatus As String, ByVal monthMiles As Double, ByVal tripDays As Long, ByVal tripCount As Long, ByVal scheduledHours As Double, ByVal cumulativeMiles As Double)
    ws.Cells(rowNumber, 1).value = vehicleName
    If IsEmpty(lastOdometer) Or CStr(lastOdometer) = "" Then
        ws.Cells(rowNumber, 2).value = ""
    Else
        ws.Cells(rowNumber, 2).value = lastOdometer
    End If
    ws.Cells(rowNumber, 3).value = checkoutStatus
    ws.Cells(rowNumber, 4).value = monthMiles
    ws.Cells(rowNumber, 5).value = tripDays
    ws.Cells(rowNumber, 6).value = tripCount
    ws.Cells(rowNumber, 7).value = scheduledHours
    ws.Cells(rowNumber, 8).value = cumulativeMiles

    ws.Range(ws.Cells(rowNumber, 2), ws.Cells(rowNumber, 6)).numberFormat = "#,##0"
    ws.Cells(rowNumber, 7).numberFormat = "0.0"
    ws.Cells(rowNumber, 8).numberFormat = "#,##0"
    ApplyBlackGrid ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 8))
End Sub

Private Sub StyleCheckoutStatus(ByVal statusCell As Range, ByVal statusText As String)
    Dim lowered As String
    lowered = LCase$(statusText)

    statusCell.Font.Bold = True
    statusCell.WrapText = True
    statusCell.VerticalAlignment = xlCenter
    If InStr(lowered, "available") > 0 Then
        statusCell.Interior.Color = RGB(226, 239, 218)
        statusCell.Font.Color = RGB(0, 97, 0)
    ElseIf InStr(lowered, "checked out") > 0 Then
        statusCell.Interior.Color = RGB(255, 235, 156)
        statusCell.Font.Color = RGB(156, 87, 0)
    ElseIf InStr(lowered, "maintenance") > 0 Or InStr(lowered, "shop") > 0 Then
        statusCell.Interior.Color = RGB(242, 220, 219)
        statusCell.Font.Color = RGB(156, 0, 6)
    Else
        statusCell.Interior.Color = RGB(217, 225, 242)
        statusCell.Font.Color = RGB(0, 83, 143)
    End If
End Sub

Private Sub StyleMonthlyMiles(ByVal milesCell As Range, ByVal monthMiles As Double)
    milesCell.Font.Bold = True
    If monthMiles >= 500 Then
        milesCell.Interior.Color = RGB(226, 239, 218)
        milesCell.Font.Color = RGB(0, 97, 0)
    Else
        milesCell.Interior.Color = RGB(255, 235, 156)
        milesCell.Font.Color = RGB(156, 87, 0)
    End If
End Sub

Private Function LastOdometerForVehicleThroughDate(ByVal mileageTable As ListObject, ByVal vehicleId As String, ByVal endDate As Date) As Variant
    Dim i As Long
    Dim rowDate As Variant
    Dim odometer As Variant
    Dim latestDate As Date
    Dim bestOdometer As Long
    Dim hasOdometer As Boolean

    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To mileageTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(mileageTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = TableValue(mileageTable, i, "Date")
            odometer = TableValue(mileageTable, i, "Odometer End")
            If IsDate(rowDate) And IsNumeric(odometer) Then
                If DateOnly(rowDate) <= endDate Then
                    If Not hasOdometer Then
                        latestDate = DateOnly(rowDate)
                        bestOdometer = CLng(odometer)
                        hasOdometer = True
                    ElseIf DateOnly(rowDate) > latestDate Then
                        latestDate = DateOnly(rowDate)
                        bestOdometer = CLng(odometer)
                    ElseIf DateOnly(rowDate) = latestDate And CLng(odometer) > bestOdometer Then
                        bestOdometer = CLng(odometer)
                    End If
                End If
            End If
        End If
    Next i
    If hasOdometer Then LastOdometerForVehicleThroughDate = bestOdometer
End Function

Private Function TripDaysForVehicleMonth(ByVal mileageTable As ListObject, ByVal vehicleId As String, ByVal monthStart As Date, ByVal monthEnd As Date) As Long
    Dim seenDates As Object
    Dim i As Long
    Dim rowDate As Variant
    Dim key As String

    Set seenDates = CreateObject("Scripting.Dictionary")
    If mileageTable.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To mileageTable.ListRows.Count
        If IsMileageMatch(mileageTable, i, monthStart, monthEnd, vehicleId) Then
            rowDate = TableValue(mileageTable, i, "Date")
            If IsDate(rowDate) Or IsNumeric(rowDate) Then
                key = CStr(CLng(DateOnly(rowDate)))
                If Not seenDates.Exists(key) Then seenDates.Add key, True
            End If
        End If
    Next i
    TripDaysForVehicleMonth = seenDates.Count
End Function

Private Function CumulativeMilesForVehicleThroughDate(ByVal mileageTable As ListObject, ByVal vehicleId As String, ByVal endDate As Date) As Double
    Dim i As Long
    Dim rowDate As Variant
    Dim miles As Variant

    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To mileageTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(mileageTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = TableValue(mileageTable, i, "Date")
            If IsDate(rowDate) Then
                If DateOnly(rowDate) <= endDate Then
                    miles = TableValue(mileageTable, i, "Miles Driven")
                    If IsNumeric(miles) Then CumulativeMilesForVehicleThroughDate = CumulativeMilesForVehicleThroughDate + CDbl(miles)
                End If
            End If
        End If
    Next i
End Function

Private Function ScheduledHoursForVehicleMonth(ByVal reservationTable As ListObject, ByVal vehicleId As String, ByVal monthStart As Date, ByVal monthEnd As Date) As Double
    Dim reservedDates As Object
    Dim i As Long
    Dim rowDate As Variant
    Dim statusText As String
    Dim dateKey As String

    If reservationTable.DataBodyRange Is Nothing Then Exit Function
    Set reservedDates = CreateObject("Scripting.Dictionary")
    For i = 1 To reservationTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(reservationTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = TableValue(reservationTable, i, "Date")
            If IsDate(rowDate) Then
                If DateOnly(rowDate) >= monthStart And DateOnly(rowDate) <= monthEnd Then
                    statusText = Trim$(CStr(TableValue(reservationTable, i, "Status")))
                    If IsAllowedCheckoutStatus(statusText) And StrComp(statusText, "Cancelled", vbTextCompare) <> 0 Then
                        dateKey = CStr(CLng(DateOnly(rowDate)))
                        If Not reservedDates.Exists(dateKey) Then reservedDates.Add dateKey, True
                    End If
                End If
            End If
        End If
    Next i
    ScheduledHoursForVehicleMonth = reservedDates.Count * 8#
End Function

Private Function CheckoutStatusForVehicle(ByVal vehicleTable As ListObject, ByVal reservationTable As ListObject, ByVal vehicleId As String, ByVal targetDate As Date) As String
    Dim vehicleStatus As String
    Dim weekText As String
    Dim weekStart As Date
    Dim weekEnd As Date

    CheckoutStatusForVehicle = "Available"
    EnsureVehicleSupportColumns vehicleTable
    weekStart = StartOfWeek(targetDate)
    weekEnd = DateAdd("d", 6, weekStart)

    vehicleStatus = Trim$(CStr(GetRelatedValue(vehicleTable, "Vehicle ID", vehicleId, "Status")))
    If StrComp(vehicleStatus, "Maintenance", vbTextCompare) = 0 Then
        CheckoutStatusForVehicle = ShopStatusText(vehicleTable, vehicleId)
        Exit Function
    End If

    weekText = CheckoutStatusWeekText(reservationTable, vehicleId, weekStart, weekEnd)
    If weekText <> "" Then
        CheckoutStatusForVehicle = weekText
        Exit Function
    End If

End Function

Private Function CheckoutStatusWeekText(ByVal reservationTable As ListObject, ByVal vehicleId As String, ByVal weekStart As Date, ByVal weekEnd As Date) As String
    Dim grouped As Object
    Dim orderedKeys As Collection
    Dim i As Long
    Dim rowDate As Variant
    Dim statusText As String
    Dim driverText As String
    Dim purposeText As String
    Dim entryText As String
    Dim key As String
    Dim item As Variant
    Dim valueText As String
    Dim splitPosition As Long
    Dim result As String

    Set grouped = CreateObject("Scripting.Dictionary")
    Set orderedKeys = New Collection
    If reservationTable.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To reservationTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(reservationTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = TableValue(reservationTable, i, "Date")
            If IsDate(rowDate) Then
                If DateOnly(rowDate) >= weekStart And DateOnly(rowDate) <= weekEnd Then
                    statusText = Trim$(CStr(TableValue(reservationTable, i, "Status")))
                    If StrComp(statusText, "Cancelled", vbTextCompare) <> 0 And _
                        Not IsShopReservationStatus(statusText) Then
                        driverText = Trim$(CStr(TableValue(reservationTable, i, "Driver Name")))
                        purposeText = Trim$(CStr(TableValue(reservationTable, i, "Purpose")))
                        entryText = DashboardStatusEntryText(statusText, driverText, purposeText)
                        If entryText <> "" Then
                            key = UCase$(entryText)
                            If Not grouped.Exists(key) Then
                                grouped.Add key, entryText & "|" & CStr(CLng(DateOnly(rowDate)))
                                orderedKeys.Add key
                            Else
                                grouped(key) = grouped(key) & "," & CStr(CLng(DateOnly(rowDate)))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    For Each item In orderedKeys
        valueText = CStr(grouped(CStr(item)))
        splitPosition = InStr(1, valueText, "|", vbBinaryCompare)
        If splitPosition > 0 Then
            If result <> "" Then result = result & vbLf
            result = result & CompactDashboardDateList(Mid$(valueText, splitPosition + 1)) & " " & Left$(valueText, splitPosition - 1)
        End If
    Next item

    CheckoutStatusWeekText = result
End Function

Private Function DashboardStatusEntryText(ByVal statusText As String, ByVal driverText As String, ByVal purposeText As String) As String
    Dim labelText As String

    If IsShopReservationStatus(statusText) Then
        DashboardStatusEntryText = ShopReservationLabel(statusText, purposeText, driverText)
        Exit Function
    End If

    If statusText = "" Or StrComp(statusText, "Scheduled", vbTextCompare) = 0 Or StrComp(statusText, "Reserved", vbTextCompare) = 0 Then
        labelText = "Reserved"
    Else
        labelText = statusText
    End If

    If Trim$(driverText) <> "" Then
        DashboardStatusEntryText = labelText & " - " & Trim$(driverText)
    ElseIf Trim$(purposeText) <> "" Then
        DashboardStatusEntryText = labelText & " - " & Trim$(purposeText)
    Else
        DashboardStatusEntryText = labelText
    End If
End Function

Private Function CompactDashboardDateList(ByVal serialList As String) As String
    Dim rawValues As Variant
    Dim dates() As Long
    Dim rawIndex As Long
    Dim dateCount As Long
    Dim i As Long
    Dim j As Long
    Dim tempValue As Long
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim result As String

    rawValues = Split(serialList, ",")
    ReDim dates(0 To UBound(rawValues))
    For rawIndex = LBound(rawValues) To UBound(rawValues)
        If IsNumeric(rawValues(rawIndex)) Then
            tempValue = CLng(rawValues(rawIndex))
            If Not DashboardDateAlreadyCollected(dates, dateCount, tempValue) Then
                dates(dateCount) = tempValue
                dateCount = dateCount + 1
            End If
        End If
    Next rawIndex

    If dateCount = 0 Then Exit Function

    For i = 0 To dateCount - 2
        For j = i + 1 To dateCount - 1
            If dates(j) < dates(i) Then
                tempValue = dates(i)
                dates(i) = dates(j)
                dates(j) = tempValue
            End If
        Next j
    Next i

    i = 0
    Do While i < dateCount
        rangeStart = dates(i)
        rangeEnd = rangeStart
        Do While i + 1 < dateCount
            If dates(i + 1) <> rangeEnd + 1 Then Exit Do
            i = i + 1
            rangeEnd = dates(i)
        Loop

        If result <> "" Then result = result & ", "
        If rangeStart = rangeEnd Then
            result = result & Format$(CDate(rangeStart), "m/d")
        Else
            result = result & Format$(CDate(rangeStart), "m/d") & "-" & Format$(CDate(rangeEnd), "m/d")
        End If
        i = i + 1
    Loop

    CompactDashboardDateList = result
End Function

Private Function DashboardDateAlreadyCollected(ByRef dates() As Long, ByVal dateCount As Long, ByVal targetDate As Long) As Boolean
    Dim index As Long

    For index = 0 To dateCount - 1
        If dates(index) = targetDate Then
            DashboardDateAlreadyCollected = True
            Exit Function
        End If
    Next index
End Function

Private Function BuildFiscalYearValidationList(ByVal mileageTable As ListObject, ByVal reservationTable As ListObject, ByRef firstFyStart As Date, ByRef lastFyStart As Date, ByRef hasFiscalDates As Boolean) As String
    Dim minMileage As Date
    Dim maxMileage As Date
    Dim hasMileage As Boolean
    Dim minReservation As Date
    Dim maxReservation As Date
    Dim hasReservation As Boolean
    Dim minDate As Date
    Dim maxDate As Date
    Dim currentFy As Date
    Dim result As String

    GetDateRange mileageTable, "Date", minMileage, maxMileage, hasMileage
    GetDateRange reservationTable, "Date", minReservation, maxReservation, hasReservation

    hasFiscalDates = hasMileage Or hasReservation
    If Not hasFiscalDates Then
        firstFyStart = FiscalYearStartForDate(Date)
        lastFyStart = firstFyStart
    Else
        If hasMileage And hasReservation Then
            minDate = minMileage
            If minReservation < minDate Then minDate = minReservation
            maxDate = maxMileage
            If maxReservation > maxDate Then maxDate = maxReservation
        ElseIf hasMileage Then
            minDate = minMileage
            maxDate = maxMileage
        Else
            minDate = minReservation
            maxDate = maxReservation
        End If

        firstFyStart = FiscalYearStartForDate(minDate)
        lastFyStart = FiscalYearStartForDate(maxDate)
    End If

    currentFy = firstFyStart
    Do While currentFy <= lastFyStart
        If result <> "" Then result = result & ","
        result = result & fiscalYearLabel(currentFy)
        currentFy = DateAdd("yyyy", 1, currentFy)
    Loop
    BuildFiscalYearValidationList = result
End Function

Private Function DashboardSelectedFiscalYearStart(ByVal dashWs As Worksheet, ByVal firstFyStart As Date, ByVal lastFyStart As Date, ByVal hasFiscalDates As Boolean) As Date
    Dim selectedText As String
    Dim candidate As Date

    selectedText = Trim$(CStr(dashWs.Range("D17").value))
    If TryFiscalYearStartFromLabel(selectedText, candidate) Then
        If Not hasFiscalDates Or (candidate >= firstFyStart And candidate <= lastFyStart) Then
            DashboardSelectedFiscalYearStart = candidate
            Exit Function
        End If
    End If

    candidate = FiscalYearStartForDate(Date)
    If hasFiscalDates Then
        If candidate < firstFyStart Or candidate > lastFyStart Then candidate = lastFyStart
    End If
    DashboardSelectedFiscalYearStart = candidate
End Function

Private Sub ApplyFiscalYearValidation(ByVal targetCell As Range, ByVal listText As String)
    On Error Resume Next
    targetCell.Validation.Delete
    On Error GoTo 0

    If Len(listText) > 0 Then
        targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listText
        targetCell.Validation.IgnoreBlank = False
        targetCell.Validation.InCellDropdown = True
    End If
End Sub

Private Function FiscalYearStartForDate(ByVal targetDate As Date) As Date
    Dim startYear As Long
    If Month(targetDate) >= 7 Then
        startYear = Year(targetDate)
    Else
        startYear = Year(targetDate) - 1
    End If
    FiscalYearStartForDate = DateSerial(startYear, 7, 1)
End Function

Private Function fiscalYearLabel(ByVal fyStart As Date) As String
    fiscalYearLabel = "FY " & CStr(Year(fyStart)) & "-" & Format$(DateAdd("yyyy", 1, fyStart), "yy")
End Function

Private Function TryFiscalYearStartFromLabel(ByVal labelText As String, ByRef fyStart As Date) As Boolean
    Dim cleaned As String
    Dim parts As Variant
    Dim startYear As Long

    cleaned = Trim$(labelText)
    If UCase$(Left$(cleaned, 2)) = "FY" Then cleaned = Trim$(Mid$(cleaned, 3))
    parts = Split(cleaned, "-")

    If UBound(parts) >= 0 Then
        If IsNumeric(Trim$(CStr(parts(0)))) Then
            startYear = CLng(Trim$(CStr(parts(0))))
            If startYear >= 1900 And startYear <= 9999 Then
                fyStart = DateSerial(startYear, 7, 1)
                TryFiscalYearStartFromLabel = True
            End If
        End If
    End If
End Function

Private Sub WriteDashboardNumber(ByVal targetCell As Range, ByVal value As Double, ByVal numberFormat As String)
    WriteDashboardValue targetCell, value, numberFormat
End Sub

Private Sub WriteDashboardValue(ByVal targetCell As Range, ByVal value As Variant, ByVal numberFormat As String)
    With targetCell.MergeArea
        .Cells(1, 1).value = value
        .numberFormat = numberFormat
    End With
End Sub

Private Function CountTableMatches(ByVal lo As ListObject, ByVal columnName As String, ByVal matchText As String) As Long
    Dim i As Long
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(lo, i, columnName))), matchText, vbTextCompare) = 0 Then CountTableMatches = CountTableMatches + 1
    Next i
End Function

Private Function MilesForVehicleDateRange(ByVal mileageTable As ListObject, ByVal vehicleId As String, ByVal startDate As Date, ByVal endDate As Date) As Double
    Dim i As Long
    Dim miles As Variant
    If mileageTable.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To mileageTable.ListRows.Count
        If IsMileageMatch(mileageTable, i, startDate, endDate, vehicleId) Then
            miles = TableValue(mileageTable, i, "Miles Driven")
            If IsNumeric(miles) Then MilesForVehicleDateRange = MilesForVehicleDateRange + CDbl(miles)
        End If
    Next i
End Function

Private Function EnsureWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Sub ClearWeeklyReservationsSheet(ByVal ws As Worksheet)
    Dim i As Long
    Dim clearRange As Range

    Set clearRange = ws.Range("A1:J40")
    ws.Rows("1:40").Hidden = False
    clearRange.UnMerge
    clearRange.Clear
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    clearRange.Font.Name = "Aptos"
    clearRange.Font.Size = 10
    clearRange.Font.Color = RGB(0, 0, 0)
    clearRange.Interior.Color = RGB(255, 255, 255)
    clearRange.VerticalAlignment = xlCenter
End Sub

Private Sub ClearReportSheet(ByVal ws As Worksheet)
    Dim i As Long

    ws.Cells.UnMerge
    ws.Cells.Clear
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Font.Name = "Aptos"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(255, 255, 255)
End Sub

Private Sub DrawMonthlyMileageBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal monthStart As Date, ByVal monthEnd As Date, ByVal vehicleTable As ListObject, ByVal mileageTable As ListObject, ByVal reservationTable As ListObject)
    Dim vehicleCount As Long
    Dim vehicleIndex As Long
    Dim col As Long
    Dim dayRow As Long
    Dim dayOffset As Long
    Dim currentDate As Date
    Dim totalRow As Long
    Dim cumulativeRow As Long
    Dim vehicleId As String
    Dim lastCol As Long
    Dim dayTrips As Long
    Dim dayEnding As Variant
    Dim shopKey As String
    Dim shopTextByKey As Object

    vehicleCount = vehicleTable.ListRows.Count
    lastCol = 2 + (vehicleCount * 2)
    Set shopTextByKey = BuildMonthlyShopStatusMap(vehicleTable, reservationTable, monthStart, monthEnd)

    ws.Cells(startRow, 1).value = "Vehicles"
    ws.Cells(startRow, 2).value = Format(monthStart, "mmmm yyyy")
    StyleMonthlyHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, lastCol))

    For vehicleIndex = 1 To vehicleCount
        col = 3 + ((vehicleIndex - 1) * 2)
        vehicleId = CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))
        ws.Range(ws.Cells(startRow, col), ws.Cells(startRow, col + 1)).Merge
        ws.Cells(startRow, col).value = MonthlyVehicleCaption(vehicleTable, vehicleIndex)
        ws.Cells(startRow, col).HorizontalAlignment = xlCenter
    Next vehicleIndex

    ws.Cells(startRow + 1, 1).value = "Day"
    ws.Cells(startRow + 1, 2).value = "Date"
    For vehicleIndex = 1 To vehicleCount
        col = 3 + ((vehicleIndex - 1) * 2)
        ws.Cells(startRow + 1, col).value = "Trip(s)"
        ws.Cells(startRow + 1, col + 1).value = "End Mileage"
    Next vehicleIndex
    StyleSubHeader ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + 1, lastCol))

    dayRow = startRow + 2
    For dayOffset = 0 To DateDiff("d", monthStart, monthEnd)
        currentDate = DateAdd("d", dayOffset, monthStart)
        ws.Cells(dayRow, 1).value = Format(currentDate, "dddd")
        ws.Cells(dayRow, 2).value = currentDate
        ws.Cells(dayRow, 2).numberFormat = "m/d/yyyy"
        If Weekday(currentDate, vbSunday) = vbSunday Or Weekday(currentDate, vbSunday) = vbSaturday Then
            ws.Range(ws.Cells(dayRow, 1), ws.Cells(dayRow, 2)).Font.Color = RGB(192, 0, 0)
        End If

        For vehicleIndex = 1 To vehicleCount
            col = 3 + ((vehicleIndex - 1) * 2)
            vehicleId = CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))
            dayTrips = TripsForDateVehicle(mileageTable, currentDate, vehicleId)
            dayEnding = EndingMileageForDateVehicle(mileageTable, currentDate, vehicleId)
            shopKey = MonthlyShopStatusKey(currentDate, vehicleId)
            If dayTrips > 0 Or (Not IsEmpty(dayEnding) And IsNumeric(dayEnding)) Then
                ResetMonthlyMileageEntryCells ws.Range(ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1))
                If dayTrips > 0 Then
                    ws.Cells(dayRow, col).value = dayTrips
                Else
                    ws.Cells(dayRow, col).ClearContents
                End If
                If Not IsEmpty(dayEnding) And IsNumeric(dayEnding) Then
                    ws.Cells(dayRow, col + 1).value = dayEnding
                Else
                    ws.Cells(dayRow, col + 1).ClearContents
                End If
                ws.Cells(dayRow, col + 1).numberFormat = "#,##0"
            ElseIf shopTextByKey.Exists(shopKey) Then
                StyleMonthlyShopPair ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1), CStr(shopTextByKey(shopKey))
            Else
                ResetMonthlyMileageEntryCells ws.Range(ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1))
                ws.Range(ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1)).ClearContents
            End If
        Next vehicleIndex
        dayRow = dayRow + 1
    Next dayOffset

    totalRow = dayRow
    cumulativeRow = dayRow + 1
    ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, 2)).Merge
    ws.Cells(totalRow, 1).value = Format(monthStart, "mmmm yyyy") & " TOTALS:"
    ws.Range(ws.Cells(cumulativeRow, 1), ws.Cells(cumulativeRow, 2)).Merge
    ws.Cells(cumulativeRow, 1).value = "CUMULATIVE MILEAGES:"
    For vehicleIndex = 1 To vehicleCount
        col = 3 + ((vehicleIndex - 1) * 2)
        vehicleId = CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))
        ws.Cells(totalRow, col).value = MonthTripsForVehicle(mileageTable, monthStart, monthEnd, vehicleId)
        ws.Cells(totalRow, col + 1).value = MonthMilesForVehicle(mileageTable, monthStart, monthEnd, vehicleId)
        ws.Cells(totalRow, col + 1).numberFormat = "#,##0"
        ws.Cells(cumulativeRow, col).value = "Ending Mileage"
        ws.Cells(cumulativeRow, col + 1).value = EndingMileageForMonthVehicle(mileageTable, monthStart, monthEnd, vehicleId)
        ws.Cells(cumulativeRow, col + 1).numberFormat = "#,##0"
    Next vehicleIndex

    StyleTotalRows ws.Range(ws.Cells(totalRow, 1), ws.Cells(cumulativeRow, lastCol))
    ApplyBlackGrid ws.Range(ws.Cells(startRow, 1), ws.Cells(cumulativeRow, lastCol))
End Sub

Private Function BuildMonthlyShopStatusMap(ByVal vehicleTable As ListObject, ByVal reservationTable As ListObject, _
    ByVal monthStart As Date, ByVal monthEnd As Date) As Object

    Dim result As Object
    Dim reservationValues As Variant
    Dim rowIndex As Long
    Dim vehicleIndex As Long
    Dim rowDate As Variant
    Dim dayDate As Date
    Dim vehicleId As String
    Dim statusText As String
    Dim vehicleStatus As String
    Dim shopStartValue As Variant
    Dim shopStart As Date
    Dim shopEnd As Date
    Dim key As String
    Dim dateIndex As Long
    Dim vehicleIdIndex As Long
    Dim statusIndex As Long
    Dim purposeIndex As Long
    Dim driverIndex As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    If Not reservationTable.DataBodyRange Is Nothing Then
        reservationValues = reservationTable.DataBodyRange.Value2
        dateIndex = reservationTable.ListColumns("Date").index
        vehicleIdIndex = reservationTable.ListColumns("Vehicle ID").index
        statusIndex = reservationTable.ListColumns("Status").index
        purposeIndex = reservationTable.ListColumns("Purpose").index
        driverIndex = reservationTable.ListColumns("Driver Name").index
        For rowIndex = 1 To UBound(reservationValues, 1)
            rowDate = reservationValues(rowIndex, dateIndex)
            If IsDate(rowDate) Or IsNumeric(rowDate) Then
                dayDate = DateOnly(rowDate)
                If dayDate >= monthStart And dayDate <= monthEnd Then
                    statusText = Trim$(CStr(reservationValues(rowIndex, statusIndex)))
                    If IsShopReservationStatus(statusText) Then
                        vehicleId = Trim$(CStr(reservationValues(rowIndex, vehicleIdIndex)))
                        key = MonthlyShopStatusKey(dayDate, vehicleId)
                        result(key) = ShopReservationLabel(statusText, reservationValues(rowIndex, purposeIndex), _
                            reservationValues(rowIndex, driverIndex))
                    End If
                End If
            End If
        Next rowIndex
    End If

    EnsureVehicleSupportColumns vehicleTable
    shopEnd = monthEnd
    For vehicleIndex = 1 To vehicleTable.ListRows.Count
        vehicleStatus = Trim$(CStr(TableValue(vehicleTable, vehicleIndex, "Status")))
        If StrComp(vehicleStatus, "Maintenance", vbTextCompare) = 0 Then
            vehicleId = Trim$(CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID")))
            shopStartValue = TableValue(vehicleTable, vehicleIndex, "Shop Start Date")
            If IsDate(shopStartValue) Then
                shopStart = DateOnly(shopStartValue)
            Else
                shopStart = Date
            End If
            If shopStart < monthStart Then shopStart = monthStart
            If shopStart <= shopEnd Then
                dayDate = shopStart
                Do While dayDate <= shopEnd
                    result(MonthlyShopStatusKey(dayDate, vehicleId)) = ShopStatusText(vehicleTable, vehicleId)
                    dayDate = DateAdd("d", 1, dayDate)
                Loop
            End If
        End If
    Next vehicleIndex

    Set BuildMonthlyShopStatusMap = result
End Function

Private Function MonthlyShopStatusKey(ByVal targetDate As Date, ByVal vehicleId As String) As String
    MonthlyShopStatusKey = CStr(CLng(DateOnly(targetDate))) & "|" & UCase$(Trim$(vehicleId))
End Function

Private Function MonthlyShopLocationLabel(ByVal shopStatus As String) As String
    shopStatus = Trim$(shopStatus)
    If InStr(1, shopStatus, "Shop - ", vbTextCompare) = 1 Then
        MonthlyShopLocationLabel = Trim$(Mid$(shopStatus, 8))
    ElseIf shopStatus = "" Or UCase$(shopStatus) = "SHOP" Or UCase$(shopStatus) = "IN SHOP" Or UCase$(shopStatus) = "MAINTENANCE" Then
        MonthlyShopLocationLabel = "Unavailable"
    Else
        MonthlyShopLocationLabel = shopStatus
    End If
End Function

Private Sub StyleMonthlyShopPair(ByVal tripCell As Range, ByVal mileageCell As Range, ByVal shopStatus As String)
    With tripCell
        .value = "IN SHOP"
        .numberFormat = "@"
    End With
    With mileageCell
        .value = MonthlyShopLocationLabel(shopStatus)
        .numberFormat = "@"
    End With
    With Union(tripCell, mileageCell)
        .Interior.Color = RGB(242, 220, 219)
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
        .Font.Italic = True
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
End Sub

Private Sub ResetMonthlyMileageEntryCells(ByVal targetRange As Range)
    With targetRange
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Font.Italic = False
        .Font.Size = 10
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
End Sub

Private Sub DrawWeeklyReservationControls(ByVal ws As Worksheet, ByVal selectedDate As Date, ByVal weekStart As Date)
    Dim yearList As String

    ws.Range("A1:J1").Merge
    ws.Range("A1").value = "WEEKLY RESERVATIONS"
    ws.Range("A1").Interior.Color = RGB(0, 83, 143)
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").HorizontalAlignment = xlCenter

    ws.Range("A2:J5").Interior.Color = RGB(245, 248, 250)
    ws.Range("A2:J5").Borders.LineStyle = xlContinuous
    ws.Range("A2:J5").Borders.Color = RGB(214, 222, 228)
    ws.Range("A2:J2").Merge
    ws.Range("A2").value = "Select week"
    ws.Range("A2").Interior.Color = RGB(255, 181, 17)
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Color = RGB(0, 0, 0)
    ws.Range("A2").HorizontalAlignment = xlLeft

    ws.Range("A3:B3").Merge
    ws.Range("A3").value = "Month"
    ws.Range("C3").value = Format(selectedDate, "mmmm")
    ws.Range("E3").value = "Year"
    ws.Range("F3").value = Year(selectedDate)
    ws.Range("A4:B4").Merge
    ws.Range("A4").value = "Date"
    ws.Range("C4:D4").Merge
    ws.Range("C4").value = selectedDate
    ws.Range("C4").numberFormat = "m/d/yyyy"
    ws.Range("A5:B5").Merge
    ws.Range("A5").value = "Showing"
    ws.Range("C5:G5").Merge
    ws.Range("C5").value = Format(weekStart, "mmmm d") & " - " & Format(DateAdd("d", 6, weekStart), "mmmm d, yyyy")

    ws.Range("A3:B5,E3").Font.Bold = True
    ws.Range("A3:B5,E3").Font.Color = RGB(0, 83, 143)
    ws.Range("A3:B5,E3").HorizontalAlignment = xlRight
    ws.Range("C3,F3,C4:D4").Interior.Color = RGB(255, 255, 255)
    ws.Range("C3,F3,C4:D4").Borders.LineStyle = xlContinuous
    ws.Range("C3,F3,C4:D4").Borders.Color = RGB(0, 83, 143)
    ws.Range("C3,F3,C4:D4").Borders.Weight = xlThin
    ws.Range("C3,F3,C4:D4").Font.Bold = True
    ws.Range("C3,F3,C4:D4").Font.Size = 12
    ws.Range("C3,F3,C4:D4").HorizontalAlignment = xlCenter
    ws.Range("D3,G3").Interior.Color = RGB(245, 248, 250)
    ws.Range("C5").Font.Bold = True
    ws.Range("C5").Font.Color = RGB(0, 83, 143)
    ws.Range("C5:G5").Interior.Color = RGB(232, 244, 252)
    ws.Range("C5:G5").HorizontalAlignment = xlLeft
    ws.Range("A2:J5").VerticalAlignment = xlCenter

    yearList = YearValidationList(Year(selectedDate) - 4, Year(selectedDate) + 4)
    On Error Resume Next
    ws.Range("C3").Validation.Delete
    ws.Range("C3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="January,February,March,April,May,June,July,August,September,October,November,December"
    ws.Range("F3").Validation.Delete
    ws.Range("F3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=yearList
    ws.Range("C4").Validation.Delete
    ws.Range("C4").Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1/1/2020", Formula2:="12/31/2035"
    ws.Range("C4").Validation.IgnoreBlank = False
    On Error GoTo 0

    ws.Range("I3").value = "Month"
    ws.Range("I3").Font.Bold = True
    ws.Range("I3").Font.Color = RGB(0, 83, 143)
    ws.Range("I3").HorizontalAlignment = xlCenter

    AddReportButton ws, "E4:E4", "CAL", "FleetToggleReservationCalendar", RGB(255, 121, 0)
    AddReportButton ws, "H3:H3", "<", "FleetPreviousReservationMonth", RGB(87, 87, 87)
    AddReportButton ws, "J3:J3", ">", "FleetNextReservationMonth", RGB(0, 83, 143)
    AddReportButton ws, "F4:G4", "PREVIOUS WEEK", "FleetPreviousReservationWeek", RGB(87, 87, 87)
    AddReportButton ws, "H4:J4", "NEXT WEEK", "FleetNextReservationWeek", RGB(0, 83, 143)
    AddReportButton ws, "H5:J5", "DASHBOARD", "FleetShowDashboard", RGB(255, 121, 0)

    ws.Rows("6:13").Hidden = True
End Sub

Private Sub EnsureWeeklyDashboardNavigationButton(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim caption As String

    On Error Resume Next
    For Each shp In ws.Shapes
        caption = vbNullString
        caption = Trim$(UCase$(shp.TextFrame.Characters.Text))
        If caption = "DASHBOARD" Then
            shp.Fill.ForeColor.RGB = RGB(255, 121, 0)
            shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            shp.TextFrame.Characters.Font.Bold = True
            shp.OnAction = "FleetShowDashboard"
        End If
    Next shp
    On Error GoTo 0
End Sub

Private Sub DrawReservationCalendarPopup(ByVal ws As Worksheet, ByVal calendarSeedDate As Date)
    Dim calendarMonth As Date
    Dim selectedDate As Date
    Dim weekStart As Date
    Dim firstGridDate As Date
    Dim currentDate As Date
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim dayOffset As Long
    Dim dayNames As Variant

    selectedDate = GetSelectedReservationDate(ws)
    weekStart = StartOfWeek(selectedDate)
    calendarMonth = DateSerial(Year(calendarSeedDate), Month(calendarSeedDate), 1)
    firstGridDate = DateAdd("d", 1 - Weekday(calendarMonth, vbSunday), calendarMonth)
    dayNames = Array("SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT")

    ws.Rows("6:13").Hidden = False
    ws.Range("A6:G13").UnMerge
    ws.Range("A6:G13").Clear
    ws.Range("A6:G13").Borders.LineStyle = xlContinuous
    ws.Range("A6:G13").Borders.Color = RGB(214, 222, 228)
    DeleteShapeIfExists ws, "btnWeeklyCalendarPrevMonth"
    DeleteShapeIfExists ws, "btnWeeklyCalendarNextMonth"

    ws.Range("A6").ClearContents
    ws.Range("B6:F6").Merge
    ws.Range("B6").numberFormat = "@"
    ws.Range("B6").value = Format(calendarMonth, "mmmm yyyy")
    ws.Range("G6").ClearContents
    ws.Range("A6:G6").Interior.Color = RGB(0, 83, 143)
    ws.Range("A6:G6").Font.Color = RGB(255, 255, 255)
    ws.Range("A6:G6").Font.Bold = True
    ws.Range("A6:G6").HorizontalAlignment = xlCenter
    AddReportButton ws, "A6:A6", "<", "FleetCalendarPreviousMonth", RGB(0, 83, 143), "btnWeeklyCalendarPrevMonth"
    AddReportButton ws, "G6:G6", ">", "FleetCalendarNextMonth", RGB(0, 83, 143), "btnWeeklyCalendarNextMonth"

    For colIndex = 1 To 7
        ws.Cells(7, colIndex).value = dayNames(colIndex - 1)
        ws.Cells(7, colIndex).Interior.Color = RGB(232, 244, 252)
        ws.Cells(7, colIndex).Font.Color = RGB(0, 83, 143)
        ws.Cells(7, colIndex).Font.Bold = True
        ws.Cells(7, colIndex).HorizontalAlignment = xlCenter
    Next colIndex

    dayOffset = 0
    For rowIndex = 8 To 13
        For colIndex = 1 To 7
            currentDate = DateAdd("d", dayOffset, firstGridDate)
            With ws.Cells(rowIndex, colIndex)
                .value = currentDate
                .numberFormat = "d"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Color = RGB(0, 0, 0)
                .Interior.Color = RGB(255, 255, 255)
                If Month(currentDate) <> Month(calendarMonth) Then
                    .Interior.Color = RGB(242, 242, 242)
                    .Font.Color = RGB(128, 128, 128)
                End If
                If currentDate >= weekStart And currentDate <= DateAdd("d", 6, weekStart) Then .Interior.Color = RGB(255, 242, 204)
                If DateOnly(currentDate) = DateOnly(selectedDate) Then
                    .Interior.Color = RGB(0, 83, 143)
                    .Font.Color = RGB(255, 255, 255)
                End If
                If DateOnly(currentDate) = Date Then
                    If DateOnly(currentDate) <> DateOnly(selectedDate) Then .Interior.Color = RGB(255, 230, 153)
                    .Borders.LineStyle = xlContinuous
                    .Borders.Color = RGB(255, 121, 0)
                    .Borders.Weight = xlMedium
                End If
            End With
            dayOffset = dayOffset + 1
        Next colIndex
    Next rowIndex
    ws.Rows("6:13").RowHeight = 22
End Sub

Private Function CalendarMonthFromHeader(ByVal ws As Worksheet) As Date
    Dim headerText As String
    headerText = Trim$(CStr(ws.Range("B6").value))
    If headerText <> "" And IsDate("1 " & headerText) Then
        CalendarMonthFromHeader = DateSerial(Year(CDate("1 " & headerText)), Month(CDate("1 " & headerText)), 1)
    Else
        CalendarMonthFromHeader = DateSerial(Year(GetSelectedReservationDate(ws)), Month(GetSelectedReservationDate(ws)), 1)
    End If
End Function

Private Sub DrawWeeklyReservationBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal weekStart As Date, ByVal vehicleTable As ListObject, ByVal reservationTable As ListObject)
    Dim vehicleCount As Long
    Dim vehicleIndex As Long
    Dim col As Long
    Dim dayIndex As Long
    Dim dayDate As Date
    Dim row As Long
    Dim lastCol As Long
    Dim vehicleId As String
    Dim shopText As String
    Dim entryText As String
    Dim purposeText As String
    Dim statusText As String
    Dim startText As String
    Dim endText As String
    Dim reservationKey As String
    Dim weekEnd As Date
    Dim reservationValues As Variant
    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim dateIndex As Long
    Dim VehicleIndexInTable As Long
    Dim driverIndex As Long
    Dim startTimeIndex As Long
    Dim endTimeIndex As Long
    Dim purposeIndex As Long
    Dim statusIndex As Long
    Dim maxLineCount As Long
    Dim displayLineCount As Long
    Dim reservationTextByKey As Object
    Dim shopTextByKey As Object

    vehicleCount = vehicleTable.ListRows.Count
    lastCol = 2 + (vehicleCount * 2)
    weekEnd = DateAdd("d", 6, weekStart)
    EnsureVehicleSupportColumns vehicleTable
    Set reservationTextByKey = CreateObject("Scripting.Dictionary")
    Set shopTextByKey = CreateObject("Scripting.Dictionary")
    reservationTextByKey.CompareMode = vbTextCompare
    shopTextByKey.CompareMode = vbTextCompare

    If Not reservationTable.DataBodyRange Is Nothing Then
        reservationValues = reservationTable.DataBodyRange.Value2
        dateIndex = reservationTable.ListColumns("Date").index
        VehicleIndexInTable = reservationTable.ListColumns("Vehicle ID").index
        driverIndex = reservationTable.ListColumns("Driver Name").index
        startTimeIndex = reservationTable.ListColumns("Start Time").index
        endTimeIndex = reservationTable.ListColumns("End Time").index
        purposeIndex = reservationTable.ListColumns("Purpose").index
        statusIndex = reservationTable.ListColumns("Status").index
        For rowIndex = 1 To UBound(reservationValues, 1)
            rowDate = reservationValues(rowIndex, dateIndex)
            If IsDate(rowDate) Or IsNumeric(rowDate) Then
                dayDate = DateOnly(CDate(rowDate))
                If dayDate >= weekStart And dayDate <= weekEnd Then
                    vehicleId = Trim$(CStr(reservationValues(rowIndex, VehicleIndexInTable)))
                    reservationKey = CStr(CLng(dayDate)) & "|" & UCase$(vehicleId)
                    statusText = Trim$(CStr(reservationValues(rowIndex, statusIndex)))
                    If IsShopReservationStatus(statusText) Then
                        If Not shopTextByKey.Exists(reservationKey) Then shopTextByKey.Add reservationKey, ShopReservationLabel(statusText, reservationValues(rowIndex, purposeIndex), reservationValues(rowIndex, driverIndex))
                    Else
                        entryText = Trim$(CStr(reservationValues(rowIndex, driverIndex)))
                        startText = ReservationDisplayTime(reservationValues(rowIndex, startTimeIndex))
                        endText = ReservationDisplayTime(reservationValues(rowIndex, endTimeIndex))
                        If startText <> "" Or endText <> "" Then
                            If entryText <> "" Then
                                entryText = startText & "-" & endText & " | " & entryText
                            Else
                                entryText = startText & "-" & endText
                            End If
                        End If
                        purposeText = Trim$(CStr(reservationValues(rowIndex, purposeIndex)))
                        If purposeText <> "" Then entryText = entryText & vbLf & purposeText
                        If entryText = "" Then entryText = "Reserved"
                        If statusText <> "" And StrComp(statusText, "Reserved", vbTextCompare) <> 0 Then entryText = entryText & " [" & statusText & "]"
                        If reservationTextByKey.Exists(reservationKey) Then
                            reservationTextByKey(reservationKey) = reservationTextByKey(reservationKey) & vbLf & entryText
                        Else
                            reservationTextByKey.Add reservationKey, entryText
                        End If
                    End If
                End If
            End If
        Next rowIndex
    End If

    ws.Cells(startRow, 1).value = UCase(Format(weekStart, "ddd"))
    ws.Cells(startRow, 2).value = Format(weekStart, "m/d")
    ws.Range(ws.Cells(startRow, 3), ws.Cells(startRow, lastCol)).Merge
    ws.Cells(startRow, 3).value = BranchCaption(vehicleTable)
    StyleWeeklyTitle ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, lastCol))

    ws.Cells(startRow + 1, 1).value = "DATE"
    ws.Cells(startRow + 1, 2).value = "DAY"
    For vehicleIndex = 1 To vehicleCount
        col = 3 + ((vehicleIndex - 1) * 2)
        ws.Range(ws.Cells(startRow + 1, col), ws.Cells(startRow + 1, col + 1)).Merge
        ws.Range(ws.Cells(startRow + 2, col), ws.Cells(startRow + 2, col + 1)).Merge
        ws.Range(ws.Cells(startRow + 3, col), ws.Cells(startRow + 3, col + 1)).Merge
        ws.Cells(startRow + 1, col).value = UCase(CStr(TableValue(vehicleTable, vehicleIndex, "Make/Model")))
        ws.Cells(startRow + 2, col).value = "Plate: " & CStr(TableValue(vehicleTable, vehicleIndex, "License Plate"))
        ws.Cells(startRow + 3, col).value = CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))
        ws.Cells(startRow + 1, col).Font.Bold = True
        ws.Range(ws.Cells(startRow + 1, col), ws.Cells(startRow + 3, col + 1)).HorizontalAlignment = xlCenter
    Next vehicleIndex
    StyleWeeklyVehicleHeader ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + 3, lastCol))

    For dayIndex = 0 To 6
        dayDate = DateAdd("d", dayIndex, weekStart)
        row = startRow + 4 + dayIndex
        ws.Cells(row, 1).value = UCase(Format(dayDate, "ddd"))
        ws.Cells(row, 2).value = dayDate
        ws.Cells(row, 2).numberFormat = "m/d"
        If DateOnly(dayDate) = Date Then
            ws.Range(ws.Cells(row, 1), ws.Cells(row, lastCol)).Interior.Color = RGB(255, 242, 204)
            ws.Range(ws.Cells(row, 1), ws.Cells(row, 2)).Font.Bold = True
        End If
        If Weekday(dayDate, vbSunday) = vbSunday Or Weekday(dayDate, vbSunday) = vbSaturday Then ws.Range(ws.Cells(row, 1), ws.Cells(row, 2)).Font.Color = RGB(192, 0, 0)

        maxLineCount = 1
        For vehicleIndex = 1 To vehicleCount
            col = 3 + ((vehicleIndex - 1) * 2)
            vehicleId = CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))
            ws.Range(ws.Cells(row, col), ws.Cells(row, col + 1)).Merge
            reservationKey = CStr(CLng(DateOnly(dayDate))) & "|" & UCase$(Trim$(vehicleId))
            If shopTextByKey.Exists(reservationKey) Then
                ws.Cells(row, col).value = shopTextByKey(reservationKey)
                StyleWeeklyShopCell ws.Cells(row, col)
            ElseIf VehicleInShopOnDate(vehicleTable, vehicleId, dayDate) Then
                ws.Cells(row, col).value = ShopStatusText(vehicleTable, vehicleId)
                StyleWeeklyShopCell ws.Cells(row, col)
            ElseIf reservationTextByKey.Exists(reservationKey) Then
                ws.Cells(row, col).value = reservationTextByKey(reservationKey)
            End If
            ws.Cells(row, col).WrapText = True
            displayLineCount = WeeklyReservationLineCount(CStr(ws.Cells(row, col).value))
            If displayLineCount > maxLineCount Then maxLineCount = displayLineCount
        Next vehicleIndex
        ws.Rows(row).RowHeight = Application.Min(112, Application.Max(28, (14 * maxLineCount) + 4))
    Next dayIndex

    ApplyBlackGrid ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 10, lastCol))
    ws.Rows(startRow + 1).RowHeight = 52
End Sub

Private Function WeeklyReservationLineCount(ByVal displayText As String) As Long
    Dim lineItem As Variant
    Dim estimatedLines As Long

    If Trim$(displayText) = "" Then
        WeeklyReservationLineCount = 1
        Exit Function
    End If

    For Each lineItem In Split(displayText, vbLf)
        estimatedLines = (Len(CStr(lineItem)) + 29) \ 30
        If estimatedLines < 1 Then estimatedLines = 1
        WeeklyReservationLineCount = WeeklyReservationLineCount + estimatedLines
    Next lineItem
End Function

Private Sub GetDateRange(ByVal lo As ListObject, ByVal dateColumn As String, ByRef minDate As Date, ByRef maxDate As Date, ByRef hasDates As Boolean)
    Dim i As Long
    Dim value As Variant
    Dim dateValue As Date

    hasDates = False
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To lo.ListRows.Count
        value = TableValue(lo, i, dateColumn)
        If IsDate(value) Then
            dateValue = DateOnly(value)
            If Not hasDates Then
                minDate = dateValue
                maxDate = dateValue
                hasDates = True
            Else
                If dateValue < minDate Then minDate = dateValue
                If dateValue > maxDate Then maxDate = dateValue
            End If
        End If
    Next i
End Sub

Private Function PreferredMonthlyMileageMonth(ByVal mileageTable As ListObject) As Date
    Dim minDate As Date
    Dim maxDate As Date
    Dim hasDates As Boolean
    Dim currentMonthStart As Date
    Dim currentMonthEnd As Date

    currentMonthStart = DateSerial(Year(Date), Month(Date), 1)
    currentMonthEnd = DateSerial(Year(Date), Month(Date) + 1, 0)
    GetDateRange mileageTable, "Date", minDate, maxDate, hasDates

    If Not hasDates Then
        PreferredMonthlyMileageMonth = currentMonthStart
    ElseIf HasMileageInMonth(mileageTable, currentMonthStart, currentMonthEnd) Then
        PreferredMonthlyMileageMonth = currentMonthStart
    Else
        PreferredMonthlyMileageMonth = DateSerial(Year(maxDate), Month(maxDate), 1)
    End If
End Function

Private Function HasMileageInMonth(ByVal mileageTable As ListObject, ByVal monthStart As Date, ByVal monthEnd As Date) As Boolean
    Dim rowIndex As Long
    Dim rowDate As Variant

    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To mileageTable.ListRows.Count
        rowDate = TableValue(mileageTable, rowIndex, "Date")
        If IsDate(rowDate) Then
            If DateOnly(rowDate) >= monthStart And DateOnly(rowDate) <= monthEnd Then
                HasMileageInMonth = True
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Sub RefreshMonthlyMileageMonth(ByVal targetMonth As Date, Optional ByVal activateSheet As Boolean = False)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim vehicleTable As ListObject
    Dim mileageTable As ListObject
    Dim reservationTable As ListObject
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim startRow As Long
    Dim lastCol As Long
    Dim blockRows As Long
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    monthStart = DateSerial(Year(targetMonth), Month(targetMonth), 1)
    monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)
    Set ws = EnsureWorksheet("Monthly Mileage")
    FleetMaterializeActiveShopHistory False
    Set vehicleTable = GetTable("tblVehicles")
    Set mileageTable = GetTable("tblMileageLog")
    Set reservationTable = GetTable("tblReservations")

    startRow = FindMonthlyMileageMonthRow(ws, monthStart)
    If startRow = 0 Then startRow = NextMonthlyMileageAppendRow(ws)

    lastCol = 2 + (vehicleTable.ListRows.Count * 2)
    blockRows = Day(monthEnd) + 4
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + blockRows - 1, lastCol))
        .UnMerge
        .Clear
    End With

    DrawMonthlyMileageBlock ws, startRow, monthStart, monthEnd, vehicleTable, mileageTable, reservationTable
    ws.Columns("A:J").ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 15.5
    ws.Columns(5).ColumnWidth = 15.5
    ws.Columns(7).ColumnWidth = 15.5
    ws.Columns(9).ColumnWidth = 15.5

    If activateSheet Then
        EnsureMonthlyMileageNavigation ws
        FreezeMonthlyMileageNavigation ws
        SelectMonthlyMileageMonth ws, monthStart
    End If

    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Err.Raise Err.Number, APP_TITLE, Err.Description
End Sub

Public Sub FleetRefreshMonthlyMileageShopStatus(Optional ByVal showMessage As Boolean = False)
    On Error GoTo CleanFail

    Dim vehicleTable As ListObject
    Dim vehicleIndex As Long
    Dim vehicleStatus As String
    Dim hasActiveShop As Boolean

    FleetMaterializeActiveShopHistory False
    Set vehicleTable = GetTable("tblVehicles")
    EnsureVehicleSupportColumns vehicleTable
    For vehicleIndex = 1 To vehicleTable.ListRows.Count
        vehicleStatus = Trim$(CStr(TableValue(vehicleTable, vehicleIndex, "Status")))
        If StrComp(vehicleStatus, "Maintenance", vbTextCompare) = 0 Then
            hasActiveShop = True
            Exit For
        End If
    Next vehicleIndex

    If hasActiveShop Then RefreshMonthlyMileageShopRange Date, MonthlyMileageProjectionEndDate()
    If showMessage Then MsgBox "Monthly Mileage shop status refreshed.", vbInformation, APP_TITLE
    Exit Sub

CleanFail:
    If showMessage Then
        MsgBox "Could not refresh Monthly Mileage shop status: " & Err.Description, vbExclamation, APP_TITLE
    Else
        Err.Raise Err.Number, APP_TITLE, Err.Description
    End If
End Sub

Private Sub RefreshMonthlyMileageShopRange(ByVal firstDate As Date, ByVal lastDate As Date)
    Dim monthCursor As Date
    Dim lastMonth As Date

    firstDate = DateOnly(firstDate)
    lastDate = DateOnly(lastDate)
    If lastDate < firstDate Then Exit Sub
    monthCursor = DateSerial(Year(firstDate), Month(firstDate), 1)
    lastMonth = DateSerial(Year(lastDate), Month(lastDate), 1)
    Do While monthCursor <= lastMonth
        RefreshMonthlyMileageShopMonth monthCursor
        monthCursor = DateAdd("m", 1, monthCursor)
    Loop
End Sub

Private Function MonthlyMileageProjectionEndDate() As Date
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim values As Variant
    Dim rowIndex As Long
    Dim monthStart As Date
    Dim latestMonth As Date

    latestMonth = DateSerial(Year(Date), Month(Date), 1)
    Set ws = EnsureWorksheet("Monthly Mileage")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    If lastRow > 0 Then
        values = ws.Range("A1:B" & CStr(lastRow)).Value2
        For rowIndex = 1 To UBound(values, 1)
            If UCase$(Trim$(CStr(values(rowIndex, 1)))) = "VEHICLES" Then
                If TryMonthStartFromValue(values(rowIndex, 2), monthStart) Then
                    If monthStart > latestMonth Then latestMonth = monthStart
                End If
            End If
        Next rowIndex
    End If

    MonthlyMileageProjectionEndDate = DateSerial(Year(latestMonth), Month(latestMonth) + 1, 0)
End Function

Private Sub RefreshMonthlyMileageShopMonth(ByVal targetMonth As Date)
    Dim ws As Worksheet
    Dim vehicleTable As ListObject
    Dim reservationTable As ListObject
    Dim shopTextByKey As Object
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim startRow As Long
    Dim vehicleIndex As Long
    Dim dayNumber As Long
    Dim dayRow As Long
    Dim col As Long
    Dim currentDate As Date
    Dim vehicleId As String
    Dim shopKey As String
    Dim tripValue As Variant
    Dim endingValue As Variant

    monthStart = DateSerial(Year(targetMonth), Month(targetMonth), 1)
    monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)
    Set ws = EnsureWorksheet("Monthly Mileage")
    Set vehicleTable = GetTable("tblVehicles")
    Set reservationTable = GetTable("tblReservations")
    startRow = FindMonthlyMileageMonthRow(ws, monthStart)
    If startRow = 0 Then
        RefreshMonthlyMileageMonth monthStart, False
        Exit Sub
    End If

    Set shopTextByKey = BuildMonthlyShopStatusMap(vehicleTable, reservationTable, monthStart, monthEnd)
    For dayNumber = 1 To Day(monthEnd)
        currentDate = DateSerial(Year(monthStart), Month(monthStart), dayNumber)
        dayRow = startRow + 1 + dayNumber
        For vehicleIndex = 1 To vehicleTable.ListRows.Count
            col = 3 + ((vehicleIndex - 1) * 2)
            vehicleId = CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))
            tripValue = ws.Cells(dayRow, col).value
            endingValue = ws.Cells(dayRow, col + 1).value
            shopKey = MonthlyShopStatusKey(currentDate, vehicleId)
            If (Not IsEmpty(tripValue) And IsNumeric(tripValue)) Or _
                (Not IsEmpty(endingValue) And IsNumeric(endingValue)) Then
                ResetMonthlyMileageEntryCells ws.Range(ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1))
            ElseIf shopTextByKey.Exists(shopKey) Then
                StyleMonthlyShopPair ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1), CStr(shopTextByKey(shopKey))
            ElseIf UCase$(Trim$(CStr(tripValue))) = "IN SHOP" Then
                ResetMonthlyMileageEntryCells ws.Range(ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1))
                ws.Range(ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1)).ClearContents
            End If
        Next vehicleIndex
    Next dayNumber
End Sub

Private Sub EnsureMonthlyMileageMonthBlock(ByVal targetMonth As Date)
    Dim ws As Worksheet
    Dim monthStart As Date

    monthStart = DateSerial(Year(targetMonth), Month(targetMonth), 1)
    Set ws = EnsureWorksheet("Monthly Mileage")
    If FindMonthlyMileageMonthRow(ws, monthStart) = 0 Then
        RefreshMonthlyMileageMonth monthStart, False
    End If
End Sub

Private Function ExistingMonthlyMileageMonthOrFallback(ByVal ws As Worksheet, ByVal preferredMonth As Date) As Date
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim headerMonth As Date
    Dim latestMonth As Date
    Dim hasLatestMonth As Boolean

    preferredMonth = DateSerial(Year(preferredMonth), Month(preferredMonth), 1)
    If FindMonthlyMileageMonthRow(ws, preferredMonth) > 0 Then
        ExistingMonthlyMileageMonthOrFallback = preferredMonth
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    For rowIndex = 1 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(rowIndex, "A").value))) = "VEHICLES" Then
            If TryMonthStartFromLabel(Trim$(CStr(ws.Cells(rowIndex, "B").value)), headerMonth) Then
                If Not hasLatestMonth Or headerMonth > latestMonth Then
                    latestMonth = headerMonth
                    hasLatestMonth = True
                End If
            End If
        End If
    Next rowIndex

    If hasLatestMonth Then
        ExistingMonthlyMileageMonthOrFallback = latestMonth
    Else
        ExistingMonthlyMileageMonthOrFallback = preferredMonth
    End If
End Function

Private Sub UpdateMonthlyMileageEntry(ByVal entryDate As Date, ByVal vehicleId As String)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim vehicleTable As ListObject
    Dim mileageTable As ListObject
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim startRow As Long
    Dim vehicleIndex As Long
    Dim col As Long
    Dim dayRow As Long
    Dim totalRow As Long
    Dim cumulativeRow As Long
    Dim dayTrips As Long
    Dim monthTrips As Long
    Dim monthMiles As Double
    Dim dayEnding As Long
    Dim monthEnding As Long
    Dim hasDayEnding As Boolean
    Dim hasMonthEnding As Boolean
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    monthStart = DateSerial(Year(entryDate), Month(entryDate), 1)
    monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)
    Set ws = EnsureWorksheet("Monthly Mileage")
    Set vehicleTable = GetTable("tblVehicles")
    Set mileageTable = GetTable("tblMileageLog")

    startRow = FindMonthlyMileageMonthRow(ws, monthStart)
    If startRow = 0 Then
        Application.EnableEvents = previousEvents
        Application.ScreenUpdating = previousScreenUpdating
        RefreshMonthlyMileageMonth monthStart, False
        Exit Sub
    End If

    vehicleIndex = VehicleIndexInTable(vehicleTable, vehicleId)
    If vehicleIndex = 0 Then GoTo CleanExit

    CollectMonthlyMileageStats mileageTable, vehicleId, entryDate, monthStart, monthEnd, _
        dayTrips, hasDayEnding, dayEnding, monthTrips, monthMiles, hasMonthEnding, monthEnding

    col = 3 + ((vehicleIndex - 1) * 2)
    dayRow = startRow + 1 + Day(entryDate)
    totalRow = startRow + 2 + Day(monthEnd)
    cumulativeRow = totalRow + 1

    ResetMonthlyMileageEntryCells ws.Range(ws.Cells(dayRow, col), ws.Cells(dayRow, col + 1))
    If dayTrips > 0 Then
        ws.Cells(dayRow, col).value = dayTrips
    Else
        ws.Cells(dayRow, col).ClearContents
    End If

    If hasDayEnding Then
        ws.Cells(dayRow, col + 1).value = dayEnding
    Else
        ws.Cells(dayRow, col + 1).ClearContents
    End If
    ws.Cells(dayRow, col + 1).numberFormat = "#,##0"

    ws.Cells(totalRow, col).value = monthTrips
    ws.Cells(totalRow, col + 1).value = monthMiles
    ws.Cells(totalRow, col + 1).numberFormat = "#,##0"
    ws.Cells(cumulativeRow, col).value = "Ending Mileage"
    If hasMonthEnding Then
        ws.Cells(cumulativeRow, col + 1).value = monthEnding
    Else
        ws.Cells(cumulativeRow, col + 1).ClearContents
    End If
    ws.Cells(cumulativeRow, col + 1).numberFormat = "#,##0"

CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Err.Raise Err.Number, APP_TITLE, Err.Description
End Sub

Private Function VehicleIndexInTable(ByVal vehicleTable As ListObject, ByVal vehicleId As String) As Long
    Dim vehicleIndex As Long

    For vehicleIndex = 1 To vehicleTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(vehicleTable, vehicleIndex, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            VehicleIndexInTable = vehicleIndex
            Exit Function
        End If
    Next vehicleIndex
End Function

Private Sub CollectMonthlyMileageStats(ByVal mileageTable As ListObject, ByVal vehicleId As String, ByVal entryDate As Date, _
    ByVal monthStart As Date, ByVal monthEnd As Date, ByRef dayTrips As Long, ByRef hasDayEnding As Boolean, _
    ByRef dayEnding As Long, ByRef monthTrips As Long, ByRef monthMiles As Double, ByRef hasMonthEnding As Boolean, _
    ByRef monthEnding As Long)

    Dim rowIndex As Long
    Dim rowDate As Variant
    Dim rowDateOnly As Date
    Dim endingValue As Variant

    If mileageTable.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To mileageTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(mileageTable, rowIndex, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
            rowDate = TableValue(mileageTable, rowIndex, "Date")
            If IsDate(rowDate) Then
                rowDateOnly = DateOnly(rowDate)
                If rowDateOnly >= monthStart And rowDateOnly <= monthEnd Then
                    monthTrips = monthTrips + TripCountForRow(mileageTable, rowIndex)
                    endingValue = TableValue(mileageTable, rowIndex, "Odometer End")
                    If IsNumeric(endingValue) Then
                        If rowDateOnly = DateOnly(entryDate) Then
                            If Not hasDayEnding Or CLng(endingValue) > dayEnding Then
                                dayEnding = CLng(endingValue)
                                hasDayEnding = True
                            End If
                        End If
                    End If
                    If rowDateOnly = DateOnly(entryDate) Then
                        dayTrips = dayTrips + TripCountForRow(mileageTable, rowIndex)
                    End If
                End If
            End If
        End If
    Next rowIndex

    monthMiles = MonthMilesForVehicle(mileageTable, monthStart, monthEnd, vehicleId)
    endingValue = EndingMileageForMonthVehicle(mileageTable, monthStart, monthEnd, vehicleId)
    If IsNumeric(endingValue) Then
        If CLng(endingValue) > 0 Then
            monthEnding = CLng(endingValue)
            hasMonthEnding = True
        End If
    End If
End Sub

Private Function FindMonthlyMileageMonthRow(ByVal ws As Worksheet, ByVal targetMonth As Date) As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim headerMonth As Date

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    For rowIndex = 1 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(rowIndex, "A").value))) = "VEHICLES" Then
            If TryMonthStartFromLabel(Trim$(CStr(ws.Cells(rowIndex, "B").value)), headerMonth) Then
                If headerMonth = DateSerial(Year(targetMonth), Month(targetMonth), 1) Then
                    FindMonthlyMileageMonthRow = rowIndex
                    Exit Function
                End If
            End If
        End If
    Next rowIndex
End Function

Private Function NextMonthlyMileageAppendRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        NextMonthlyMileageAppendRow = 1
    Else
        NextMonthlyMileageAppendRow = lastCell.row + 2
    End If
End Function

Private Sub SelectMonthlyMileageMonth(ByVal ws As Worksheet, ByVal targetMonth As Date)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim headerMonth As Date
    Dim latestRow As Long
    Dim latestMonth As Date

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    For rowIndex = 1 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(rowIndex, "A").value))) = "VEHICLES" Then
            If TryMonthStartFromLabel(Trim$(CStr(ws.Cells(rowIndex, "B").value)), headerMonth) Then
                If latestRow = 0 Or headerMonth > latestMonth Then
                    latestRow = rowIndex
                    latestMonth = headerMonth
                End If
                If headerMonth = targetMonth Then
                    ScrollToMonthlyMileageRow ws, rowIndex
                    Exit Sub
                End If
            End If
        End If
    Next rowIndex

    If latestRow = 0 Then
        For rowIndex = 1 To lastRow
            If TryMonthStartFromLabel(Trim$(CStr(ws.Cells(rowIndex, "B").value)), headerMonth) Then
                If latestRow = 0 Or headerMonth > latestMonth Then
                    latestRow = rowIndex
                    latestMonth = headerMonth
                End If
                If headerMonth = targetMonth Then
                    ScrollToMonthlyMileageRow ws, rowIndex
                    Exit Sub
                End If
            End If
        Next rowIndex
    End If

    If latestRow > 0 Then ScrollToMonthlyMileageRow ws, latestRow
End Sub

Private Sub ScrollToMonthlyMileageRow(ByVal ws As Worksheet, ByVal rowIndex As Long)
    ws.Activate
    On Error Resume Next
    ActiveWindow.ScrollRow = rowIndex
    ActiveWindow.ScrollColumn = 1
    If Application.Visible Then ws.Cells(rowIndex, "B").Select
    On Error GoTo 0
End Sub

Private Sub EnsureMonthlyMileageNavigation(ByVal ws As Worksheet)
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim dashboardButton As Shape

    On Error GoTo CleanExit
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error Resume Next
    Set dashboardButton = ws.Shapes("btnMonthlyDashboardTop")
    On Error GoTo CleanExit

    ws.Rows(1).RowHeight = 24
    If dashboardButton Is Nothing Then
        AddReportButton ws, "A1:B1", "DASHBOARD", "FleetShowDashboard", RGB(255, 121, 0), "btnMonthlyDashboardTop"
        Set dashboardButton = ws.Shapes("btnMonthlyDashboardTop")
    End If

    On Error Resume Next
    dashboardButton.Fill.ForeColor.RGB = RGB(255, 121, 0)
    dashboardButton.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    dashboardButton.TextFrame.Characters.Font.Size = 10
    dashboardButton.Placement = xlFreeFloating
    On Error GoTo CleanExit

CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    On Error GoTo 0
End Sub

Private Sub FreezeMonthlyMileageNavigation(ByVal ws As Worksheet)
    If Not Application.Visible Then Exit Sub
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Function TableValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    TableValue = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).index).value
End Function

Private Function TableHasColumn(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    On Error Resume Next
    TableHasColumn = (lo.ListColumns(columnName).index > 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TripCountForRow(ByVal mileageTable As ListObject, ByVal rowIndex As Long) As Long
    Dim value As Variant

    TripCountForRow = 1
    If TableHasColumn(mileageTable, "Trip Count") Then
        value = TableValue(mileageTable, rowIndex, "Trip Count")
        If IsNumeric(value) Then
            TripCountForRow = CLng(value)
        End If
    End If
End Function

Private Function MonthlyVehicleCaption(ByVal vehicleTable As ListObject, ByVal rowIndex As Long) As String
    MonthlyVehicleCaption = CStr(TableValue(vehicleTable, rowIndex, "Make/Model")) & " (" & CStr(TableValue(vehicleTable, rowIndex, "Vehicle ID")) & ")"
End Function

Private Function TripsForDateVehicle(ByVal mileageTable As ListObject, ByVal targetDate As Date, ByVal vehicleId As String) As Long
    Dim i As Long
    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To mileageTable.ListRows.Count
        If IsMileageMatch(mileageTable, i, targetDate, targetDate, vehicleId) Then TripsForDateVehicle = TripsForDateVehicle + TripCountForRow(mileageTable, i)
    Next i
End Function

Private Function EndingMileageForDateVehicle(ByVal mileageTable As ListObject, ByVal targetDate As Date, ByVal vehicleId As String) As Variant
    Dim i As Long
    Dim endingValue As Variant
    Dim bestEnding As Long
    Dim hasEnding As Boolean

    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To mileageTable.ListRows.Count
        If IsMileageMatch(mileageTable, i, targetDate, targetDate, vehicleId) Then
            endingValue = TableValue(mileageTable, i, "Odometer End")
            If IsNumeric(endingValue) Then
                If Not hasEnding Then
                    bestEnding = CLng(endingValue)
                    hasEnding = True
                ElseIf CLng(endingValue) > bestEnding Then
                    bestEnding = CLng(endingValue)
                End If
            End If
        End If
    Next i
    If hasEnding Then EndingMileageForDateVehicle = bestEnding
End Function

Private Function MonthTripsForVehicle(ByVal mileageTable As ListObject, ByVal monthStart As Date, ByVal monthEnd As Date, ByVal vehicleId As String) As Long
    Dim i As Long
    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To mileageTable.ListRows.Count
        If IsMileageMatch(mileageTable, i, monthStart, monthEnd, vehicleId) Then MonthTripsForVehicle = MonthTripsForVehicle + TripCountForRow(mileageTable, i)
    Next i
End Function

Private Function MonthMilesForVehicle(ByVal mileageTable As ListObject, ByVal monthStart As Date, ByVal monthEnd As Date, ByVal vehicleId As String) As Double
    Dim i As Long
    Dim rowDate As Date
    Dim miles As Variant
    Dim endingValue As Variant
    Dim firstDate As Date
    Dim latestDate As Date
    Dim firstEnding As Long
    Dim latestEnding As Long
    Dim matchingRows As Long
    Dim summedMiles As Double
    Dim hasFirstEnding As Boolean
    Dim hasLatestEnding As Boolean
    Dim useSourceTotals As Boolean
    Dim notesText As String

    If mileageTable.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To mileageTable.ListRows.Count
        If StrComp(Trim$(CStr(TableValue(mileageTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 _
                And IsDate(TableValue(mileageTable, i, "Date")) Then
            rowDate = DateOnly(TableValue(mileageTable, i, "Date"))
            notesText = CStr(TableValue(mileageTable, i, "Notes"))
            endingValue = TableValue(mileageTable, i, "Odometer End")

            If rowDate >= monthStart And rowDate <= monthEnd Then
                matchingRows = matchingRows + 1
                miles = TableValue(mileageTable, i, "Miles Driven")
                If IsNumeric(miles) Then summedMiles = summedMiles + CDbl(miles)
                If InStr(1, notesText, "monthly source adjustment", vbTextCompare) > 0 _
                        Or InStr(1, notesText, "weekly source", vbTextCompare) > 0 Then
                    useSourceTotals = True
                End If

                If IsNumeric(endingValue) Then
                    If CLng(endingValue) > 0 And InStr(1, notesText, "odometer value lower than prior known ending", vbTextCompare) = 0 Then
                        If Not hasFirstEnding Or rowDate < firstDate _
                                Or (rowDate = firstDate And CLng(endingValue) < firstEnding) Then
                            firstDate = rowDate
                            firstEnding = CLng(endingValue)
                            hasFirstEnding = True
                        End If
                        If Not hasLatestEnding Or rowDate > latestDate _
                                Or (rowDate = latestDate And CLng(endingValue) > latestEnding) Then
                            latestDate = rowDate
                            latestEnding = CLng(endingValue)
                            hasLatestEnding = True
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If useSourceTotals Or matchingRows = 0 Or Not hasLatestEnding Then
        MonthMilesForVehicle = summedMiles
    ElseIf hasFirstEnding And latestEnding >= firstEnding Then
        MonthMilesForVehicle = latestEnding - firstEnding
    Else
        MonthMilesForVehicle = summedMiles
    End If
End Function

Private Function FiscalYearMilesForVehicle(ByVal mileageTable As ListObject, ByVal vehicleId As String, _
    ByVal fiscalYearStart As Date, ByVal selectedMonthStart As Date) As Double

    Dim monthCursor As Date
    Dim monthEnd As Date

    monthCursor = DateSerial(Year(fiscalYearStart), Month(fiscalYearStart), 1)
    Do While monthCursor <= selectedMonthStart
        monthEnd = DateSerial(Year(monthCursor), Month(monthCursor) + 1, 0)
        FiscalYearMilesForVehicle = FiscalYearMilesForVehicle + _
            MonthMilesForVehicle(mileageTable, monthCursor, monthEnd, vehicleId)
        monthCursor = DateAdd("m", 1, monthCursor)
    Loop
End Function

Private Function EndingMileageForMonthVehicle(ByVal mileageTable As ListObject, ByVal monthStart As Date, ByVal monthEnd As Date, ByVal vehicleId As String) As Variant
    Dim i As Long
    Dim endingValue As Variant
    Dim rowDate As Date
    Dim latestDate As Date
    Dim hasEnding As Boolean
    Dim bestEnding As Long
    Dim notesText As String

    If mileageTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To mileageTable.ListRows.Count
        If IsDate(TableValue(mileageTable, i, "Date")) Then
            rowDate = DateOnly(TableValue(mileageTable, i, "Date"))
            If rowDate <= monthEnd And StrComp(Trim$(CStr(TableValue(mileageTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
                endingValue = TableValue(mileageTable, i, "Odometer End")
                notesText = CStr(TableValue(mileageTable, i, "Notes"))
                If IsNumeric(endingValue) Then
                    If CLng(endingValue) > 0 And InStr(1, notesText, "odometer value lower than prior known ending", vbTextCompare) = 0 Then
                        If Not hasEnding Then
                            bestEnding = CLng(endingValue)
                            latestDate = rowDate
                            hasEnding = True
                        ElseIf rowDate > latestDate Then
                            bestEnding = CLng(endingValue)
                            latestDate = rowDate
                        ElseIf rowDate = latestDate And CLng(endingValue) > bestEnding Then
                            bestEnding = CLng(endingValue)
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If hasEnding Then EndingMileageForMonthVehicle = bestEnding
End Function

Private Function IsMileageMatch(ByVal mileageTable As ListObject, ByVal rowIndex As Long, ByVal startDate As Date, ByVal endDate As Date, ByVal vehicleId As String) As Boolean
    Dim rowDate As Variant
    rowDate = TableValue(mileageTable, rowIndex, "Date")
    If Not IsDate(rowDate) Then Exit Function
    If DateOnly(rowDate) < startDate Or DateOnly(rowDate) > endDate Then Exit Function
    If StrComp(Trim$(CStr(TableValue(mileageTable, rowIndex, "Vehicle ID"))), vehicleId, vbTextCompare) <> 0 Then Exit Function
    IsMileageMatch = True
End Function

Private Function StartOfWeek(ByVal targetDate As Date) As Date
    StartOfWeek = DateAdd("d", 1 - Weekday(targetDate, vbSunday), targetDate)
End Function

Private Function DateOnly(ByVal value As Variant) As Date
    DateOnly = DateSerial(Year(CDate(value)), Month(CDate(value)), Day(CDate(value)))
End Function

Private Function TryDashboardDateRange(ByVal value As Variant, ByRef startDate As Date, ByRef endDate As Date) As Boolean
    Dim textValue As String
    Dim splitAt As Long
    Dim leftText As String
    Dim rightText As String

    If IsDate(value) Then
        startDate = DateOnly(value)
        endDate = startDate
        TryDashboardDateRange = True
        Exit Function
    End If

    textValue = Trim$(CStr(value))
    If textValue = "" Then Exit Function
    textValue = Replace(textValue, " to ", " - ", 1, -1, vbTextCompare)
    splitAt = InStr(1, textValue, " - ", vbTextCompare)
    If splitAt = 0 Then Exit Function

    leftText = Trim$(Left$(textValue, splitAt - 1))
    rightText = Trim$(Mid$(textValue, splitAt + 3))
    If IsDate(leftText) And IsDate(rightText) Then
        startDate = DateOnly(CDate(leftText))
        endDate = DateOnly(CDate(rightText))
        If endDate < startDate Then SwapDates startDate, endDate
        TryDashboardDateRange = True
    End If
End Function

Private Function DashboardDateRangeText(ByVal startDate As Date, ByVal endDate As Date) As String
    If DateOnly(startDate) = DateOnly(endDate) Then
        DashboardDateRangeText = Format(startDate, "m/d/yyyy")
    Else
        DashboardDateRangeText = Format(startDate, "m/d/yyyy") & " - " & Format(endDate, "m/d/yyyy")
    End If
End Function

Private Sub SwapDates(ByRef firstDate As Date, ByRef secondDate As Date)
    Dim tempDate As Date
    tempDate = firstDate
    firstDate = secondDate
    secondDate = tempDate
End Sub

Private Function DashboardCalendarStateSheet() As Worksheet
    Set DashboardCalendarStateSheet = EnsureWorksheet("DashboardData")
End Function

Private Function DashboardCalendarAllowsRange() As Boolean
    DashboardCalendarAllowsRange = (UCase$(Trim$(CStr(DashboardCalendarStateSheet().Range("P2").value))) = "TRUE")
End Function

Private Function DashboardCalendarRangeMode() As Boolean
    DashboardCalendarRangeMode = (UCase$(Trim$(CStr(DashboardCalendarStateSheet().Range("P3").value))) = "RANGE")
End Function

Private Function MonthNumberFromName(ByVal monthText As String) As Long
    Dim monthIndex As Long

    monthText = Trim$(monthText)
    For monthIndex = 1 To 12
        If StrComp(monthText, Format$(DateSerial(2000, monthIndex, 1), "mmmm"), vbTextCompare) = 0 _
            Or StrComp(monthText, Format$(DateSerial(2000, monthIndex, 1), "mmm"), vbTextCompare) = 0 Then
            MonthNumberFromName = monthIndex
            Exit Function
        End If
    Next monthIndex
End Function

Private Function YearValidationList(ByVal firstYear As Long, ByVal lastYear As Long) As String
    Dim yearNumber As Long
    Dim result As String

    If firstYear < 2020 Then firstYear = 2020
    If lastYear > 2035 Then lastYear = 2035
    For yearNumber = firstYear To lastYear
        If result <> "" Then result = result & ","
        result = result & CStr(yearNumber)
    Next yearNumber
    YearValidationList = result
End Function

Private Function BranchCaption(ByVal vehicleTable As ListObject) As String
    Dim branch As String
    If vehicleTable.ListRows.Count = 0 Then
        BranchCaption = "Fleet Reservations"
        Exit Function
    End If

    branch = CStr(TableValue(vehicleTable, 1, "Branch"))
    If UCase$(branch) = "DEMO FLEET" Then
        BranchCaption = "DEMO FLEET"
    ElseIf branch <> "" Then
        BranchCaption = branch & " Fleet Reservations"
    Else
        BranchCaption = "Fleet Reservations"
    End If
End Function

Private Function ReservationTextForDateVehicle(ByVal reservationTable As ListObject, ByVal targetDate As Date, ByVal vehicleId As String) As String
    Dim i As Long
    Dim rowDate As Variant
    Dim result As String
    Dim entryText As String
    Dim purposeText As String
    Dim statusText As String

    If reservationTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To reservationTable.ListRows.Count
        rowDate = TableValue(reservationTable, i, "Date")
        If IsDate(rowDate) Then
            If DateOnly(rowDate) = targetDate Then
                If StrComp(Trim$(CStr(TableValue(reservationTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
                    entryText = CStr(TableValue(reservationTable, i, "Driver Name"))
                    purposeText = Trim$(CStr(TableValue(reservationTable, i, "Purpose")))
                    statusText = Trim$(CStr(TableValue(reservationTable, i, "Status")))
                    If IsShopReservationStatus(statusText) Then GoTo NextReservationRow
                    If purposeText <> "" Then
                        If Trim$(entryText) <> "" Then
                            entryText = entryText & vbLf & purposeText
                        Else
                            entryText = purposeText
                        End If
                    End If
                    If Trim$(entryText) = "" And (statusText = "" Or StrComp(statusText, "Reserved", vbTextCompare) = 0 Or StrComp(statusText, "Scheduled", vbTextCompare) = 0) Then entryText = "Reserved"
                    If statusText <> "" And StrComp(statusText, "Reserved", vbTextCompare) <> 0 Then entryText = entryText & " [" & statusText & "]"
                    If result <> "" Then result = result & vbLf
                    result = result & entryText
                End If
            End If
        End If
NextReservationRow:
    Next i
    ReservationTextForDateVehicle = result
End Function

Private Function ShopReservationTextForDateVehicle(ByVal reservationTable As ListObject, ByVal targetDate As Date, ByVal vehicleId As String) As String
    Dim i As Long
    Dim rowDate As Variant
    Dim statusText As String

    If reservationTable.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To reservationTable.ListRows.Count
        rowDate = TableValue(reservationTable, i, "Date")
        If IsDate(rowDate) Then
            If DateOnly(rowDate) = targetDate Then
                If StrComp(Trim$(CStr(TableValue(reservationTable, i, "Vehicle ID"))), vehicleId, vbTextCompare) = 0 Then
                    statusText = Trim$(CStr(TableValue(reservationTable, i, "Status")))
                    If IsShopReservationStatus(statusText) Then
                        ShopReservationTextForDateVehicle = ShopReservationLabel(statusText, _
                            TableValue(reservationTable, i, "Purpose"), TableValue(reservationTable, i, "Driver Name"))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Sub StyleMonthlyHeader(ByVal rng As Range)
    rng.Interior.Color = RGB(0, 83, 143)
    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
End Sub

Private Sub StyleSubHeader(ByVal rng As Range)
    rng.Interior.Color = RGB(217, 217, 217)
    rng.Font.Color = RGB(0, 0, 0)
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
End Sub

Private Sub StyleTotalRows(ByVal rng As Range)
    rng.Interior.Color = RGB(226, 239, 218)
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
End Sub

Private Sub StyleWeeklyTitle(ByVal rng As Range)
    rng.Interior.Color = RGB(0, 83, 143)
    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
End Sub

Private Sub StyleWeeklyVehicleHeader(ByVal rng As Range)
    rng.Interior.Color = RGB(232, 244, 252)
    rng.Font.Color = RGB(0, 0, 0)
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
End Sub

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub DeleteShapesByPrefix(ByVal ws As Worksheet, ByVal namePrefix As String)
    Dim shapeIndex As Long

    On Error Resume Next
    For shapeIndex = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(shapeIndex).Name, Len(namePrefix)) = namePrefix Then
            ws.Shapes(shapeIndex).Delete
        End If
    Next shapeIndex
    On Error GoTo 0
End Sub

Private Sub AddReportButton(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, ByVal macroName As String, ByVal fillColor As Long, Optional ByVal shapeName As String = "")
    Dim rng As Range
    Dim shp As Shape

    Set rng = ws.Range(address)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    If shapeName <> "" Then shp.Name = shapeName
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.HorizontalAlignment = xlCenter
    shp.TextFrame.VerticalAlignment = xlCenter
    shp.OnAction = macroName
End Sub

Private Sub ApplyBlackGrid(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
End Sub

Public Sub FleetProtectOutputSheets()
    Const sheetPassword As String = "VehicleTrackerDemo"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Monthly Mileage")
    ws.Unprotect Password:=sheetPassword
    ws.Cells.Locked = True
    ws.Protect Password:=sheetPassword, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    Set ws = ThisWorkbook.Worksheets("Weekly Reservations")
    ws.Unprotect Password:=sheetPassword
    ws.Cells.Locked = True
    ws.Range("C3,F3,C4:D4,L3").Locked = False
    ws.Protect Password:=sheetPassword, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub


Private Sub RefreshWeeklyReservationsForDate(ByVal ws As Worksheet, ByVal selectedDate As Date)
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo CleanFail
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Range("C4").value = DateOnly(selectedDate)
    RefreshWeeklyReservationsGrid ws, selectedDate
    If Not (Application.ActiveSheet Is ws) Then ApplySoftwareWindowChrome ws

CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Err.Raise errorNumber, APP_TITLE, errorDescription
End Sub

Private Sub ProtectWeeklyReservationsSheet(ByVal ws As Worksheet)
    Const sheetPassword As String = "VehicleTrackerDemo"
    On Error Resume Next
    ws.Unprotect Password:=sheetPassword
    ws.Range("A1:M40").Locked = True
    ws.Range("C3,F3,C4:D4,L3").Locked = False
    ws.Protect Password:=sheetPassword, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Private Sub AlignDashboardStaticShapes()
    Dim ws As Worksheet
    Dim item As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Rows(16).RowHeight = 34
    ws.Rows(18).Hidden = True
    ws.Rows("27:28").Hidden = True
    ws.Rows(30).Hidden = False
    ws.Rows(30).RowHeight = 26
    StyleDashboardEntryHeader ws.Range("A5:D5")
    StyleDashboardEntryHeader ws.Range("E5:H5")
    StyleDashboardReportSelector ws
    StyleDashboardShopHeader ws.Range("A12:H12")
    FitDashboardShapeToRange ws, "btnCALD6D6", "D6", 2, 0
    FitDashboardShapeToRange ws, "btnCALH6H6", "H6", 2, 0
    FitDashboardShapeToRange ws, "btnSAVEMILEAGEA10B10", "A10:B10", 0, 2, 2, 0
    FitDashboardShapeToRange ws, "btnCLEARC10D10", "C10:D10", 2, 0, 2, 0
    FitDashboardShapeToRange ws, "btnMONTHLYMILEAGEA11D11", "A11:D11", 22, 22, 4, 4
    FitDashboardShapeToRange ws, "btnSAVECHECKOUTE11F11", "E11:F11", 0, 2
    FitDashboardShapeToRange ws, "btnCLEARG11H11", "G11:H11", 2, 0
    FitDashboardShapeToRange ws, "btnWEEKLYCHECKOUTBOARDE12H15", "E15:H15", 44, 44, 4, 4
    FitDashboardShapeToRange ws, "btnMARKINSHOPA15B15", "A15:B15", 0, 2
    FitDashboardShapeToRange ws, "btnMARKAVAILABLEC15D15", "C15:D15", 2, 0
    AlignDashboardUtilityShapes ws
    StyleDashboardSaveShape ws.Shapes("btnSAVEMILEAGEA10B10")
    StyleDashboardSaveShape ws.Shapes("btnSAVECHECKOUTE11F11")
    For Each item In Array("btnCALD6D6", "btnCALH6H6", "btnSAVEMILEAGEA10B10", "btnCLEARC10D10", "btnMONTHLYMILEAGEA11D11", "btnSAVECHECKOUTE11F11", "btnCLEARG11H11", "btnWEEKLYCHECKOUTBOARDE12H15", "btnMARKINSHOPA15B15", "btnMARKAVAILABLEC15D15")
        StyleDashboardShapeOutline ws.Shapes(CStr(item))
    Next item
    AlignDashboardPictureShapes ws
    FleetApplyDashboardPolish ws
    On Error GoTo 0
End Sub

Private Sub AlignDashboardPictureShapes(ByVal ws As Worksheet)
    PositionDashboardHeaderPicture ws, "v3LogHeaderIcon", "v2LogHeader", 24, 5
    PositionDashboardHeaderPicture ws, "v3ReserveHeaderIcon", "v2CheckoutHeader", 24, 10
    PositionDashboardHeaderPicture ws, "v3ShopHeaderIcon", "v2ShopHeader", 24, 5
    PositionDashboardHeaderPicture ws, "v3ReservationsHeaderIcon", "v2ReservationsHeader", 24, 10
    PositionDashboardHeaderPicture ws, "v3ReviewHeaderIcon", "v2ReviewTitleBand", 24, 6
    PositionDashboardCellPicture ws, "v3MileageDateIcon", "C6", 16, 16, 7
    PositionDashboardCellPicture ws, "v3CheckoutDateIcon", "G6", 16, 16, 7
    PositionDashboardBodyPicture ws, "v3ReservationBodyIcon", "E13:H14", 46, 18, 8
End Sub

Private Sub PositionDashboardHeaderPicture(ByVal ws As Worksheet, ByVal pictureName As String, _
    ByVal headerName As String, ByVal iconSize As Double, ByVal leftInset As Double)
    Dim pic As Shape
    Dim headerShape As Shape
    Set pic = ws.Shapes(pictureName)
    Set headerShape = ws.Shapes(headerName)
    pic.LockAspectRatio = msoFalse
    pic.Left = headerShape.Left + leftInset
    pic.Top = headerShape.Top + ((headerShape.Height - iconSize) / 2)
    pic.Width = iconSize
    pic.Height = iconSize
    pic.Placement = xlMove
    pic.ZOrder msoBringToFront
End Sub

Private Sub PositionDashboardCellPicture(ByVal ws As Worksheet, ByVal pictureName As String, _
    ByVal address As String, ByVal iconWidth As Double, ByVal iconHeight As Double, ByVal rightInset As Double)
    Dim pic As Shape
    Dim targetCell As Range
    Set pic = ws.Shapes(pictureName)
    Set targetCell = ws.Range(address)
    pic.LockAspectRatio = msoFalse
    pic.Left = targetCell.Left + targetCell.Width - iconWidth - rightInset
    pic.Top = targetCell.Top + ((targetCell.Height - iconHeight) / 2)
    pic.Width = iconWidth
    pic.Height = iconHeight
    pic.Placement = xlMove
    pic.ZOrder msoBringToFront
End Sub

Private Sub PositionDashboardBodyPicture(ByVal ws As Worksheet, ByVal pictureName As String, _
    ByVal address As String, ByVal iconSize As Double, ByVal leftInset As Double, ByVal topInset As Double)
    Dim pic As Shape
    Dim bodyRange As Range
    Set pic = ws.Shapes(pictureName)
    Set bodyRange = ws.Range(address)
    pic.LockAspectRatio = msoFalse
    pic.Left = bodyRange.Left + leftInset
    pic.Top = bodyRange.Top + topInset
    pic.Width = iconSize
    pic.Height = iconSize
    pic.Placement = xlMove
    pic.ZOrder msoBringToFront
End Sub

Private Sub StyleDashboardEntryHeader(ByVal rng As Range)
    rng.Interior.Color = RGB(0, 52, 101)
    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 34, 68)
        .Weight = xlMedium
    End With
End Sub

Private Sub StyleDashboardReportSelector(ByVal ws As Worksheet)
    Dim inputRange As Range
    With ws.Range("A17:C17")
        .Interior.Color = RGB(0, 34, 68)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A17:H17").Borders
        .LineStyle = xlContinuous
        .Color = RGB(215, 220, 225)
        .Weight = xlThin
    End With
    ws.Range("C17").ClearContents
    ws.Range("F17").Interior.Color = RGB(245, 246, 248)
    ws.Range("F17").Font.Color = RGB(25, 30, 35)
    For Each inputRange In Union(ws.Range("D17:E17"), ws.Range("G17")).Areas
        inputRange.Interior.Color = RGB(255, 242, 204)
        inputRange.Font.Color = RGB(0, 0, 0)
        With inputRange.Borders
            .LineStyle = xlContinuous
            .Color = RGB(215, 220, 225)
            .Weight = xlThin
        End With
    Next inputRange
    If Not ws.Range("H17").MergeCells Then
        With ws.Range("H17")
            .ClearContents
            .Interior.Color = RGB(0, 34, 68)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
End Sub

Private Sub StyleDashboardShopHeader(ByVal rng As Range)
    rng.Interior.Color = RGB(255, 255, 255)
    rng.Font.Color = RGB(255, 255, 255)
    With rng.Borders
        .LineStyle = xlNone
    End With
End Sub

Private Sub FitDashboardShapeToRange(ByVal ws As Worksheet, ByVal shapeName As String, ByVal address As String, _
    Optional ByVal leftInset As Double = 0, Optional ByVal rightInset As Double = 0, _
    Optional ByVal topInset As Double = 2, Optional ByVal bottomInset As Double = 2)
    Dim rng As Range
    Dim shp As Shape
    Set rng = ws.Range(address)
    Set shp = ws.Shapes(shapeName)
    shp.Left = rng.Left + leftInset
    shp.Top = rng.Top + topInset
    shp.Width = rng.Width - leftInset - rightInset
    shp.Height = rng.Height - topInset - bottomInset
    shp.Placement = xlMove
End Sub

Private Sub AlignDashboardUtilityShapes(ByVal ws As Worksheet)
    Dim rng As Range
    Dim names As Variant
    Dim index As Long
    Dim gap As Double
    Dim buttonWidth As Double
    Dim shp As Shape
    Set rng = ws.Range("A30:H30")
    names = Array("btnCLEARFIELDSA16H16", "btnUndoLastMileage", "btnUndoLastCheckout")
    gap = 2
    buttonWidth = (rng.Width - (2 * gap)) / 3
    For index = 0 To 2
        Set shp = ws.Shapes(CStr(names(index)))
        shp.Left = rng.Left + (index * (buttonWidth + gap))
        shp.Top = rng.Top + 3
        shp.Width = buttonWidth
        shp.Height = rng.Height - 6
        shp.Placement = xlMove
    Next index
End Sub

Private Sub StyleDashboardSaveShape(ByVal shp As Shape)
    shp.Fill.ForeColor.RGB = RGB(0, 122, 45)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
End Sub

Private Sub StyleDashboardShapeOutline(ByVal shp As Shape)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(0, 59, 92)
    shp.Line.Weight = 1.25
    shp.TextFrame.HorizontalAlignment = xlCenter
    shp.TextFrame.VerticalAlignment = xlCenter
End Sub


Private Sub RefreshWeeklyReservationsGrid(ByVal ws As Worksheet, ByVal selectedDate As Date)
    Dim vehicleTable As ListObject
    Dim reservationTable As ListObject
    Dim weekStart As Date
    Dim gridRange As Range

    Set vehicleTable = GetTable("tblVehicles")
    Set reservationTable = GetTable("tblReservations")
    EnsureVehicleSupportColumns vehicleTable
    selectedDate = DateOnly(selectedDate)
    weekStart = StartOfWeek(selectedDate)
    ws.Range("C3").value = Format(selectedDate, "mmmm")
    ws.Range("F3").value = Year(selectedDate)
    ws.Range("C4").value = selectedDate
    ws.Range("C4").numberFormat = "m/d/yyyy"
    ws.Range("C5").value = Format(weekStart, "mmmm d") & " - " & Format(DateAdd("d", 6, weekStart), "mmmm d, yyyy")
    Set gridRange = ws.Range("A15:J25")
    gridRange.UnMerge
    gridRange.Clear
    gridRange.Font.Name = "Aptos"
    gridRange.Font.Size = 10
    gridRange.Font.Color = RGB(0, 0, 0)
    gridRange.Interior.Color = RGB(255, 255, 255)
    gridRange.VerticalAlignment = xlCenter
    DrawWeeklyReservationBlock ws, 15, weekStart, vehicleTable, reservationTable
    StoreWeeklyReservationsCache ws, weekStart, WeeklyReservationSignature(reservationTable, vehicleTable, weekStart)
    ProtectWeeklyReservationsSheet ws
End Sub

