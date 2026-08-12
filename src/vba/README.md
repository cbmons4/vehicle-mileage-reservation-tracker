# VBA Source Map

This directory contains the complete plain-text export of the workbook's VBA project. It allows reviewers to assess Charles Monsanto's implementation without opening the macro-enabled workbook.

## Architecture

The project is intentionally lightweight: Excel `ListObject` tables provide local storage, worksheet modules dispatch user events, and three standard modules contain the application logic. There is no server, external database, add-in, API, or external data connection.

| File | Responsibility |
|---|---|
| `modFleetMileageTool.bas` | Entry validation, table operations, reservation conflict rules, maintenance history, reports, calendars, navigation, sheet protection, and Excel state controls |
| `modFleetDashboardLayout.bas` | Dashboard layout, selector popups, action styling, and visual alignment |
| `modFleetUndoActions.bas` | Recent-entry menus, mileage correction, grouped multi-day reservation edit/delete, conflict revalidation, and dependent recalculation |
| `ThisWorkbook.cls` | Workbook-open initialization and protection setup |

`modFleetUndoActions.bas` retains its original filename for workbook compatibility. Its current role is a controlled correction workflow rather than a simple undo operation.

## Transaction Pattern

User actions follow a consistent sequence:

1. Read values from controlled dashboard inputs.
2. Validate required fields, data types, dates, times, odometers, and selected records.
3. Apply business rules before changing a source table.
4. Write or revise the relevant `ListObject` rows.
5. Recalculate affected odometers, maintenance state, metrics, and reports.
6. Restore Excel events, calculation, and screen state on success or failure.

The 15 most recent mileage and reservation records are loaded back into the dashboard instead of requiring direct table edits. Multi-day reservations remain grouped while their date range is expanded, contracted, moved, or deleted.

## Reservation Rules

- One reservation row represents one vehicle-date; an inclusive multi-day range creates multiple associated rows.
- Multiple reservations may share a vehicle-date when their times do not overlap.
- Time intervals use half-open `[start, end)` logic: 8:00 AM-12:00 PM may be followed by 12:00 PM-5:00 PM.
- An overlap exists when `newStart < existingEnd` and `existingStart < newEnd`.
- Maintenance blocks the affected vehicle-date regardless of reservation time.
- Grouped edits exclude their existing rows from self-conflict checks before the replacement range is committed.
- Grouped deletion removes every dated row in the selected multi-day reservation.

## Reports and Historical State

The VBA derives weekly reservations, monthly mileage, ending odometers, trip days, trip counts, scheduled hours, fiscal-year totals, and vehicle status from structured tables. Dated maintenance records remain available after return to service and flow into weekly, monthly, and mileage-log views.

## Worksheet Modules

Worksheet document modules use codename-plus-sheet filenames: `Sheet1_Dashboard.cls`, `Sheet2_Setup.cls`, `Sheet3_Vehicles.cls`, `Sheet4_Drivers.cls`, `Sheet5_Reservations.cls`, `Sheet6_Mileage_Log.cls`, `Sheet7_DashboardData.cls`, `Sheet8_Monthly_Mileage.cls`, `Sheet9_Weekly_Reservations.cls`, and `Sheet10_Import_Summary.cls`. Some contain event procedures; others are intentionally empty. Keeping event handlers thin makes the standard modules the primary review surface for business logic.

## Reviewing the Export

The `.bas` and `.cls` files are reviewable text copies. The executable project remains embedded in `Vehicle Mileage and Reservation Tracker.xlsm`.

Excel document modules cannot be imported as ordinary class modules without reconnecting them to their workbook objects. Use this export for code review, version history, testing reference, and controlled maintenance.

All public demonstration data is synthetic; the source map contains no organization-specific operating context.
