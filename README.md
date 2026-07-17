# VBA Source Map

This directory contains a complete export of the workbook's VBA project so the automation can be reviewed in GitHub without opening the binary workbook.

## Core Modules

| File | Responsibility |
|---|---|
| `modFleetMileageTool.bas` | Data entry, validation, reporting, maintenance status, calendars, navigation, protection, and performance controls |
| `modFleetDashboardLayout.bas` | Dashboard layout, selector popups, and visual alignment |
| `modFleetUndoActions.bas` | Reversible mileage and reservation transactions |
| `ThisWorkbook.cls` | Workbook-open initialization |

## Worksheet Modules

`Sheet1.cls` through `Sheet10.cls` are the exported worksheet document modules. They are included for completeness; some contain event procedures while others are intentionally empty.

## Reviewing or Reusing the Export

The `.bas` and `.cls` files are plain-text exports. The executable copy of the project remains embedded in `Vehicle Mileage and Reservation Tracker.xlsm`.

Excel document modules cannot be imported as ordinary class modules without reconnecting them to their corresponding workbook objects. Use the source files primarily for review, version history, and controlled maintenance.

