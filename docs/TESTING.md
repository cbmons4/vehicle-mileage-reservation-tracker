# Testing and Verification

Charles Monsanto designed the release checks around the application's highest-risk behaviors: transaction integrity, reservation boundaries, grouped corrections, historical status, report synchronization, Excel state restoration, and public-data privacy. Transaction tests run against a disposable workbook copy so the delivered synthetic demonstration data remains unchanged.

## Release Regression Coverage

The 2026-07-20 release candidate passed checks covering:

- workbook startup, dashboard initialization, and view navigation;
- presence and readability of the embedded and exported VBA projects;
- loading of synthetic vehicle, driver, reservation, maintenance, and mileage tables;
- mileage save, edit, delete, vehicle-odometer recalculation, and report refresh;
- single-day and multi-day reservation save, edit, and delete;
- grouped multi-day edits that expand, contract, or move the date range;
- grouped deletion without orphaned reservation dates;
- multiple non-overlapping bookings for one vehicle on the same day;
- conflict rejection for actual time overlaps and maintenance dates;
- dated maintenance history in the mileage log, monthly report, and weekly view;
- weekly reservations, monthly mileage, trip counts, scheduled hours, and fiscal-year totals;
- target-aware dashboard formatting at `>=500` miles and correct below-target formatting for the current synthetic dataset;
- expected transaction row counts after corrections;
- formula-error scanning and macro-enabled package integrity; and
- restoration of Excel events, calculation, and screen updating after automation.

The 2026-08-11 maintenance release additionally reconciled every vehicle-month in the generated mileage history against the raw mileage table, including first-to-last monthly odometer baselines, single-entry fallbacks, anomalous lower-odometer rows, zero placeholders, carried-forward positive ending mileage, and FYTD totals derived from corrected monthly results.

## Reservation Boundary Cases

The reservation suite verifies half-open interval logic, where a booking occupies `[start, end)`:

| Existing booking | Proposed booking | Expected result |
|---|---|---|
| 8:00 AM-12:00 PM | 12:00 PM-5:00 PM | Allowed: adjacent boundary |
| 8:00 AM-12:00 PM | 7:00 AM-8:00 AM | Allowed: adjacent boundary |
| 8:00 AM-12:00 PM | 11:00 AM-1:00 PM | Blocked: overlap |
| 8:00 AM-12:00 PM | 9:00 AM-10:00 AM | Blocked: contained interval |
| 8:00 AM-12:00 PM | 8:00 AM-12:00 PM | Blocked: duplicate interval |

Additional cases verify end time after start time, inclusive multi-day date expansion, multiple same-day bookings, maintenance conflicts on any date in a range, and conflict checks during grouped edits.

## Privacy and Package Checks

The workbook, exported source, and public repository text were checked for:

- organization-specific names, acronyms, identifiers, email domains, and branding;
- real staff, vehicle, mileage, reservation, location, and maintenance records;
- personal machine paths and temporary build references;
- sensitive strings in visible sheets, hidden sheets, exported VBA, and the embedded VBA project;
- external workbook relationships, hyperlinks, data connections, and custom properties;
- embedded thumbnails or stale package content;
- missing or corrupt macro-enabled package parts;
- broken Markdown links; and
- common formula errors.

The portfolio edition uses synthetic demonstration data and has no external data connections. No measured time or cost savings are claimed.

## Reviewer Walkthrough

1. Open the `.xlsm` file in desktop Excel and enable macros.
2. Save, edit, and delete a mileage entry; confirm the odometer and reports recalculate.
3. Create two adjacent same-day bookings, such as 8:00 AM-12:00 PM and 12:00 PM-5:00 PM; confirm both are accepted.
4. Attempt an overlapping booking; confirm it is blocked with no table write.
5. Create a multi-day reservation, edit its range, then delete it; confirm the dated rows remain grouped throughout.
6. Place a vehicle in maintenance and return it to service; confirm current availability changes while dated history remains visible.
7. Review weekly reservations, monthly mileage, scheduled hours, and fiscal-year totals after each relevant transaction.
