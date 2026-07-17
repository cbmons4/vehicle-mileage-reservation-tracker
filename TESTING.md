# Testing and Verification

The portfolio workbook was tested through a disposable copy so save and undo operations could be exercised without changing the delivered demonstration data.

## Functional Tests

The following workflows passed:

- workbook opens to the portfolio dashboard;
- VBA project is present and viewable;
- synthetic vehicle and driver tables load correctly;
- dashboard navigation to Monthly Mileage and back;
- dashboard navigation to Weekly Reservations and back;
- previous-week and next-week navigation;
- mileage save, vehicle odometer update, and mileage undo;
- reservation save and reservation undo;
- formula error scan;
- Excel events and screen updating restored after automation; and
- macro-enabled package integrity.

## Measured Navigation Times

| Operation | Time |
|---|---:|
| Open workbook | 4.600 s |
| Dashboard to Monthly Mileage | 0.252 s |
| Monthly Mileage to Dashboard | 0.045 s |
| Dashboard to Weekly Reservations | 0.049 s |
| Next reservation week | 0.299 s |
| Previous reservation week | 0.309 s |
| Weekly Reservations to Dashboard | 0.332 s |

Timings are from one Windows desktop Excel test environment and are included as an implementation reference, not as a service-level guarantee.

## Privacy and Package Checks

The delivered workbook and repository were checked for:

- organization-specific names, identifiers, email domains, and branding;
- personal machine paths and temporary build references;
- external workbook relationships and hyperlinks;
- custom document properties and embedded thumbnails;
- residual operational records;
- missing or corrupt VBA package parts; and
- common formula errors.

The portfolio workbook uses synthetic demonstration records and contains no external data connections.

## Manual Review Checklist

Before demonstrating the workbook:

1. Download the `.xlsm` file from the repository.
2. Open it in desktop Excel and enable macros.
3. Log a sample mileage entry and undo it.
4. Create a sample reservation and undo it.
5. Mark a vehicle in maintenance and confirm that it appears unavailable.
6. Navigate between the dashboard, monthly mileage, and weekly reservations views.
