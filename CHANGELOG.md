# Changelog

## 2.1.8 - 2026-08-12

- Expanded the controlled Edit Mileage and Edit Reservation menus to the 15 most recent transactions.
- Added and verified dedicated callbacks for entries 1 through 15, including grouped multi-day reservations.
- Re-ran formula, button-target, workbook-protection, privacy, and navigation-performance audits without changing the synthetic operational data.

## 2.1.7 - 2026-08-12

- Added synthetic reservations across every August week through August 31 for a complete calendar demonstration.
- Expanded the August mileage example so Demo SUV reaches 600 miles and displays the green compliant state.
- Verified Scheduled Hours against the operational branch implementation: each unique valid reserved date contributes eight hours, while cancelled and maintenance rows are excluded.
- Retained the zero-mile Demo Crossover example with its July ending odometer carried forward into August.
- Re-ran workbook integrity, formula, protection, privacy, visual, and navigation-performance audits.

## 2.1.6 - 2026-08-11

- Revised monthly mileage to subtract the latest valid ending odometer through the prior month from the latest valid ending odometer in the current month.
- Retained the prior ending odometer and a zero monthly total when a vehicle has no new reading in the selected month.
- Added a sanitized August demonstration month with three active vehicles and one zero-mile carry-forward example.
- Rebuilt and reconciled all operational history and recalculated FYTD totals from the revised monthly results.

## 2.1.5 - 2026-08-11

- Corrected monthly mileage totals to use the first and last valid ending-odometer readings within each month, with controlled fallbacks for single-entry and imported adjustment rows.
- Carried the latest positive ending odometer forward when a later month has no new ending reading.
- Corrected FYTD mileage to sum the reconciled monthly totals instead of raw trip-mile values.
- Rebuilt and reconciled every historical month in the demonstration workbook and re-ran package, macro, navigation, formula-error, and data-integrity checks.

## 2.1.4 - 2026-07-21

- Added the verified four-week initial functional release timeline and aligned the public narrative around learn-first, evidence-based, targeted process improvement.

## 2.1.3 - 2026-07-20

- Normalized heading and body spacing, aligned cards, captions, and bullets, and completed final two-page visual consistency QA.

## 2.1.2 - 2026-07-20

- Rebuilt the public two-page portfolio overview for cleaner executive readability, corrected measured text flow and graphics, and refreshed visual QA.

## 2.1.1 - 2026-07-20

- Updated the synthetic Demo SUV example to 530 July miles, demonstrating the green `>=500` target threshold.
- Refreshed the workbook screenshots and portfolio overview visuals.

## 2.1.0 - 2026-07-20

- Prepared the final public portfolio release and aligned all reviewer-facing documentation.
- Documented support for multiple non-overlapping same-day bookings and half-open `[start, end)` boundary logic.
- Clarified multi-day reservation storage and grouped edit/delete behavior.
- Strengthened coverage of maintenance history, reporting, controls, VBA architecture, testing, and synthetic-data privacy.
- Removed organization-specific context and avoided unmeasured savings claims.

## 2.0.0 - 2026-07-20

- Rebuilt the public demonstration workbook from the finalized application workflow.
- Replaced undo-only controls with recent-entry edit and delete workflows.
- Added grouped editing for multi-day reservations.
- Added scheduled-hours reporting to the dashboard.
- Added dated maintenance history across monthly, weekly, and mileage-log views.
- Added dedicated reservation-calendar month navigation.
- Refreshed the synthetic demonstration data and screenshots.
- Expanded functional tests for save, edit, delete, recalculation, and report synchronization.
- Re-ran workbook, VBA, metadata, package, and repository privacy audits.

## 1.0.0 - 2026-07-17

- Published a sanitized portfolio edition with synthetic demonstration data.
- Included the macro-enabled workbook and complete exported VBA source.
- Added dashboard, monthly mileage, and weekly reservation screenshots.
- Added portfolio, architecture, testing, privacy, and usage documentation.
