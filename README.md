# Fire Timesheet

A browser-based timesheet for entering daily hours, viewing pay-period totals, and calculating FLSA overtime. The app is built with plain HTML, CSS, and JavaScript; it has no build step or package dependencies.

## Run

Open `index.html` in a modern browser. To serve it locally instead, run a static web server from this directory, then open the local URL it provides.

## Use

- Enter hours in the blue cells. Row totals, pay-period totals, pay-code summaries, and the FLSA calculation update as you type.
- Select **Update date ranges** to change the start and end dates for all three pay periods and the FLSA work period. Pay periods must be in chronological order and cannot overlap.
- Select **Load sample data** to load the example hours, or **Clear all** to remove entered hours. Use **Print** for a printable timesheet.

Hours and date ranges are saved separately in the current browser's local storage. They are not synced between browsers or devices. Clearing browser storage removes saved data. Loading sample data changes the hours but does not reset the date ranges.

## Calculation notes

- Regular summary hours are `112 - leave hours`.
- FLSA OT ½-rate hours are `max(0, regular hours in the FLSA work period - 182)`.
- Following the workbook convention, the FLSA start date is excluded from the counted days; hours are counted from the following day through the selected end date.
- FLSA OT ½-rate hours are included in the third pay period's OT ½ total.
