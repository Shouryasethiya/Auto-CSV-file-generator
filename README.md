# Attendance CSV Generator

A Google Apps Script that turns a wide-format Google Sheets attendance grid (one column per class, per day) into a clean, long-format CSV ready for import into a database or LMS — for a single day at a time, filtered and mapped automatically.

No downloads, no manual copy-pasting, no re-running a Python script every day. It runs from a menu button inside the Google Sheet itself and always reads the live data.

## Output format

```
student_email,subject_id,attend_date,status,attend_time
```

| Column | Example | Notes |
|---|---|---|
| `student_email` | `aarushv371@gmail.com` | Only students with a real email are included |
| `subject_id` | `102` | Looked up from the mapping tab by class code |
| `attend_date` | `2026-08-31` | Normalized to `YYYY-MM-DD` |
| `status` | `present` | One of `present`, `absent`, `late` |
| `attend_time` | `09:30` | Class start time, normalized to `HH:MM` |

## Expected sheet layout

**Attendance tab** — matches a typical multi-header attendance export:

| Row | Content |
|---|---|
| 1 | Date, merged across all of that day's class columns (e.g. `31-Aug`) |
| 2 | Class code (e.g. `U26AIMLB102`) |
| 3 | Class timing (e.g. `9:30-10:30`) |
| 4 | Subject name (not used in the output, informational only) |
| 5 | Blank |
| 6+ | Student rows |

**Columns:**

| Column | Content |
|---|---|
| A | Serial No. |
| B | Name |
| C | Phone Number |
| E | Email — the one actually used to export |
| F onward | One column per class, per day, containing `Present` / `Absent` / `Late` |

**Subject mapping tab** — two columns, header row in row 1:

| Class Code | subject_id |
|---|---|
| U26AIMLB102 | 101 |
| U26AIMLPC103 | 102 |

If your sheet's layout differs (different columns, different starting row, mapping columns in the other order), everything is adjustable in one place — see [Configuration](#configuration).

## Setup

1. Open your Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Delete any placeholder code in `Code.gs` and paste in the contents of [`attendance_csv_generator.gs`](./attendance_csv_generator.gs).
4. Edit the `CONFIGURATION` block at the top of the script to match your actual tab names and column layout (see below).
5. Save, then close the Apps Script tab and reload your Google Sheet.
6. A new **Attendance Tools** menu appears next to Help. Click **Attendance Tools → Generate CSV**.
7. On first run, Google will show a one-time authorization prompt since it's a script tied to your own account — click through **Advanced → Go to project (unsafe) → Allow**. This is expected for any script you write yourself and only happens once.

## Usage

1. Click **Attendance Tools → Generate CSV**.
2. Enter the date to export, in `YYYY-MM-DD` format (e.g. `2026-08-31`), even though the sheet header itself shows it as `31-Aug`.
3. If any class code that day isn't in the mapping tab, you'll get a warning listing exactly which codes are missing, with the option to continue anyway or cancel and fix the mapping first.
4. The CSV downloads straight to your computer as `attendance_YYYY-MM-DD.csv`, grouped by class (all students for one class, then the next class), rather than by student.

## Configuration

All adjustable values live in one block at the top of the script:

```javascript
var ATTENDANCE_SHEET_NAME = 'Attendance';       // your attendance tab's exact name
var MAPPING_SHEET_NAME    = 'Subject Mapping';  // your mapping tab's exact name

var MAPPING_CLASS_CODE_COL = 1; // column in mapping tab holding the class code
var MAPPING_SUBJECT_ID_COL = 2; // column in mapping tab holding the subject_id

var SNO_COL   = 1; // A
var NAME_COL  = 2; // B
var PHONE_COL = 3; // C
var EMAIL_COL = 5; // E -- the email column actually used

var DATE_ROW       = 1;
var CLASS_CODE_ROW = 2;
var TIMING_ROW      = 3;
var SUBJECT_ROW     = 4;

var FIRST_DATA_ROW  = 6; // first row of student data
var FIRST_CLASS_COL = 6; // first class column (F)
```

Tab names must match **exactly** — including spacing and capitalization — or the script will alert you that it can't find the tab.

## Behavior notes

- Merged date cells (row 1) are forward-filled internally, so every class column resolves to the correct date even though only the leftmost cell in a merged block holds a value.
- Rows with a blank or `#N/A` email are skipped entirely.
- Attendance cells with no recognizable status (blank, or anything other than present/absent/late) are skipped.
- Class codes with no matching row in the subject mapping tab are skipped, with a warning before the file downloads — the script never guesses or leaves `subject_id` blank.

## License

Use and modify freely for your own attendance workflow.
