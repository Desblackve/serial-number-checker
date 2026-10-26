Creator: Desblack "Dess"
Contact: desblackve@gmail.com


SERIAL NUMBER CHECKER
=====================

Version 2.1A


1. WHAT THE APP DOES
--------------------

Serial Number Checker has two modes (buttons at the top left):

- CHECKER: loads a list of serial numbers (S/N) from a CSV or Excel file
  and checks, by scanning or typing, which ones are there, which are
  missing and which are extra. When you finish, it exports a colored
  Excel report.
- CREATE S/N LIST: builds a new list from what you scan, organized by
  pallets (in columns or sheets) or, with Rack mode, by the locations of
  a rack. It exports it to Excel or CSV.

The interface is available in Spanish, English and Portuguese. The
README button (top right) opens this guide inside the app.


2. QUICK EXAMPLES
-----------------

EXAMPLE 1 - Check an inventory against a file

  1. Press "Load CSV / Excel" and choose the file (for example,
     inventory_october.xlsx).
  2. Scan the items one by one. Each read shows:
       Green      FOUND             (Sheet Lot1 · row 7)
       Yellow     ALREADY SCANNED   (the counter does not go up)
       Red        NOT FOUND         (saved in Extras)
  3. Look at the counter: "128 / 340" and "38% done · 212 pending".
  4. Press "Export report": inventory_october_resultado.xlsx is created
     with the S/Ns in green (scanned), red (pending) and blue
     (duplicate), plus the Summary, Scans and Extras sheets.

EXAMPLE 2 - I scanned the wrong code

  If you scan the wrong code, press Ctrl+Z (or the "Undo" button in
  Latest scans): the last scan is cancelled and the counter goes back to
  what it was.

EXAMPLE 3 - I have the S/Ns in an email or another Excel sheet

  Copy the list (one per line) and paste it with Ctrl+V into the scan
  field. The app processes all of them at once and shows a summary:
  "Pasted 25: 22 found · 2 not found · 1 repeated".

EXAMPLE 4 - Build a list of pallets

  1. Switch to "Create S/N List" mode.
  2. Scan the S/Ns of the first pallet: they are saved in column A
     (Pallet 1).
  3. Press "Next pallet" (or F2): the next ones go to column B
     (Pallet 2), and so on.
  4. If you want a pallet on a separate sheet, press "New sheet" next to
     the sheet tabs.
  5. Press "Export list" (or Ctrl+E) to save it to Excel or CSV.

EXAMPLE 5 - Fill a rack location by location (Rack mode)

  1. In "Create S/N List", tick "Rack mode".
  2. Type the location (for example, DC05), the rows and columns of the
     rack (for example, 6 and 8) and the order: Horizontal or Vertical.
  3. Press "Add new rack": DC05-1 is created.
  4. Scan: the counter shows the next location, for example (0,0), and
     each S/N is saved there. Then (1,0), (2,0)...
  5. When the rack is full, the app warns "Rack DC05-1 is complete".
     Press "Add new rack" to continue with DC05-2.

EXAMPLE 6 - Check that each S/N goes on the right pallet

  1. In "Create S/N List", tick "Validate list" and press "Load list" to
     choose the CSV or Excel file with the S/Ns that belong on the
     pallets.
  2. Scan as usual. If an S/N is not in that list, it is not recorded and
     "S/N NOT VALID FOR THE PALLET" appears in red.
  3. The counter shows how many you have and how many are left, for
     example "Validated 45 of 60 · 15 left".

EXAMPLE 7 - See the progress from a distance

  Press F11 (or "Big view"): the screen shows only the counter and the
  result in a very large size. F11 or Esc to go back.


3. WHAT'S NEW
-------------

VERSION 2.1A - Validate list and sheet names

- Validate list (Create list): only S/Ns from a loaded list are
  recorded; the others get a red alert and are not added.
- With Validate list you can also use "A/B/U alternatives".
- The first sheet is called "Pallet 1" and any sheet can be renamed by
  double-clicking its tab.
- A short description appears when you hover over "Rack mode",
  "Validate list" and the sheet tabs.
- When exporting to Excel, columns have the exact width for the longest
  S/N.
- Checker: copy the pending S/N with Ctrl+C or with right-click → "Copy
  S/N".

VERSION 2.0 - New design and Rack mode

- New design: dark cards over a line-pattern background, buttons with
  icons and a dark title bar on Windows.
- Windows executable (SerialNumberChecker.exe): Python no longer needs
  to be installed.
- README button: this guide can be read inside the app.
- Rack mode: a list guided by the (x,y) locations of each rack.
- Extras: S/Ns that are not in the file are kept in their own tab and
  exported on a separate sheet.
- Undo the last scan (Ctrl+Z or the "Undo" button).
- Sheet and row: when an S/N is found, the app shows where it is in the
  file.
- Format filter: if another code on the box is scanned (for example, one
  with 13 characters), it warns "DOES NOT LOOK LIKE AN S/N" and does not
  count it.
- Paste several S/Ns at once with Ctrl+V.
- Counter with a progress bar, pace (S/Ns per minute) and estimated time
  to finish.
- Pending / Extras tabs with a search box.
- Big view (F11).
- The report includes the Summary, Scans (time of each read) and Extras
  sheets.
- Print (Ctrl+P): summary, pending and extras (Checker) or the current
  sheet / all sheets (Create list).
- Create list: "New sheet" button, columns that can be stretched and the
  Backspace key to delete the selected record.
- More S/N headers are recognized (see section 15).
- The session is saved automatically: when you open the app again you
  continue where you left off.


4. CHECKER MODE, STEP BY STEP
-----------------------------

  1. Open Serial Number Checker.
  2. If needed, change the language in the selector at the top right
     (Español, English or Português). The change is immediate.
  3. Press "Load CSV / Excel" and select the file with the S/Ns.
  4. The file name, the values read and the unique S/Ns are shown.
  5. If the file is a report exported before, the app asks whether you
     want to resume where you left off (section 9).
  6. Scan or type an S/N in the scan field.
  7. At 17 characters the check is automatic. You can also press ENTER.
  8. Repeat with all the S/Ns you want to check.
  9. When you finish, press "Export report".
  10. To start from scratch, use the "Reset" button (section 16).

The "A/B/U alternatives" box in the top bar turns the alternative search
of section 5 on or off.


5. MATCHING RULES
-----------------

- First, an exact match with a loaded S/N is searched for.
- If there is none and "A/B/U alternatives" is ticked, the seventh
  character is replaced with A, B and U. If an alternative matches, the
  S/N that actually appears in the list is marked as found (the message
  says "variant" and what was read).
- An S/N already found does not add to the counter again.
- If the code read does not have the length of any S/N in the file, the
  app warns "DOES NOT LOOK LIKE AN S/N" (probably another code on the box
  was scanned) and it is not saved as an extra.

IMPORTANT: replacing the seventh character is an alternative search. It
does not guarantee that the serial number is genuine and it is not a
complete format validation.


6. COLORS, SOUNDS AND MESSAGES
------------------------------

In the app:

- GREEN: S/N found. Confirmation sound. The message says whether it was
  exact or a variant, and on which sheet and row of the file it is.
- YELLOW: S/N already scanned. Another sound. The counter does not go up
  and the app stays yellow until the next scan or until you press Esc.
- RED: S/N not found (saved in Extras) or a code that does not look like
  an S/N. Error sound.
- The border of the scan field flashes in the color of the result.
- Below the scan field there is a legend with the report colors.

In the Excel report:

- GREEN: S/N found during the check.
- RED: S/N not found yet.
- BLUE: S/N duplicated in the file. The first occurrence keeps its green
  or red and the repetitions (on any sheet) are blue.

If it was found through an A/B/U alternative, the real S/N in the list
is colored.


7. COUNTER, PROGRESS AND PACE
-----------------------------

The counter shows the S/Ns found out of the total unique S/Ns:

    128 / 340
    38% done · 212 pending
    12/min · ~18 min left

- An S/N repeated in the file counts only once.
- When a file is loaded, the app shows how many values were read and how
  many are unique. Example: "12 values read · 10 unique S/Ns" means there
  are 2 repeated values in the file.
- The pace and estimated time appear after a few scans and are
  calculated with the last 10 minutes.


8. PENDING, EXTRAS AND HISTORY
------------------------------

- LATEST SCANS: the last 5 S/Ns read and their result. The "Undo" button
  (or Ctrl+Z) cancels the last scan; it can be repeated.
- PENDING: the S/Ns in the file that have not been scanned yet.
- EXTRAS: the scanned S/Ns that are not in the file, with how many times
  they were read and the time of the first read.
- COPY: select an S/N in Pending or Extras and press Ctrl+C, or
  right-click it and choose "Copy S/N".
- SEARCH: filters the open tab as you type. Enter or Esc return the
  cursor to the scan field.


9. RESUME AN EXPORTED REPORT
----------------------------

If you load an XLSX report exported before by the app, you can continue
where you left off.

- The app detects the green S/Ns in the report and asks in the yellow
  alert whether you want to resume.
- "Resume": those S/Ns are marked as found. "Start from scratch": the
  file is loaded without progress.
- The Summary, Scans and Extras sheets of the report are ignored when
  reading and replaced when exporting again.
- CSV files cannot be resumed because they do not keep colors.

NOTE: the exact green of the report is recognized. If a file has cells
with that same green for another reason (for example, Excel's "Good"
style), it will ask as well; in that case choose "Start from scratch".


10. EXPORT REPORT
-----------------

"Export report" (or Ctrl+E) saves an Excel file (.xlsx) with the result.

- If the loaded file is XLSX, the report keeps its sheets.
- If it is CSV, an XLSX is created so the cells can be colored.
- The S/N columns are colored green, red or blue (section 6).
- The S/N columns have the exact width for the longest S/N.
- Three sheets are added:
    Summary      file, totals, found, pending, extras and the times of the
                 first and last read.
    Scans        each read with its time, what was read and the result.
    Extras       the S/Ns that are not in the file (if any).
- The original file is not modified: the report is saved separately and
  cannot be saved over the original.
- Exception: if you are resuming a report (section 9), you can save the
  progress over that same report. Windows will ask you to confirm that
  you want to replace it.
- The suggested name is file_name_resultado.xlsx.


11. CREATE S/N LIST MODE
------------------------

Used to build a new list with what you scan. It accepts any text (case
is kept and there is no length limit).

- DESTINATION of the next pallet:
    New column         each pallet in the next column (default).
    New sheet          each pallet on its own sheet.
    From the cell      starts at the cell selected in the viewer.
- NAME (optional): title of the pallet or sheet. If left empty, "Pallet
  1", "Pallet 2"... are used.
- "Next pallet" (F2): starts the next pallet.
- "New sheet" (next to the sheet tabs): creates a new sheet and jumps to
  it.
- SHEET NAMES: the first one is called "Pallet 1". Double-click the tab
  of any sheet to rename it (Enter to save, Esc to cancel). The new name
  is used when exporting and printing.
- VALIDATE LIST: checks that each S/N is in a loaded list (section 12).
- VIEWER: shows the sheet exactly as it will be exported. Click to select
  a cell, double-click to edit it. Columns are stretched by dragging
  their border in the letter row; double-click the border to fit the
  text.
- DUPLICATES: they come in blue and the app stays yellow with an alert:
  "Delete duplicates" or "Hide alert".
- DELETE: Backspace or Del with the field empty, or the "Delete" button,
  deletes the selected record. "Undo" (Ctrl+Z) brings it back.
- EXPORT LIST (Ctrl+E): to Excel (with the same sheets and colors as the
  viewer) or to CSV (S/N, pallet, sheet, cell and whether it is a
  duplicate). In Excel the columns have the exact width for the longest
  S/N; a CSV file does not keep column widths.
- A value that starts with "=" is saved as text, not as a formula.


12. VALIDATE LIST
-----------------

To make sure only the right S/Ns go on the pallets. It is turned on with
the "Validate list" box in the top bar of Create list mode.

- LOAD LIST: choose a CSV or Excel file with the expected S/Ns (the same
  headers as in Checker are recognized, section 15).
- NO LIST: until you load one, the app shows a yellow alert and does not
  record anything you scan.
- VALID S/N: it goes on the pallet as usual.
- DUPLICATE S/N: it comes in blue, as always.
- S/N NOT IN THE LIST: it is not recorded and "S/N NOT VALID FOR THE
  PALLET" appears in red, with the error sound.
- LEFT: the counter shows "Validated X of Y · Z left".
- A/B/U ALTERNATIVES: with Validate list on, this box also appears, with
  the same rule as in Checker (section 5).
- RACK MODE: Rack mode only places machines and does not validate, so
  each option locks the other (grey): while Rack mode is on, "Validate
  list" cannot be ticked, and while Validate list is on, "Rack mode"
  cannot be ticked.


13. RACK MODE
-------------

To store items in racks, location by location. It is turned on with the
"Rack mode" box in the top bar of Create list mode.

- LOCATION: the physical place, for example DC05. The racks are named
  DC05-1, DC05-2, DC05-3...
- ROWS and COLUMNS: size of the rack (6 rows and 8 columns by default).
- COORDINATES (x,y): x is the column and y the row, starting at (0,0) at
  the top left. The viewer counts from 0.
- ORDER: Horizontal fills (0,0), (1,0), (2,0)... row by row. Vertical
  fills (0,0), (0,1), (0,2)... column by column.
- "ADD NEW RACK" (or F2, or next to the tabs): creates the next rack of
  the location and starts at (0,0).
- GUIDE: the counter shows the next location in large size and the
  viewer marks it with a green box. Each scan is saved there and moves
  on.
- JUMP: click an empty location in the viewer and the next scan goes
  there.
- RACK COMPLETE: when the last location is filled the app warns "Rack
  DC05-1 is complete" and does not record anything else in that rack
  until you add a new one.
- EXPORT: one sheet per rack drawn like the shelf (x coordinates at the
  top and y on the left) and a "List" sheet with Location, Rack, Position
  (x,y), X, Y, S/N and whether it is a duplicate.


14. PRINT AND BIG VIEW
----------------------

PRINT ("Print" button or Ctrl+P): the app asks in the yellow alert what
to print.

- Checker: tick Summary, Pending and/or Extras and press "Print".
- Create list: "Current sheet" or "All sheets".
- It prints on the Windows default printer. If the computer has Excel,
  it prints with Excel; otherwise, with Notepad.

BIG VIEW (F11): only the counter and the last result in large size, to
see them from a distance. You can keep scanning. F11 or Esc to go back.


15. SUPPORTED FILES AND HEADERS
-------------------------------

Formats:

- XLSX: Excel files. All the sheets of the workbook are read.
- CSV: a single table. The separator (comma, semicolon, tab or |) is
  detected automatically.

Headers recognized as an S/N column (regardless of case, accents, spaces
or punctuation):

    SN, S/N, NS, N/S
    SERIAL, SERIAL NO, SERIAL NR, SERIAL NBR, SERIAL NUM, SERIAL NUMBER
    SERIE, N SERIE, Nº DE SERIE, NO SERIE, NO. DE SERIE
    NUM SERIE, NUM. DE SERIE, NÚMERO SERIE, NÚMERO DE SERIE
    N SERIAL, NÚMERO SERIAL

Reading rules:

- If there are several S/N headers (in different columns or sheets), all
  of them are read.
- Columns that are not S/N columns are not used.
- Empty cells and NONE or NAN values are ignored.
- If a cell has a formula, the calculated value saved by Excel is used.
- Numbers are read without decimals (123456.0 is read as 123456).


16. SHORTCUTS, RESET AND SESSION
--------------------------------

Keyboard shortcuts:

    Enter        check / record
    Ctrl+Z       undo
    Ctrl+E       export
    Ctrl+P       print
    Ctrl+V       paste several S/Ns
    Ctrl+C       copy the S/N selected in Pending or Extras
    F11          big view
    Esc          dismiss the alert / leave the big view
    F2           next pallet or new rack (Create list)
    Backspace    delete the selected record (Create list, with the field
                 empty). Del does the same.

RESET: the "Reset" button (bottom right) clears the session of the
active mode. If there is data, the first click shows "Sure?" and you
must click again within 3 seconds. The original file is not deleted from
your computer.

SESSION: progress is saved automatically. If you close the app, when you
open it again you get back the loaded file, what was scanned, the extras
and the list or racks being created. Even so, export the report before
resetting.


17. REQUIREMENTS AND INSTALLATION
---------------------------------

EXECUTABLE (recommended):

- Unzip the whole folder and open SerialNumberChecker.exe. Do not
  separate it from the _internal folder.
- It does not need Python or any other installation.
- If Windows shows "Windows protected your PC", press "More info" and
  "Run anyway". It happens because the program is not digitally signed.

FROM THE CODE (optional):

- Python 3 for Windows (includes Tkinter) and the openpyxl library:

      pip install openpyxl

- Then run serial_number_checker.py.
- To build the executable, use build_exe.bat.

Sounds work on Windows. The app icon is arthur.ico.


18. RECOMMENDATIONS
-------------------

- Make sure the loaded file is the right one.
- Check that the S/N columns have a recognized header (section 15).
- Remember that all the sheets of XLSX files are read.
- If an S/N does not appear as found, check that it is written correctly
  and that it exists in an S/N column.
- Review the Extras tab before exporting.
- Export the report before resetting.
- Keep an unmodified copy of the original file.


19. GitHub
----------

Project, updates and downloads:

    https://github.com/Desblackve/serial-number-checker

------------------------------------------------------------
