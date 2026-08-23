================================================================
  SCHOOL TIMETABLE GENERATOR — BUILD INSTRUCTIONS
  Package into single .exe for Windows (offline, no install)
================================================================

REQUIREMENTS (on your Windows PC):
--------------------------------------
  Python 3.10+ (from python.org)
  Install dependencies once:
    pip install pillow openpyxl python-docx reportlab pyinstaller

BUILD STEPS (run on Windows):
--------------------------------------
1. Copy these two files to a folder, e.g. C:\TimetableApp\
     - timetable_app.py
     - build_windows.spec

2. Open Command Prompt in that folder:
     cd C:\TimetableApp

3. Build the EXE:
     pyinstaller build_windows.spec

4. Find the finished app at:
     dist\SchoolTimetableGenerator.exe

5. The .exe is fully standalone — copy it anywhere, no Python needed.

QUICK BUILD (one-liner, no spec file needed):
--------------------------------------
  pyinstaller --onefile --windowed --name "SchoolTimetableGenerator" ^
    --hidden-import PIL._tkinter_finder ^
    --hidden-import openpyxl --hidden-import docx ^
    --hidden-import reportlab ^
    timetable_app.py

================================================================
  APP FEATURES
================================================================
  ① Enter school name and upload logo
  ② Choose section (Pre School/Primary/Middle/Senior/College)
  ③ Set opening time, number of periods, period duration,
     break duration, break after period N
  ④ Select classes (PG, KG, Nursery, Prep, 1-12) and sections (A-G)
  ⑤ Assign class in-charge teacher per class-section
  ⑥ Add subjects with multiple teachers (workload balanced)
  ⑦ Generate CLASS-WISE or TEACHER-WISE timetable
  ⑧ Preview on screen with color-coded table
  ⑨ Download as Excel (.xlsx), Word (.docx), or PDF

  RULES APPLIED:
  - Class in-charge always gets Period 1
  - Period 1 is 10 minutes longer than others
  - Multi-teacher subjects: workload distributed equally
  - Every subject appears at least once per class
  - Break/recess inserted after the chosen period
  - Single master timetable (applies to all days of week)
================================================================
