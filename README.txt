TO be able to run or test this script, Python 3 should have been installed through the Software Centre.

RUNNING:
To use the hours budgeting script, use the RUN batch file.
This ensures you have the correct python packages installed.

TESTING:
To test, use the TEST batch file. These tests have become obsolete from version 1 to version 2 of the software. Version 1 is not available in this repo.

REQUIREMENTS AND DEFINITIONS:
DEF: "Accumulated hours" in the hours budgeting file, which may be so mentioned in the log files, is data which is formatted with default/black colour and are the hours calculated from the SAP extract file.
REQ: The employee names in the budgeting file must be copied directly from the SAP extract table. DO NOT type in the names yourself and do not remove diacritics or hyphenation.
REQ: The row of employee names in the budgeting file must be located between the cells containing the data exactly as displayed below in between the ---, including new lines:
(start cell, first column)
---
MA Name
Startdatum
---
(end cell)
---
Anzahl
MA
---
These have been chosen due to their presence in all budgeting files.
DEF: Planned/predicted hours are those with font colour "White, background 1, darker 35%" (Weiß, Hintergrund 1, dunkler 35%). This definition is important for the logging of compared planned to recorded hours.