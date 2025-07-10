Executable distribution available in /dist/
Built on Python 3.13.2. For further requirements, see /phb_app/requirements.txt.

DEFINITIONS:
DEF: "Accumulated hours" in the hours budgeting file, which may be so mentioned in the log files, is data which is formatted with default/black colour and are the hours calculated from the SAP extract file.
DEF: Planned/predicted hours are those with font colour "White, background 1, darker 35%" (Weiß, Hintergrund 1, dunkler 35%). This definition is important for the logging of compared planned to recorded hours.

REQUIREMENTS:
REQ: The employee names in the budgeting file must be copied directly from the SAP extract table. DO NOT type in the names yourself and do not remove diacritics or hyphenation.
REQ: The row of employee names in the budgeting file must be located between the cells containing the data exactly as displayed below, including new lines, excluding indents:
(start cell, first column)
    Anställds namn
    Datum
(end cell)
    Antal
    anställda
REQ: All budgeting files must contain these above mentioned cells and positioned similarly.
REQ: The contents of these above mentioned cells may be changed 