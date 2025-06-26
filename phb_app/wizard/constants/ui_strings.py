'''
String constants. English.
'''

from pathlib import Path
from enum import StrEnum, auto
import git

GUI_TITLE = "Project Hours Budgeting Wizard"

#           --- INTRO PAGE ---

INTRO_TITLE = "Introduction"

INTRO_SUBTITLE = "Welcome to the Project Hours Budgeting Wizard!"

INTRO_MESSAGE = """
This wizard allows the user to budget the hours of many employees of a project at once.

There are three steps:

1: Select the input files (SAP extracts or external timesheets) and the output budgeting file.

2: Select the employees of the project.

3: Review the process, checking/deleting the log file or restoring the budgeting file to its previous state before the writing of data.

NOTE: Close the output Excel file which will be used in this process!
"""

IMAGE_LOAD_FAIL = "Failed to load image."

def find_git_root():
    '''Find the root of the project.'''
    repo = git.Repo(Path(__file__).resolve(), search_parent_directories=True)
    return Path(repo.working_tree_dir)

APP_ROOT = find_git_root() / "phb_app"
IMAGES_DIR = APP_ROOT / "images"

#           --- IO SELECTION PAGE ---

IO_FILE_TITLE = "Input/Output Files"

IO_FILE_SUBTITLE = """
Manage the input timesheets and output project hours budgeting file.
NOTE: The selected date of the output file is the row in which the hours will be written but also will affect whether hours are accumulated at all from the input files.
"""

INPUT_FILE_INSTRUCTION_TEXT = "Select the timesheets as the input files."

OUTPUT_FILE_INSTRUCTION_TEXT = "Select the project hours budgeting file as the target output file."

ADD_FILE = "Add File"
EXCEL_FILE = "Excel files (*.xlsx)"

#           --- PROJECT SELECTION PAGE ---

PROJECT_SELECTION_TITLE = "Project Selection"

PROJECT_SELECTION_SUBTITLE = "Select the projects in which the hours were booked."

PROJECT_SELECTION_INSTRUCTIONS = """
<p><strong>Select one or more project IDs.<strong></p>
<p>The project decription is purely to help in the recognition of the wished after project ID. It will not be used in any calculations.</p>
"""

#           --- EMPLOYEE SELECTION PAGE ---

EMPLOYEE_SELECTION_TITLE = "Employee Selection"

EMPLOYEE_SELECTION_SUBTITLE = "Select the employees whose hours will be budgeted."

EMPLOYEE_SELECTION_INSTRUCTIONS = """
<p><strong>Select one or more employees by name.</strong></p>
<p><em>Coord</em> indicates the cell coordinate in the output budgeting file where the employee's name stands.</p>
"""

#           --- SUMMARY PAGE ---

SUMMARY_TITLE = "Summary"

SUMMARY_SUBTITLE = "Select the employees whose hours will be budgeted."

IO_SUMMARY = "IO Summary."

SUMMARY_INSTRUCTIONS = """
<p>Check all details before selecting which employee's hours should be recorded.</P>
<p>Missing hours will be omitted. Red predicted hours imply that hours have already been recorded and thus will not be overwritten. The project ID column displays from where the hours were taken.</p>
"""

class IORole(StrEnum):
    '''Input or output roles.'''

    INPUTS = auto()
    OUTPUT = auto()
    PROJECT_TABLE = auto()
    EMPLOYEE_TABLE = auto()
    SUMMARY_IO_TABLE = auto()
    SUMMARY_DATA_TABLE = auto()

class SpecialStrings(StrEnum):
    '''Enum for selecting worksheets.'''

    SELECT_WORKSHEET = "<select worksheet>"
    XLSX = ".xlsx"
    UTF_8 = "utf-8"
    DATA_ONLY_EXCEL = "_wizard_data_only.xlsx"
    ZERO_HOURS = "0.00"
    MISSING = "Missing"

class QPropName(StrEnum):
    '''QWizard property names.'''

    MANAGED_WORKBOOKS = "Managed Workbooks Property"
    SELECTED_PROJECTS = "Selected Projects"

class ButtonNames(StrEnum):
    '''Enum of names to display on buttons.'''

    ADD = "Add"
    REMOVE = "Remove"
    SELECT_ALL = "Select all"
    DESELECT_ALL = "Deselect all"

class LogTableHeaders(StrEnum):
    '''Summary table headers in summary selection.'''

    EMPLOYEE = "Employee"
    PREDICTED_HOURS = "Predicted Hours"
    ACCUMULATED_HOURS = "Accumulated Hours"
    DEVIATION = "Deviation"
    PROJECT_ID = "Project ID"
    COORDINATE = "Coordinate"

    @classmethod
    def list_all_values(cls) -> list[int | str]:
        '''Returns a list of all member values.'''
        return [member.value for member in cls]

class YamlEnum(StrEnum):
    '''Enum of top level yaml config entries.'''

    COUNTRIES = "countries"
    DEVIATIONS = "deviations"
    ROW_ANCHORS = "row_anchors"

class CountriesEnum(StrEnum):
    '''Enum of countries.'''

    # AUSTRIA = "Austria"
    # CROATIA = "Croatia"
    # DENMARK = "Denmark"
    ENGLAND = "England"
    # FINLAND = "Finland"
    # FRANCE = "France"
    GERMANY = "Germany"
    # GREECE = "Greece"
    # HUNGARY = "Hungary"
    # ICELAND = "Iceland"
    # ITALY = "Italy"
    # NORWAY = "Norway"
    # POLAND = "Poland"
    # PORTUGAL = "Portugal"
    ROMANIA = "Romania"
    # SLOVENIA = "Slovenia"
    # SPAIN = "Spain"
    # SWEDEN = "Sweden"
    # SWITZERLAND = "Switzerland"

NON_NAMES = [
    'Anställds namn\nDatum',
    'MA Name\nStartdatum',
    'GWR REG',
    'Verhandlung',
    'h/Mnt RUM',
    'h/Mnt REG',
    'h/Mnt\ngesamt',
    'Antal\nanställda',
    'Anzahl\nMA'
]
