'''
String constants. English.
'''

from pathlib import Path
import git

###########################
### General GUI Strings ###
###########################

GUI_TITLE = "Project Hours Budgeting Wizard"

INTRO_TITLE = "Introduction"

INTRO_SUBTITLE = "Welcome to the Project Hours Budgeting Wizard!"

INTRO_MESSAGE = """
This wizard allows the user to budget the hours of many employees of a project at once.

There are three steps:

1: Select the input files (SAP extracts or external timesheets) and the output budgeting file.

2: Select the employees of the project.

3: Review the process, checking/deleting the log file or restoring the budgeting file to its previous state before the writing of data.

NOTE: Close all Excel files which will be used in this process!
"""

IMAGE_LOAD_FAIL = "Failed to load image."

def find_git_root():
    '''Find the root of the project.'''
    repo = git.Repo(Path(__file__).resolve(), search_parent_directories=True)
    return Path(repo.working_tree_dir)

APP_ROOT = find_git_root() / "phb_app"
IMAGES_DIR = APP_ROOT / "images"

########################
### IO File  Strings ###
########################

IO_FILE_TITLE = "Input/Output Files"

IO_FILE_SUBTITLE = """
Manage the input timesheets and output project hours budgeting file.
NOTE: The selected date of the output file is the row in which the hours will be written but also will affect whether hours are accumulated at all from the input files.
"""

I_FILE_INSTRUCTION_TEXT = "Select the timesheets as the input files."

O_FILE_INSTRUCTION_TEXT = "Select the project hours budgeting file as the target output file."

ADD_FILE = "Add File"
EXCEL_FILE = "Excel files (*.xlsx)"

#################################
### Project Selection Strings ###
#################################

###################################
### Employee Selection  Strings ###
###################################

########################
### Summary  Strings ###
########################
