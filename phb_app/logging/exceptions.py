'''
Module Name
---------
PHB Wizard Exceptions

Author
-------
Karl Goran Antony Zuvela

Description
-----------
PHB Wizard specific exceptions handling.
'''
#           --- Standard libraries ---
from datetime import datetime
# First party libraries
import phb_app.data.io_management as io

##############################
### IOSelection Exceptions ###
##############################

class WorkbookLoadError(Exception):
    '''Custom exception for the error when loading a workbook.'''

class WorkbookAlreadyTracked(Exception):
    '''Custom exception for when a workbook with the same file name is already in the list.'''

    def __init__(self, file_name: str):
        super().__init__(f"Workbook {file_name} is already being tracked.")

class IncorrectWorksheetSelected(Exception):
    '''Custom exception for when a non hours budgeting worksheet is selected.
    This exception should be used for the re-raising of MissingEmployeeRow
    in the IOSelection page validation.'''

    def __init__(self):
        super().__init__("The incorrect worksheet was selected.")

class CountryIdentifiersNotInFilename(Exception):
    '''Custom exception for when the file name does not contain any country identifying
    patterns.'''

    def __init__(self, file_name: str):
        message = (f"Workbook {file_name} does not contain the required patterns "
                    "in its file name to identify its country of origin.")
        super().__init__(message)

class FileAlreadySelected(Exception):
    '''Custom exception for when the user attempts to add the same file twice.'''

    def __init__(self, file_name: str):
        super().__init__(f"Workbook {file_name} has already been selected.")

class TooManyOutputFilesSelected(Exception):
    '''Custom exception for when the output table has more than one output files selected.'''

    def __init__(self):
        super().__init__("Only one output file may be selected!")

class BudgetingDatesNotFound(Exception):
    '''Custom exception for when the chosen month and year are not in the output file.'''

    def __init__(self, dropdown_handler: io.SelectedText, file: str):
        super().__init__(f"{dropdown_handler.month} or {dropdown_handler.year} not found in sheet {dropdown_handler.worksheet} of file {file}.")

######################################
### Employee Management Exceptions ###
######################################

class MissingEmployeeRow(Exception):
    '''Custom exception for when the row containing the employees in the output
    file is not found. Strings are represented with preserved escape characters.'''

    def __init__(self, start_anchor: str = "", end_anchor: str = ""):
        self.start_anchor = start_anchor
        self.end_anchor = end_anchor
        if self.start_anchor and self.end_anchor:
            message = ("The selected worksheet contains neither "
                       f"{repr(self.start_anchor)} nor {repr(self.end_anchor)}.")
        elif self.start_anchor:
            message = f"The selected worksheet does not contain {repr(self.start_anchor)}."
        elif self.end_anchor:
            message = f"The selected worksheet does not contain {repr(self.end_anchor)}."
        else:
            message = "An unexpected error occured."
        super().__init__(message)

class EmployeeRowAnchorsMisalignment(Exception):
    '''Custom exception for when both employee row anchor cells are found
    but are not in the same row. Strings are represented with preserved escape characters.'''

    def __init__(self, start_anchor: str, end_anchor: str):
        self.start_anchor = start_anchor
        self.end_anchor = end_anchor
        message = (f"Both {repr(self.start_anchor)} and {repr(self.end_anchor)} are present, "
                       "but not in the same row.")
        super().__init__(message)

###############################
### Input Search Exceptions ###
###############################

class ItemNotFound(Exception):
    '''Custom base exception for when searching input worksheets and the item is not found.'''

    def __init__(self, item: str, item_val: int|str|datetime):
        self.item = item
        self.item_val = item_val
        message = f"{item} {self.item_val} not found in worksheet."
        super().__init__(message)

class EmployeeNotFound(ItemNotFound):
    '''Custom exception for when the selected employee is not found in the input file.'''

    def __init__(self, emp_name: str):
        super().__init__("Employee", emp_name)

class DateNotFound(Exception):
    '''Custom exception for when the given date is not found in the input worksheet.'''

    def __init__(self, date: datetime):
        super().__init__("Date", date)

class ProjectNotFound(Exception):
    '''Custom exception for when the project number is not found in the input worksheet.'''

    def __init__(self, proj_id: int|str):
        super().__init__("Project ID", proj_id)
