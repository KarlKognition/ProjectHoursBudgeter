'''
Module Name
---------
PHB Wizard Constant - Integer Enums

Author
-------
Karl Goran Antony Zuvela

Description
-----------
PHB Wizard integer enums.
'''
from enum import IntEnum, auto

class BaseTableHeaders(IntEnum):
    '''Base enum class with overridden string dunder method
    and general class method.'''
    @classmethod
    def cap_members_list(cls) -> list[str]:
        '''Return a list of capitalised enum members.'''
        return [name.replace('_', ' ').title() for name in cls.__members__]

    @classmethod
    def list_all_values(cls) -> list[int | str]:
        '''Returns a list of all member values.'''
        return [member.value for member in cls]

class WizardPageIDs(BaseTableHeaders):
    '''Enumerated pages in order of appearance.'''

    EXPLANATION_PAGE = 0
    I_O_SELECTION_PAGE = auto()
    PROJECT_SELECTION_PAGE = auto()
    EMPLOYEE_SELECTION_PAGE = auto()
    SUMMARY_PAGE = auto()

class InputTableHeaders(BaseTableHeaders):
    '''Input table headers in IOSelection.'''

    FILENAME = 0
    COUNTRY = auto()
    WORKSHEET = auto()

class OutputTableHeaders(BaseTableHeaders):
    '''Output table headers in IOSelection.'''

    FILENAME = 0
    WORKSHEET = auto()
    MONTH = auto()
    YEAR = auto()

class OutputFile(BaseTableHeaders):
    '''Enum for table rows or columns.'''

    FIRST_ENTRY = 0
    SECOND_ENTRY = auto()

class ProjectIDTableHeaders(BaseTableHeaders):
    '''Project ID headers in project selection.'''

    PROJECT_ID = 0
    DESCRIPTION = auto()
    FILENAME = auto()

class EmployeeTableHeaders(BaseTableHeaders):
    '''Employee table headers in employee selection.'''

    EMPLOYEE = 0
    WORKSHEET = auto()
    COORDINATE = auto()

class SummaryIOTableHeaders(BaseTableHeaders):
    '''Summary IO table headers in summary selection.'''
    INPUT_WORKBOOKS = 0
    OUTPUT_WORKBOOK = auto()
    SELECTED_DATE = auto()

class SummaryDataTableHeaders(BaseTableHeaders):
    '''Summary data table headers in summary selection.'''

    EMPLOYEE = 0
    PREDICTED_HOURS = auto()
    ACCUMULATED_HOURS = auto()
    DEVIATION = auto()
    PROJECT_ID = auto()
    OUTPUT_WORKSHEET = auto()
    COORDINATE = auto()

#           --- CONSTANTS ---

CONST_0 = 0
CONST_1 = 1
IO_SUMMARY_ROW_COUNT = 3
