'''
Module Name
---------
Project Hours Budgeting Data Classes (phb_dataclasses)

Version
-------
Date-based Version: 20250304
Author: Karl Goran Antony Zuvela

Description
-----------
Provides several data classes for the Project Hours Budgeting Wizard.
'''

from enum import auto
from dataclasses import dataclass
from PyQt6.QtWidgets import QTableWidget
# First party library imports
import phb_app.wizard.constants.integer_enums as ie

type ColWidths = dict[ie.BaseTableHeaders, int]

class WizardPageIDs(ie.BaseTableHeaders):
    '''Enumerated pages in order of appearance.'''

    EXPLANATION_PAGE = 0
    I_O_SELECTION_PAGE = auto()
    PROJECT_SELECTION_PAGE = auto()
    EMPLOYEE_SELECTION_PAGE = auto()
    SUMMARY_PAGE = auto()

class EmployeeTableHeaders(ie.BaseTableHeaders):
    '''Employee table headers in employee selection.'''

    EMPLOYEE = 0
    WORKSHEET = auto()
    COORDINATE = auto()

class SummaryTableHeaders(ie.BaseTableHeaders):
    '''Summary table headers in summary selection.'''

    EMPLOYEE = 0
    PREDICTED_HOURS = auto()
    ACCUMULATED_HOURS = auto()
    DEVIATION = auto()
    PROJECT_ID = auto()
    OUTPUT_WORKSHEET = auto()
    COORDINATE = auto()

INPUT_COLUMN_WIDTHS = {
    ie.InputTableHeaders.FILENAME: 250,
    ie.InputTableHeaders.COUNTRY: 150,
    ie.InputTableHeaders.WORKSHEET: 100

}
OUTPUT_COLUMN_WIDTHS = {
    ie.OutputTableHeaders.FILENAME: 250,
    ie.OutputTableHeaders.WORKSHEET: 150,
    ie.OutputTableHeaders.MONTH: 60,
    ie.OutputTableHeaders.YEAR: 60
}

DEFAULT_PADDING = 5

@dataclass
class RowHandler:
    '''Data class for managing the row in the table.'''
    table: QTableWidget
    row_position: int
    file_name: str
    file_path: str
