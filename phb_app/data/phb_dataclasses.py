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

from datetime import datetime
from enum import auto
from typing import Iterator
from dataclasses import dataclass, field
from PyQt6.QtWidgets import QTableWidget
# First party library imports
import phb_app.data.common as common
import phb_app.logging.exceptions as ex

type ColWidths = dict[common.BaseTableHeaders, int]

class WizardPageIDs(common.BaseTableHeaders):
    '''Enumerated pages in order of appearance.'''

    EXPLANATION_PAGE = 0
    I_O_SELECTION_PAGE = auto()
    PROJECT_SELECTION_PAGE = auto()
    EMPLOYEE_SELECTION_PAGE = auto()
    SUMMARY_PAGE = auto()

###########################
#### IO Selection Enum ####
###########################



#################################
#### Employee Selection Enum ####
#################################

class EmployeeTableHeaders(common.BaseTableHeaders):
    '''Employee table headers in employee selection.'''

    EMPLOYEE = 0
    WORKSHEET = auto()
    COORDINATE = auto()

######################
#### Summary Enum ####
######################

class SummaryTableHeaders(common.BaseTableHeaders):
    '''Summary table headers in summary selection.'''

    EMPLOYEE = 0
    PREDICTED_HOURS = auto()
    ACCUMULATED_HOURS = auto()
    DEVIATION = auto()
    PROJECT_ID = auto()
    OUTPUT_WORKSHEET = auto()
    COORDINATE = auto()

##################
#### Log Enum ####
##################



#######################
#### Column widths ####
#######################

INPUT_COLUMN_WIDTHS = {
    common.InputTableHeaders.FILENAME: 250,
    common.InputTableHeaders.COUNTRY: 150,
    common.InputTableHeaders.WORKSHEET: 100

}
OUTPUT_COLUMN_WIDTHS = {
    common.OutputTableHeaders.FILENAME: 250,
    common.OutputTableHeaders.WORKSHEET: 150,
    common.OutputTableHeaders.MONTH: 60,
    common.OutputTableHeaders.YEAR: 60
}

DEFAULT_PADDING = 5

############################################
#### To be replaced with a yaml handler ####
############################################

NON_NAMES = [
    'MA Name\nStartdatum',
    'GWR REG',
    'Verhandlung',
    'h/Mnt RUM',
    'h/Mnt REG',
    'h/Mnt\ngesamt',
    'Anzahl\nMA'
]

#############################
#### Workbook management ####
#############################



#######################
#### IO management ####
#######################

@dataclass
class RowHandler:
    '''Data class for managing the row in the table.'''
    table: QTableWidget
    row_position: int
    file_name: str
    file_path: str

########################
#### Log management ####
########################

@dataclass
class FileMetaData:
    '''Encapsulates all necessary log data.'''
    log_file_path: str
    selected_date: datetime
    input_workbooks: list[str]
    output_file_name: str
    output_worksheet_name: str

@dataclass
class TableStructure:
    '''Encapsulates table-related metadata.'''
    headers: list[str]
    col_widths: list[int]
