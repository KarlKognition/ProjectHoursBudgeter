'''
Package
-------
File Handling Utilities

Module Name
---------
Employee Management Utilities

Version
-------
Date-based Version: 20250425
Author: Karl Goran Antony Zuvela

Description
-----------
Provides file handling utility functions in the project hours budgeting wizard.
'''

import zipfile
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.utils.exceptions import (
    ReadOnlyWorkbookException,
    InvalidFileException
)
from PyQt6.QtWidgets import QTableWidget
import phb_app.data.location_management as loc
import phb_app.logging.exceptions as ex
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.workbook_management as wm

def get_origin_from_file_name(file_name: str, country_data: loc.CountryData, countries_enum: st.CountriesEnum) -> str:
    '''
    Checks the workbook's origin. Returns the country of origin.
    '''
    # Pattern in all German timesheets
    german_patterns = country_data.get_locale_by_country(countries_enum.GERMANY.value).file_patterns
    # Pattern in all Romanian timesheets
    romanian_patterns = country_data.get_locale_by_country(countries_enum.ROMANIA.value).file_patterns
    if all(pattern in file_name.lower() for pattern in german_patterns):
        return countries_enum.GERMANY.value
    if all(pattern in file_name.lower() for pattern in romanian_patterns):
        return countries_enum.ROMANIA.value
    raise ex.CountryIdentifiersNotInFilename(file_name)

def is_workbook_open(file_path:str) -> bool:
    '''Check if the file is already in use.'''
    try:
        with open(file_path, 'r+', encoding='utf-8'):
            pass
    except (OSError, IOError):
        return True
    return False

def try_load_workbook(managed_workbook: wm.ManagedWorkbook, wb_type: wm.ManagedOutputWorkbook) -> Workbook:
    '''
    Template for attempting to load the workbook.
    '''
    file_name = managed_workbook.file_name
    file_path = managed_workbook.file_path
    try:
        if isinstance(managed_workbook, wb_type):
            # We only care about the output workbook being open
            # as we will not write to the input workbook
            with open(file_path, 'r+', encoding='utf-8'):
                pass
        return load_workbook(file_path)
    except ReadOnlyWorkbookException as e:
        raise ex.WorkbookLoadError(f"Workbook '{file_name}' is read-only: {str(e)}.") from e
    except InvalidFileException as e:
        raise ex.WorkbookLoadError(f"Invalid Excel file '{file_name}': {str(e)}.") from e
    except FileNotFoundError as e:
        raise ex.WorkbookLoadError(f"File '{file_name}' not found.") from e
    except PermissionError as e:
        raise ex.WorkbookLoadError(f"Permission denied: Close '{file_name}' if open.") from e
    except zipfile.BadZipFile as e:
        raise ex.WorkbookLoadError(f"Corrupted Excel file '{file_name}': {str(e)}") from e

def is_file_already_in_table(file_path: str, col: int, table: QTableWidget) -> bool:
    '''
    Check if the file already exists in the given table
    and return a respective boolean.
    '''
    return any(
        table.item(row, col)
        and table.item(row, col).text() == file_path
        for row in range(table.rowCount())
    )

