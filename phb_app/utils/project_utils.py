'''
Package
-------
General Function Utilities

Module Name
---------
Employee Management Utilities

Version
-------
Date-based Version: 20250304
Author: Karl Goran Antony Zuvela

Description
-----------
Provides utility functions for the management
of employees in the project hours budgeting wizard.
'''

from typing import TYPE_CHECKING
from PyQt6.QtCore import QModelIndex
from PyQt6.QtWidgets import QTableWidget
import phb_app.wizard.constants.integer_enums as ie

if TYPE_CHECKING:
    import phb_app.data.workbook_management as wm

def set_selected_project_ids(mngd_wb: "wm.ManagedInputWorkbook", table: QTableWidget, rows: list[QModelIndex], headers: ie.ProjectIDTableHeaders) -> None:
    '''
    Sets the selected projects IDs as references from the selectable IDs.
    '''
    for row in rows:
        # Get the project ID and file name from each row
        proj_id = table.item(row.row(), headers.PROJECT_ID).text()
        file_name = table.item(row.row(), headers.FILENAME).text()
        if file_name == mngd_wb.file_name:
            # If the file name of the managed input workbook is correct,
            # make a reference by key from selected to selectable project IDs
            mngd_wb.managed_sheet_object.selected_project_ids[proj_id] = mngd_wb.managed_sheet_object.selectable_project_ids[proj_id]
