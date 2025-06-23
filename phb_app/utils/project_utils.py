'''
Package
-------
General Function Utilities

Module Name
---------
Project ID Utilities

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
from PyQt6.QtWidgets import QTableWidget
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

if TYPE_CHECKING:
    import phb_app.data.workbook_management as wm

def set_project_ids_each_input_wb(wb_mngr: "wm.WorkbookManager") -> None:
    '''Set project IDs for each input workbook.'''

    for wb_ctx in wb_mngr.yield_workbook_ctxs_by_role(st.IORole.INPUTS):
        wb_ctx.worksheet_service.set_selectable_project_ids(
            wb_ctx.locale_data.filter_headers.proj_id,
            wb_ctx.locale_data.filter_headers.description,
            wb_ctx.locale_data.filter_headers.name)

def set_selected_project_ids(wb_ctx: "wm.InputWorkbookContext", table: QTableWidget, row: int, headers: ie.ProjectIDTableHeaders) -> None:
    '''Sets the selected projects IDs as references from the selectable IDs.'''
    # Get the project ID and file name from each row
    proj_id = table.item(row, headers.PROJECT_ID).text()
    if proj_id not in wb_ctx.managed_sheet.selected_project_ids:
        wb_ctx.managed_sheet.selected_project_ids[proj_id] = wb_ctx.managed_sheet.selectable_project_ids[proj_id]
