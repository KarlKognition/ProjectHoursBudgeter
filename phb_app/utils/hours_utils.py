'''
Package
-------
Managers

Module Name
---------
Hours Utilities

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Utility functions the the calculation of hours of each employee
in the project hours budgeting wizard.
'''
from typing import Optional, TYPE_CHECKING
from datetime import datetime
from functools import lru_cache
from openpyxl.styles import Font
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
#           --- First party libraries ---
import phb_app.data.employee_management as em
import phb_app.utils.employee_utils as eu
import phb_app.wizard.constants.ui_strings as st

if TYPE_CHECKING:
    import phb_app.data.workbook_management as wm

def compute_predicted_hours(out_wb_ctx: "wm.OutputWorkbookContext") -> None:
    '''Compute the predicted hours for each employee in the output workbook. Selected rows are purely for cacheing purposes.'''
    coords = tuple(out_wb_ctx.managed_sheet.selected_employees.keys())
    pre_hours = eu.find_predicted_hours(
        coords,
        out_wb_ctx.managed_sheet.selected_date.row,
        out_wb_ctx.mngd_wb.file_path,
        out_wb_ctx.managed_sheet.selected_sheet.sheet_name
    )
    eu.set_employee_hours(
        out_wb_ctx.managed_sheet.selected_employees,
        out_wb_ctx.managed_sheet.selected_date.row
    )
    out_wb_ctx.worksheet_service.set_predicted_hours(pre_hours)
    out_wb_ctx.worksheet_service.set_predicted_hours_colour()

def compute_accumulated_hours_for_selected_employees(wbs: "wm.WorkbookManager", out_wb_ctx: "wm.OutputWorkbookContext") -> None:
    """Compute the hours for each selected employee in the output workbook. Selected rows are purely for cacheing purposes."""
    for emp in out_wb_ctx.worksheet_service.yield_from_selected_employees():
        _sum_hours_selected_employee(wbs, out_wb_ctx, emp)
        _format_accumulated_hours(emp)
        emp.hours.set_deviation()

def _sum_hours_selected_employee(wbs: "wm.WorkbookManager", out_wb_ctx: "wm.OutputWorkbookContext", sel_emp: em.Employee) -> None:
    '''Sum the hours of each employee by project ID if they are
    found in the given worksheets.'''
    selected_date = out_wb_ctx.managed_sheet.selected_date
    # Get the selected employee objects
    for in_wb in wbs.yield_workbook_ctxs_by_role(st.IORole.INPUTS):
        # Get the localised filter headings from the managed input workbook
        employee_name_col = in_wb.managed_sheet.indexed_headers.get(in_wb.locale_data.filter_headers.name)
        proj_id_col = in_wb.managed_sheet.indexed_headers.get(in_wb.locale_data.filter_headers.proj_id)
        hours_col = in_wb.managed_sheet.indexed_headers.get(in_wb.locale_data.filter_headers.hours)
        date_col = in_wb.managed_sheet.indexed_headers.get(in_wb.locale_data.filter_headers.date)
        # Selected project ID iterator
        proj_id_dict = in_wb.managed_sheet.selected_project_ids
        # Go through each row of the selected worksheet, skipping the header row (row 1)
        for row in in_wb.managed_sheet.selected_sheet.sheet_object.iter_rows(min_row=2):
            # Skip rows with missing data
            if (not row[employee_name_col].value or
                not row[proj_id_col].value or
                not row[hours_col].value or
                not row[date_col].value
            ):
                continue
            # Get the value in the row with the given column header
            employee_name_val: str = row[employee_name_col].value
            proj_id_val: str|int = row[proj_id_col].value
            hours_val: float = row[hours_col].value
            date_val: datetime = row[date_col].value
            # Skip rows with non-matching date and project ID
            if (date_val.month != selected_date.month or
                date_val.year != selected_date.year or
                proj_id_val not in proj_id_dict.keys()):
                continue
            # Get the first employee with the matching name
            if employee_name_val == sel_emp.name:
                # Match found!
                if proj_id_val not in sel_emp.found_projects:
                    sel_emp.found_projects[proj_id_val] = proj_id_dict[proj_id_val]
                if sel_emp.hours.accumulated_hours is None:
                    # Init recorded hours to 0 if the selected employee is found
                    # in the search for the first time
                    sel_emp.hours.accumulated_hours = 0
                # Accumulate found hours
                sel_emp.hours.accumulated_hours += hours_val

def _format_accumulated_hours(emp: em.Employee) -> None:
    '''Format the accumulated hours.'''
    if emp.hours.accumulated_hours is None or emp.hours.accumulated_hours == 0.0:
        emp.hours.acc_hours_colour = QColor(Qt.GlobalColor.red)

def format_summary_data_row_hours(hours: Optional[float], text: str) -> str:
    '''Formats hours for the summary data table.'''
    if not hours:
        return text
    return f"{hours:.2f}"

def format_log_row_hours(hours: Optional[float], text: str, colour: QColor) -> str:
    '''Formats hours for the summary data table.'''
    if hours is None:
        return text
    formatted = f"{hours:.2f}"
    return f"*{formatted}*" if colour == QColor(Qt.GlobalColor.red) else formatted

def write_hours_to_output_file(output_file: "wm.OutputWorkbookContext") -> None:
    '''Write recorded hours to output budgeting file.'''
    emp_dict = output_file.managed_sheet.selected_employees
    date_row = output_file.managed_sheet.selected_date.row
    sheet = output_file.managed_sheet.selected_sheet.sheet_object
    for emp_coord in emp_dict.keys():
        acc_hours = emp_dict.get(emp_coord).hours.accumulated_hours
        if acc_hours is not None:
            # If the employee is not missing in the input file
            # get the associated coordinate to write hours in
            # the output file
            hours_coord = next(eu.yield_hours_coord(emp_coord, date_row))
            cell = sheet[hours_coord]
            cell.value = acc_hours
            cell.font = Font(name='Arial', size=12, color='FF000000')
            cell.number_format = '0.00 "h"'
