'''
Package
-------
Managers

Module Name
---------
Hours Utilities

Version
-------
Date-based Version: 20250210
Author: Karl Goran Antony Zuvela

Description
-----------
Utility functions the the calculation of hours of each employee
in the project hours budgeting wizard.
'''

from datetime import datetime
from openpyxl.styles import Font
import phb_app.data.phb_dataclasses as dc
import phb_app.utils.general_func_utils as gu

def sum_hours_selected_employee(workbooks: dc.WorkbookManager) -> None:
    '''Sum the hours of each employee by project ID if they are
    found in the given worksheets.'''

    # Get the first (only) managed output workbook
    out_wb = next(workbooks.yield_workbooks_by_type(dc.ManagedOutputWorkbook))
    selected_date = out_wb.managed_sheet_object.selected_date
    # Get the selected employee objects
    sel_emps = out_wb.managed_sheet_object.selected_employees.values()
    for in_wb in workbooks.yield_workbooks_by_type(dc.ManagedInputWorkbook):
        # Get the localised filter headings from the managed input workbook
        employee_name_col = in_wb.managed_sheet_object.indexed_headers.get(
            in_wb.locale_data.filter_headers.name)
        proj_id_col = in_wb.managed_sheet_object.indexed_headers.get(
            in_wb.locale_data.filter_headers.proj_id)
        hours_col = in_wb.managed_sheet_object.indexed_headers.get(
            in_wb.locale_data.filter_headers.hours)
        date_col = in_wb.managed_sheet_object.indexed_headers.get(
            in_wb.locale_data.filter_headers.date)
        # Selected project ID iterator
        proj_id_dict = in_wb.managed_sheet_object.selected_project_ids
        # Go through each row of the selected worksheet, skipping the header row (row 1)
        for row in in_wb.managed_sheet_object.selected_sheet.sheet_object.iter_rows(min_row=2):
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
            emp = next((e for e in sel_emps if e.name == employee_name_val), None)
            if emp:
                # Match found!
                if proj_id_val not in emp.found_projects:
                    emp.found_projects[proj_id_val] = proj_id_dict[proj_id_val]
                if emp.hours.accumulated_hours is None:
                    # Init recorded hours to 0 if the selected employee is found
                    # in the search for the first time
                    emp.hours.accumulated_hours = 0
                # Accumulate found hours
                emp.hours.accumulated_hours += hours_val

def write_hours_to_output_file(output_file: dc.ManagedOutputWorkbook) -> None:
    '''Write recorded hours to output budgeting file.'''

    emp_dict = output_file.managed_sheet_object.selected_employees
    date_row = output_file.managed_sheet_object.selected_date.row
    sheet = output_file.managed_sheet_object.selected_sheet.sheet_object
    for emp_coord in emp_dict.keys():
        acc_hours = emp_dict.get(emp_coord).hours.accumulated_hours
        if acc_hours is not None:
            # If the employee is not missing in the input file
            # get the associated coordinate to write hours in
            # the output file
            hours_coord = next(gu.yield_hours_coord(emp_coord, date_row))
            cell = sheet[hours_coord]
            cell.value = acc_hours
            cell.font = Font(name='Arial', size=12, color='FF000000')
            cell.number_format = '0.00 "h"'
