'''
Package
-------
PHB Wizard

Module Name
---------
Employee Utilities

Author
-------
Karl Goran Antony Zuvela

Description
-----------
All functions necessary for managing employee data.
'''

from typing import Iterator
from functools import lru_cache
from PyQt6.QtWidgets import QTableWidget
import openpyxl.utils as xlutils
import xlwings as xw
from openpyxl.worksheet.worksheet import Worksheet
#           --- First party libraries ---
import phb_app.data.header_management as hm
import phb_app.data.workbook_management as wm
import phb_app.wizard.constants.integer_enums as ie
import phb_app.logging.exceptions as ex
import phb_app.data.employee_management as emp

# Cache the function results
# Only cache one result to minimise memory use
@lru_cache(maxsize=1)
def set_employee_range(sheet_obj: Worksheet, emp_range: emp.EmployeeRange, anchors: emp.EmployeeRowAnchors) -> None:
    '''
    Finds the row range where the employee names should be located.
    '''
    # Temp variables to hold the cell data
    start_anchor_temp = None
    end_anchor_temp = None
    # Loop through every row
    for row in sheet_obj.iter_rows():
        # Loop through every cell of that row
        for cell in row:
            if cell.value == anchors.start_anchor:
                start_anchor_temp = cell
            elif cell.value == anchors.end_anchor:
                end_anchor_temp = cell
    if start_anchor_temp and end_anchor_temp:
        if start_anchor_temp.row == end_anchor_temp.row:
            emp_range.start_cell = start_anchor_temp.coordinate
            emp_range.end_cell = end_anchor_temp.coordinate
        else:
            raise ex.EmployeeRowAnchorsMisalignment(anchors.start_anchor, anchors.end_anchor)
    else:
        if not start_anchor_temp and not end_anchor_temp:
            raise ex.MissingEmployeeRow(anchors.start_anchor, anchors.end_anchor)
        if not start_anchor_temp and end_anchor_temp:
            raise ex.MissingEmployeeRow(anchors.start_anchor)
        if start_anchor_temp and not end_anchor_temp:
            raise ex.MissingEmployeeRow(anchors.end_anchor)

def yield_hours_coord(coord: str, row: int) -> Iterator[str]:
    '''
    Yields the coordinate for where the hours are located per employee
    for the given date (row).
    '''
    # The first item ([0] -> col) in the tuple from `coordinate_from_string` is used
    yield f"{str(xlutils.cell.coordinate_from_string(coord)[0])}{str(row)}"

def find_predicted_hours(emp_dict: dict[str, emp.Employee], row: int, file_path: str, sheet_name: str) -> dict[str, int]:
    '''
    Goes through all given coordinates of a worksheet, computes
    any formulae and returns the hours by employee name coordinate.
    '''
    # Do not diplay Excel while computing
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets[sheet_name]
    # Prepare a dictionary of coord:hours
    pre_hours = {}
    for emp_coord, employee in emp_dict.items():
        # Create a coordinate from the date's row and employee's column
        hours_coord = next(yield_hours_coord(emp_coord, row))
        # Save the computed value
        pre_hours[emp_coord] = sheet.range(hours_coord).value
        # Save the hours coordinate
        employee.hours.hours_coord = hours_coord
    wb.close()
    return pre_hours

def compute_selected_employees(table: QTableWidget, wb_ctx: "wm.OutputWorkbookContext") -> None:
    """
    Find selected employees in the table and set them as selected in the managed output workbook.
    """
    selected_rows = table.selectionModel().selectedRows()
    selected_employees = [
        (table.item(row.row(), ie.EmployeeTableHeaders.COORDINATE).text(),
        table.item(row.row(), ie.EmployeeTableHeaders.EMPLOYEE).text())
        for row in selected_rows
    ]
    wb_ctx.worksheet_service.set_selected_employees(selected_employees)

def pop_unselected_employees(table: QTableWidget, out_wb_ctx: "wm.OutputWorkbookContext") -> None:
    """Pop unselected employees from the managed output workbook."""
    unselected_coords = []
    for row in range(table.rowCount()):
        if not table.selectionModel().isRowSelected(row):
            coord = table.item(row, ie.SummaryDataTableHeaders.COORDINATE).text()
            unselected_coords.append(coord)
    # Pop the unselected employees from the selected employees dictionary
    for coord in unselected_coords:
        out_wb_ctx.managed_sheet.selected_employees.pop(coord, None)
