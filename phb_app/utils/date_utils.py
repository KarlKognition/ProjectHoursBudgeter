'''
Module Name
---------
PHB Wizard Date Utilities

Author
-------
Karl Goran Antony Zuvela

Description
-----------
PHB Wizard date utility functions.
'''
#           --- Standard libraries ---
from datetime import datetime
from typing import TYPE_CHECKING
from functools import lru_cache
#           --- Third party libraries ---
import xlwings as xw
#           --- First party libraries ---
import phb_app.data.months_dict as md
import phb_app.logging.exceptions as ex

if TYPE_CHECKING:
    import phb_app.data.io_management as io
    import phb_app.data.workbook_management as wm

def abbr_month(month_number: int, months: dict) -> str:
    '''
    Change a month number to its short name.
    '''
    for month_key, month_value in months.items():
        if int(month_value) == int(month_number):
            return month_key
    return ""

# Cache the function results
# Only cache one result to minimise memory use
@lru_cache(maxsize=1)
def get_budgeting_dates(file_path: str, sheet_name: str) -> list[tuple[int, int, int]]:
    '''
    Returns a list of tuples, each tuple containing the month, year and coordinate
    retreived from the possible budgeting dates in the budgeting file.
    '''
    # Do not diplay Excel while computing
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets[sheet_name]
    # 52 was chosen because the dates begin there in every file.
    # This could be automated to find yet another dubious entry
    # and the key for that entry could be kept in a yaml config
    # for the user to define. Note: expand stops at the first
    # empty cell.
    col1 = sheet.range('A8').expand('down')
    months_years_rows = [(cell.value.month, cell.value.year, cell.row)
                         for cell in col1
                         if isinstance(cell.value, datetime)]
    wb.close()
    # Excel will quit if no other workbooks are open
    return months_years_rows

def set_budgeting_date(wb_ctx: "wm.OutputWorkbookContext", dropdown_text: "io.SelectedText") -> None:
    '''Sets the budgeting date with the row it is located in the worksheet.'''
    # Convert dates to integers and put in a tuple
    month_year = (md.LOCALIZED_MONTHS_SHORT.get(dropdown_text.month), int(dropdown_text.year))
    budgeting_dates = get_budgeting_dates(wb_ctx.mngd_wb.file_path, dropdown_text.worksheet)
    for tup in budgeting_dates:
        if tup[:2] == month_year:
            month, year, row = tup
            break
    else:
        raise ex.BudgetingDatesNotFound(dropdown_text, wb_ctx.mngd_wb.file_name)
    selected_date = wb_ctx.managed_sheet.selected_date
    selected_date.month, selected_date.year, selected_date.row = month, year, row
