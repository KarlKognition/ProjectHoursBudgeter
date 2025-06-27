'''
Module Name
---------
Wizard Logger

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Logs wizard summary results.
'''

## Imports
from os import path
from datetime import datetime
from typing import Optional, Callable
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QTableWidget
# First party library imports
import phb_app.utils.date_utils as du
import phb_app.data.months_dict as md
import phb_app.data.workbook_management as wm
import phb_app.data.log_management as lm
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.employee_management as emp
import phb_app.data.header_management as hm

def get_time_stamp() -> str:
    '''Returns a formatted current time stamp.'''

    return datetime.now().strftime('%Y_%m_%d_-_%H-%M-%S')

def generate_log_file_name(output_dir: str,
                           date: datetime) -> str:
    '''Generates a properly formatted log file name.'''

    timestamp = get_time_stamp()
    month = du.abbr_month(date.month, md.LOCALIZED_MONTHS_SHORT)
    return path.join(output_dir, f"log_output_for_{month}_{date.year}__{timestamp}.txt")

def get_file_data(wb_mng: wm.WorkbookManager) -> lm.FileMetaData:
    '''Gets all required log information and returns it as a LogData object.'''

    out_wb = wb_mng.get_output_workbook_ctx()
    output_file_name = out_wb.mngd_wb.file_name
    output_worksheet_name = out_wb.managed_sheet.selected_sheet.sheet_name
    output_dir = path.dirname(out_wb.mngd_wb.file_path)
    selected_date = out_wb.managed_sheet.selected_date
    log_file_path = generate_log_file_name(output_dir, selected_date)
    input_workbooks = [wb.mngd_wb.file_name for wb in wb_mng.yield_workbook_ctxs_by_role(st.IORole.INPUTS)]
    return lm.FileMetaData(
        log_file_path=log_file_path,
        selected_date=selected_date,
        input_workbooks=input_workbooks,
        output_file_name=output_file_name,
        output_worksheet_name=output_worksheet_name
    )

def get_table_structure(table: QTableWidget,
                        max_proj_id_list_len: int) -> lm.TableStructure:
    '''Gets table headers and column widths.'''

    headers = st.LogTableHeaders.list_all_values()
    col_widths = calculate_column_widths(table, headers, max_proj_id_list_len)
    return lm.TableStructure(
        headers=headers,
        col_widths=col_widths
    )

def get_max_project_id_list_length(employees: list[emp.Employee]) -> int:
    '''Returns the max length of the list of project IDs from all
    selected employees. If the list is empty, a default length of 0 is returned.'''

    return max((len(", ".join(str(proj_id)
                for proj_id
                in emp.found_projects.keys()))
                for emp in employees),
                default=0
              )

def get_max_table_item_length(table: QTableWidget,
                              col: int) -> int:
    '''Returns the max length of the table item from all
    selected employees. If the item is empty, a default length of 0 is returned.'''

    return max((len(table.item(row, col).text())
                for row in range(table.rowCount())
                if table.item(row, col)),
                default=0
              )

def add_spacings(header: str,
                 max_table_item_length: int) -> int:
    '''Adds spacings together.'''

    return max(len(header), max_table_item_length) + hm.DEFAULT_PADDING

def calculate_column_widths(table: QTableWidget,
                            headers: list[str],
                            max_proj_id_list_len: int) -> list[int]:
    '''Calculates column widths based on the longest value in each column.'''

    col_widths = []
    for col, header in enumerate(headers):
        log_headers = st.LogTableHeaders.list_all_values()
        if header not in log_headers:
            continue
        if st.LogTableHeaders.PROJECT_ID in header:
            col_width = add_spacings(header, max_proj_id_list_len)
        else:
            max_table_item_length = get_max_table_item_length(table, col)
            col_width = add_spacings(header, max_table_item_length)
        col_widths.append(col_width)
    return col_widths

def format_hours_wrapper(text: str) -> Callable[[Optional[float], Qt.GlobalColor], str]:
    '''Returns a function that formats hours according to the text given.'''

    def format_hours(hours: Optional[float],
                     colour: Qt.GlobalColor) -> str:
        '''The function to be returned with the correct formatting procedure.'''

        if hours is None:
            return text
        formatted = f"{hours:.2f}"
        return f"*{formatted}*" if colour == Qt.GlobalColor.red else formatted
    return format_hours

def get_employee_data(wb_mng: wm.WorkbookManager) -> list[emp.Employee]:
    '''Gets employee names, project IDs and coordinates from the summary table.'''

    out_wb = wb_mng.get_output_workbook_ctx()
    return list(out_wb.managed_sheet.selected_employees.values())

def format_row(employees: list[emp.Employee],
               table_structure: lm.TableStructure) -> list[str]:
    '''Formats a single row of table data with correct spacing.'''

    # Create two functions which deal with "None" hours and "missing" employees
    format_predicted_hours = format_hours_wrapper(st.SpecialStrings.ZERO_HOURS)
    format_accumulated_hours = format_hours_wrapper(st.SpecialStrings.MISSING)
    formatted_rows = []
    # For each employee, get the name, predicted hours, accumulated hours, project IDs,
    # then put in a row value list and format the spacing of each list element using
    # the table structure
    for employee in employees:
        name = employee.name
        predicted_hours = format_predicted_hours(
            employee.hours.predicted_hours, employee.hours.pre_hours_colour)
        accumulated_hours = format_accumulated_hours(
            employee.hours.accumulated_hours, employee.hours.acc_hours_colour)
        project_info = ", ".join(
            f"{proj_id}" for proj_id in employee.found_projects.keys()) if employee.found_projects else " "
        coord = employee.hours.hours_coord
        deviation = employee.hours.deviation
        row_values = [name, predicted_hours, accumulated_hours, deviation, project_info, coord]
        formatted_row = "".join(
            str(value).rjust(width) for value, width in zip(row_values, table_structure.col_widths))
        formatted_rows.append(formatted_row)
    return formatted_rows

def write_log_file(file_meta: lm.FileMetaData,
                   table_structure: lm.TableStructure,
                   employees: list[emp.Employee]) -> None:
    '''Writes the formatted log to a text file.'''

    with open(file_meta.log_file_path, "w", encoding="utf-8") as log_file:
        datetime_now_str = get_time_stamp()
        month = du.abbr_month(file_meta.selected_date.month, md.LOCALIZED_MONTHS_SHORT)
        log_file.write(f"* Selected date: {month} {file_meta.selected_date.year}; Log created: {datetime_now_str}\n\n")
        log_file.write(f"* Input workbook(s): {'\n'.join(file_meta.input_workbooks)}\n")
        log_file.write(f"* Output workbook: {file_meta.output_file_name}\n")
        log_file.write(f"* Output worksheet: {file_meta.output_worksheet_name}\n\n")
        header_line = "".join(table_structure.headers[col].rjust(table_structure.col_widths[col])
                              for col in range(len(table_structure.headers)))
        log_file.write(header_line + "\n")
        log_file.write("-" * len(header_line) + "\n")
        formatted_rows = format_row(employees, table_structure)
        for row in formatted_rows:
            log_file.write(row + "\n")

def print_log(table: QTableWidget,
              wb_mng: wm.WorkbookManager) -> None:
    '''Coordinates the log file generation and writing.'''

    file_meta = get_file_data(wb_mng)
    employees = get_employee_data(wb_mng)
    max_proj_id_list_len = get_max_project_id_list_length(employees)
    table_structure = get_table_structure(table, max_proj_id_list_len)
    write_log_file(file_meta, table_structure, employees)
