'''
Package
-------
Utilities

Module Name
---------
Worksheet Management Utilities

Version
-------
Date-based Version: 20250219
Author: Karl Goran Antony Zuvela

Description
-----------
Provides utility functions for the management
of worksheets in the project hours budgeting wizard.
'''

from openpyxl.worksheet.worksheet import Worksheet


def locate_target_row_employee_names(self, sheet_obj: Worksheet) -> bool:
    '''Finds the initial range where the employee names can be located.'''

    sheet = self.selected_sheet['sheet_object']
    both_coords_found = False
    # Locate 'MA Name\nStartdatum' in the first column
    for row in sheet.iter_rows(max_col=1):
        if row[0].value == "MA Name\nStartdatum":
            self.employee_range['start_cell'] = row[0].coordinate
            target_row = sheet[self.employee_range['start_cell']].row
            for cell in sheet[target_row]:
                if cell.value == "Anzahl\nMA":
                    self.employee_range['end_cell'] = cell.coordinate
                    both_coords_found = True
                    break
        if both_coords_found:
            break
    else:
        print("Error: Neither 'MA Name\\nStartdatum' nor 'Anzahl\\nMA' found.")
    return both_coords_found
