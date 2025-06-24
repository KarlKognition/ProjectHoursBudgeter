'''
Package
-------
PHB Wizard

Module Name
---------
Employee Selection Page

Version
-------
Date-based Version: 20250304
Author: Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the employee selection page.
'''

from PyQt6.QtWidgets import QWizardPage, QLabel, QTableWidget, QPushButton, QHBoxLayout
import phb_app.data.header_management as hm
import phb_app.data.io_management as io
import phb_app.data.workbook_management as wm
import phb_app.utils.employee_utils as eu
import phb_app.utils.hours_utils as hu
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

class EmployeeSelectionPage(QWizardPage):
    '''Page for selecting the employees whose hours will be budgeted.'''
    def __init__(self, managed_workbooks: wm.WorkbookManager):
        super().__init__()

        self.wb_mgmt = managed_workbooks
        pu.set_titles(self, st.PROJECT_SELECTION_TITLE, st.PROJECT_SELECTION_SUBTITLE)
        self.employee_panel = io.IOControls(
            page=self,
            role=st.IORole.EMPLOYEE_TABLE,
            label=QLabel(st.EMPLOYEE_SELECTION_INSTRUCTIONS),
            table=pu.create_table(self, ie.EmployeeTableHeaders, QTableWidget.SelectionMode.MultiSelection, hm.EMPLOYEE_COLUMN_WIDTHS),
            buttons=[QPushButton(st.ButtonNames.SELECT_ALL, self), QPushButton(st.ButtonNames.DESELECT_ALL, self)]
        )
        self.emp_ctx = io.EntryContext(self.employee_panel, io.EmployeeTableContext())
        pu.setup_page(self, [pu.create_interaction_panel(self.employee_panel)], QHBoxLayout())

#           --- QWizard function overrides ---

    def initializePage(self) -> None:
        '''Override page initialisation. Setup page on each visit.'''
        io.EntryHandler(self.emp_ctx)
        pu.connect_buttons(self, self.emp_ctx)
        pu.populate_employee_table(self, self.emp_ctx, self.wb_mgmt)

    def cleanupPage(self):
        '''Clean up if the back button is pressed.'''

        # Clear the table
        self.employees_table.clear()
        self.employees_table.setRowCount(0)
        # Clear the selected employees of the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(wm.ManagedOutputWorkbook))
        out_wb.managed_sheet_object.clear_predicted_hours()
        for in_wb in self.managed_workbooks.yield_workbooks_by_type(wm.ManagedInputWorkbook):
            in_wb.managed_sheet_object.selected_project_ids.clear()

    def isComplete(self) -> bool:
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''

        return bool(self.employee_panel.table.selectionModel().selectedRows())

    def validatePage(self) -> bool:
        '''Override the page validation.
        Set the property.'''

        ## Selected employees
        # Set the selected employees in the managed output workbook
        selected_rows = self.employees_table.selectionModel().selectedRows()
        selected_employees = [
            (self.employees_table
            .item(row.row(), hm.EmployeeTableHeaders.COORDINATE)
            .text(),
            self.employees_table
            .item(row.row(), hm.EmployeeTableHeaders.EMPLOYEE)
            .text()
            )
            for row in selected_rows
        ]
        # This would be changed to `get_worksheet_by_name` if in the future, more than
        # one worksheet would be chosen. If across several workbooks too, make and use
        # get workbook by type and name, then get worksheet by name.
        # Here: Get the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(wm.ManagedOutputWorkbook))
        out_wb.managed_sheet_object.set_selected_employees(selected_employees)
        ## Predicted hours per employee
        # Find hours
        pre_hours = eu.find_predicted_hours(
            out_wb.managed_sheet_object.selected_employees,
            out_wb.managed_sheet_object.selected_date.row,
            out_wb.file_path,
            out_wb.managed_sheet_object.selected_sheet.sheet_name
        )
        # Set hours
        out_wb.managed_sheet_object.set_predicted_hours(pre_hours)
        # Set hours colouring
        out_wb.managed_sheet_object.set_predicted_hours_colour()
        ## Recorded hours
        # Accumulate and set
        hu.sum_hours_selected_employee(self.managed_workbooks)
        for emp in out_wb.managed_sheet_object.selected_employees.values():
            emp.hours.set_deviation()
        # Validation complete
        return True
