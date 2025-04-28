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

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWizardPage,
    QLabel,
    QTableWidget,
    QPushButton,
    QVBoxLayout,
    QTableWidgetItem
)
import phb_app.data.common as common
import phb_app.data.phb_dataclasses as dc (
    EmployeeTableHeaders,
    WorkbookManager,
    QPropName,
    EmployeeTableHeaders,
    ButtonNames,
    ManagedOutputWorkbook,
    ManagedInputWorkbook,
    NON_NAMES
)
import phb_app.utils.general_func_utils as gu
import phb_app.utils.hours_utils as hutils

class EmployeeSelectionPage(QWizardPage):
    '''Page for selecting the employees whose hours will be budgeted.'''
    def __init__(self):
        super().__init__()

        self.setTitle("Employee Selection")
        self.setSubTitle("Select the employees whose hours will be budgeted.")
        # Headers
        self.headers = dc.EmployeeTableHeaders.cap_members_list()

        ## Init property
        self.managed_workbooks: dc.WorkbookManager

        self.setup_widgets()
        self.setup_layout()
        self.setup_connections()

    #######################################
    ### QWizard setup function override ###
    #######################################

    def initializePage(self) -> None:
        '''Retrieve fields from other pages.'''

        # Get workbooks from IOSelection
        self.managed_workbooks = self.wizard().property(
            dc.QPropName.MANAGED_WORKBOOKS.value)
        ## Table data
        # Table headers
        self.employees_table.setHorizontalHeaderLabels(self.headers)
        # Add all data to the table
        self.populate_table(self.employees_table)

    #######################
    ### Setup functions ###
    #######################

    def setup_widgets(self) -> None:
        '''Init all widgets.'''

        ## Instructions
        # Selection
        self.employee_instructions_label = QLabel(
            "<p><strong>Select one or more employees by name.</strong></p>"
            "<p><em>Coord</em> indicates the cell coordinate in the output budgeting "
            "file where the employee's name stands.</p>"
            )
        self.employee_instructions_label.setTextFormat(Qt.TextFormat.RichText)

        ## Projects table
        # Start with 0 rows and 2 columns
        self.employees_table = QTableWidget(0, len(dc.EmployeeTableHeaders))
        # Allow multiple row selections in the table
        self.employees_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.employees_table.setSelectionMode(QTableWidget.SelectionMode.MultiSelection)
        # Adjust the width of each column
        self.employees_table.setColumnWidth(dc.EmployeeTableHeaders.EMPLOYEE.value, 450)
        self.employees_table.setColumnWidth(dc.EmployeeTableHeaders.WORKSHEET.value, 150)
        self.employees_table.setColumnWidth(dc.EmployeeTableHeaders.COORDINATE.value, 80)

        ## Table selection buttons
        # Add deselect all button
        self.select_all_employees_button = QPushButton(dc.ButtonNames.SELECT_ALL.value, self)
        # Add deselect all button
        self.deselect_all_employees_button = QPushButton(dc.ButtonNames.DESELECT_ALL.value, self)

    def setup_layout(self) -> None:
        '''Init layout.'''

        ## Layout types
        # Main layout
        main_layout = QVBoxLayout()

        ## Add widgets
        # To main layout
        main_layout.addWidget(self.employee_instructions_label)
        main_layout.addWidget(self.employees_table)
        main_layout.addWidget(self.select_all_employees_button)
        main_layout.addWidget(self.deselect_all_employees_button)

        ## Display
        self.setLayout(main_layout)

    def setup_connections(self):
        '''Connect buttons to their respective actions.'''

        ## Connect functions
        # Select all
        self.select_all_employees_button.clicked.connect(
            self.employees_table.selectAll
        )
        # Deselect all
        self.deselect_all_employees_button.clicked.connect(
            self.employees_table.clearSelection
        )
        # Check isComplete
        self.employees_table.itemSelectionChanged.connect(
            self.completeChanged.emit
        )

    ############################
    ### Supporting functions ###
    ############################

    def populate_table(self, table: QTableWidget) -> None:
        '''Populate the table with the employees with
        coordinates from the given worksheet.'''

        # Get the first and only output workbook's file name
        wb_name = next(self.managed_workbooks.yield_workbooks_by_type(common.ManagedOutputWorkbook)).file_name
        wb = self.managed_workbooks.get_workbook_by_name(wb_name)
        ws = wb.managed_sheet_object.selected_sheet.sheet_object
        # Add employee per row. Skip empty cells from output file
        row_position = 0
        for col in range(
            wb.managed_sheet_object.employee_range.start_col_idx,
            wb.managed_sheet_object.employee_range.end_col_idx + 1
        ):
            cell = ws.cell(
                row=wb.managed_sheet_object.employee_range.start_row_idx,
                column=col
            )
            ## Populate row. No items editable after setting
            if cell.value and cell.value not in dc.NON_NAMES:
                employee_item = QTableWidgetItem(cell.value)
                # Insert row 0-indexed then add employee name
                table.insertRow(row_position)
                employee_item.setFlags(employee_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(
                    row_position,
                    dc.EmployeeTableHeaders.EMPLOYEE.value,
                    employee_item
                )
                # Add the worksheet name where the employee was found
                desc_text = wb.managed_sheet_object.selected_sheet.sheet_name
                # Add Excel cell coordinate where the employee is located
                desc_item = QTableWidgetItem(desc_text)
                desc_item.setFlags(desc_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(
                    row_position,
                    dc.EmployeeTableHeaders.WORKSHEET.value,
                    desc_item
                )
                # Add Excel cell coordinate where the employee is located
                coord_item = QTableWidgetItem(cell.coordinate)
                coord_item.setFlags(coord_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(
                    row_position,
                    dc.EmployeeTableHeaders.COORDINATE.value,
                    coord_item
                )
                row_position += 1

    ##################################
    ### QWizard function overrides ###
    ##################################

    def cleanupPage(self):
        '''Clean up if the back button is pressed.'''

        # Clear the table
        self.employees_table.clear()
        self.employees_table.setRowCount(0)
        # Clear the selected employees of the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(common.ManagedOutputWorkbook))
        out_wb.managed_sheet_object.clear_predicted_hours()
        for in_wb in self.managed_workbooks.yield_workbooks_by_type(common.ManagedInputWorkbook):
            in_wb.managed_sheet_object.selected_project_ids.clear()

    def isComplete(self) -> bool:
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''

        return bool(self.employees_table.selectionModel().selectedRows())

    def validatePage(self) -> bool:
        '''Override the page validation.
        Set the property.'''

        ## Selected employees
        # Set the selected employees in the managed output workbook
        selected_rows = self.employees_table.selectionModel().selectedRows()
        selected_employees = [
            (self.employees_table
            .item(row.row(), dc.EmployeeTableHeaders.COORDINATE.value)
            .text(),
            self.employees_table
            .item(row.row(), dc.EmployeeTableHeaders.EMPLOYEE.value)
            .text()
            )
            for row in selected_rows
        ]
        # This would be changed to `get_worksheet_by_name` if in the future, more than
        # one worksheet would be chosen. If across several workbooks too, make and use
        # get workbook by type and name, then get worksheet by name.
        # Here: Get the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(common.ManagedOutputWorkbook))
        out_wb.managed_sheet_object.set_selected_employees(selected_employees)
        ## Predicted hours per employee
        # Find hours
        pre_hours = gu.find_predicted_hours(
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
        hutils.sum_hours_selected_employee(self.managed_workbooks)
        for emp in out_wb.managed_sheet_object.selected_employees.values():
            emp.hours.set_deviation()
        # Validation complete
        return True
