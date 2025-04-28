'''
Package
-------
PHB Wizard

Module Name
---------
Summary Page

Version
-------
Date-based Version: 20250304
Author: Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the summary page.
'''

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWizardPage, QLabel, QTableWidget, QPushButton,
    QVBoxLayout, QTableWidgetItem
)
import phb_app.data.common as common
import phb_app.data.phb_dataclasses as dc
import phb_app.utils.general_func_utils as gu
import phb_app.logging.logger as logger

class SummaryPage(QWizardPage):
    '''Page for displaying the summary of results from the actions taken by
    the background functionality.'''
    def __init__(self):
        super().__init__()

        self.setTitle("Summary")
        self.setSubTitle(
            "Check all details before selecting which employee's hours should be recorded. "
            "Missing hours will be omitted. Red predicted hours implies hours have already "
            "been recorded and thus will not be overwritten. The project ID column displays from "
            "where the hours were taken."
        )
        # Headers
        self.headers = dc.SummaryTableHeaders.cap_members_list()
        self.wb_out: common.ManagedOutputWorkbook = None

        ## Init property
        # To be populated at page init
        self.selected_projects = []
        self.managed_workbooks: dc.WorkbookManager

        self.setup_widgets()
        self.setup_layout()
        self.setup_connections()
        self.setFinalPage(True)

    #######################################
    ### QWizard setup function override ###
    #######################################

    def initializePage(self) -> None:
        '''Retrieve fields from other pages.'''

        # Get workbooks from IOSelection
        self.managed_workbooks = self.wizard().property(
            dc.QPropName.MANAGED_WORKBOOKS.value)
        ## Info label text
        # Input names text
        input_names = "".join(
            f"<p style='margin:0;'> {wb.file_name}</p>"
            for wb
            in self.managed_workbooks.yield_workbooks_by_type(common.ManagedInputWorkbook))
        in_details_text = (
            "<table border='0'>"
                "<tr>"
                    "<td>"
                        "<strong>Input workbook(s)</strong>: "
                    "</td>"
                    "<td>"
                        f"{input_names}"
                    "</td>"
                "</tr>"
            "</table>"
        )
        # Get the first and only output workbook
        self.wb_out = next(self.managed_workbooks.yield_workbooks_by_type(
            common.ManagedOutputWorkbook))
        file_out_name = self.wb_out.file_name
        # Output name text
        out_details_text = (
            f"<p><strong>Output file name</strong>: {file_out_name}</p>"
        )
        # Selected date text
        date = self.wb_out.managed_sheet_object.selected_date
        month = gu.german_abbr_month(date.month, common.MONATE_KURZ_DE)
        selected_date_text = (
            f"<p><strong>Selected date</strong>: {month} {date.year}</p>"
        )
        # Set text
        self.selected_date_label.setText(selected_date_text)
        self.in_details_label.setText(in_details_text)
        self.out_details_label.setText(out_details_text)
        ## Table data
        # Table headers
        self.summary_table.setHorizontalHeaderLabels(self.headers)
        # Add all data to the table
        self.populate_table(self.summary_table)
        self.summary_table.viewport().update()
        # Get selected projects from the Project Selection page
        self.selected_projects = self.wizard().property(
            dc.QPropName.SELECTED_PROJECTS.value)

    #######################
    ### Setup functions ###
    #######################

    def setup_widgets(self) -> None:
        '''Init all widgets.'''

        ## Summary
        # Date
        self.selected_date_label = QLabel()
        self.selected_date_label.setTextFormat(Qt.TextFormat.RichText)
        # In
        self.in_details_label = QLabel()
        self.in_details_label.setTextFormat(Qt.TextFormat.RichText)
        # Out
        self.out_details_label = QLabel()
        self.out_details_label.setTextFormat(Qt.TextFormat.RichText)

        ## Projects table
        # Start with 0 rows and 2 columns
        self.summary_table = QTableWidget(0, len(dc.SummaryTableHeaders))
        # Allow multiple row selections in the table
        self.summary_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.summary_table.setSelectionMode(QTableWidget.SelectionMode.MultiSelection)
        # Adjust the width of each column
        self.summary_table.setColumnWidth(dc.SummaryTableHeaders.EMPLOYEE.value, 250)
        self.summary_table.setColumnWidth(dc.SummaryTableHeaders.PREDICTED_HOURS.value, 100)
        self.summary_table.setColumnWidth(dc.SummaryTableHeaders.ACCUMULATED_HOURS.value, 120)
        self.summary_table.setColumnWidth(dc.SummaryTableHeaders.DEVIATION.value, 160)
        self.summary_table.setColumnWidth(dc.SummaryTableHeaders.PROJECT_ID.value, 450)
        self.summary_table.setColumnWidth(dc.SummaryTableHeaders.OUTPUT_WORKSHEET.value, 150)
        self.summary_table.setColumnWidth(dc.SummaryTableHeaders.COORDINATE.value, 80)

        ## Table selection control buttons
        # Add deselect all button
        self.select_employees_button = QPushButton(dc.ButtonNames.SELECT_ALL.value, self)
        # Add deselect all button
        self.deselect_employees_button = QPushButton(dc.ButtonNames.DESELECT_ALL.value, self)

    def setup_layout(self) -> None:
        '''Init layout.'''

        ## Layout types
        # Main layout
        main_layout = QVBoxLayout()

        ## Add widgets
        # To main layout
        main_layout.addWidget(self.selected_date_label)
        main_layout.addWidget(self.in_details_label)
        main_layout.addWidget(self.out_details_label)
        main_layout.addWidget(self.summary_table)
        main_layout.addWidget(self.select_employees_button)
        main_layout.addWidget(self.deselect_employees_button)

        ## Display
        self.setLayout(main_layout)

    def setup_connections(self):
        '''Connect buttons to their respective actions.'''

        ## Connect functions
        # Select all
        self.select_employees_button.clicked.connect(
            self.summary_table.selectAll
        )
        # Deselect all
        self.deselect_employees_button.clicked.connect(
            self.summary_table.clearSelection
        )
        # Check isComplete
        self.summary_table.itemSelectionChanged.connect(
            self.completeChanged.emit
        )

    ############################
    ### Supporting functions ###
    ############################

    def populate_table(self, table: QTableWidget) -> None:
        '''Populate the table with the selected employees with their
        predicted and recorded hours per project ID.'''

        # Add selected employee per row with respective hours
        for row_position, emp in enumerate(
            self.wb_out.managed_sheet_object.yield_from_selected_employee()):
            ## Populate row. No items editable after setting
            # Insert row 0-indexed then add employee name
            table.insertRow(row_position)
            # Add the employee name
            employee_item = QTableWidgetItem(emp.name)
            employee_item.setFlags(employee_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                dc.EmployeeTableHeaders.EMPLOYEE.value,
                employee_item
            )
            # Add the predicted hours
            if emp.hours.predicted_hours:
                pre_hours_text = f"{emp.hours.predicted_hours:.2f}"
            else:
                pre_hours_text = common.SpecialStrings.ZERO_HOURS.value
            pre_hours_item = QTableWidgetItem(pre_hours_text)
            pre_hours_item.setForeground(emp.hours.pre_hours_colour)
            pre_hours_item.setTextAlignment(Qt.AlignmentFlag.AlignRight)
            pre_hours_item.setFlags(pre_hours_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                dc.SummaryTableHeaders.PREDICTED_HOURS.value,
                pre_hours_item
            )
            # Add the accumulated hours
            if emp.hours.accumulated_hours is None:
                acc_hours_text = common.SpecialStrings.MISSING.value
                font_colour = Qt.GlobalColor.red
            else:
                acc_hours_text = f"{emp.hours.accumulated_hours:.2f}"
                font_colour = Qt.GlobalColor.black
            acc_hours_item = QTableWidgetItem(str(acc_hours_text))
            acc_hours_item.setForeground(font_colour)
            acc_hours_item.setTextAlignment(Qt.AlignmentFlag.AlignRight)
            acc_hours_item.setFlags(acc_hours_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                dc.SummaryTableHeaders.ACCUMULATED_HOURS.value,
                acc_hours_item
            )
            # Add the hours deviation
            dev_item_text = emp.hours.deviation
            dev_item = QTableWidgetItem(dev_item_text)
            dev_item.setTextAlignment(Qt.AlignmentFlag.AlignRight)
            dev_item.setFlags(acc_hours_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                dc.SummaryTableHeaders.DEVIATION.value,
                dev_item
            )
            # Add the related project ID
            proj_id_text = "\n".join(
                f"{id_item}: {'; '.join(desc)}" for id_item, desc in emp.found_projects.items())
            proj_id_item = QTableWidgetItem(proj_id_text)
            proj_id_item.setFlags(proj_id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                dc.SummaryTableHeaders.PROJECT_ID.value,
                proj_id_item
            )
            # Resize the row according to the number of rows
            # in the project ID cell
            table.resizeRowToContents(row_position)
            # Add the worksheet name where the employee was found
            desc_text = self.wb_out.managed_sheet_object.selected_sheet.sheet_name
            # Add Excel cell coordinate where the employee is located
            desc_item = QTableWidgetItem(desc_text)
            desc_item.setFlags(desc_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                dc.SummaryTableHeaders.OUTPUT_WORKSHEET.value,
                desc_item
            )
            # Add Excel cell coordinate where the employee is located
            emp_coord = next(
                (c for c, e in self.wb_out.managed_sheet_object.selected_employees.items()
                 if e.name == emp.name),
                None)
            coord_item = QTableWidgetItem(emp_coord)
            coord_item.setFlags(coord_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                dc.SummaryTableHeaders.COORDINATE.value,
                coord_item
            )

    ##################################
    ### QWizard function overrides ###
    ##################################

    def cleanupPage(self):
        '''Clean up if the back button is pressed.'''

        # Clear the table
        self.summary_table.clear()
        self.summary_table.setRowCount(0)
        # Clear the selected employees from the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(common.ManagedOutputWorkbook))
        out_wb.managed_sheet_object.clear_predicted_hours()

    def isComplete(self) -> bool:
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''

        return bool(self.summary_table.selectionModel().selectedRows())

    def validatePage(self) -> bool:
        '''Override the page validation.
        Set the property.'''

        # Set the selected employees in the managed output workbook
        unselected_coords = []
        for row in range(self.summary_table.rowCount()):
            if not self.summary_table.selectionModel().isRowSelected(row):
                coord = self.summary_table.item(row, dc.SummaryTableHeaders.COORDINATE.value).text()
                unselected_coords.append(coord)
        # This would be changed to `get_worksheet_by_name` if in the future more than
        # one worksheet would be chosen for recording hours. If across several workbooks
        # too, make and use get workbook by type and name, then get worksheet by name.
        # Here: Get the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(common.ManagedOutputWorkbook))
        # Pop the unselected employees from the selected employees dictionary
        for coord in unselected_coords:
            out_wb.managed_sheet_object.selected_employees.pop(coord, None)
        # Print the summary page to an Excel file as a log
        logger.print_log(self.summary_table, self.managed_workbooks)
        # Validation complete
        return True
