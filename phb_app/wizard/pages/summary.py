'''
Package
-------
PHB Wizard

Module Name
---------
Summary Page

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the summary page.
'''

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QWizardPage, QLabel, QTableWidget, QPushButton, QVBoxLayout, QTableWidgetItem
#           --- First party libraries ---
import phb_app.data.header_management as hm
import phb_app.data.io_management as io
import phb_app.data.months_dict as md
import phb_app.data.workbook_management as wm
import phb_app.utils.date_utils as du
import phb_app.utils.page_utils as pu
import phb_app.logging.logger as logger
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

class SummaryPage(QWizardPage):
    '''Page for displaying the summary of results'''
    def __init__(self, managed_workbooks: wm.WorkbookManager):
        super().__init__()
        self.wb_mgmt = managed_workbooks
        pu.set_titles(self, st.SUMMARY_TITLE, st.SUMMARY_SUBTITLE)
        self.summary_io_panel = io.IOControls(
            page=self,
            role=st.IORole.SUMMARY_IO_TABLE,
            label=QLabel(st.IO_SUMMARY),
            table=pu.create_table(
                page=self,
                table_headers=ie.SummaryIOTableHeaders,
                selection_mode=QTableWidget.SelectionMode.NoSelection,
                col_widths=hm.SUMMARY_IO_COLUMN_WIDTH,
                vertical_headers=True
            ),
            buttons=None
        )
        self.summary_data_panel = io.IOControls(
            page=self,
            role=st.IORole.SUMMARY_DATA_TABLE,
            label=QLabel(st.SUMMARY_INSTRUCTIONS),
            table=pu.create_table(
                page=self,
                table_headers=ie.SummaryDataTableHeaders,
                selection_mode=QTableWidget.SelectionMode.MultiSelection,
                col_widths=hm.SUMMARY_DATA_COLUMN_WIDTHS
            ),
            buttons=[QPushButton(st.ButtonNames.SELECT_ALL, self), QPushButton(st.ButtonNames.DESELECT_ALL, self)]
        )
        self.sum_io_ctx = io.EntryContext(self.summary_io_panel, io.SummaryIOContext())
        self.sum_data_ctx = io.EntryContext(self.summary_data_panel, io.SummaryDataContext())
        pu.setup_page(
            page=self,
            widgets=[pu.create_interaction_panel(self.summary_io_panel), pu.create_interaction_panel(self.summary_data_panel)],
            layout_type=QVBoxLayout())
        self.setFinalPage(True)

#           --- QWizard function overrides ---

    def initializePage(self) -> None:
        '''Override page initialisation. Setup page on each visit.'''

        ## Info label text
        # Input names text
        input_names = "".join(
            f"<p style='margin:0;'> {wb.file_name}</p>"
            for wb
            in self.managed_workbooks.yield_workbooks_by_type(wm.ManagedInputWorkbook))
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
            wm.ManagedOutputWorkbook))
        file_out_name = self.wb_out.file_name
        # Output name text
        out_details_text = (
            f"<p><strong>Output file name</strong>: {file_out_name}</p>"
        )
        # Selected date text
        date = self.wb_out.managed_sheet_object.selected_date
        month = du.german_abbr_month(date.month, md.LOCALIZED_MONTHS_SHORT)
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
            st.QPropName.SELECTED_PROJECTS)

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
                hm.EmployeeTableHeaders.EMPLOYEE,
                employee_item
            )
            # Add the predicted hours
            if emp.hours.predicted_hours:
                pre_hours_text = f"{emp.hours.predicted_hours:.2f}"
            else:
                pre_hours_text = st.SpecialStrings.ZERO_HOURS
            pre_hours_item = QTableWidgetItem(pre_hours_text)
            pre_hours_item.setForeground(emp.hours.pre_hours_colour)
            pre_hours_item.setTextAlignment(Qt.AlignmentFlag.AlignRight)
            pre_hours_item.setFlags(pre_hours_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                hm.SummaryTableHeaders.PREDICTED_HOURS,
                pre_hours_item
            )
            # Add the accumulated hours
            if emp.hours.accumulated_hours is None:
                acc_hours_text = st.SpecialStrings.MISSING
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
                hm.SummaryTableHeaders.ACCUMULATED_HOURS,
                acc_hours_item
            )
            # Add the hours deviation
            dev_item_text = emp.hours.deviation
            dev_item = QTableWidgetItem(dev_item_text)
            dev_item.setTextAlignment(Qt.AlignmentFlag.AlignRight)
            dev_item.setFlags(acc_hours_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                hm.SummaryTableHeaders.DEVIATION,
                dev_item
            )
            # Add the related project ID
            proj_id_text = "\n".join(
                f"{id_item}: {'; '.join(desc)}" for id_item, desc in emp.found_projects.items())
            proj_id_item = QTableWidgetItem(proj_id_text)
            proj_id_item.setFlags(proj_id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                hm.SummaryTableHeaders.PROJECT_ID,
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
                hm.SummaryTableHeaders.OUTPUT_WORKSHEET,
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
                hm.SummaryTableHeaders.COORDINATE,
                coord_item
            )

    ##################################
    ### QWizard function overrides ###
    ##################################

    def cleanupPage(self) -> None: # pylint: disable=invalid-name
        '''Clean up if the back button is pressed.'''

        # Clear the table
        self.summary_table.clear()
        self.summary_table.setRowCount(0)
        # Clear the selected employees from the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(wm.ManagedOutputWorkbook))
        out_wb.managed_sheet_object.clear_predicted_hours()

    def isComplete(self) -> bool: # pylint: disable=invalid-name
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''

        return bool(self.summary_data_panel.table.selectionModel().selectedRows())

    def validatePage(self) -> bool: # pylint: disable=invalid-name
        '''Override the page validation.
        Set the property.'''

        # Set the selected employees in the managed output workbook
        unselected_coords = []
        for row in range(self.summary_table.rowCount()):
            if not self.summary_table.selectionModel().isRowSelected(row):
                coord = self.summary_table.item(row, hm.SummaryTableHeaders.COORDINATE).text()
                unselected_coords.append(coord)
        # This would be changed to `get_worksheet_by_name` if in the future more than
        # one worksheet would be chosen for recording hours. If across several workbooks
        # too, make and use get workbook by type and name, then get worksheet by name.
        # Here: Get the first and only output workbook
        out_wb = next(self.managed_workbooks.yield_workbooks_by_type(wm.ManagedOutputWorkbook))
        # Pop the unselected employees from the selected employees dictionary
        for coord in unselected_coords:
            out_wb.managed_sheet_object.selected_employees.pop(coord, None)
        # Print the summary page to an Excel file as a log
        logger.print_log(self.summary_table, self.managed_workbooks)
        # Validation complete
        return True
