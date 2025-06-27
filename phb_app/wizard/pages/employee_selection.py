'''
Package
-------
PHB Wizard

Module Name
---------
Employee Selection Page

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the employee selection page.
'''
from typing import Optional
from PyQt6.QtWidgets import QWizardPage, QLabel, QTableWidget, QPushButton, QHBoxLayout
#           --- First party libraries ---
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
        self.out_wb_ctx: Optional[wm.OutputWorkbookContext] = None
        pu.set_titles(self, st.PROJECT_SELECTION_TITLE, st.PROJECT_SELECTION_SUBTITLE)
        self.employee_panel = io.IOControls(
            page=self,
            role=st.IORole.EMPLOYEE_TABLE,
            label=QLabel(st.EMPLOYEE_SELECTION_INSTRUCTIONS),
            table=pu.create_col_header_table(
                page=self,
                table_headers=ie.EmployeeTableHeaders,
                selection_mode=QTableWidget.SelectionMode.MultiSelection,
                tab_widths=hm.EMPLOYEE_COLUMN_WIDTHS
            ),
            buttons=[QPushButton(st.ButtonNames.SELECT_ALL, self), QPushButton(st.ButtonNames.DESELECT_ALL, self)]
        )
        self.emp_ctx = io.EntryContext(self.employee_panel, io.EmployeeTableContext())
        pu.setup_page(
            page=self,
            widgets=[pu.create_interaction_panel(self.employee_panel)],
            layout_type=QHBoxLayout()
        )

#           --- QWizard function overrides ---

    def initializePage(self) -> None: # pylint: disable=invalid-name
        '''Override page initialisation. Setup page on each visit.'''
        self.out_wb_ctx = self.wb_mgmt.get_output_workbook_ctx()
        io.set_row_configurator(self.emp_ctx)
        pu.connect_buttons(self, self.emp_ctx)
        pu.populate_employee_table(self, self.emp_ctx, self.wb_mgmt)

    def cleanupPage(self) -> None: # pylint: disable=invalid-name
        '''Clean up if the back button is pressed.'''
        self.employee_panel.table.clear()
        self.employee_panel.table.setRowCount(0)

    def isComplete(self) -> bool: # pylint: disable=invalid-name
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''
        check = bool(self.employee_panel.table.selectionModel().selectedRows())
        eu.compute_selected_employees(self.employee_panel.table, self.out_wb_ctx)
        return check
    
    def validatePage(self) -> bool: # pylint: disable=invalid-name
        '''Override the page validation.'''
        hu.compute_predicted_hours(self.out_wb_ctx)
        hu.compute_hours_for_selected_employees(self.wb_mgmt, self.out_wb_ctx)
        return True
