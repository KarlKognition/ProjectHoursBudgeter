'''
Package
-------
PHB Wizard

Module Name
---------
Project Selection Page

Version
-------
Date-based Version: 20250304
Author: Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the employee selection page.
'''

from PyQt6.QtWidgets import QWizardPage, QLabel, QTableWidget, QPushButton, QHBoxLayout
# First party libraries
import phb_app.data.workbook_management as wm
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.io_management as io
import phb_app.data.header_management as hm
import phb_app.utils.project_utils as pro

class ProjectSelectionPage(QWizardPage):
    '''Page for selecting the projects in which the hours were booked.'''
    def __init__(self, managed_workbooks: wm.WorkbookManager):
        super().__init__()
        self.wb_mgmt = managed_workbooks
        pu.set_titles(self, st.PROJECT_SELECTION_TITLE, st.PROJECT_SELECTION_SUBTITLE)
        self.project_panel = io.IOControls(
            page=self,
            role=st.IORole.PROJECT_TABLE,
            label=QLabel(st.PROJECT_SELECTION_INSTRUCTIONS),
            table=pu.create_table(ie.ProjectIDTableHeaders, QTableWidget.SelectionMode.MultiSelection, hm.PROJECT_COLUMN_WIDTHS),
            buttons=[QPushButton(st.ButtonNames.DESELECT_ALL, self)]
        )
        pu.setup_page(self, [pu.create_interaction_panel(self.project_panel)], QHBoxLayout())

#           --- QWizard function overrides ---

    def initializePage(self) -> None:
        '''Retrieve fields from other pages.'''
        pro.set_project_ids_each_input_wb(self.wb_mgmt)
        proj_ctx = io.EntryContext(self.project_panel, io.ProjectTableContext())
        io.EntryHandler(proj_ctx)
        pu.connect_buttons(self, proj_ctx)
        pu.populate_selection_table(self, proj_ctx, self.wb_mgmt)

    def cleanupPage(self):
        '''Override the page cleanup.
        Clear the table if the back button is pressed.'''
        self.project_panel.table.clear()
        self.project_panel.table.setRowCount(0)

    def isComplete(self) -> bool:
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''
        return pu.check_completion(self.project_panel)
