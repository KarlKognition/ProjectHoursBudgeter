'''
Package
-------
PHB Wizard

Module Name
---------
Project Selection Page

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the employee selection page.
'''

from PyQt6.QtWidgets import QWizardPage, QLabel, QTableWidget, QPushButton, QHBoxLayout
# First party libraries
import phb_app.data.header_management as hm
import phb_app.data.io_management as io
import phb_app.data.workbook_management as wm
import phb_app.utils.page_utils as pu
import phb_app.utils.project_utils as pro
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

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
            table=pu.create_col_header_table(
                page=self,
                table_headers=ie.ProjectIDTableHeaders,
                selection_mode=QTableWidget.SelectionMode.MultiSelection,
                tab_widths=hm.PROJECT_COLUMN_WIDTHS
            ),
            buttons=[QPushButton(st.ButtonNames.DESELECT_ALL, self)]
        )
        self.proj_ctx = io.EntryContext(self.project_panel, io.ProjectTableContext())
        pu.setup_page(
            page=self,
            widgets=[pu.create_interaction_panel(self.project_panel)],
            layout_type=QHBoxLayout()
        )

#           --- QWizard function overrides ---

    def initializePage(self) -> None: # pylint: disable=invalid-name
        '''Override page initialisation. Setup page on each visit.'''
        pro.set_project_ids_each_input_wb(self.wb_mgmt)
        io.set_row_configurator(self.proj_ctx)
        pu.connect_buttons(self, self.proj_ctx)
        pu.populate_project_table(self, self.proj_ctx, self.wb_mgmt)

    def cleanupPage(self) -> None: # pylint: disable=invalid-name
        '''Override the page cleanup.
        Clear the table if the back button is pressed.'''
        pu.clean_up_table(self.project_panel.table)

    def isComplete(self) -> bool: # pylint: disable=invalid-name
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''
        check = pu.check_completion(self.project_panel)
        rows = self.project_panel.table.selectionModel().selectedRows()
        for wb_ctx in self.wb_mgmt.yield_workbook_ctxs_by_role(st.IORole.INPUTS):
            pro.set_selected_project_ids(wb_ctx, self.project_panel.table, rows, ie.ProjectIDTableHeaders)
        return check
