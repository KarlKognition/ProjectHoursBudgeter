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
from typing import Optional
from PyQt6.QtWidgets import QWizardPage, QLabel, QTableWidget, QPushButton, QVBoxLayout
#           --- First party libraries ---
import phb_app.data.header_management as hm
import phb_app.data.io_management as io
import phb_app.data.workbook_management as wm
import phb_app.utils.employee_utils as eu
import phb_app.utils.page_utils as pu
import phb_app.logging.logger as logger
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

class SummaryPage(QWizardPage):
    '''Page for displaying the summary of results'''
    def __init__(self, managed_workbooks: wm.WorkbookManager):
        super().__init__()
        self.wb_mgmt = managed_workbooks
        self.out_wb_ctx: Optional[wm.OutputWorkbookContext] = None
        pu.set_titles(self, st.SUMMARY_TITLE, st.SUMMARY_SUBTITLE)
        self.summary_io_panel = io.IOControls(
            page=self,
            role=st.IORole.SUMMARY_IO_TABLE,
            label=QLabel(st.IO_SUMMARY),
            table=pu.create_row_header_table(
                page=self,
                table_headers=ie.SummaryIOTableHeaders,
                selection_mode=QTableWidget.SelectionMode.NoSelection
            ),
            buttons=None
        )
        self.summary_data_panel = io.IOControls(
            page=self,
            role=st.IORole.SUMMARY_DATA_TABLE,
            label=QLabel(st.SUMMARY_INSTRUCTIONS),
            table=pu.create_col_header_table(
                page=self,
                table_headers=ie.SummaryDataTableHeaders,
                selection_mode=QTableWidget.SelectionMode.MultiSelection,
                tab_widths=hm.SUMMARY_DATA_COLUMN_WIDTHS
            ),
            buttons=[QPushButton(st.ButtonNames.SELECT_ALL, self), QPushButton(st.ButtonNames.DESELECT_ALL, self)]
        )
        self.sum_io_ctx = io.EntryContext(self.summary_io_panel, io.SummaryIOContext())
        self.sum_data_ctx = io.EntryContext(self.summary_data_panel, io.SummaryDataContext())
        pu.setup_page(
            page=self,
            widgets=[pu.create_interaction_panel(self.summary_io_panel, 130), pu.create_interaction_panel(self.summary_data_panel)],
            layout_type=QVBoxLayout())
        self.setFinalPage(True)

#           --- QWizard function overrides ---

    def initializePage(self) -> None: # pylint: disable=invalid-name
        '''Override page initialisation. Setup page on each visit.'''
        self.out_wb_ctx = self.wb_mgmt.get_output_workbook_ctx()
        io.set_row_configurator(self.sum_io_ctx)
        pu.populate_io_summary_table(self, self.sum_io_ctx, self.wb_mgmt)
        io.set_row_configurator(self.sum_data_ctx)
        pu.connect_buttons(self, self.sum_data_ctx)
        pu.populate_summary_data_table(self, self.sum_data_ctx, self.wb_mgmt)

    def cleanupPage(self) -> None: # pylint: disable=invalid-name
        '''Clean up if the back button is pressed.'''
        self.summary_io_panel.table.clearContents()
        self.summary_io_panel.table.setColumnCount(0)
        self.summary_data_panel.table.clearContents()
        self.summary_data_panel.table.setRowCount(0)
        self.out_wb_ctx.worksheet_service.clear_selected_employees()

    def isComplete(self) -> bool: # pylint: disable=invalid-name
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''

        return bool(self.summary_data_panel.table.selectionModel().selectedRows())

    def validatePage(self) -> bool: # pylint: disable=invalid-name
        '''Override the page validation.'''
        eu.pop_unselected_employees(self.summary_data_panel.table, self.out_wb_ctx)
        logger.print_log(self.summary_data_panel.table, self.wb_mgmt)
        # Validation complete
        return True
