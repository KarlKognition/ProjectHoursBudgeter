'''
Package
-------
PHB Wizard

Module Name
---------
IO Selection Page

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the Input-Output selection page.
'''
# Third party libaries
from PyQt6.QtWidgets import (
    QWizardPage, QHBoxLayout, QPushButton,
    QTableWidget, QLabel
)
# First party libraries
import phb_app.data.header_management as hm
import phb_app.data.io_management as io
import phb_app.data.location_management as loc
import phb_app.data.workbook_management as wm
import phb_app.logging.error_manager as em
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

class IOSelectionPage(QWizardPage):
    '''Page for the selection of input and output files.
    Input files require the selection of worksheet containing the data and the file's origin.
    The ability to select multiple workbooks is provided by an \'add\' button.
    Output files require the selection of the output month and year, if the default gained
    from its file name is not desired. One or more project numbers must be selected for
    the output file.'''

    def __init__(
        self,
        country_data: loc.CountryData,
        wb_mgmt: wm.WorkbookManager
        ) -> None:

        super().__init__()
        pu.setup_error_panel(st.IORole.INPUTS)
        pu.setup_error_panel(st.IORole.OUTPUT)
        pu.set_titles(self, st.IO_FILE_TITLE, st.IO_FILE_SUBTITLE)
        self.input_panel = io.IOControls(
            page=self,
            role=st.IORole.INPUTS,
            label=QLabel(st.INPUT_FILE_INSTRUCTION_TEXT),
            table=pu.create_col_header_table(
                page=self,
                table_headers=ie.InputTableHeaders,
                selection_mode=QTableWidget.SelectionMode.MultiSelection,
                tab_widths=hm.INPUT_COLUMN_WIDTHS
            ),
            buttons=[QPushButton(st.ButtonNames.ADD, self), QPushButton(st.ButtonNames.REMOVE, self)],
            error_panel=em.error_panels[st.IORole.INPUTS]
        )
        self.output_panel = io.IOControls(
            page=self,
            role=st.IORole.OUTPUT,
            label=QLabel(st.OUTPUT_FILE_INSTRUCTION_TEXT),
            table=pu.create_col_header_table(
                page=self,
                table_headers=ie.OutputTableHeaders,
                selection_mode=QTableWidget.SelectionMode.SingleSelection,
                tab_widths=hm.OUTPUT_COLUMN_WIDTHS
            ),
            buttons=[QPushButton(st.ButtonNames.ADD, self), QPushButton(st.ButtonNames.REMOVE, self)],
            error_panel=em.error_panels[st.IORole.OUTPUT]
        )
        pu.setup_page(
            page=self,
            widgets=[pu.create_interaction_panel(self.input_panel), pu.create_interaction_panel(self.output_panel)],
            layout_type=QHBoxLayout()
        )
        in_ctx = io.EntryContext(self.input_panel, data=io.IOFileContext(country_data=country_data))
        io.set_row_configurator(in_ctx)
        pu.connect_buttons(self, in_ctx, wb_mgmt)
        out_ctx = io.EntryContext(self.output_panel, data=io.IOFileContext())
        io.set_row_configurator(out_ctx)
        pu.connect_buttons(self, out_ctx, wb_mgmt)

    def isComplete(self) -> bool: # pylint: disable=invalid-name
        '''Override the page completion.
        Check if both tables have at least one row selected
        and no error messages are displayed.'''
        return pu.check_completion(self.input_panel) and pu.check_completion(self.output_panel)
