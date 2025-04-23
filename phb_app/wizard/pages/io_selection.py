'''
Package
-------
PHB Wizard

Module Name
---------
IO Selection Page

Version
-------
Date-based Version: 20250304
Author: Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the Input-Output selection page.
'''
# Third party libaries
from PyQt6.QtWidgets import (
    QWizardPage,
    QHBoxLayout,
    QWidget,
    QPushButton,
    QTableWidget,
    QLabel
)
# First party libraries
from phb_app.logging.error_manager import ErrorManager
from phb_app.wizard.constants.ui_strings import (
    IO_FILE_TITLE,
    IO_FILE_SUBTITLE,
    I_FILE_INSTRUCTION_TEXT,
    O_FILE_INSTRUCTION_TEXT
)
import phb_app.utils.page_utils as pu
from phb_app.data.phb_dataclasses import (
    CountryData,
    WorkbookManager,
    InputTableHeaders,
    OutputTableHeaders,
    ButtonNames,
    FileDialogHandler,
    IORole,
    IOControls,
    INPUT_COLUMN_WIDTHS,
    OUTPUT_COLUMN_WIDTHS
)
#################################################################################
class IOSelectionPage(QWizardPage):
    '''Page for the selection of input and output files.
    Input files require the selection of worksheet containing the data and the file's origin.
    The ability to select multiple workbooks is provided by an \'add\' button.
    Output files require the selection of the output month and year, if the default gained
    from its file name is not desired. One or more project numbers must be selected for
    the output file.'''

    def __init__(self, country_data: CountryData, error_manager: ErrorManager, managed_workbooks: WorkbookManager) -> None:
        super().__init__()
        self.managed_workbooks = managed_workbooks
        error_manager.error_panels[IORole.INPUT_FILE] = QWidget()
        error_manager.error_panels[IORole.OUTPUT_FILE] = QWidget()
        pu.set_titles(self, IO_FILE_TITLE, IO_FILE_SUBTITLE)
        self.input_panel = IOControls(
            role=IORole.INPUT_FILE,
            label=QLabel(I_FILE_INSTRUCTION_TEXT),
            table=pu.create_table(InputTableHeaders, QTableWidget.SelectionMode.MultiSelection, INPUT_COLUMN_WIDTHS),
            buttons=[QPushButton(ButtonNames.ADD, self), QPushButton(ButtonNames.REMOVE, self)],
            error_panel=error_manager.error_panels[IORole.INPUT_FILE]
        )
        self.output_panel = IOControls(
            role=IORole.OUTPUT_FILE,
            label=QLabel(O_FILE_INSTRUCTION_TEXT),
            table=pu.create_table(OutputTableHeaders, QTableWidget.SelectionMode.SingleSelection, OUTPUT_COLUMN_WIDTHS),
            buttons=[QPushButton(ButtonNames.ADD, self), QPushButton(ButtonNames.REMOVE, self)],
            error_panel=error_manager.error_panels[IORole.OUTPUT_FILE]
        )
        pu.setup_page(self, [pu.create_interaction_panel(self.input_panel), pu.create_interaction_panel(self.output_panel)], QHBoxLayout())
        pu.connect_buttons(self, FileDialogHandler(self.input_panel, country_data, self.managed_workbooks, error_manager))
        pu.connect_buttons(self, FileDialogHandler(self.output_panel, country_data, self.managed_workbooks, error_manager))

    def isComplete(self) -> bool:
        '''Override the page completion.
        Check if both tables have at least one row selected
        and no error messages are displayed.'''
        return pu.check_completion(panels=(self.input_panel, self.output_panel))
