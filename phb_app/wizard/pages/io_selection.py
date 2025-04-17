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
    QComboBox,
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
import phb_app.utils.page_setup_utils as setup
from phb_app.data.phb_dataclasses import (
    CountryData,
    WorkbookManager,
    InputTableHeaders,
    OutputTableHeaders,
    ButtonNames,
    SpecialStrings,
    OutputFile,
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
        self.country_data = country_data
        self.error_manager = error_manager
        self.managed_workbooks = managed_workbooks
        self.error_manager.error_panels[IORole.INPUT_FILE] = QWidget()
        self.error_manager.error_panels[IORole.OUTPUT_FILE] = QWidget()
        setup.set_titles(self, IO_FILE_TITLE, IO_FILE_SUBTITLE)
        self.input_panel = IOControls(
            role=IORole.INPUT_FILE,
            label=QLabel(I_FILE_INSTRUCTION_TEXT),
            table=setup.create_table(InputTableHeaders, QTableWidget.SelectionMode.MultiSelection, INPUT_COLUMN_WIDTHS),
            buttons=[QPushButton(ButtonNames.ADD, self), QPushButton(ButtonNames.REMOVE, self)],
            error_panel=error_manager.error_panels[IORole.INPUT_FILE]
        )
        self.output_panel = IOControls(
            role=IORole.OUTPUT_FILE,
            label=QLabel(O_FILE_INSTRUCTION_TEXT),
            table=setup.create_table(OutputTableHeaders, QTableWidget.SelectionMode.SingleSelection, OUTPUT_COLUMN_WIDTHS),
            buttons=[QPushButton(ButtonNames.ADD, self), QPushButton(ButtonNames.REMOVE, self)],
            error_panel=error_manager.error_panels[IORole.OUTPUT_FILE]
        )
        setup.set_page(self, [setup.create_interaction_panel(self.input_panel), setup.create_interaction_panel(self.output_panel)], QHBoxLayout())

    ##################################
    ### QWizard function overrides ###
    ##################################

    # def cleanupPage(self) -> None:
    #     '''Override the page cleanup.
    #     Remove file duplicates when the back button is pressed.'''

    def isComplete(self) -> bool:
        '''Override the page completion.
        Check if both tables have at least one row selected
        and no error messages are displayed.'''

        output_item: QComboBox = self.output_panel.table.cellWidget(
                OutputFile.FIRST_ENTRY.value,
                OutputTableHeaders.WORKSHEET.value
        )
        complete = (
            self.input_panel.table.rowCount() >= 1 and
            self.output_panel.table.rowCount() >= 1 and
            output_item is not None and
            output_item.currentText() != SpecialStrings.SELECT_WORKSHEET and
            not self.error_manager.errors
        )
        return complete
