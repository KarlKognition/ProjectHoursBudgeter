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
from phb_app.protocols_callables.customs import (
    IOControls
)
from phb_app.logging.error_manager import ErrorManager
from phb_app.wizard.constants.ui_strings import (
    IO_TITLE,
    IO_SUBTITLE,
    INPUT_INSTRUCTION_TEXT,
    OUTPUT_INSTRUCTION_TEXT
)
import phb_app.utils.setup_utils as setup
from phb_app.data.phb_dataclasses import (
    WorkbookManager,
    InputTableHeaders,
    OutputTableHeaders,
    ButtonNames,
    SpecialStrings,
    OutputFile,
    QPropName,
    IOTable
)

##################################
class IOSelectionPage(QWizardPage):
    '''Page for the selection of input and output files.
    Input files require the selection of worksheet containing the data and the file's origin.
    The ability to select multiple workbooks is provided by an \'add\' button.
    Output files require the selection of the output month and year, if the default gained
    from its file name is not desired. One or more project numbers must be selected for
    the output file.'''

    INPUT_COLUMN_WIDTHS = {
        InputTableHeaders.FILENAME: 250,
        InputTableHeaders.COUNTRY: 150,
        InputTableHeaders.WORKSHEET: 100,
    }
    OUTPUT_COLUMN_WIDTHS = {
        OutputTableHeaders.FILENAME: 250,
        OutputTableHeaders.WORKSHEET: 150,
        OutputTableHeaders.MONTH: 60,
        OutputTableHeaders.YEAR: 60,
    }

    def __init__(self, country_data, error_manager: ErrorManager, managed_workbooks: WorkbookManager) -> None:
        super().__init__()

        error_manager.error_panel[IOTable.INPUT] = QWidget()
        error_manager.error_panel[IOTable.OUTPUT] = QWidget()
        self.setTitle(IO_TITLE)
        self.setSubTitle(IO_SUBTITLE)
        self.input_panel = IOControls(
            role=IOTable.INPUT,
            label=QLabel(INPUT_INSTRUCTION_TEXT),
            table=setup.create_table(InputTableHeaders, QTableWidget.SelectionMode.MultiSelection, self.INPUT_COLUMN_WIDTHS),
            buttons=[QPushButton(ButtonNames.ADD, self), QPushButton(ButtonNames.REMOVE, self)],
            error_panel=error_manager.error_panel[IOTable.INPUT]
        )
        self.output_panel = IOControls(
            role=IOTable.OUTPUT,
            label=QLabel(OUTPUT_INSTRUCTION_TEXT),
            table=setup.create_table(OutputTableHeaders, QTableWidget.SelectionMode.SingleSelection, self.OUTPUT_COLUMN_WIDTHS),
            buttons=[QPushButton(ButtonNames.ADD, self), QPushButton(ButtonNames.REMOVE, self)],
            error_panel=error_manager.error_panel[IOTable.OUTPUT]
        )
        setup.join_panels(self, layout=QHBoxLayout(), containers=(setup.create_interaction_panel(self.input_panel), setup.create_interaction_panel(self.output_panel)))

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

        output_item: QComboBox = self.output_table.cellWidget(
                OutputFile.FIRST_ENTRY.value,
                OutputTableHeaders.WORKSHEET.value
        )
        complete = (
            self.input_table.rowCount() >= 1 and
            self.output_table.rowCount() >= 1 and
            output_item is not None and
            output_item.currentText() != SpecialStrings.SELECT_WORKSHEET.value and
            not self.error_manager.errors
        )
        return complete

    def validatePage(self) -> bool:
        '''Override the page validation.
        Set the property.'''

        # Set the QWizard property with the collected workbooks
        # for use in the next page.
        self.wizard().setProperty(
            QPropName.MANAGED_WORKBOOKS.value,
            self.get_workbooks()
        )
        # Validation complete
        return True
