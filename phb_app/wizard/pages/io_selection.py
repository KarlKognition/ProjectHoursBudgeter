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
    QWidget,
    QComboBox,
    QPushButton,
    QTableWidget,
    QLabel
)
# First party libraries
from phb_app.protocols_callables.content import (
    IOControls,
    ErrorControls
)
from phb_app.wizard.constants.ui_strings import (
    IO_TITLE,
    IO_SUBTITLE,
    INPUT_INSTRUCTION_TEXT,
    OUTPUT_INSTRUCTION_TEXT
)
from phb_app.protocols_callables.content import (
    SetupWidgets,
    SetupLayout,
    SetupMainButtonConnections
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

    def __init__(self, country_data):
        super().__init__()

        # List of file paths for workbook object creation and deletion
        self.managed_workbooks = WorkbookManager()
        ## Init actions
        self.setTitle(IO_TITLE)
        self.setSubTitle(IO_SUBTITLE)
        self.input_controls = IOControls(
            buttons=QPushButton(...),
            table=QTableWidget(),
            label=QLabel(...),
            col_widths=self.INPUT_COLUMN_WIDTHS
        )
        self.output_controls = IOControls(
            buttons=QPushButton(...),
            table=QTableWidget(),
            label=QLabel(...),
            col_widths=self.OUTPUT_COLUMN_WIDTHS
        )

        self.error_controls = ErrorControls(
            error_panel=QWidget()
        )

        ## Instruction labels
        instr = setup.Instructions()
        instr.input_label = QLabel(INPUT_INSTRUCTION_TEXT)
        instr.output_label = QLabel(OUTPUT_INSTRUCTION_TEXT)
        input_table = setup.create_table(InputTableHeaders, QTableWidget.SelectionMode.MultiSelection)
        output_table = setup.create_table(OutputTableHeaders, QTableWidget.SelectionMode.SingleSelection)
        
        ## Buttons
        add_input_file_button = QPushButton(ButtonNames.ADD.value, self)
        remove_input_file_button = QPushButton(ButtonNames.REMOVE.value, self)
        add_output_file_button = QPushButton(ButtonNames.ADD.value, self)
        remove_output_file_button = QPushButton(ButtonNames.REMOVE.value, self)

        layout = setup.setup_layout(
            input_controls=self.input_controls,
            output_controls=self.output_controls,
            error_controls=self.error_controls
        )
        self.setLayout(layout)

        setup.setup_main_button_connections(
            input_controls=self.input_controls,
            output_controls=self.output_controls,
            managed_workbooks=self.managed_workbooks
        )

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
