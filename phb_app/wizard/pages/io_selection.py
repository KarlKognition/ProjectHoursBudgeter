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

## Imports
# Standard libraries
from os import path
from datetime import datetime
from dateutil.relativedelta import relativedelta
# Third party libaries
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWizardPage,
    QHBoxLayout,
    QWidget,
    QLabel,
    QPushButton,
    QFileDialog,
    QComboBox,
    QVBoxLayout,
    QTableWidget,
    QTableWidgetItem
)
# First party libraries
from phb_app.logging.error_manager import ErrorManager
import phb_app.utils.func_utils as futils
from phb_app.data.phb_dataclasses import (
    WorkbookManager,
    CountryData,
    IOTableStartHeader,
    InputTableHeaders,
    OutputTableHeaders,
    ButtonNames,
    SpecialStrings,
    IOTable,
    ManagedInputWorkbook,
    CountriesEnum,
    OutputFile,
    QPropName,
    MONTHS
)
from phb_app.logging.exceptions import (
    FileAlreadySelected,
    TooManyOutputFilesSelected,
    CountryIdentifiersNotInFilename,
    IncorrectWorksheetSelected,
    MissingEmployeeRow,
    BudgetingDatesNotFound
)
from phb_app.protocols.phb_Protocols import (
    ConfigureRow,
    AddWorkbook
)

##################################
class IOSelectionPage(QWizardPage):
    '''Page for the selection of input and output files.
    Input files require the selection of worksheet containing the data and the file's origin.
    The ability to select multiple workbooks is provided by an \'add\' button.
    Output files require the selection of the output month and year, if the default gained
    from its file name is not desired. One or more project numbers must be selected for
    the output file.'''

    def __init__(self):
        super().__init__()

        # List of file paths for workbook object creation and deletion
        self.managed_workbooks = WorkbookManager()
        ## Init actions
        self.setTitle("Input/Output Files")
        self.setSubTitle(
            "Manage the input timesheets and output project hours budgeting file. "
            "NOTE: The selected date of the output file is the row in which the hours will be written "
            "but also will affect whether hours are accumulated at all from the input files."
        )
        self.country_data = CountryData()
        self.setup_widgets()
        self.setup_layout()
        self.setup_main_button_connections()

    #######################
    ### Setup functions ###
    #######################

    def setup_widgets(self) -> None:
        '''Init all widgets.'''

        ## Error labels
        # Panel for error messages
        self.error_panel = QWidget()
        # Vertical layout of errors
        self.error_panel.setLayout(QVBoxLayout())
        self.error_manager = ErrorManager(self.error_panel)
        ## Input labels
        # Instructions
        self.input_instructions_label = QLabel(
            "Select the German PS18 and/or external timesheets as the input files.")

        ## Output labels
        # Instructions
        self.output_instructions_label = QLabel(
            "Select the project hours budgeting file as the target output file."
            )

        ## Input table
        # Start with 0 rows and so many columns as per headers
        self.input_table = QTableWidget(
            0,
            len(IOTableStartHeader) + len(InputTableHeaders)
        )
        # Allow multiple row selections in the table
        self.input_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.input_table.setSelectionMode(QTableWidget.SelectionMode.MultiSelection)
        # Table headers
        self.input_table.setHorizontalHeaderLabels(
            IOTableStartHeader.cap_members_list() + InputTableHeaders.cap_members_list())
        # Header alignment
        self.input_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
        # Adjust the width of each column
        self.input_table.setColumnWidth(IOTableStartHeader.FILENAME.value, 250)
        self.input_table.setColumnWidth(InputTableHeaders.COUNTRY.value, 150)
        self.input_table.setColumnWidth(InputTableHeaders.WORKSHEET.value, 100)

        ## Output table
        # Start with 0 rows and so many columns as per headers
        self.output_table = QTableWidget(
            0,
            len(IOTableStartHeader) + len(OutputTableHeaders)
        )
        # Allow row selection
        self.output_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        # Table headers
        self.output_table.setHorizontalHeaderLabels(
            IOTableStartHeader.cap_members_list() + OutputTableHeaders.cap_members_list())
        # Header alignment
        self.output_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
        # Adjust the width of each column
        self.output_table.setColumnWidth(IOTableStartHeader.FILENAME.value, 250)
        self.output_table.setColumnWidth(OutputTableHeaders.WORKSHEET.value, 150)
        self.output_table.setColumnWidth(OutputTableHeaders.MONTH.value, 60)
        self.output_table.setColumnWidth(OutputTableHeaders.YEAR.value, 60)

        ## File control buttons
        # Add input file button
        self.add_input_file_button = QPushButton(ButtonNames.ADD.value, self)
        # Remove input file button
        self.remove_input_file_button = QPushButton(ButtonNames.REMOVE.value, self)
        # Add output file button
        self.add_output_file_button = QPushButton(ButtonNames.ADD.value, self)
        # Remove output file button
        self.remove_output_file_button = QPushButton(ButtonNames.REMOVE.value, self)

    def setup_layout(self) -> None:
        '''Init layout.'''

        ## Layout types
        # Main layout
        page_layout = QVBoxLayout()
        selection_container = QHBoxLayout() # All stuff input/putput
        # Input layout
        input_container = QTableWidget()
        input_container.setMaximumWidth(485)
        input_panel = QVBoxLayout()
        input_button_layout = QHBoxLayout()
        # Output layout
        output_container = QTableWidget()
        output_container.setMaximumWidth(485)
        output_panel = QVBoxLayout()
        output_button_layout = QHBoxLayout()

        ## Add widgets
        # To input button layout
        input_button_layout.addWidget(self.add_input_file_button)
        input_button_layout.addWidget(self.remove_input_file_button)
        # To input layout
        input_panel.addWidget(self.input_instructions_label)
        input_panel.addWidget(self.input_table)
        # To output button layout
        output_button_layout.addWidget(self.add_output_file_button)
        output_button_layout.addWidget(self.remove_output_file_button)
        # To outputs layout
        output_panel.addWidget(self.output_instructions_label)
        output_panel.addWidget(self.output_table)

        ## Add layouts heirarchically
        input_panel.addLayout(input_button_layout)
        input_container.setLayout(input_panel)
        output_panel.addLayout(output_button_layout)
        output_container.setLayout(output_panel)

        ## Containers
        # Add to selection container
        selection_container.addWidget(input_container)
        selection_container.addWidget(output_container)
        # Main container
        page_layout.addLayout(selection_container)
        page_layout.addWidget(self.error_panel)

        ## Display
        self.setLayout(page_layout)

    def setup_main_button_connections(self):
        '''Connect buttons to their respective actions.'''

        ## Connect functions
        # Add input files
        self.add_input_file_button.clicked.connect(
            lambda: self.add_file_dialog(
                QFileDialog.FileMode.ExistingFiles,
                self.input_table,
                self.managed_workbooks
            )
        )
        # Add output file
        self.add_output_file_button.clicked.connect(
            lambda: self.add_file_dialog(
                QFileDialog.FileMode.ExistingFile,
                self.output_table,
                self.managed_workbooks
            )
        )

        ## Remove files
        #Input files
        self.remove_input_file_button.clicked.connect(
            lambda: self.remove_selected_file(
                self.input_table,
                self.managed_workbooks
            )
        )
        # Output file
        self.remove_output_file_button.clicked.connect(
            lambda: self.remove_selected_file(
                self.output_table,
                self.managed_workbooks
            )
        )

    ############################
    ### Supporting functions ###
    ############################

    def add_file_dialog(self,
                        file_mode: QFileDialog.FileMode,
                        target_table: QTableWidget,
                        wb_manager: WorkbookManager
                        ) -> None:
        '''Add files to either the input or output selection tables.'''

        ## Set up file dialog box
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("Add File")
        # Only search for Excel files
        file_dialog.setNameFilter("Excel files (*.xlsx)")
        # Allow multiple file selection
        file_dialog.setFileMode(file_mode)
        file_dialog.setViewMode(QFileDialog.ViewMode.Detail)
        # Select files if add button clicked
        if file_dialog.exec():
            if target_table == self.input_table:
                selected_files = file_dialog.selectedFiles()
                # Add row in input table
                self.populate_table(
                    target_table,
                    selected_files,
                    wb_manager.add_input_workbook, # -> add_workbook
                    self.configure_input_row
                )
            elif target_table == self.output_table:
                selected_file = file_dialog.selectedFiles()
                # Add row in output table
                self.populate_table(
                    target_table,
                    selected_file,
                    wb_manager.add_output_workbook, # -> add_workbook
                    self.configure_output_row
                )

    def populate_table(self,
                       table: QTableWidget,
                       files: list[str] | str,
                       add_workbook: AddWorkbook,
                       configure_row: ConfigureRow
                       ) -> None:
        '''Populate the table with selected file(s).'''

        for file_path in files:
            file_name = path.basename(file_path)
            try:
                # Check if the file is already in either table
                file_exists = futils.is_file_already_in_table(
                    file_name,
                    IOTableStartHeader.FILENAME.value,
                    self.input_table,
                    self.output_table
                )
                # Insert row into table at the end of all existing rows
                row_position = table.rowCount()
                table.insertRow(row_position)
                # Item is uneditable.
                file_item = QTableWidgetItem(file_name)
                file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                # Display the file name in first column
                table.setItem(
                    row_position,
                    IOTableStartHeader.FILENAME.value,
                    file_item
                )
                # If there are 2 or more rows in the output table: error
                if table is self.output_table and table.rowCount() >= 2:
                    # Highlight it red
                    self.highlight_bad_item(file_item)
                    raise TooManyOutputFilesSelected()
                # If file is a duplicate: error
                if file_exists:
                    # Highlight it red
                    self.highlight_bad_item(file_item)
                    raise FileAlreadySelected(file_name)
                # Delegate row configuration to the provided callable
                if table is self.output_table:
                    # Create a tracked managed workbook without formulae
                    add_workbook(file_path, True)
                    # Create a tracked managed workbook with formulae
                    add_workbook(file_path, False)
                else:
                    add_workbook(file_path)
                configure_row(table, row_position, file_name, file_path)
            except Exception as e:
                file_role = self.check_file_role(table)
                self.error_manager.add_error(file_name, file_role, e)
            # Update UI
            self.completeChanged.emit()

    def configure_input_row(self,
                            table: QTableWidget,
                            row_position: int,
                            file_name: str,
                            file_path: str) -> None:
        '''Configure a row for the input table. File path is unused.'''

        ## Workbook entry
        # Work with the current workbook
        workbook_entry = self.managed_workbooks.get_workbook_by_name(file_name)
        try:
            ## Country
            # Get location from the file name
            self.update_country_details_in_table(workbook_entry)
            country = workbook_entry.locale_data.country
            country_item = QTableWidgetItem(country)
            # Country name is uneditable once displayed
            country_item.setFlags(country_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(
                row_position,
                InputTableHeaders.COUNTRY.value,
                country_item)
            ## Worksheet
            workbook_entry.init_input_worksheet()
            worksheet_item = QTableWidgetItem(workbook_entry.managed_sheet_object.selected_sheet.sheet_name)
            # Worksheet name is uneditable once displayed
            worksheet_item.setFlags(worksheet_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            # Display which worksheet is selected
            table.setItem(
                row_position,
                InputTableHeaders.WORKSHEET.value,
                worksheet_item)
        except CountryIdentifiersNotInFilename as e:
            item = table.item(row_position, IOTableStartHeader.FILENAME.value)
            self.highlight_bad_item(item)
            raise e
        # Update UI
        self.completeChanged.emit()

    def configure_output_row(self,
                             table: QTableWidget,
                             row_position: int,
                             file_name: str,
                             file_path: str) -> None:
        '''Configure a row for the output table.'''

        ## Workbook entry
        # Work with current workbook
        # wf = with formulae
        workbook_entry_wf = self.managed_workbooks.get_workbook_by_name(file_name)
        # Init the worksheet with formulae
        workbook_entry_wf.init_output_worksheet()
        ## Dropdowns
        # Worksheets
        worksheet_dropdown = QComboBox()
        worksheet_dropdown.addItem(SpecialStrings.SELECT_WORKSHEET.value)
        worksheet_dropdown.addItems(workbook_entry_wf.managed_sheet_object.sheet_names)
        table.setCellWidget(
            row_position,
            OutputTableHeaders.WORKSHEET.value,
            worksheet_dropdown
        )
        # Item for error and highlighting tracking
        file_item = table.item(row_position, IOTableStartHeader.FILENAME.value)
        # Update the selected worksheet upon dropdown menu selection
        def handle_selections() -> None:
            '''Provide the try-except block with the set selected sheet
            function to the connect function of the worksheet dropdown.'''
            # Update managed worksheet data
            try:
                file_role = self.check_file_role(table)
                # Remove any associated error
                self.error_manager.remove_error(
                    file_name,
                    file_role
                )
                self.remove_highlighting(file_item)
                # Get the sheet name visible in the dropdown
                selected_sheet = worksheet_dropdown.currentText()
                # Set the selected sheet in the managed sheet object
                workbook_entry_wf.managed_sheet_object.set_selected_sheet(
                    sheet_name=selected_sheet,
                    sheet_object=workbook_entry_wf.workbook_object[selected_sheet]
                )
                # Get the dates visible in the dropdowns
                selected_month = month_dropdown.currentText()
                selected_year = year_dropdown.currentText()
                # Set the budgeting date info for the selected sheet
                workbook_entry_wf.managed_sheet_object.set_budgeting_date(
                    file_path,
                    selected_sheet,
                    selected_month,
                    selected_year
                )
            except (IncorrectWorksheetSelected, BudgetingDatesNotFound, MissingEmployeeRow, KeyError) as e:
                self.highlight_bad_item(file_item)
                self.error_manager.add_error(file_name, file_role, e)
            # Update UI
            self.completeChanged.emit()
        # One month before for when the hours are to be budgeted
        date_minus_one_month = datetime.now() + relativedelta(months=-1)
        # Two years later to keep the list of years a bit in the future
        current_year_plus_two = (datetime.now() + relativedelta(years=+2)).year
        # Default values
        default_year = str(date_minus_one_month.year)
        default_month = futils.german_abbr_month(date_minus_one_month.month, MONTHS)
        # Year
        year_dropdown = QComboBox()
        # Add the range of years
        year_dropdown.addItems([str(year) for year in range(2020, current_year_plus_two)])
        # Display the current year as default
        year_dropdown.setCurrentText(default_year)
        table.setCellWidget(row_position, OutputTableHeaders.YEAR.value, year_dropdown)
        # Month
        month_dropdown = QComboBox()
        # Use months defined in the PHB data class module
        month_dropdown.addItems(list(MONTHS.keys()))
        month_dropdown.setCurrentText(default_month)
        table.setCellWidget(row_position, OutputTableHeaders.MONTH.value, month_dropdown)
        ## Dropdown action connections
        # Connect the sheet selection
        worksheet_dropdown.currentIndexChanged.connect(handle_selections)
        # Connect the year selection
        year_dropdown.currentIndexChanged.connect(handle_selections)
        # Connect the month selection
        month_dropdown.currentIndexChanged.connect(handle_selections)

    def remove_selected_file(self, table: QTableWidget, wb_manager: WorkbookManager) -> None:
        '''Remove the currently selected file(s) from the table.'''

        # Selected rows
        selected_items = table.selectionModel().selectedRows()
        # Search for selected rows in reverse so that the row numbering
        # doesn't change with each deletion.
        for row in sorted(
            [item.row() for item in selected_items],
            reverse=True
        ):
            file_name = table.item(row, IOTableStartHeader.FILENAME.value).text()
            # Remove the workbook
            wb_manager.remove_workbook(file_name)
            # Remove selected rows
            table.removeRow(row)
            # Check the file role
            file_role = self.check_file_role(table)
            # Remove any associated error
            self.error_manager.remove_error(
                file_name,
                file_role
                )
            # Update UI
            self.completeChanged.emit()

    def highlight_bad_item(self, item: QTableWidgetItem) -> None:
        '''Highlight table items causing errors.'''

        item.setBackground(Qt.GlobalColor.red)
        item.setForeground(Qt.GlobalColor.white)

    def remove_highlighting(self, item: QTableWidgetItem) -> None:
        '''Removes the highlighting after error correction.'''

        item.setBackground(Qt.GlobalColor.white)
        item.setForeground(Qt.GlobalColor.black)

    def check_file_role(self, table: QTableWidget) -> str:
        '''Checks and returns whether the file is an inout or output.'''
        if table is self.input_table:
            return IOTable.INPUT.value
        return IOTable.OUTPUT.value

    def update_country_details_in_table(self, workbook_entry: ManagedInputWorkbook) -> None:
        '''Update country details.'''

        # Check that the workbook is an input workbook
        if isinstance(workbook_entry, ManagedInputWorkbook):
            # Get country name from file name
            country_name = futils.get_origin_from_file_name(
                workbook_entry.file_name,
                self.country_data,
                CountriesEnum)
            # Set the local data of the workbook
            workbook_entry.set_locale_data(self.country_data, country_name)

    def get_workbooks(self) -> WorkbookManager:
        '''Return list of workbooks.'''

        return self.managed_workbooks

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
