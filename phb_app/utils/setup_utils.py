'''
Package
-------
PHB Wizard

Module Name
---------
Page Setup

Version
-------
Date-based Version: 20250504
Author: Karl Goran Antony Zuvela

Description
-----------
All functions necessary for setting up the wizard pages.
'''
## Imports
# Standard libraries
from os import path
from datetime import datetime
from dateutil.relativedelta import relativedelta
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWizardPage,
    QBoxLayout,
    QHBoxLayout,
    QComboBox,
    QWidget,
    QFileDialog,
    QVBoxLayout,
    QTableWidget,
    QTableWidgetItem
)
# First party libraries
import phb_app.utils.func_utils as futils
from phb_app.data.phb_dataclasses import (
    WorkbookManager,
    BaseTableHeaders,
    InputTableHeaders,
    OutputTableHeaders,
    IORole,
    ManagedInputWorkbook,
    SpecialStrings,
    CountriesEnum,
    IOControls,
    ColWidths,
    MONATE_KURZ_DE
)
from phb_app.logging.exceptions import (
    FileAlreadySelected,
    TooManyOutputFilesSelected,
    CountryIdentifiersNotInFilename,
    IncorrectWorksheetSelected,
    MissingEmployeeRow,
    BudgetingDatesNotFound
)
from phb_app.protocols_callables.customs import (
    ConfigureRow,
    AddWorkbook
)
from phb_app.wizard.constants.ui_strings import (
    ADD_FILE,
    EXCEL_FILE
)

#############################
### Table and Panel Setup ###
#############################

def create_table(table_headers: BaseTableHeaders, selection_mode: QTableWidget.SelectionMode, col_widths: ColWidths) -> QTableWidget:
    '''Create a table widget with the given headers and selection mode.'''
    table = QTableWidget(0, len(table_headers))
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(selection_mode)
    table.setHorizontalHeaderLabels(table_headers.cap_members_list())
    table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
    for header, width in col_widths.items():
        table.setColumnWidth(header.value, width)
    return table

def create_interaction_panel(panel: IOControls) -> QWidget:
    '''Set up the table with the given buttons and column widths.'''
    container = QWidget()
    layout = QVBoxLayout()
    layout.addWidget(panel.label)
    layout.addWidget(panel.table)
    buttons_layout = QHBoxLayout()
    for button in panel.buttons:
        buttons_layout.addWidget(button)
    layout.addLayout(buttons_layout)
    layout.addWidget(panel.error_panel)
    container.setLayout(layout)
    return container

def set_page(*containers: QWidget, page: QWizardPage, layout: QBoxLayout) -> None:
    '''Create the final layout for the page.'''
    for container in containers:
        layout.addWidget(container)
    page.setLayout(layout)

###############
### Buttons ###
###############

def connect_buttons(page: QWizardPage, panel: IOControls, managed_workbooks: WorkbookManager) -> None:
    '''Connect buttons to their respective actions dynamically.'''

    # Define a function dispatch table
    def add_action():
        add_file_dialog(page, QFileDialog.FileMode.ExistingFiles, panel, managed_workbooks)

    def remove_action():
        remove_selected_file(page, panel, managed_workbooks)

    action_dispatch = {
        "add_button": add_action,
        "remove_button": remove_action,
        "select_all_button": panel.table.selectAll,
        "deselect_all_button": panel.table.clearSelection
    }

    # Iterate over buttons and connect them dynamically
    for button in panel.buttons:
        action = action_dispatch.get(button.objectName())
        if action:
            button.clicked.connect(action)

#####################
### File Handling ###
#####################

def setup_file_dialog(page: QWizardPage, file_mode: QFileDialog.FileMode) -> QFileDialog:
    '''Set up and return a file dialog.'''
    file_dialog = QFileDialog(page)
    file_dialog.setWindowTitle(ADD_FILE)
    file_dialog.setNameFilter(EXCEL_FILE)
    file_dialog.setFileMode(file_mode)
    file_dialog.setViewMode(QFileDialog.ViewMode.Detail)
    return file_dialog

def handle_file_selection(panel: IOControls, selected_files: list[str], page: QWizardPage, wb_manager: WorkbookManager) -> None:
    '''Handle file selection and populate the appropriate table.'''
    if panel.role == IORole.INPUT_FILE:
        populate_table(page, panel, selected_files, wb_manager.add_input_workbook, configure_input_row)
    elif panel.role == IORole.OUTPUT_FILE:
        populate_table(page, panel, selected_files, wb_manager.add_output_workbook, configure_output_row)

def add_file_dialog(page: QWizardPage, file_mode: QFileDialog.FileMode, panel: IOControls, wb_manager: WorkbookManager) -> None:
    '''Add files to either the input or output selection tables.'''
    file_dialog = setup_file_dialog(page, file_mode)
    if file_dialog.exec():
        selected_files = file_dialog.selectedFiles()
        handle_file_selection(panel, selected_files, page, wb_manager)

def remove_selected_file(page: QWizardPage, panel: IOControls, wb_manager: WorkbookManager) -> None:
    '''Remove the currently selected file(s) from the table.'''
    # Selected rows
    selected_items = panel.table.selectionModel().selectedRows()
    # Search for selected rows in reverse so that the row numbering
    # doesn't change with each deletion.
    for row in sorted([item.row() for item in selected_items], reverse=True):
        file_name = panel.table.item(row, InputTableHeaders.FILENAME).text()
        # Remove the workbook
        wb_manager.remove_workbook(file_name)
        # Remove selected rows
        panel.table.removeRow(row)
        # Remove any associated error
        panel.error_manager.remove_error(file_name, panel.role)
        # Update UI
        page.completeChanged.emit()

##############################
### Input Table Population ###
##############################

def check_file_validity(file_name: str, panel: IOControls, file_item: QTableWidgetItem) -> None:
    '''Check if the file is valid and highlight errors if necessary.'''
    file_exists = futils.is_file_already_in_table(file_name, InputTableHeaders.FILENAME, panel.table, self.output_table)
    if panel.table == IORole.INPUT_FILE and panel.table.rowCount() >= 2:
        highlight_bad_item(file_item)
        raise TooManyOutputFilesSelected()
    if file_exists:
        highlight_bad_item(file_item)
        raise FileAlreadySelected(file_name)

def insert_file_row(panel: IOControls, file_name: str) -> int:
    '''Insert a new row for the file and return the row position.'''
    row_position = panel.table.rowCount()
    panel.table.insertRow(row_position)
    file_item = QTableWidgetItem(file_name)
    file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    panel.table.setItem(row_position, InputTableHeaders.FILENAME, file_item)
    return row_position

def populate_table(page: QWizardPage, panel: IOControls, files: list[str], add_workbook: AddWorkbook, configure_row: ConfigureRow) -> None:
    '''Populate the table with selected file(s).'''
    for file_path in files:
        file_name = path.basename(file_path)
        try:
            file_item = QTableWidgetItem(file_name)
            check_file_validity(file_name, panel, file_item)
            row_position = insert_file_row(panel, file_name)
            add_workbook(file_path)
            configure_row(panel.table, row_position, file_name, file_path)
        except Exception as e:
            error_manager.add_error(file_name, panel.role, e)
        page.completeChanged.emit()

def set_country_item(table: QTableWidget, row_position: int, country: str) -> None:
    '''Set the country item in the table.'''
    country_item = QTableWidgetItem(country)
    country_item.setFlags(country_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    table.setItem(row_position, InputTableHeaders.COUNTRY, country_item)

def set_worksheet_item(table: QTableWidget, row_position: int, sheet_name: str) -> None:
    '''Set the worksheet item in the table.'''
    worksheet_item = QTableWidgetItem(sheet_name)
    worksheet_item.setFlags(worksheet_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    table.setItem(row_position, InputTableHeaders.WORKSHEET, worksheet_item)

def configure_input_row(page: QWizardPage, table: QTableWidget, row_position: int, file_name: str, file_path: str) -> None:
    '''Configure a row for the input table.'''
    workbook_entry = self.managed_workbooks.get_workbook_by_name(file_name)
    try:
        update_country_details_in_table(workbook_entry)
        set_country_item(table, row_position, workbook_entry.locale_data.country)
        workbook_entry.init_input_worksheet()
        set_worksheet_item(table, row_position, workbook_entry.managed_sheet_object.selected_sheet.sheet_name)
    except CountryIdentifiersNotInFilename as e:
        item = table.item(row_position, InputTableHeaders.FILENAME)
        self.highlight_bad_item(item)
        raise e
    page.completeChanged.emit()

##############################
### Output Table Population ###
##############################

def create_dropdown(items: list[str], default: str) -> QComboBox:
    '''Create a dropdown with the given items and default selection.'''
    dropdown = QComboBox()
    dropdown.addItems(items)
    dropdown.setCurrentText(default)
    return dropdown

def setup_dropdowns(table: QTableWidget, row_position: int, default_year: str, default_month: str) -> tuple[QComboBox, QComboBox, QComboBox]:
    '''Set up year, month, and worksheet dropdowns.'''
    year_dropdown = create_dropdown([str(year) for year in range(2020, datetime.now().year + 3)], default_year)
    month_dropdown = create_dropdown(list(MONATE_KURZ_DE.keys()), default_month)
    table.setCellWidget(row_position, OutputTableHeaders.YEAR.value, year_dropdown)
    table.setCellWidget(row_position, OutputTableHeaders.MONTH.value, month_dropdown)
    return year_dropdown, month_dropdown

def configure_output_row(self, table: QTableWidget, row_position: int, file_name: str, file_path: str) -> None:
    '''Configure a row for the output table.'''
    workbook_entry_wf = self.managed_workbooks.get_workbook_by_name(file_name)
    workbook_entry_wf.init_output_worksheet()
    worksheet_dropdown = create_dropdown(
        [SpecialStrings.SELECT_WORKSHEET.value] + workbook_entry_wf.managed_sheet_object.sheet_names,
        SpecialStrings.SELECT_WORKSHEET.value
    )
    table.setCellWidget(row_position, OutputTableHeaders.WORKSHEET.value, worksheet_dropdown)
    default_year = str((datetime.now() + relativedelta(months=-1)).year)
    default_month = futils.german_abbr_month((datetime.now() + relativedelta(months=-1)).month, MONATE_KURZ_DE)
    year_dropdown, month_dropdown = setup_dropdowns(table, row_position, default_year, default_month)

    def handle_selections() -> None:
        try:
            file_role = self.check_file_role(table)
            self.error_manager.remove_error(file_name, file_role)
            self.remove_highlighting(table.item(row_position, InputTableHeaders.FILENAME))
            selected_sheet = worksheet_dropdown.currentText()
            selected_month = month_dropdown.currentText()
            selected_year = year_dropdown.currentText()
            workbook_entry_wf.managed_sheet_object.set_selected_sheet(
                sheet_name=selected_sheet,
                sheet_object=workbook_entry_wf.workbook_object[selected_sheet]
            )
            workbook_entry_wf.managed_sheet_object.set_budgeting_date(
                file_path, selected_sheet, selected_month, selected_year
            )
        except (IncorrectWorksheetSelected, BudgetingDatesNotFound, MissingEmployeeRow, KeyError) as e:
            self.highlight_bad_item(table.item(row_position, InputTableHeaders.FILENAME))
            self.error_manager.add_error(file_name, file_role, e)
        page.completeChanged.emit()

    worksheet_dropdown.currentIndexChanged.connect(handle_selections)
    year_dropdown.currentIndexChanged.connect(handle_selections)
    month_dropdown.currentIndexChanged.connect(handle_selections)

##############################
### Highlighting Functions ###
##############################

def highlight_bad_item(item: QTableWidgetItem) -> None:
    '''Highlight table items causing errors.'''

    item.setBackground(Qt.GlobalColor.red)
    item.setForeground(Qt.GlobalColor.white)

def remove_highlighting(item: QTableWidgetItem) -> None:
    '''Removes the highlighting after error correction.'''

    item.setBackground(Qt.GlobalColor.white)
    item.setForeground(Qt.GlobalColor.black)

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