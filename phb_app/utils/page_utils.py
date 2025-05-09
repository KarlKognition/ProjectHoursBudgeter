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
from datetime import datetime
from typing import Optional, TYPE_CHECKING
from dateutil.relativedelta import relativedelta
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import (
    QWizardPage, QBoxLayout, QHBoxLayout, QComboBox,
    QWidget, QLabel, QFileDialog, QVBoxLayout,
    QTableWidget, QTableWidgetItem
)
# First party libraries
import phb_app.utils.date_utils as du
import phb_app.utils.employee_utils as eu
import phb_app.utils.file_handling_utils as fu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.data.header_management as hm
import phb_app.logging.exceptions as ex
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.location_management as loc
import phb_app.data.months_dict as md
import phb_app.logging.error_manager as em

if TYPE_CHECKING:
    import phb_app.data.io_management as io
    import phb_app.data.workbook_management as wm

###################
### Title Setup ###
###################

def set_titles(page: QWizardPage, title: str, subtitle: str) -> None:
    '''Set the title and subtitle for the page.'''
    page.setTitle(title)
    page.setSubTitle(subtitle)

def _set_watermark(watermark: QLabel, directory: str, error: str) -> None:
    '''Set watermark.'''
    watermark_file = f"{directory}\\budget_watermark.jpg"
    watermark_pixmap = QPixmap(watermark_file)
    if watermark_pixmap.isNull():
        watermark.setStyleSheet("color: red;")
        watermark.setPixmap(QPixmap())  # Clear the pixmap
        watermark.setText(error)
        return
    watermark.setPixmap(watermark_pixmap)
    watermark.setScaledContents(False)

def create_watermark_label(directory: str, error: str) -> QLabel:
    '''Create and return the watermark label.'''
    watermark_label = QLabel()
    _set_watermark(watermark_label, directory, error)
    return watermark_label

def create_intro_message(text: str) -> QLabel:
    '''Create and return the introduction message label.'''
    intro_message = QLabel(text)
    intro_message.setWordWrap(True)
    intro_message.setAlignment(Qt.AlignmentFlag.AlignLeft)
    return intro_message

def setup_page(page: QWizardPage, widgets: list[QWidget], layout_type: QBoxLayout, spacing: int = 6, margins: tuple[int, int, int, int] = (9, 9, 9, 9)) -> None:
    '''Set up the layout for a page with the given widgets and layout type.
    The default spacing and margins from Qt are set.'''
    layout_type.setSpacing(spacing)
    layout_type.setContentsMargins(*margins)
    for widget in widgets:
        layout_type.addWidget(widget)
    page.setLayout(layout_type)

#############################
### Table and Panel Setup ###
#############################

def setup_error_panel(error_manager: em.ErrorManager, role: st.IORole) -> QWidget:
    '''Set up the error panel for the given role with vertical layout.'''
    error_manager.error_panels[role] = QWidget()
    error_manager.error_panels[role].setLayout(QVBoxLayout())
    return error_manager.error_panels[role]

def create_table(table_headers: ie.BaseTableHeaders, selection_mode: QTableWidget.SelectionMode, col_widths: hm.ColWidths) -> QTableWidget:
    '''Create a table widget with the given headers and selection mode.'''
    table = QTableWidget(0, len(table_headers))
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(selection_mode)
    table.setHorizontalHeaderLabels(table_headers.cap_members_list())
    table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
    for header, width in col_widths.items():
        table.setColumnWidth(header, width)
    return table

def create_interaction_panel(panel: "io.IOControls") -> QWidget:
    '''Set up the table with the given buttons and column widths.'''
    container = QWidget()
    layout = QVBoxLayout()
    layout.addWidget(panel.label)
    layout.addWidget(panel.table)
    buttons_layout = QHBoxLayout()
    # Add all assigned buttons
    for button in panel.buttons:
        buttons_layout.addWidget(button)
    layout.addLayout(buttons_layout)
    layout.addWidget(panel.error_panel)
    container.setLayout(layout)
    return container

#####################
### File Handling ###
#####################

def _setup_file_dialog(page: QWizardPage, file_mode: QFileDialog.FileMode) -> QFileDialog:
    '''Set up and return a file dialog.'''
    file_dialog = QFileDialog(page)
    file_dialog.setWindowTitle(st.ADD_FILE)
    file_dialog.setNameFilter(st.EXCEL_FILE)
    file_dialog.setFileMode(file_mode)
    file_dialog.setViewMode(QFileDialog.ViewMode.Detail)
    return file_dialog

def _add_file_dialog(page: QWizardPage, file_mode: QFileDialog.FileMode, file_handler: "io.FileDialogHandler") -> None:
    '''Add files to either the input or output selection tables.'''
    file_dialog = _setup_file_dialog(page, file_mode)
    if file_dialog.exec():
        _populate_table(page, file_dialog.selectedFiles(), file_handler)

def _remove_selected_file(page: QWizardPage, file_handler: "io.FileDialogHandler") -> None:
    '''Remove the currently selected file(s) from the table.'''
    # Selected rows
    selected_items = file_handler.panel.table.selectionModel().selectedRows()
    # Search for selected rows in reverse so that the row numbering
    # doesn't change with each deletion.
    for row in sorted([item.row() for item in selected_items], reverse=True):
        file_name = file_handler.panel.table.item(row, ie.InputTableHeaders.FILENAME).text()
        # Remove the workbook
        file_handler.workbook_manager.remove_workbook(file_name)
        # Remove selected rows
        file_handler.panel.table.removeRow(row)
        # Remove any associated error
        file_handler.error_manager.remove_error(file_name, file_handler.panel.role)
        # Update UI
        page.completeChanged.emit()

###############
### Buttons ###
###############

def connect_buttons(page: QWizardPage, file_handler: "io.FileDialogHandler") -> None:
    '''Connect buttons to their respective actions dynamically.'''
    action_dispatch = {
        st.ButtonNames.ADD: lambda: _add_file_dialog(page, QFileDialog.FileMode.ExistingFiles, file_handler),
        st.ButtonNames.REMOVE: lambda: _remove_selected_file(page, file_handler),
        st.ButtonNames.SELECT_ALL: file_handler.panel.table.selectAll,
        st.ButtonNames.DESELECT_ALL: file_handler.panel.table.clearSelection
    }
    # Iterate over buttons and connect them dynamically
    for button in file_handler.panel.buttons:
        if action := action_dispatch.get(button.text()):
            button.clicked.connect(action)

##############################
### Input Table Population ###
##############################

def _check_file_validity(file_name: str, panel: "io.IOControls", file_item: QTableWidgetItem) -> None:
    '''Check if the file is valid and highlight errors if necessary.'''
    file_exists = fu.is_file_already_in_table(file_name, ie.InputTableHeaders.FILENAME, panel.table)
    if panel.table == st.IORole.OUTPUT_FILE and panel.table.rowCount() >= 2:
        highlight_bad_item(file_item)
        raise ex.TooManyOutputFilesSelected()
    if file_exists:
        highlight_bad_item(file_item)
        raise ex.FileAlreadySelected(file_name)

def _insert_file_row(panel: "io.IOControls", file_name: str) -> int:
    '''Insert a new row for the file and return the row position.'''
    row_position = panel.table.rowCount()
    panel.table.insertRow(row_position)
    file_item = QTableWidgetItem(file_name)
    file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    panel.table.setItem(row_position, ie.InputTableHeaders.FILENAME, file_item)
    return row_position

def set_country_item(table: QTableWidget, row_position: int, country: str) -> None:
    '''Set the country item in the table.'''
    country_item = QTableWidgetItem(country)
    country_item.setFlags(country_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    table.setItem(row_position, ie.InputTableHeaders.COUNTRY, country_item)

def set_worksheet_item(table: QTableWidget, row_position: int, sheet_name: str) -> None:
    '''Set the worksheet item in the table.'''
    worksheet_item = QTableWidgetItem(sheet_name)
    worksheet_item.setFlags(worksheet_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    table.setItem(row_position, ie.InputTableHeaders.WORKSHEET, worksheet_item)

def update_country_details_in_table(data: loc.CountryData, workbook_entry: "wm.ManagedInputWorkbook") -> None:
    '''
    Update country details.
    '''
    # Get country name from file name
    country_name = fu.get_origin_from_file_name(workbook_entry.file_name, data, st.CountriesEnum)
    # Set the local data of the workbook
    workbook_entry.set_locale_data(data, country_name)

###############################
### Output Table Population ###
###############################

def _create_dropdown(items: list[str], default: str) -> QComboBox:
    '''Create a dropdown with the given items and default selection.'''
    dropdown = QComboBox()
    dropdown.addItem(default)
    dropdown.addItems(items)
    dropdown.setCurrentIndex(0)
    return dropdown

def get_initialised_managed_workbook(file_handler: "io.FileDialogHandler") -> "wm.ManagedOutputWorkbook":
    '''Initialise and return the output worksheet.'''
    workbook_entry = file_handler.workbook_manager.get_workbook_by_name(file_handler.file_name)
    workbook_entry.init_output_worksheet()
    return workbook_entry

def create_worksheet_dropdown(workbook_entry: "wm.ManagedOutputWorkbook") -> QComboBox:
    '''Create a dropdown for selecting a worksheet.'''
    return _create_dropdown(workbook_entry.managed_sheet_object.sheet_names, st.SpecialStrings.SELECT_WORKSHEET)

def create_year_dropdown() -> QComboBox:
    '''
    Create a dropdown for selecting a year. The default year is the current year unless it is December,
    then due to the default month being last month, the default year is set to the previous year.
    '''
    default_year = str((datetime.now() + relativedelta(months=-1)).year)
    return _create_dropdown([str(year) for year in range(2020, datetime.now().year + 3)], default_year)

def create_month_dropdown() -> QComboBox:
    '''
    Create a dropdown for selecting a month. The default month is set to the previous month.
    This is the month in which hours are yet to be analysed.
    '''
    default_month = du.german_abbr_month((datetime.now() + relativedelta(months=-1)).month, md.LOCALIZED_MONTHS_SHORT)
    return _create_dropdown(list(md.LOCALIZED_MONTHS_SHORT.keys()), default_month)

def setup_dropdowns(table:QTableWidget, row_position: int, dropdowns: "io.DropdownHandler") -> None:
    '''Set up year, month, and worksheet dropdowns.'''
    table.setCellWidget(row_position, ie.OutputTableHeaders.WORKSHEET, dropdowns.worksheet_dropdown)
    table.setCellWidget(row_position, ie.OutputTableHeaders.MONTH, dropdowns.month_dropdown)
    table.setCellWidget(row_position, ie.OutputTableHeaders.YEAR, dropdowns.year_dropdown)

def handle_selection_error(row_position: int, file_handler: "io.FileDialogHandler", error: Exception) -> None:
    '''Handle errors during selection and update the UI accordingly.'''
    col: dict[st.IORole, ie.InputTableHeaders] = {
        st.IORole.INPUT_FILE: ie.InputTableHeaders.FILENAME,
        st.IORole.OUTPUT_FILE: ie.OutputTableHeaders.FILENAME
    }
    highlight_bad_item(file_handler.panel.table.item(row_position, col[file_handler.panel.role]))
    file_handler.error_manager.add_error(file_handler.file_name, file_handler.panel.role, error)

def handle_dropdown_selection(file_handler: "io.FileDialogHandler", row_position: int, dropdowns: "io.DropdownHandler") -> None:
    '''Handle dropdown selection and update the selected sheet accordingly.'''
    # Remove the error if it is still there (if retrying)
    clear_row_error_status(file_handler, row_position, ie.OutputTableHeaders.FILENAME)
    dropdowns.update_current_text()
    try:
        eu.set_selected_sheet(file_handler, dropdowns.current_text.worksheet)
        du.set_budgeting_date(file_handler, dropdowns.current_text)
    except (ex.EmployeeRowAnchorsMisalignment, ex.MissingEmployeeRow, ex.BudgetingDatesNotFound,
            KeyError) as exc:
        handle_selection_error(row_position, file_handler, exc)

########################
### Table Population ###
########################

def _populate_table(page: QWizardPage, files: list[str], file_handler: "io.FileDialogHandler") -> None:
    '''Populate the table with selected file(s).'''
    for file_path in files:
        file_handler.set_file_path_and_name(file_path)
        try:
            file_item = QTableWidgetItem(file_handler.file_name)
            _check_file_validity(file_handler.file_name, file_handler.panel, file_item)
            row_position = _insert_file_row(file_handler.panel, file_handler.file_name)
            file_handler.workbook_manager.add_workbook(file_path, file_handler.panel.role)
            file_handler.configure_row(row_position)
        except (ex.FileAlreadySelected, ex.TooManyOutputFilesSelected, ex.CountryIdentifiersNotInFilename,
                ex.IncorrectWorksheetSelected, ex.BudgetingDatesNotFound, ex.WorkbookAlreadyTracked,
                ex.WorkbookLoadError, KeyError) as exc:
            handle_selection_error(row_position, file_handler, exc)
        page.completeChanged.emit()

######################################
### Exception Styling and Handling ###
######################################

def highlight_bad_item(item: QTableWidgetItem) -> None:
    '''Highlight table items causing errors.'''

    item.setBackground(Qt.GlobalColor.red)
    item.setForeground(Qt.GlobalColor.white)

def remove_highlighting(item: QTableWidgetItem) -> None:
    '''Removes the highlighting after error correction.'''

    item.setBackground(Qt.GlobalColor.white)
    item.setForeground(Qt.GlobalColor.black)

def clear_row_error_status(file_handler: "io.FileDialogHandler", row_position: int, header: ie.BaseTableHeaders) -> None:
    '''Clear the error status of a row in the table.'''
    file_handler.error_manager.remove_error(file_handler.file_name, file_handler.panel.role)
    remove_highlighting(file_handler.panel.table.item(row_position, header))

##################################
### QWizard function overrides ###
##################################

def get_combo_box(panel: "io.IOControls", row_position: int, col: int) -> Optional[QComboBox]:
    '''Get the combo box from the table.'''
    return panel.table.cellWidget(row_position, col) if panel.table.cellWidget(row_position, col) else None

def check_completion(panels: tuple["io.IOControls"]) -> bool:
    '''Return the appropriate isComplete function based on the page type.'''
    in_panel, out_panel = (panel for panel in panels if panel.role in {st.IORole.INPUT_FILE, st.IORole.OUTPUT_FILE})
    def io_selection_complete() -> bool:
        '''Check if both tables have at least one row selected
        and no error messages are displayed. Errors are added as QLabel widgets. No QLable, no error.'''
        output_item = get_combo_box(out_panel, ie.OutputFile.FIRST_ENTRY, ie.OutputTableHeaders.WORKSHEET)
        return (
            all(panel.table.rowCount() >= 1 for panel in (in_panel, out_panel))
            and output_item is not None
            and output_item.currentText() != st.SpecialStrings.SELECT_WORKSHEET
            and all(not panel.error_panel.findChildren(QLabel) for panel in (in_panel, out_panel))
        )
    def table_selection_complete() -> bool:
        '''Check if the table has at least one row selected. There is only one panel in this case.'''
        return bool(panels[0].table.selectionModel().selectedRows())
    if in_panel or out_panel:
        return io_selection_complete()
    return table_selection_complete()