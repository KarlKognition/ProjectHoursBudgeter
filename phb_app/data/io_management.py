'''
Module Name
---------
PHB Wizard IO Management

Version
-------
Date-based Version: 20250428
Author: Karl Goran Antony Zuvela

Description
-----------
PHB Wizard text selection management.
'''

from dataclasses import dataclass, field
from typing import Optional, Callable, TYPE_CHECKING, Union
from os import path
from PyQt6.QtWidgets import (
    QWidget, QTableWidget, QComboBox, QLabel, QWizardPage, QTableWidgetItem
)
import phb_app.templating.protocols as protocols
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.location_management as loc
import phb_app.logging.error_manager as em
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.templating.types as types

if TYPE_CHECKING:
    import phb_app.data.workbook_management as wm

@dataclass
class SelectedText:
    '''Data class for managing the selected text in the dropdowns.'''
    year: Optional[str] = None
    month: Optional[str] = None
    worksheet: Optional[str] = None

@dataclass
class IOControls:
    '''Data class for managing the IO controls.'''
    page: QWizardPage
    role: st.IORole
    label: QLabel
    table: QTableWidget
    buttons: types.ButtonsList
    error_panel: Optional[QWidget] = None

@dataclass
class DropdownHandler:
    '''Data class for managing the dropdowns.'''
    year_dropdown: QComboBox
    month_dropdown: QComboBox
    worksheet_dropdown: QComboBox
    current_text: SelectedText = field(default_factory=SelectedText)

    def __post_init__(self) -> None:
        self.update_current_text()
    
    def update_current_text(self) -> None:
        '''Update the current text of the dropdowns.'''
        self.current_text.year = self.year_dropdown.currentText()
        self.current_text.month = self.month_dropdown.currentText()
        self.current_text.worksheet = self.worksheet_dropdown.currentText()

    def connect_dropdowns(self, func: Callable) -> None:
        '''Connect the dropdowns to the given function.'''
        for dropdown in (self.year_dropdown, self.month_dropdown, self.worksheet_dropdown):
            dropdown.currentTextChanged.connect(func)

@dataclass
class InputTableItems:
    '''Data class for managing the input table items.'''
    file_name: Optional[QTableWidgetItem] = None
    country: Optional[QTableWidgetItem] = None
    sheet_name: Optional[QTableWidgetItem] = None

@dataclass
class InputFileHandler:
    '''Data class for managing the data in the table.'''
    file_path: Optional[str] = None
    file_name: Optional[str] = None
    country_data: Optional[loc.CountryData] = None
    project_identifiers : Optional[types.ProjectsTup] = None
    table_items: Optional[InputTableItems] = field(default_factory=InputTableItems)

    def set_table_country(self, country: str) -> None:
        '''Set the table item for the given file name and country.'''
        self.table_items.country = QTableWidgetItem(country)

    def set_table_sheet_name(self, sheet_name: str) -> None:
        '''Set the table item for the given sheet name.'''
        self.table_items.sheet_name = QTableWidgetItem(sheet_name)

@dataclass
class EntryHandler:
    '''Data class for managing the file dialog.'''
    panel: IOControls
    workbook_manager: "wm.WorkbookManager"
    error_manager: Optional[em.ErrorManager] = None
    data: InputFileHandler = field(default_factory=InputFileHandler)
    workbook_entry: Optional[Union["wm.ManagedInputWorkbook", "wm.ManagedOutputWorkbook"]] = None
    configure_row: Optional[protocols.ConfigureRow] = None

    def __post_init__(self) -> None:
        configure_row_dispatch = {
            st.IORole.INPUTS: self.configure_input_file_row,
            st.IORole.OUTPUT: self.configure_output_file_row,
            st.IORole.PROJECT_TABLE: self.configure_project_row
        }
        self.configure_row = configure_row_dispatch.get(self.panel.role)

    def set_file_path_and_name(self, file_path: str) -> None:
        '''Set the file path and name.'''
        self.data.file_path = file_path
        self.data.file_name = path.basename(file_path)
        self.data.table_items.file_name = QTableWidgetItem(self.data.file_name)

    def configure_input_file_row(self, row_position: int) -> None:
        '''Configure the input row in the table.'''
        self.workbook_entry = self.workbook_manager.get_workbook_by_name(self.data.file_name)
        pu.update_handlers_country_details(self.data.country_data, self.workbook_entry)
        self.data.set_table_country(self.workbook_entry.locale_data.country)
        pu.insert_row_data(self.panel.table, self.data.table_items.country, row_position, ie.InputTableHeaders.COUNTRY)
        self.workbook_entry.init_input_worksheet()
        self.data.set_table_sheet_name(self.workbook_entry.managed_sheet_object.selected_sheet.sheet_name)
        pu.insert_row_data(self.panel.table, self.data.table_items.sheet_name, row_position, ie.InputTableHeaders.WORKSHEET)

    def configure_output_file_row(self, row_position: int) -> None:
        '''Configure the output row in the table.'''
        self.workbook_entry = pu.get_initialised_managed_output_workbook(self)
        dropdowns = DropdownHandler(pu.create_year_dropdown(), pu.create_month_dropdown(), pu.create_worksheet_dropdown(self.workbook_entry))
        pu.setup_dropdowns(self.panel.table, row_position, dropdowns)
        def connection_wrapper() -> None:
            '''Connect functionality to the dropdowns.'''
            pu.handle_dropdown_selection(self, row_position, dropdowns)
            self.panel.page.completeChanged.emit()
        dropdowns.connect_dropdowns(connection_wrapper)

    def configure_project_row(self, row_position: int) -> None:
        '''Configure the project row in the table.'''
        
