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
    QWidget, QTableWidget, QComboBox, QLabel,
    QPushButton, QWizardPage
)
import phb_app.protocols_callables.protocols as protocols
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.location_management as loc
import phb_app.logging.error_manager as em
import phb_app.utils.page_utils as pu

if TYPE_CHECKING:
    import phb_app.data.workbook_management as wm

@dataclass
class SelectedText:
    '''Data class for managing the selected text in the dropdowns.'''
    year: Optional[str] = None
    month: Optional[str] = None
    worksheet: Optional[str] = None

type ButtonsList = list[QPushButton]

@dataclass
class IOControls:
    '''Data class for managing the IO controls.'''
    page: QWizardPage
    role: st.IORole
    label: QLabel
    table: QTableWidget
    buttons: ButtonsList
    error_panel: QWidget

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
        self.year_dropdown.currentTextChanged.connect(func)
        self.month_dropdown.currentTextChanged.connect(func)
        self.worksheet_dropdown.currentTextChanged.connect(func)

@dataclass
class FileDialogHandler:
    '''Data class for managing the file dialog.'''
    panel: IOControls
    workbook_manager: "wm.WorkbookManager"
    error_manager: Optional[em.ErrorManager]
    data: Optional[loc.CountryData] = None
    workbook_entry: Optional[Union["wm.ManagedInputWorkbook", "wm.ManagedOutputWorkbook"]] = None
    file_path: Optional[str] = None
    file_name: Optional[str] = None
    configure_row: Optional[protocols.ConfigureRow] = None

    def __post_init__(self) -> None:
        configure_row_dispatch = {
            st.IORole.INPUT_FILE: self.configure_input_row,
            st.IORole.OUTPUT_FILE: self.configure_output_row
        }
        self.configure_row = configure_row_dispatch.get(self.panel.role)

    def set_file_path_and_name(self, file_path: str) -> None:
        '''Set the file path and name.'''
        self.file_path = file_path
        self.file_name = path.basename(file_path)

    def configure_input_row(self, row_position: int) -> None:
        '''Configure the input row in the table.'''
        self.workbook_entry = self.workbook_manager.get_workbook_by_name(self.file_name)
        pu.update_country_details_in_table(self.data, self.workbook_entry)
        pu.set_country_item(self.panel.table, row_position, self.workbook_entry.locale_data.country)
        self.workbook_entry.init_input_worksheet()
        pu.set_worksheet_item(self.panel.table, row_position, self.workbook_entry.managed_sheet_object.selected_sheet.sheet_name)

    def configure_output_row(self, row_position: int) -> None:
        '''Configure the output row in the table.'''
        self.workbook_entry = pu.get_initialised_managed_workbook(self)
        dropdowns = DropdownHandler(pu.create_year_dropdown(), pu.create_month_dropdown(), pu.create_worksheet_dropdown(self.workbook_entry))
        pu.setup_dropdowns(self.panel.table, row_position, dropdowns)
        def connection_wrapper() -> None:
            '''Connect functionality to the dropdowns.'''
            pu.handle_dropdown_selection(self, row_position, dropdowns)
            self.panel.page.completeChanged.emit()
        dropdowns.connect_dropdowns(connection_wrapper)
