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
from typing import Optional, TYPE_CHECKING, Callable
from PyQt6.QtWidgets import (
    QWidget, QTableWidget, QComboBox, QLabel, QWizardPage, QTableWidgetItem
)
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.location_management as loc
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.templating.types as t
import phb_app.utils.project_utils as pro

if TYPE_CHECKING:
    import phb_app.data.workbook_management as wm

#           --- DATA CONTAINERS ---

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
    buttons: t.ButtonsList
    error_panel: Optional[QWidget] = None

@dataclass
class Dropdowns:
    '''Data class for managing the dropdowns.'''
    year: QComboBox
    month: QComboBox
    worksheet: QComboBox
    current_text: SelectedText = field(default_factory=SelectedText)

@dataclass
class InputTableItems:
    '''Data class for managing the input table items.'''
    file_name: Optional[QTableWidgetItem] = None
    country: Optional[QTableWidgetItem] = None
    sheet_name: Optional[QTableWidgetItem] = None

@dataclass
class ProjectTableItems:
    '''Data class for managing the project table items.'''
    project_id: Optional[QTableWidgetItem] = None
    project_identifiers: Optional[QTableWidgetItem] = None
    file_name: Optional[QTableWidgetItem] = None

@dataclass
class IOFileContext:
    '''Data class for managing the data in the table.'''
    file_name: Optional[str] = None
    country_data: Optional[loc.CountryData] = None
    table_items: Optional[InputTableItems] = field(default_factory=InputTableItems)

@dataclass
class ProjectTableContext:
    '''Data class for managing the project table.'''
    project_id: Optional[t.ProjectId] = None
    project_identifiers: Optional[t.ProjectsTup] = None
    file_name: Optional[str] = None
    table_items: Optional[ProjectTableItems] = field(default_factory=ProjectTableItems)

type FileHandlerData = IOFileContext | ProjectTableContext

@dataclass
class EntryContext:
    '''Data class for managing the file dialog.'''
    panel: IOControls
    data: Optional[FileHandlerData] = None
    configure_row: Callable[[int, Optional["wm.WorkbookManager"]], None] = None

#           --- SERVICE CLASS ---

class EntryHandler:
    """Class for handling entries per panel."""
    def __init__(self, ent_ctx: EntryContext) -> None:
        self.ent_ctx = ent_ctx
        self.set_row_configurator()

    def set_row_configurator(self) -> None:
        '''Initialise the entry handler.'''
        self.ent_ctx.configure_row = {
            st.IORole.INPUTS:           self._configure_input_file_row,
            st.IORole.OUTPUT:           self._configure_output_file_row,
            st.IORole.PROJECT_TABLE:    self._configure_project_row
            }.get(self.ent_ctx.panel.role)

    def _configure_input_file_row(self, row: int, wb_mngr: "wm.WorkbookManager") -> None:
        '''Configure the input row in the table.'''
        wb_ctx = wb_mngr.get_workbook_ctx_by_file_name_and_role(self.ent_ctx.panel.role, self.ent_ctx.data.file_name)
        import phb_app.data.workbook_management as wm # pylint: disable=import-outside-toplevel
        wm.init_input_worksheet(wb_ctx)
        pu.update_handlers_country_details(self.ent_ctx.data.country_data, wb_ctx)
        self.ent_ctx.data.table_items.country = QTableWidgetItem(wb_ctx.locale_data.country)
        pu.insert_data_widget(self.ent_ctx.panel.table, self.ent_ctx.data.table_items.country, row, ie.InputTableHeaders.COUNTRY)
        self.ent_ctx.data.table_items.sheet_name = QTableWidgetItem(wb_ctx.managed_sheet.selected_sheet.sheet_name)
        pu.insert_data_widget(self.ent_ctx.panel.table, self.ent_ctx.data.table_items.sheet_name, row, ie.InputTableHeaders.WORKSHEET)

    def _configure_output_file_row(self, row: int, wb_mngr: "wm.WorkbookManager") -> None:
        '''Configure the output row in the table.'''
        wb_ctx = wb_mngr.get_workbook_ctx_by_file_name_and_role(self.ent_ctx.panel.role, self.ent_ctx.data.file_name)
        import phb_app.data.workbook_management as wm # pylint: disable=import-outside-toplevel
        wm.init_output_worksheet(wb_ctx)
        dropdowns = Dropdowns(pu.create_year_dropdown(), pu.create_month_dropdown(), pu.create_worksheet_dropdown(wb_ctx))
        pu.setup_dropdowns(self.ent_ctx.panel.table, row, dropdowns)
        def connection_wrapper() -> None:
            '''Connect functionality to the dropdowns.'''
            pu.handle_dropdown_selection(self.ent_ctx, wb_ctx, row, dropdowns)
            self.ent_ctx.panel.page.completeChanged.emit()
        connect_dropdowns(dropdowns, connection_wrapper)

    def _configure_project_row(self, row: int, wb_mngr: "wm.WorkbookManager" = None) -> None:
        '''Configure the project row in the table.'''
        # Only input workbooks have project IDs, so we can safely assume the role is INPUTS
        self.ent_ctx.data.table_items.project_id = QTableWidgetItem(self.ent_ctx.data.project_id)
        pu.insert_data_widget(self.ent_ctx.panel.table, self.ent_ctx.data.table_items.project_id, row, ie.ProjectIDTableHeaders.PROJECT_ID)
        self.ent_ctx.data.table_items.project_identifiers = QTableWidgetItem(self.ent_ctx.data.project_identifiers)
        pu.insert_data_widget(self.ent_ctx.panel.table, self.ent_ctx.data.table_items.project_identifiers, row, ie.ProjectIDTableHeaders.DESCRIPTION)
        self.ent_ctx.data.table_items.file_name = QTableWidgetItem(self.ent_ctx.data.file_name)
        pu.insert_data_widget(self.ent_ctx.panel.table, self.ent_ctx.data.table_items.file_name, row, ie.ProjectIDTableHeaders.FILENAME)

#           --- MODULE SERVICE FUNCTIONS ---

def join_project_identifiers(id_tup: t.ProjectsTup) -> str:
    '''Public module level. Set the project identifiers.'''
    return "\n".join(proj_item for proj_item in id_tup[ie.ProjectIDTableHeaders.DESCRIPTION])

def update_current_text(dd: Dropdowns) -> None:
    '''Public module level. Update the current text of the dropdown QComboboxes.'''
    dd.current_text.year = dd.year.currentText()
    dd.current_text.month = dd.month.currentText()
    dd.current_text.worksheet = dd.worksheet.currentText()

def connect_dropdowns(dd: Dropdowns, func: Callable) -> None:
    '''Public module level. Connect the dropdowns to the given function.'''
    for dropdown in (dd.year, dd.month, dd.worksheet):
        dropdown.currentTextChanged.connect(func)
