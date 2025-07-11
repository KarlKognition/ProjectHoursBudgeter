'''
Module Name
---------
PHB Wizard IO Management

Author
-------
Karl Goran Antony Zuvela

Description
-----------
PHB Wizard text selection management.
'''
#           --- Standard libraries ---
from uuid import UUID
from dataclasses import dataclass, field
from typing import Optional, TYPE_CHECKING, Callable, Protocol, Union
#           --- Third party libraries ---
from PyQt6.QtCore import QModelIndex
from PyQt6.QtWidgets import QWidget, QTableWidget, QComboBox, QLabel, QWizardPage, QTableWidgetItem
#           --- First party libraries ---
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.employee_management as em
import phb_app.data.location_management as loc
import phb_app.data.worksheet_management as ws
import phb_app.utils.hours_utils as hu
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.templating.types as t

if TYPE_CHECKING:
    import phb_app.data.workbook_management as wm

#           --- PROTOCOLS ---

class ConfigRowWithEntCtx(Protocol):
    """Protocol for a configuration row that does not require a workbook manager."""
    def __call__(self, ent_ctx: "EntryContext", row: int) -> None: ...

class ConfigRowWithEntCtxWbCtx(Protocol):
    """Protocol for a configuration row that does not require a workbook manager."""
    def __call__(self, ent_ctx: "EntryContext", row: int, wb_ctx: Union["wm.InputWorkbookContext", "wm.OutputWorkbookContext"]) -> None: ...

class ConfigRowWithEntCtxEmpHrs(Protocol):
    """Protocol for a configuration row that requires an employee context."""
    def __call__(self, ent_ctx: "EntryContext", row: int, emp_hrs: em.EmployeeHours) -> None: ...

#           --- DATA CONTAINERS ---

@dataclass(slots=True)
class SelectedText:
    '''Data class for managing the selected text in the dropdowns.'''
    year: Optional[str] = None
    month: Optional[str] = None
    worksheet: Optional[str] = None

@dataclass(slots=True)
class IOControls:
    '''Data class for managing the IO controls.'''
    page: QWizardPage
    role: st.IORole
    label: QLabel
    table: QTableWidget
    buttons: t.ButtonsList
    error_panel: Optional[QWidget] = None

@dataclass(slots=True)
class Dropdowns:
    '''Data class for managing the dropdowns.'''
    year: QComboBox
    month: QComboBox
    worksheet: QComboBox
    current_text: SelectedText = field(default_factory=SelectedText)

@dataclass(slots=True)
class InputTableItems:
    '''Data class for managing the input table items.'''
    file_name: Optional[QTableWidgetItem] = None
    country: Optional[QTableWidgetItem] = None
    sheet_name: Optional[QTableWidgetItem] = None
    uuid: Optional[QTableWidgetItem] = None

@dataclass(slots=True)
class ProjectTableItems:
    '''Data class for managing the project table items.'''
    project_id: Optional[QTableWidgetItem] = None
    project_identifiers: Optional[QTableWidgetItem] = None
    file_name: Optional[QTableWidgetItem] = None

@dataclass(slots=True)
class EmployeeTableItems:
    '''Data class for managing the employee table items.'''
    employee: Optional[QTableWidgetItem] = None
    worksheet: Optional[QTableWidgetItem] = None
    coord: Optional[QTableWidgetItem] = None

@dataclass(slots=True)
class SummaryIOTableItems:
    '''Data class for managing the summary IO table items.'''
    in_file_names: Optional[QTableWidgetItem] = None
    out_file_names: Optional[QTableWidgetItem] = None
    date: Optional[QTableWidgetItem] = None

@dataclass(slots=True)
class SummaryDataTableItems:
    '''Data class for managing the summary data table items.'''
    emp_name: Optional[QTableWidgetItem] = None
    pred_hrs: Optional[QTableWidgetItem] = None
    acc_hrs: Optional[QTableWidgetItem] = None
    dev: Optional[QTableWidgetItem] = None
    proj_id: Optional[QTableWidgetItem] = None
    out_ws: Optional[QTableWidgetItem] = None
    coord: Optional[QTableWidgetItem] = None

@dataclass(slots=True)
class IOFileContext:
    '''Data class for managing the data in the table.'''
    file_name: Optional[str] = None
    country_data: Optional[loc.CountryData] = None
    uuid: Optional[UUID] = None
    table_items: Optional[InputTableItems] = field(default_factory=InputTableItems)

@dataclass(slots=True)
class ProjectTableContext:
    '''Data class for managing the project table.'''
    project_id: Optional[t.ProjectId] = None
    project_identifiers: Optional[t.ProjectsTup] = None
    file_name: Optional[str] = None
    table_items: Optional[ProjectTableItems] = field(default_factory=ProjectTableItems)

@dataclass(slots=True)
class EmployeeTableContext:
    '''Data class for managing the employee table.'''
    emp_name: Optional[str] = None
    worksheet: Optional[str] = None
    coord: Optional[str] = None
    table_items: Optional[EmployeeTableItems] = field(default_factory=EmployeeTableItems)

@dataclass(slots=True)
class SummaryIOContext:
    '''Data class for managing the summary table.'''
    in_file_names: Optional[str] = None
    out_file_names: Optional[str] = None
    date: Optional[str] = None
    table_items: Optional[SummaryIOTableItems] = field(default_factory=SummaryIOTableItems)

@dataclass(slots=True)
class SummaryDataContext:
    '''Data class for managing the summary table.'''
    emp_name: Optional[str] = None
    pred_hrs: Optional[float] = None
    acc_hrs: Optional[float] = None
    dev: Optional[str] = None
    proj_id: Optional[t.ProjectId] = None
    out_ws_name: Optional[str] = None
    coord: Optional[str] = None
    table_items: Optional[SummaryDataTableItems] = field(default_factory=SummaryDataTableItems)

type FileHandlerData = (
    IOFileContext | ProjectTableContext | EmployeeTableContext | SummaryIOContext | SummaryDataContext
)

@dataclass(slots=True)
class EntryContext:
    '''Data class for managing the file dialog.'''
    panel: IOControls
    data: Optional[FileHandlerData] = None
    configure_row: Optional[ConfigRowWithEntCtx | ConfigRowWithEntCtxEmpHrs] = None

#           --- MODULE SERVICE FUNCTIONS ---

def set_row_configurator(ent_ctx: EntryContext) -> None:
    '''Initialise the entry context row configurator.'''
    ent_ctx.configure_row = {
        st.IORole.INPUTS:               _configure_input_file_row,
        st.IORole.OUTPUT:               _configure_output_file_row,
        st.IORole.PROJECT_TABLE:        _configure_project_row,
        st.IORole.EMPLOYEE_TABLE:       _configure_employee_row,
        st.IORole.SUMMARY_IO_TABLE:     _configure_summary_io_row,
        st.IORole.SUMMARY_DATA_TABLE:   _configure_summary_data_row
    }.get(ent_ctx.panel.role)

def configure_error_row(
        ent_ctx: EntryContext,
        row: int
    ) -> None:
    '''Configure the error row in the table.'''
    ent_ctx.data.table_items.file_name = QTableWidgetItem(ent_ctx.data.file_name)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.file_name, row, ie.InputTableHeaders.FILENAME)
    ent_ctx.data.table_items.uuid = QTableWidgetItem(str(ent_ctx.data.uuid))
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.uuid, row, ie.InputTableHeaders.UNIQUE_ID)

def _configure_input_file_row(
        ent_ctx: EntryContext,
        row: int,
        wb_ctx: Union["wm.InputWorkbookContext", "wm.OutputWorkbookContext"]
    ) -> None:
    '''Configure the input row in the table.'''
    import phb_app.data.workbook_management as wm # pylint: disable=import-outside-toplevel
    ws.init_input_worksheet(wb_ctx)
    pu.update_handlers_country_details(ent_ctx.data.country_data, wb_ctx)
    ent_ctx.data.table_items.country = QTableWidgetItem(wb_ctx.locale_data.country)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.country, row, ie.InputTableHeaders.COUNTRY)
    ent_ctx.data.table_items.sheet_name = QTableWidgetItem(wb_ctx.managed_sheet.selected_sheet.sheet_name)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.sheet_name, row, ie.InputTableHeaders.WORKSHEET)
    ent_ctx.data.table_items.uuid = QTableWidgetItem(str(wb_ctx.mngd_wb.uuid))
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.uuid, row, ie.InputTableHeaders.UNIQUE_ID)

def _configure_output_file_row(
        ent_ctx: EntryContext,
        row: int,
        wb_ctx: Union["wm.InputWorkbookContext", "wm.OutputWorkbookContext"]
    ) -> None:
    '''Configure the output row in the table.'''
    import phb_app.data.workbook_management as wm # pylint: disable=import-outside-toplevel
    ws.init_output_worksheet(wb_ctx)
    dropdowns = Dropdowns(pu.create_year_dropdown(), pu.create_month_dropdown(), pu.create_worksheet_dropdown(wb_ctx))
    pu.setup_dropdowns(ent_ctx.panel.table, row, dropdowns)
    ent_ctx.data.table_items.uuid = QTableWidgetItem(str(wb_ctx.mngd_wb.uuid))
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.uuid, row, ie.InputTableHeaders.UNIQUE_ID)
    def connection_wrapper() -> None:
        '''Connect functionality to the dropdowns.'''
        pu.handle_dropdown_selection(ent_ctx, wb_ctx, row, dropdowns)
        ent_ctx.panel.page.completeChanged.emit()
    connect_dropdowns(dropdowns, connection_wrapper)

def _configure_project_row(
        ent_ctx: EntryContext,
        row: int
    ) -> None:
    '''Configure the project row in the table.'''
    # Only input workbooks have project IDs, so we can safely assume the role is INPUTS
    ent_ctx.data.table_items.project_id = QTableWidgetItem(ent_ctx.data.project_id)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.project_id, row, ie.ProjectIDTableHeaders.PROJECT_ID)
    ent_ctx.data.table_items.project_identifiers = QTableWidgetItem(ent_ctx.data.project_identifiers)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.project_identifiers, row, ie.ProjectIDTableHeaders.DESCRIPTION)
    ent_ctx.data.table_items.file_name = QTableWidgetItem(ent_ctx.data.file_name)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.file_name, row, ie.ProjectIDTableHeaders.FILENAME)

def _configure_employee_row(
        ent_ctx: EntryContext,
        row: int
    ) -> None:
    '''Configure the employee row in the table.'''
    ent_ctx.data.table_items.employee = QTableWidgetItem(ent_ctx.data.emp_name)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.employee, row, ie.EmployeeTableHeaders.EMPLOYEE)
    ent_ctx.data.table_items.worksheet = QTableWidgetItem(ent_ctx.data.worksheet)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.worksheet, row, ie.EmployeeTableHeaders.WORKSHEET)
    ent_ctx.data.table_items.coord = QTableWidgetItem(ent_ctx.data.coord)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.coord, row, ie.EmployeeTableHeaders.COORDINATE)

def _configure_summary_io_row(
        ent_ctx: EntryContext,
        col: int
    ) -> None:
    '''Configure the summary IO row in the table. There is only one column.'''
    ent_ctx.data.table_items.in_file_names = QTableWidgetItem(ent_ctx.data.in_file_names)
    pu.insert_col_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.in_file_names, ie.SummaryIOTableHeaders.INPUT_WORKBOOKS, col)
    ent_ctx.data.table_items.out_file_names = QTableWidgetItem(ent_ctx.data.out_file_names)
    pu.insert_col_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.out_file_names, ie.SummaryIOTableHeaders.OUTPUT_WORKBOOK, col)
    ent_ctx.data.table_items.date = QTableWidgetItem(ent_ctx.data.date)
    pu.insert_col_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.date, ie.SummaryIOTableHeaders.SELECTED_DATE, col)
    ent_ctx.panel.table.resizeColumnsToContents()

def _configure_summary_data_row(
        ent_ctx: EntryContext,
        row: int,
        hours: em.EmployeeHours
    ) -> None:
    '''Configure the summary data row in the table.'''
    ent_ctx.data.table_items.emp_name = QTableWidgetItem(ent_ctx.data.emp_name)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.emp_name, row, ie.SummaryDataTableHeaders.EMPLOYEE)
    ent_ctx.data.table_items.pred_hrs = QTableWidgetItem(hu.format_summary_data_row_hours(ent_ctx.data.pred_hrs, st.SpecialStrings.ZERO_HOURS))
    ent_ctx.data.table_items.pred_hrs.setForeground(hours.pre_hours_colour)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.pred_hrs, row, ie.SummaryDataTableHeaders.PREDICTED_HOURS)
    ent_ctx.data.table_items.acc_hrs = QTableWidgetItem(hu.format_summary_data_row_hours(ent_ctx.data.acc_hrs, st.SpecialStrings.MISSING))
    ent_ctx.data.table_items.acc_hrs.setForeground(hours.acc_hours_colour)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.acc_hrs, row, ie.SummaryDataTableHeaders.ACCUMULATED_HOURS)
    ent_ctx.data.table_items.dev = QTableWidgetItem(ent_ctx.data.dev)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.dev, row, ie.SummaryDataTableHeaders.DEVIATION)
    ent_ctx.data.table_items.proj_id = QTableWidgetItem(ent_ctx.data.proj_id)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.proj_id, row, ie.SummaryDataTableHeaders.PROJECT_ID)
    ent_ctx.data.table_items.out_ws = QTableWidgetItem(ent_ctx.data.out_ws_name)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.out_ws, row, ie.SummaryDataTableHeaders.OUTPUT_WORKSHEET)
    ent_ctx.data.table_items.coord = QTableWidgetItem(ent_ctx.data.coord)
    pu.insert_row_data_widget(ent_ctx.panel.table, ent_ctx.data.table_items.coord, row, ie.SummaryDataTableHeaders.COORDINATE)

def join_str_list(formatter: str, items: t.StrList) -> str:
    '''Public module function. Joins the items into a single string.'''
    if not items:
        return ""
    return formatter.join(item for item in items)

def join_found_projects(formatter: str, projects: t.ProjectsTup) -> str:
    '''Public module function. Joins the projects into a single string.'''
    if not projects:
        return ""
    return formatter.join(f"{project_id}: {descriptions}" for project_id, descriptions in projects)

def update_current_text(dd: Dropdowns) -> None:
    '''Public module function. Update the current text of the dropdown QComboboxes.'''
    dd.current_text.year = dd.year.currentText()
    dd.current_text.month = dd.month.currentText()
    dd.current_text.worksheet = dd.worksheet.currentText()

def connect_dropdowns(dd: Dropdowns, func: Callable) -> None:
    '''Public module function. Connect the dropdowns to the given function.'''
    for dropdown in (dd.year, dd.month, dd.worksheet):
        dropdown.currentTextChanged.connect(func)

def get_selected_rows(table: QTableWidget) -> list[QModelIndex]:
    '''Public module function. Get the selected rows from the table.'''
    return table.selectionModel().selectedRows() if table.selectionModel() else []

def get_selected_row_indices(selected_rows: list[QModelIndex]) -> tuple[int, ...]:
    '''Public module function. Check if the row selection has changed.'''
    return tuple(row.row() for row in selected_rows)
