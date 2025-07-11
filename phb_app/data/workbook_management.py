"""
Package
-------
Data Handling

Module Name
---------
Workbook Management

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Provides workbook management.
"""
#           --- Standard libraries ---
from os import path
from typing import Optional, Iterator
from uuid import UUID, uuid4
from dataclasses import dataclass
#           --- Third party libraries ---
from openpyxl import Workbook
#           --- First party libraries ---
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.location_management as loc
import phb_app.logging.exceptions as ex
import phb_app.data.worksheet_management as ws
import phb_app.utils.file_handling_utils as fu

#           --- DATA CONTAINERS ---

@dataclass(slots=True)
class ManagedWorkbook:
    """Parent data class for workbook data."""
    file_path: str
    file_name: str
    uuid: UUID
    workbook_object: Workbook

@dataclass(slots=True)
class InputWorkbookContext:
    """Context data class for managing an input workbook."""
    mngd_wb: ManagedWorkbook
    locale_data: Optional[loc.InputLocaleData] = None
    managed_sheet: Optional[ws.InputWorksheetContext] = None
    worksheet_service: Optional[ws.InputWorksheetService] = None

# Set eq to false to allow lru caching of hours_utils.compute_predicted_hours(...)
@dataclass(eq=False, slots=True)
class OutputWorkbookContext:
    """Context data class for managing an output workbook."""
    mngd_wb: ManagedWorkbook
    managed_sheet: Optional[ws.OutputWorksheetContext] = None
    worksheet_service: Optional[ws.OutputWorksheetService] = None

#           --- MODULE FACTORY FUNCTIONS ---

def get_file_name_from_path(file_path: str) -> str:
    """Public module level. Returns the file name from a file path."""
    return path.basename(file_path)

def set_uuid() -> UUID:
    """Public module level. Returns a new UUID."""
    return uuid4()

def _create_managed_workbook_from_file(file_path: str, writable: bool = False) -> ManagedWorkbook:
    """Private module level. Creates a ManagedWorkbook instance from a file path, loading the workbook object."""
    file_name = get_file_name_from_path(file_path)
    uuid = set_uuid()
    workbook_object = fu.try_load_workbook(file_path, file_name, writable=writable)
    return ManagedWorkbook(file_path, file_name, uuid, workbook_object)

def _create_input_context(file_path: str) -> InputWorkbookContext:
    """Private module level. Creates an InputWorkbookContext for the given file path."""
    core = _create_managed_workbook_from_file(file_path)
    return InputWorkbookContext(mngd_wb=core)

def _create_output_context(file_path: str) -> OutputWorkbookContext:
    """Private module level. Creates an OutputWorkbookContext for the given file path."""
    core = _create_managed_workbook_from_file(file_path, writable=True)
    return OutputWorkbookContext(mngd_wb=core)

def create_wb_context_by_role(file_path: str, role: st.IORole) -> InputWorkbookContext | OutputWorkbookContext:
    """Public module level. Creates a context for the given file path and role."""
    dispatch = {
        st.IORole.INPUTS: _create_input_context,
        st.IORole.OUTPUT: _create_output_context
    }
    try:
        return dispatch[role](file_path)
    except KeyError as exc:
        raise ValueError(f"Invalid role: {role}") from exc

#           --- INPUT SERVICE MODULE FUNCTIONS ---

def set_locale_data(
    context: InputWorkbookContext,
    country_data: loc.CountryData,
    country_name: str
    ) -> None:
    """Public module level. Sets the locale data for the input workbook."""
    # A deep copy is not required because the data object will not change
    context.locale_data = next(
        (locale for locale in country_data.countries
        if locale.country == country_name),
        None)

#           --- OUTPUT SERVICE MODULE FUNCTIONS ---

def save_output_workbook(context: OutputWorkbookContext) -> None:
    """Public module level. Saves the workbook with its given file path."""
    context.mngd_wb.workbook_object.save(context.mngd_wb.file_path)

#            --- WORKBOOK MANAGER ---

class WorkbookManager:
    """Class for tracking workbooks."""

    __slots_ = ('workbooks_ctxs',)

    def __init__(self) -> None:
        self.workbooks_ctxs: dict[st.IORole, list[InputWorkbookContext | OutputWorkbookContext]] = {
            st.IORole.INPUTS: [],
            st.IORole.OUTPUT: []
        }

    def add_workbook(
        self,
        role: st.IORole,
        ctx: InputWorkbookContext | OutputWorkbookContext
    ) -> None:
        """Adds a workbook context to the manager by role."""
        wb_ctx = self.get_workbook_ctx_by_role_and_uuid(role, ctx.mngd_wb.uuid)
        self.workbooks_ctxs[role].append(ctx)
        if wb_ctx:
            raise ex.WorkbookAlreadyTracked(ctx.mngd_wb.file_name)

    def get_output_workbook_ctx(self) -> OutputWorkbookContext | None:
        """Retrieves the first output workbook context."""
        output_list = self.workbooks_ctxs.get(st.IORole.OUTPUT, [])
        return output_list[0] if output_list else None

    def get_workbook_ctx_by_role_and_uuid(
        self,
        role: st.IORole,
        uuid: UUID
    ) -> InputWorkbookContext | OutputWorkbookContext | None:
        """Retrieves the first workbook context by file name and role."""
        return next((ctx for ctx in self.workbooks_ctxs[role] if ctx.mngd_wb.uuid == uuid), None)

    def yield_workbook_ctxs_by_role(
        self,
        role: st.IORole
    ) -> Iterator[InputWorkbookContext | OutputWorkbookContext | None]:
        """Yields workbooks by role. An empty list is returned if no workbooks are found."""
        yield from self.workbooks_ctxs.get(role, [])

    def get_wb_names_list_by_role(self, role: st.IORole) -> list[str]:
        """Returns a list of workbook names by role."""
        return [ctx.mngd_wb.file_name for ctx in self.workbooks_ctxs.get(role, [])]

    def remove_wb_ctx_by_uuid(self, role: st.IORole, uuid: UUID) -> None:
        """Removes a workbook context by file name and role."""
        ctx = self.get_workbook_ctx_by_role_and_uuid(role, uuid)
        if ctx:
            self.workbooks_ctxs[role].remove(ctx)
            del ctx
