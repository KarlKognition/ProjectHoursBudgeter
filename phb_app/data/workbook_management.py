"""
Package
-------
Data Handling

Module Name
---------
Workbook Management

Version
-------
Date-based Version: 10062025
Author: Karl Goran Antony Zuvela

Description
-----------
Provides workbook management.
"""
from os import path
from typing import Optional, Iterator
from dataclasses import dataclass
from openpyxl import Workbook
# First party libraries
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.location_management as loc
import phb_app.logging.exceptions as ex
import phb_app.data.worksheet_management as ws
import phb_app.utils.file_handling_utils as fu

            ### DATA CONTAINERS ###

@dataclass
class ManagedWorkbook:
    """Parent data class for workbook data."""
    file_path: str
    file_name: str
    workbook_object: Workbook

@dataclass
class InputWorkbookContext:
    """Context data class for managing an input workbook."""
    mngd_wb: ManagedWorkbook
    locale_data: Optional[loc.InputLocaleData] = None
    managed_sheet: Optional[ws.ManagedInputWorksheet] = None
    worksheet_service: Optional[ws.InputWorksheetService] = None

@dataclass
class OutputWorkbookContext:
    """Context data class for managing an output workbook."""
    mngd_wb: ManagedWorkbook
    managed_sheet: Optional[ws.ManagedOutputWorksheet] = None
    worksheet_service: Optional[ws.OutputWorksheetService] = None

            ### FACTORY FUNCTIONS ###

def _create_managed_workbook_from_file(file_path: str, writable: bool = False) -> ManagedWorkbook:
    """Creates a ManagedWorkbook instance from a file path, loading the workbook object."""
    file_name = path.basename(file_path)
    workbook_object = fu.try_load_workbook(file_path, file_name, writable=writable)
    return ManagedWorkbook(file_path, file_name, workbook_object)

def _create_input_context(file_path: str) -> InputWorkbookContext:
    """Creates an InputWorkbookContext for the given file path."""
    core = _create_managed_workbook_from_file(file_path)
    return InputWorkbookContext(mngd_wb=core)

def _create_output_context(file_path: str) -> OutputWorkbookContext:
    """Creates an OutputWorkbookContext for the given file path."""
    core = _create_managed_workbook_from_file(file_path, writable=True)
    return OutputWorkbookContext(mngd_wb=core)

def create_wb_context_by_role(file_path: str, role: st.IORole) -> InputWorkbookContext | OutputWorkbookContext:
    """Creates a context for the given file path and role."""
    dispatch = {
        st.IORole.INPUTS: _create_input_context,
        st.IORole.OUTPUT: _create_output_context
    }
    try:
        return dispatch[role](file_path)
    except KeyError as exc:
        raise ValueError(f"Invalid role: {role}") from exc

            ### INPUT SERVICE FUNCTIONS ###

def set_locale_data(
    context: InputWorkbookContext,
    country_data: loc.CountryData,
    country_name: str
    ) -> None:
    """Sets the locale data for the input workbook."""
    # A deep copy is not required because the data object will not change
    context.locale_data = next(
        (locale for locale in country_data.countries
        if locale.country == country_name),
        None)

def init_input_worksheet(context: InputWorkbookContext) -> None:
    """Init input worksheet."""
    sheetnames = context.mngd_wb.workbook_object.sheetnames
    # If there is only one sheet, use that;
    # otherwise, use the locale data's expected sheet name.
    sheet_name = (
        sheetnames[0]
        if len(sheetnames) <= 1
        else context.locale_data.exp_sheet_name
    )
    sheet_obj = context.mngd_wb.workbook_object[sheet_name]
    managed_sheet = ws.create_managed_input_worksheet(sheet_name, sheet_obj)
    service = ws.InputWorksheetService(worksheet=managed_sheet)
    context.managed_sheet = managed_sheet
    context.worksheet_service = service
    service.index_headers()

            ### OUTPUT SERVICE FUNCTIONS ###

def init_output_worksheet(context: OutputWorkbookContext) -> None:
    """Init output worksheet."""
    worksheet = ws.create_managed_output_worksheet()
    service = ws.OutputWorksheetService(worksheet=worksheet)
    service.set_sheet_names(context.mngd_wb.workbook_object.sheetnames)
    context.managed_sheet = worksheet
    context.worksheet_service = service

def save_output_workbook(context: OutputWorkbookContext) -> None:
    """Saves the workbook with its given file path."""
    context.mngd_wb.workbook_object.save(context.mngd_wb.file_path)

            ### MANAGER ###

class WorkbookManager:
    """Class for tracking workbooks."""

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
        wb_ctx = self.get_workbook_ctx_by_file_name_and_role(role, ctx.mngd_wb.file_name)
        if wb_ctx:
            raise ex.WorkbookAlreadyTracked(ctx.mngd_wb.file_name)
        self.workbooks_ctxs[role].append(ctx)

    def get_output_workbook_ctx(self) -> OutputWorkbookContext | None:
        """Retrieves the first output workbook context."""
        output_list = self.workbooks_ctxs.get(st.IORole.OUTPUT, [])
        return output_list[0] if output_list else None

    def get_workbook_ctx_by_file_name_and_role(
        self,
        role: st.IORole,
        name: str
    ) -> InputWorkbookContext | OutputWorkbookContext | None:
        """Retrieves the first workbook context by file name and role."""
        return next((ctx for ctx in self.workbooks_ctxs[role] if ctx.mngd_wb.file_name == name), None)

    def yield_workbook_ctxs_by_role(
        self,
        role: st.IORole
    ) -> Iterator[InputWorkbookContext | OutputWorkbookContext | None]:
        """Yields workbooks by role. An empty list is returned if no workbooks are found."""
        yield from self.workbooks_ctxs.get(role, [])

    def remove_wb_ctx_by_file_name(self, role: st.IORole, name: str) -> None:
        """Removes a workbook context by file name and role."""
        ctx = self.get_workbook_ctx_by_file_name_and_role(role, name)
        if ctx:
            self.workbooks_ctxs[role].remove(ctx)
            del ctx
