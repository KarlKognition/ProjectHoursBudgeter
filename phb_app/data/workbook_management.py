from os import path
from typing import Optional, Iterator
from dataclasses import dataclass, field
from openpyxl import Workbook
# First party libraries
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.location_management as loc
import phb_app.logging.exceptions as ex
import phb_app.data.worksheet_management as ws
import phb_app.utils.file_handling_utils as fu

@dataclass
class ManagedWorkbook:
    '''Parent data class for workbook data.'''
    file_path: str
    file_name: Optional[str] = None
    workbook_object: Optional[Workbook] = None

    def __post_init__(self):
        self.file_name = path.basename(self.file_path)

    def __eq__(self, other):
        '''Compare data objects by file names and data_only.'''
        if isinstance(other, ManagedWorkbook):
            return self.file_name == other.file_name
        return False

@dataclass
class ManagedOutputWorkbook(ManagedWorkbook):
    '''Child data class for output workbook data (budget).'''
    managed_sheet_object: Optional[ws.ManagedOutputWorksheet] = None

    def __post_init__(self):
        super().__post_init__()
        self.workbook_object = fu.try_load_workbook(self, ManagedOutputWorkbook)

    def init_output_worksheet(self) -> None:
        '''Init output worksheet.'''
        self.managed_sheet_object = ws.ManagedOutputWorksheet()
        self.managed_sheet_object.set_sheet_names(self.workbook_object.sheetnames)

    def save_output_workbook(self) -> None:
        '''Saves the workbook with its given file path.'''
        self.workbook_object.save(self.file_path)

@dataclass
class ManagedInputWorkbook(ManagedWorkbook):
    '''Child data class for input workbook data (timesheets).
    The data is provided from a deep copy made from the Location
    Manager's retrieved data.'''
    locale_data: Optional[loc.InputLocaleData] = None
    managed_sheet_object: Optional[ws.ManagedInputWorksheet] = None

    def __post_init__(self):
        super().__post_init__()
        self.workbook_object = fu.try_load_workbook(self, ManagedInputWorkbook)

    def set_locale_data(self, country_data: loc.CountryData, country_name: str) -> None:
        '''Sets the locale data.'''
        # A deep copy is not required because the data object will not change
        self.locale_data = next(
            (locale for locale in country_data.countries
            if locale.country == country_name),
            None)

    def init_input_worksheet(self) -> None:
        '''Init input worksheet.'''
        if len(self.workbook_object.sheetnames) <= 1:
            sheet_name = self.workbook_object.sheetnames[0]
        else:
            sheet_name = self.locale_data.exp_sheet_name
        self.managed_sheet_object = ws.ManagedInputWorksheet(ws.SelectedSheet(sheet_name, self.workbook_object[sheet_name]))
        self.managed_sheet_object.index_headers()

Mng_wb_dispatch: dict[st.IORole, ManagedInputWorkbook | ManagedOutputWorkbook] = {
    st.IORole.INPUTS: ManagedInputWorkbook,
    st.IORole.OUTPUT: ManagedOutputWorkbook
}

@dataclass
class WorkbookManager:
    '''Data class for tracking workbooks.'''

    workbooks: list[ManagedInputWorkbook | ManagedOutputWorkbook] = field(default_factory=list)

    def try_add_workbook(self, workbook: ManagedInputWorkbook | ManagedOutputWorkbook) -> None:
        '''Try to add the workbook to tracking.'''
        retrieved_wb = self.get_workbook_by_name(workbook.file_name)
        if retrieved_wb != workbook:
            self.workbooks.append(workbook)
        else:
            raise ex.WorkbookAlreadyTracked(workbook.file_name)

    def add_workbook(self, file_path: str, role: st.IORole) -> None:
        '''Add input workbooks by filename.'''
        workbook = Mng_wb_dispatch[role](file_path=file_path)
        self.try_add_workbook(workbook)

    def get_workbook_by_name(self, file_name: str) -> ManagedInputWorkbook | ManagedOutputWorkbook | None:
        '''Retrieves the first workbook in the list by file name.'''
        return next((wb for wb in self.workbooks if file_name == wb.file_name), None)

    def yield_workbooks_by_type(self, wb_class: ManagedInputWorkbook | ManagedOutputWorkbook) -> Iterator[ManagedInputWorkbook | ManagedOutputWorkbook]:
        '''Returns a generator of workbooks in the list by type.'''
        for wb in self.workbooks:
            if isinstance(wb, wb_class):
                yield wb

    def remove_workbook(self, file_name: str) -> None:
        '''Add input/output workbooks by filename.'''
        if file_name:
            wb = self.get_workbook_by_name(file_name)
            # Delete workbook object if it was created
            if wb:
                self.workbooks.remove(wb)
                del wb
