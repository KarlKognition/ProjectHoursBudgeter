'''
Package
-------
PHB Wizard

Module Name
---------
Employee Management

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Data classes for managing employee names and hours in the given worksheet.
'''
#           --- Standard libraries ---
from dataclasses import dataclass, field
from typing import Optional
#           --- Third party libraries ---
from PyQt6.QtGui import QColor
from openpyxl import utils as xlutils
#           --- First party libraries ---
import phb_app.data.yaml_handler as yh
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.hours_deviation as hd

@dataclass(eq=False, slots=True) # Set eq to false to allow lru caching of file_handling_utils.locate_employee_range(...)
class EmployeeRowAnchors(yh.YamlHandler):
    '''Data class to define the anchor strings of the row containing the employee names.'''
    start_anchor: str = ""
    end_anchor: str = ""

    def __post_init__(self):
        yh.YamlHandler.__init__(self)
        self._load_yaml_data()

    def _process_yaml(self, yaml_data) -> None:
        '''Processes the yaml data.'''
        anchors: Optional[dict[str, str]] = yaml_data.get(st.YamlEnum.ROW_ANCHORS, {})
        for key, value in anchors.items():
            if hasattr(self, key):
                setattr(self, key, value)

@dataclass(eq=False, slots=True) # Set eq to false to allow lru caching of file_handling_utils.locate_employee_range(...)
class EmployeeRange:
    '''Data class for the range within which the employee names are located.'''
    start_cell: str = ""
    end_cell: str = ""

    @property
    def start_col_idx(self) -> int:
        '''Get the column integer of the start cell.'''
        return xlutils.coordinate_to_tuple(self.start_cell)[1]

    @property
    def end_col_idx(self) -> int:
        '''Get the column integer of the end cell.'''
        return xlutils.coordinate_to_tuple(self.end_cell)[1]

    @property
    def start_row_idx(self) -> int:
        '''Get the row integer of the start cell.'''
        return xlutils.coordinate_to_tuple(self.start_cell)[0]

    @property
    def end_row_idx(self) -> int:
        '''Get the row integer of the end cell.'''
        return xlutils.coordinate_to_tuple(self.end_cell)[0]

@dataclass(slots=True)
class EmployeeHours:
    '''Data class for managing the predicted and accumulated hours
    per employee.'''
    predicted_hours: Optional[float|int] = None
    pre_hours_colour: QColor = field(default_factory=lambda: st.DEFAULT_FONT_COLOUR)
    accumulated_hours: Optional[float|int] = None
    acc_hours_colour: QColor = field(default_factory=lambda: st.DEFAULT_FONT_COLOUR)
    hours_coord: Optional[str] = None
    thresholds: hd.HoursDeviation = field(default_factory=hd.HoursDeviation)
    deviation: Optional[str] = None

    def set_deviation(self) -> None:
        '''Sets the deviation by percentage between predicted and accumulated hours.'''
        pre_hrs = self.predicted_hours
        acc_hrs = self.accumulated_hours
        if not isinstance(acc_hrs, (float, int)):
            acc_hrs = 0
        if acc_hrs == 0 and pre_hrs == 0:
            self.deviation = " "
        else:
            # Use only absolute values to ensure correct calculation of 1 - frac
            pre_hrs = abs(pre_hrs)
            acc_hrs = abs(acc_hrs)
            frac = (min(acc_hrs, pre_hrs) / max(acc_hrs, pre_hrs))
            if 1 - frac >= self.thresholds.strong_dev:
                self.deviation = "Warning! Strong deviation!"
            elif 1 - frac >= self.thresholds.weak_dev:
                self.deviation = "Weak deviation"
            else:
                self.deviation = "Negligible"

@dataclass(slots=True)
class Employee:
    '''Data class for managing employee name location and related hours
    in the given worksheet. Projects where the employee registered hours
    will be saved here.'''
    name: str
    found_projects: dict[int|str, list[str]] = field(default_factory=dict)
    hours: EmployeeHours = field(default_factory=EmployeeHours)
