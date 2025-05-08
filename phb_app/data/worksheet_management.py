from dataclasses import dataclass, field
from typing import Optional, Iterator
from openpyxl.worksheet.worksheet import Worksheet
from PyQt6.QtCore import Qt
import phb_app.data.selected_date as sd
import phb_app.data.employee_management as emp
import phb_app.templating.types as types

@dataclass
class SelectedSheet:
    '''Data class for worksheet names.'''
    sheet_name: str
    sheet_object: Worksheet

@dataclass
class ManagedWorksheet:
    '''Parent data class for worksheet data.'''
    selected_sheet: Optional[SelectedSheet] = None
    sheet_names: list[str] = field(default_factory=list)

@dataclass
class ManagedInputWorksheet(ManagedWorksheet):
    '''Child data class for input worksheet data.'''
    # The project number is not provided at init/postinit
    # key: project ID (network number [int] or PSP element [str]),
    # value: list of short project desciptions
    selectable_project_ids: types.ProjectsDict = field(default_factory=types.ProjectsDict)
    # Dictionary of project IDs (network number [int] or PSP element [str]): descriptions
    selected_project_ids: types.ProjectsDict = field(default_factory=types.ProjectsDict)
    indexed_headers: dict[str, int] = field(default_factory=dict)

    def index_headers(self) -> None:
        '''Retrieves the table headers from the given worksheet
        and maps them to their column indices.'''
        # sheet object [1] is row one, where the headers are
        for idx, cell in enumerate(self.selected_sheet.sheet_object[1]):
            if isinstance(cell.value, str):
                self.indexed_headers[cell.value] = idx

    def yield_from_project_id_and_desc(self) -> Iterator[types.ProjectsTup]:
        '''Yields from the project ID and description,
        one at a time in a tuple.'''
        yield from self.selectable_project_ids.items()

    def set_selectable_project_ids(self, proj_id: str, proj_desc: str, name: str) -> None:
        '''Extracts all project IDs in the selected worksheet and saves the 
        data in the selectable project IDs dictionary.'''
        # Get the worksheet object
        sheet_object = self.selected_sheet.sheet_object
        # Get the project ID column index
        project_id_col = self.indexed_headers.get(proj_id)
        # Get the project description column index
        project_desc_col = self.indexed_headers.get(proj_desc)
        # Get the project person column index
        person_col = self.indexed_headers.get(name)
        for row in sheet_object.iter_rows(min_row=2):
            # Get the contents of each row under the given headers
            id_value = row[project_id_col].value
            desc_value = row[project_desc_col].value
            person_value = row[person_col].value
            # Only add the project ID and description if they exist alongside an employee
            if id_value and desc_value and person_value:
                id_value = str(id_value)
                desc_value = str(desc_value)
                person_value = str(person_value)
                if id_value not in self.selectable_project_ids:
                    # If the ID is not in the dictionary, make a new list for descriptions
                    self.selectable_project_ids[id_value] = []
                if desc_value not in self.selectable_project_ids[id_value]:
                    # Append the descriptions per ID, if it has not already been appended
                    self.selectable_project_ids[id_value].append(desc_value)

@dataclass
class ManagedOutputWorksheet(ManagedWorksheet):
    '''Child data class for output worksheet data.'''
    # Replaced through IOSelection
    selected_date: sd.SelectedDate = field(default_factory=sd.SelectedDate)
    employee_row_anchors: emp.EmployeeRowAnchors = field(default_factory=emp.EmployeeRowAnchors)
    # Start cell: coord, end cell: coord
    employee_range: emp.EmployeeRange = field(default_factory=emp.EmployeeRange)
    # Key: Cell coord; value: Employee object
    selected_employees: dict[str, emp.Employee] = field(default_factory=dict)

    def set_sheet_names(self, sheet_names) -> None:
        '''Set sheet names.'''
        self.sheet_names = sheet_names

    def set_selected_employees(self, coord_name: list[tuple[str, str]]) -> None:
        '''Save the coordinate in the worksheet with the selected employee.'''
        for coord, name in coord_name:
            self.selected_employees[coord] = emp.Employee(name)

    def set_predicted_hours(self, coord_hours: dict[str, float|int]) -> None:
        '''Sets the attribute with the predicted hours found in the worksheet.'''
        for coord, hours in coord_hours.items():
            employee = self.selected_employees.get(coord)
            if employee:
                if not isinstance(hours, (float, int)):
                    hours = 0.0
                employee.hours.predicted_hours = hours

    def set_predicted_hours_colour(self) -> None:
        '''Set the color formatting of the employee's predicted hours.
        Predicted hours may be coloured as the user wishes but recorded
        hours must be set to black (default, theme or RGB).'''
        for employee in self.yield_from_selected_employee():
            if employee.hours.hours_coord:
                # Check each employee's predicted hours colour
                font_colour = self.selected_sheet.sheet_object[employee.hours.hours_coord].font.color
                if not font_colour:
                    # Default (black): Already recorded.
                    # Show red to get the user's attention in summary
                    employee.hours.pre_hours_colour = Qt.GlobalColor.red
                elif (
                    font_colour.type == 'theme' and
                    font_colour.theme == 1 and
                    font_colour.tint == 0.0
                ):
                    # Black theme: Already recorded.
                    # Show red to get the user's attention in summary
                    employee.hours.pre_hours_colour = Qt.GlobalColor.red
                elif (
                    font_colour.type == 'rgb' and
                    font_colour.rgb == 'FF000000'
                ):
                    # Black RGB: Already recorded.
                    # Show red to get the user's attention in summary
                    employee.hours.pre_hours_colour = Qt.GlobalColor.red

    def clear_predicted_hours(self) -> None:
        '''Clears the recorded employee names and respective hours.'''
        self.selected_employees.clear()

    def yield_from_selected_employee(self) -> Iterator[emp.Employee]:
        '''Yields from the selected employee.'''
        yield from self.selected_employees.values()
