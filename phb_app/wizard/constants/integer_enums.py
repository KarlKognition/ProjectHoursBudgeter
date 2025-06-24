from enum import IntEnum, auto

class BaseTableHeaders(IntEnum):
    '''Base enum class with overridden string dunder method
    and general class method.'''
    @classmethod
    def cap_members_list(cls) -> list[str]:
        '''Return a list of capitalised enum members.'''
        return [name.replace('_', ' ').title() for name in cls.__members__]

    @classmethod
    def list_all_values(cls) -> list[int | str]:
        '''Returns a list of all member values.'''
        return [member.value for member in cls]

class InputTableHeaders(BaseTableHeaders):
    '''Input table headers in IOSelection.'''

    FILENAME = 0
    COUNTRY = auto()
    WORKSHEET = auto()

class OutputTableHeaders(BaseTableHeaders):
    '''Output table headers in IOSelection.'''

    FILENAME = 0
    WORKSHEET = auto()
    MONTH = auto()
    YEAR = auto()

class OutputFile(BaseTableHeaders):
    '''Enum for table rows or columns.'''

    FIRST_ENTRY = 0
    SECOND_ENTRY = auto()

class ProjectIDTableHeaders(BaseTableHeaders):
    '''Project ID headers in project selection.'''

    PROJECT_ID = 0
    DESCRIPTION = auto()
    FILENAME = auto()

class EmployeeTableHeaders(BaseTableHeaders):
    '''Employee table headers in employee selection.'''

    EMPLOYEE = 0
    WORKSHEET = auto()
    COORDINATE = auto()
