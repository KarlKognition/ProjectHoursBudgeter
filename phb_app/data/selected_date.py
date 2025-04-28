from dataclasses import dataclass
from typing import Optional

@dataclass(init=False)
class SelectedDate:
    '''Data class for the located date for which hours will be recorded per employee
    in the budgeting file.'''
    month: Optional[int]
    year: Optional[int]
    row: Optional[int]