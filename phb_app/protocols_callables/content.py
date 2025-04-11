# pylint: disable=missing-docstring
# pylint: disable=line-too-long
# pylint: disable=too-few-public-methods
'''
Module Name
---------
Project Hours Budgeting Protocols

Version
-------
Date-based Version: 202502010
Author: Karl Goran Antony Zuvela

Description
-----------
Provides typesetting signatures.
'''

from dataclasses import (
    dataclass,
    field
)
from typing import (
    Callable,
    TypeAlias,
    Protocol,
    Optional,
    runtime_checkable,
    TYPE_CHECKING
)
from PyQt6.QtWidgets import (
    QVBoxLayout,
    QWidget,
    QTableWidget,
    QPushButton,
    QLabel
)
from phb_app.data.phb_dataclasses import (
    WorkbookManager
)
from phb_app.logging.error_manager import ErrorManager

if TYPE_CHECKING:
    from phb_app.data.phb_dataclasses import BaseTableHeaders

@dataclass
class Instructions:
    input_label: Optional[QLabel] = None
    output_label: Optional[QLabel] = None

@dataclass
class IOControls:
    buttons: list[QPushButton]
    table: QTableWidget
    label: QLabel
    col_widths: dict["BaseTableHeaders", int]

@dataclass
class ErrorControls:
    error_panel: QWidget
    error_manager: Optional[ErrorManager] = field(init=False, default=None)

    def __post_init__(self) -> None:
        self.error_manager = ErrorManager(self.error_panel)

SetupWidgets: TypeAlias = Callable[[IOControls], None]
SetupLayout: TypeAlias = Callable[[list[IOControls], ErrorControls], QVBoxLayout]
SetupMainButtonConnections: TypeAlias = Callable[[IOControls, WorkbookManager], None]

@runtime_checkable
class ConfigureRow(Protocol):
    '''Configure row signature'''

    def __call__(self, table: QTableWidget, row_position: int, file_name: str, file_path: str) -> None: ...

class AddWorkbook(Protocol):
    '''Add workbook signature'''

    def __call__(self, file_path: str, data_only: Optional[bool] = False) -> None: ...
