# pylint: disable=missing-docstring
# pylint: disable=C0301
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
    runtime_checkable
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

@dataclass
class Instructions:
    input_label: Optional[QLabel] = None
    output_label: Optional[QLabel] = None

@dataclass
class IOControls:
    add_input_button: QPushButton
    remove_input_button: QPushButton
    input_table: QTableWidget
    input_label: QLabel

@dataclass
class OutputControls:
    output_table: QTableWidget
    output_label: QLabel

@dataclass
class ErrorControls:
    error_panel: QWidget
    error_manager: Optional[ErrorManager] = field(init=False, default=None)

    def __post_init__(self) -> None:
        self.error_manager = ErrorManager(self.error_panel)

SetupWidgets: TypeAlias = Callable[[Instructions, IOControls, OutputControls], None]
SetupLayout: TypeAlias = Callable[[IOControls, OutputControls, ErrorControls], QVBoxLayout]
SetupMainButtonConnections: TypeAlias = Callable[[IOControls, WorkbookManager], None]

@runtime_checkable
class ConfigureRow(Protocol):
    '''Configure row signature'''

    def __call__(self, table: QTableWidget, row_position: int, file_name: str, file_path: str) -> None: ...

class AddWorkbook(Protocol):
    '''Add workbook signature'''

    def __call__(self, file_path: str, data_only: Optional[bool] = False) -> None: ...
