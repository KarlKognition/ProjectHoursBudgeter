# pylint: disable=missing-docstring

from typing import Protocol
from PyQt6.QtWidgets import (
    QTableWidget,
    QTableWidgetItem
)

class HasTable(Protocol):

    @property
    def table(self) -> QTableWidget: ...

class TableHighlighting(Protocol):

    def highlight_bad_item(self, item: QTableWidgetItem) -> None: ...
    def remove_highlighting(self, item: QTableWidgetItem) -> None: ...

class ErrorReporting(Protocol):

    def add_error(self, file_name:str, file_role: str, error: Exception) -> None: ...
    def remove_error(self, _name: str, file_role: str) -> None: ...
