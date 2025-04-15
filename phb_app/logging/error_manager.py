'''
Module Name
---------
Project Hours Budgeting Error Manager

Version
-------
Date-based Version: 20250224
Author: Karl Goran Antony Zuvela

Description
-----------
Tracks errors for the Project Hours Budgeting Wizard.
'''

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QWidget, QLabel

class ErrorManager:
    '''Manages the errors using QWizard QLabels.'''
    def __init__(self) -> None:
        # UI container for error messages
        self.error_panels: dict[str, QWidget] = {}
        self.errors: dict[tuple[str, str], dict[str, QLabel]] = {}

    def add_error(self,
                  file_name:str,
                  file_role: str,
                  error: Exception) -> None:
        '''Add the error per file.'''

        key = (file_name, file_role)
        if key not in self.errors:
            self.errors[key] = {}
        # Set the error message in a label
        label = QLabel(str(error))
        # Allow word wrapping
        label.setWordWrap(True)
        # Red styled error
        self.style_error_message(label)
        # Put the label in the UI
        self.error_panels.layout().addWidget(label)
        # Track the label
        self.errors[key][str(error)] = label

    def remove_error(self,
                     file_name: str,
                     file_role: str) -> None:
        '''Remove the error and label per file.'''

        key = (file_name, file_role)
        if key in self.errors:
            for label in self.errors[key].values():
                # Remove the message and file name from tracking
                self.error_panels.layout().removeWidget(label)
                # Remove the associated label
                label.deleteLater()
            # If no error for this role, remove entry
            del self.errors[key]

    def style_error_message(self,
                            label: QLabel) -> None:
        '''Error messages are highlighted red and have white text.'''

        palette = label.palette()
        palette.setColor(label.foregroundRole(), Qt.GlobalColor.red)
        label.setPalette(palette)
