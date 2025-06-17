'''
Module Name
---------
Project Hours Budgeting Error Management functionality

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
import phb_app.wizard.constants.ui_strings as st

# UI container for error messages
error_panels: dict[st.IORole, QWidget] = {}
errors: dict[tuple[str, str], dict[str, QLabel]] = {}

def add_error(file_name:str, file_role: st.IORole, error: Exception) -> None:
    '''Add the error per file.'''

    key = (file_name, file_role)
    if key not in errors:
        errors[key] = {}
    # Set the error message in a label
    label = QLabel(str(error))
    # Allow word wrapping
    label.setWordWrap(True)
    # Red styled error
    style_error_message_red(label)
    # Put the label in the UI
    error_panels[file_role].layout().addWidget(label)
    # Track the label
    errors[key][str(error)] = label

def remove_error(file_name: str, file_role: st.IORole) -> None:
    '''Remove the error and label per file.'''

    key = (file_name, file_role)
    if key in errors:
        for label in errors[key].values():
            # Remove the message and file name from tracking
            error_panels[file_role].layout().removeWidget(label)
            # Remove the associated label
            label.deleteLater()
        # If no error for this role, remove entry
        del errors[key]

def style_error_message_red(label: QLabel) -> None:
    '''Error messages are highlighted red and have white text.'''

    palette = label.palette()
    palette.setColor(label.foregroundRole(), Qt.GlobalColor.red)
    label.setPalette(palette)
