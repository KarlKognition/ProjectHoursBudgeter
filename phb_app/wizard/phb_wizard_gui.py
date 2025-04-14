'''
Module Name
---------
Project Hours Budgeting Wizard GUI

Version
-------
Date-based Version: 202502005
Author: Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the stages of the GUI.
'''

## Imports
# Third party libraries
from PyQt6.QtGui import QGuiApplication
from PyQt6.QtWidgets import QWizard
# First party libraries
from phb_app.wizard.pages.employee_selection import EmployeeSelectionPage
from phb_app.wizard.pages.io_selection import IOSelectionPage
from phb_app.wizard.pages.project_selection import ProjectSelectionPage
from phb_app.wizard.pages.explanation import ExplanationPage
from phb_app.wizard.pages.summary import SummaryPage
from phb_app.data.phb_dataclasses import (
    QPropName,
    ManagedOutputWorkbook,
    WorkbookManager
)
import phb_app.utils.hours_utils as hutils
from phb_app.wizard.constants.ui_strings import GUI_TITLE


########################
class PHBWizard(QWizard):
    '''Main GUI interface for the Auto Hours Collector.'''

    def __init__(self, country_data, error_manager, workbook_manager):
        super().__init__()

        self.init_main_window(country_data, error_manager, workbook_manager)

    def init_main_window(self, country_data, error_manager, workbook_manager):
        '''Init main window.'''

        self.setWindowTitle(GUI_TITLE)
        self.setGeometry(0, 0, 1000, 600)
        # Centre the main window
        self.move(QGuiApplication.primaryScreen().availableGeometry().center()
                  - self.frameGeometry().center())

        # Created wizard pages
        self.addPage(ExplanationPage())
        self.addPage(IOSelectionPage(country_data, error_manager, workbook_manager))
        self.addPage(ProjectSelectionPage())
        self.addPage(EmployeeSelectionPage())
        self.addPage(SummaryPage())

        self.setWizardStyle(QWizard.WizardStyle.ModernStyle)

    def accept(self):
        '''Extend the functionality of the Finish button.'''

        managed_workbooks = self.property(
            QPropName.MANAGED_WORKBOOKS.value)
        if isinstance(managed_workbooks, WorkbookManager):
            # Get the first and only output workbook
            wb_out = next(managed_workbooks.yield_workbooks_by_type(
                ManagedOutputWorkbook))
            hutils.write_hours_to_output_file(wb_out)
            wb_out.save_output_workbook()
        return super().accept()
