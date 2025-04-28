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
import phb_app.wizard.pages.employee_selection as esp
import phb_app.wizard.pages.io_selection as iosp
import phb_app.wizard.pages.project_selection as psp
import phb_app.wizard.pages.explanation as ep
import phb_app.wizard.pages.summary as sp
import phb_app.data.common as common
import phb_app.logging.error_manager as em
import phb_app.data.phb_dataclasses as dc
import phb_app.utils.hours_utils as hutils
import phb_app.wizard.constants.ui_strings as st


########################
class PHBWizard(QWizard):
    '''Main GUI interface for the Auto Hours Collector.'''

    def __init__(self, country_data: common.CountryData, error_manager: em.ErrorManager, workbook_manager: dc.WorkbookManager):
        super().__init__()

        self.setWindowTitle(st.GUI_TITLE)
        self.setGeometry(0, 0, 1000, 600)
        # Centre the main window
        self.move(QGuiApplication.primaryScreen().availableGeometry().center()
                  - self.frameGeometry().center())

        # Created wizard pages
        self.addPage(ep.ExplanationPage())
        self.addPage(iosp.IOSelectionPage(country_data, error_manager, workbook_manager))
        self.addPage(psp.ProjectSelectionPage())
        self.addPage(esp.EmployeeSelectionPage())
        self.addPage(sp.SummaryPage())

        self.setWizardStyle(QWizard.WizardStyle.ModernStyle)

    def accept(self):
        '''Extend the functionality of the Finish button.'''

        managed_workbooks = self.property(
            dc.QPropName.MANAGED_WORKBOOKS.value)
        if isinstance(managed_workbooks, dc.WorkbookManager):
            # Get the first and only output workbook
            wb_out = next(managed_workbooks.yield_workbooks_by_type(
                common.ManagedOutputWorkbook))
            hutils.write_hours_to_output_file(wb_out)
            wb_out.save_output_workbook()
        return super().accept()
