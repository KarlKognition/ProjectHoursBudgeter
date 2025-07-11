'''
Module Name
---------
Project Hours Budgeting Wizard GUI

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the stages of the GUI.
'''
#           --- Third party libraries ---
from PyQt6.QtGui import QGuiApplication
from PyQt6.QtWidgets import QWizard
#           --- First party libraries ---
import phb_app.wizard.pages.employee_selection as ems
import phb_app.wizard.pages.io_selection as iosp
import phb_app.wizard.pages.project_selection as ps
import phb_app.wizard.pages.explanation as ep
import phb_app.wizard.pages.summary as sp
import phb_app.data.workbook_management as wm
import phb_app.data.location_management as loc
import phb_app.utils.hours_utils as hu
import phb_app.wizard.constants.ui_strings as st


class PHBWizard(QWizard):
    '''Main GUI interface for the Auto Hours Collector.'''

    def __init__(
        self,
        country_data: loc.CountryData,
        workbook_manager: wm.WorkbookManager
        ) -> None:

        super().__init__()
        self.workbook_manager = workbook_manager
        self.setWindowTitle(st.GUI_TITLE)
        self.setGeometry(0, 0, 1000, 600)
        # Centre the main window
        self.move(QGuiApplication.primaryScreen().availableGeometry().center()
                  - self.frameGeometry().center())

        # Created wizard pages
        self.addPage(ep.ExplanationPage())
        self.addPage(iosp.IOSelectionPage(country_data, self.workbook_manager))
        self.addPage(ps.ProjectSelectionPage(self.workbook_manager))
        self.addPage(ems.EmployeeSelectionPage(self.workbook_manager))
        self.addPage(sp.SummaryPage(self.workbook_manager))

        self.setWizardStyle(QWizard.WizardStyle.ModernStyle)

    def accept(self) -> bool:
        '''Extend the functionality of the Finish button.'''
        # Get the first and only output workbook
        wb_out_ctx = self.workbook_manager.get_output_workbook_ctx()
        hu.write_hours_to_output_file(wb_out_ctx)
        wm.save_output_workbook(wb_out_ctx)
        return super().accept()
