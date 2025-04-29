
'''
Package
-------
Project Hours Budgeting

Module Name
---------
Main

Version
-------
Date-based Version: 202502010
Author: Karl Goran Antony Zuvela

Description
-----------
The main entry point to the Project Hours Budgeting Wizard.
'''

## Imports
# Standard libraries
import sys
# Third party libraries
from PyQt6.QtWidgets import QApplication
# First party libraries
import phb_app.data.location_management as loc
import phb_app.data.workbook_management as wm
import phb_app.logging.error_manager as em
import phb_app.wizard.phb_wizard_gui as wg

def main():
    '''Main entry point to program.'''

    app = QApplication(sys.argv)
    window = wg.PHBWizard(loc.CountryData(), em.ErrorManager(), wm.WorkbookManager())
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
