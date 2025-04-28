
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
import phb_app.data.common as common
import phb_app.data.phb_dataclasses as dc
import phb_app.logging.error_manager as em
import phb_app.wizard.phb_wizard_gui as wg


        #######################
        ## Start the wizard! ##
        #######################

def main():
    '''Main entry point to program.'''

    app = QApplication(sys.argv)
    window = wg.PHBWizard(common.CountryData(), em.ErrorManager(), dc.WorkbookManager())
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
