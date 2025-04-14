
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
from phb_app.data.phb_dataclasses import CountryData, WorkbookManager
from phb_app.logging.error_manager import ErrorManager
from phb_app.wizard.phb_wizard_gui import PHBWizard


        #######################
        ## Start the wizard! ##
        #######################

def main():
    '''Main entry point to program.'''

    app = QApplication(sys.argv)
    window = PHBWizard(CountryData(), ErrorManager(), WorkbookManager())
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
