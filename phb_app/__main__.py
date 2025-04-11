
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
from phb_app.data.phb_dataclasses import CountryData
from phb_app.wizard.phb_wizard_gui import PHBWizard


        #######################
        ## Start the wizard! ##
        #######################

def main():
    '''Main entry point to program.'''

    country_data = CountryData()
    app = QApplication(sys.argv)
    window = PHBWizard(country_data)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
