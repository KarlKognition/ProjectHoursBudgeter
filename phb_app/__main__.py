
'''
Package
-------
Project Hours Budgeting

Module Name
---------
Main

Author
-------
Karl Goran Antony Zuvela

Description
-----------
The main entry point to the Project Hours Budgeting Wizard.
'''
#           --- Standard libraries ---
import sys
#           --- Third party libraries ---
from PyQt6.QtWidgets import QApplication
#           --- First party libraries ---
import phb_app.data.location_management as loc
import phb_app.data.workbook_management as wm
import phb_app.wizard.phb_wizard_gui as wg
import phb_app.wizard.constants.ui_strings as st

def main():
    '''Main entry point to program.'''

    app = QApplication(sys.argv)
    st.set_app_default_font_theme(app)
    window = wg.PHBWizard(loc.CountryData(), wm.WorkbookManager())
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
