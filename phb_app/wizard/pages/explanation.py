'''
Package
-------
PHB Wizard

Module Name
---------
IO Selection Page

Version
-------
Date-based Version: 202502010
Author: Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the Input-Output selection page.
'''

from PyQt6.QtWidgets import QWizardPage, QHBoxLayout
import phb_app.wizard.constants.ui_strings as st
import phb_app.utils.page_utils as pu

class ExplanationPage(QWizardPage):
    '''Explanation of the wizard's main use case and the steps involved.'''
    def __init__(self):
        super().__init__()

        self.init_intro_page()

    def init_intro_page(self):
        '''Init the introduction page.'''
        # Set titles
        pu.set_titles(self, st.INTRO_TITLE, st.INTRO_SUBTITLE)

        # Create widgets
        watermark_label = pu.create_watermark_label(st.IMAGES_DIR, st.IMAGE_LOAD_FAIL)
        intro_message = pu.create_intro_message(st.INTRO_MESSAGE)

        # Set up layout
        pu.setup_page(self, [watermark_label, intro_message], QHBoxLayout(), spacing= 35, margins=(25, 25, 25, 25))
