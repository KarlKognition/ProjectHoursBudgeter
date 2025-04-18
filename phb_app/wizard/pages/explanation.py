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
from phb_app.wizard.constants.ui_strings import INTRO_MESSAGE, INTRO_TITLE, INTRO_SUBTITLE, IMAGE_LOAD_FAIL, IMAGES_DIR
import phb_app.utils.page_utils as putils

class ExplanationPage(QWizardPage):
    '''Explanation of the wizard's main use case and the steps involved.'''
    def __init__(self):
        super().__init__()

        self.init_intro_page()

    def init_intro_page(self):
        '''Init the introduction page.'''
        # Set titles
        putils.set_titles(self, INTRO_TITLE, INTRO_SUBTITLE)

        # Create widgets
        watermark_label = putils.create_watermark_label(IMAGES_DIR, IMAGE_LOAD_FAIL)
        intro_message = putils.create_intro_message(INTRO_MESSAGE)

        # Set up layout
        putils.setup_page(self, [watermark_label, intro_message], QHBoxLayout(), spacing= 35, margins=(25, 25, 25, 25))
