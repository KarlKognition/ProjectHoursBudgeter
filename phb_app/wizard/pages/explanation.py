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

from pathlib import Path
import git
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QWizardPage, QHBoxLayout, QLabel

def find_git_root():
    '''Find the root of the project.'''
    repo = git.Repo(Path(__file__).resolve(), search_parent_directories=True)
    return Path(repo.working_tree_dir)

APP_ROOT = find_git_root() / "phb_app"
IMAGES_DIR = APP_ROOT / "images"

class ExplanationPage(QWizardPage):
    '''Explanation of the wizard's main use case and the steps involved.'''
    def __init__(self):
        super().__init__()

        self.init_intro_page()

    def init_intro_page(self):
        '''Init the introduction page.'''

        main_layout = QHBoxLayout() # Horizontal box layout
        main_layout.setSpacing(35) # Horizontal spacing between widgets
        main_layout.setContentsMargins(25, 25, 25, 25)
        self.setTitle("Introdution")
        self.setSubTitle("Welcome to the Project Hours Budgeting Wizard!")

        ## Watermark
        # Prepare a label which will receive the pixmap
        watermark_label = QLabel(self)
        self.set_watermark(watermark_label)

        ## Intro message
        intro_message = QLabel("This wizard allows the user to budget the hours of many "
                         "employees of a project at once.\n"
                         "There are three steps:\n\n"
                         "1: Select the input files (SAP extracts "
                         "or external timesheets) and the output budgeting file.\n\n"
                         "2: Select the employees of the project.\n\n"
                         "3: Review the process, checking/deleting the log file or "
                         "restoring the budgeting file to its previous state before "
                         "the writing of data.\n\n"
                         "NOTE: Close all Excel files which will be used in this process!")
        intro_message.setWordWrap(True)
        intro_message.setAlignment(Qt.AlignmentFlag.AlignLeft)

        ## Layout
        main_layout.addWidget(watermark_label)
        main_layout.addWidget(intro_message)
        main_layout.addStretch()
        self.setLayout(main_layout)

    def set_watermark(self, watermark:QLabel):
        '''Set watermark.'''
        watermark_file = f"{IMAGES_DIR}\\budget_watermark.jpg"
        watermark_pixmap = QPixmap(watermark_file)
        if watermark_pixmap.isNull():
            print("Failed to load image.")
        watermark.setPixmap(watermark_pixmap)
        watermark.setScaledContents(False)
