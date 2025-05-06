'''
Package
-------
PHB Wizard

Module Name
---------
Project Selection Page

Version
-------
Date-based Version: 20250304
Author: Karl Goran Antony Zuvela

Description
-----------
Constructs and manages the employee selection page.
'''

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWizardPage, QLabel, QTableWidget, QPushButton,
    QVBoxLayout, QTableWidgetItem
)
import phb_app.data.workbook_management as wm
import phb_app.utils.project_utils as pro
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

class ProjectSelectionPage(QWizardPage):
    '''Page for selecting the projects in which the hours were booked.'''
    def __init__(self, managed_workbooks: wm.WorkbookManager):
        super().__init__()

        self.setTitle(st.PROJECT_SELECTION_TITLE)
        self.setSubTitle(st.PROJECT_SELECTION_SUBTITLE)
        # Headers
        self.headers = ie.ProjectIDTableHeaders.cap_members_list()

        ## Init property
        # To be populated at page init
        self.managed_workbooks: wm.WorkbookManager

        self.setup_widgets()
        self.setup_layout()
        self.setup_connections()

    #######################################
    ### QWizard setup function override ###
    #######################################

    def initializePage(self) -> None:
        '''Retrieve fields from other pages.'''

        # Make sure the table headers are displayed
        # Get workbooks from IOSelection
        self.managed_workbooks = self.wizard().property(
            st.QPropName.MANAGED_WORKBOOKS)
        # Set the IDs of every project in each workbook
        self.set_each_workbooks_project_ids()
        ## Table data
        # Table headers
        self.projects_table.setHorizontalHeaderLabels(self.headers)
        # Add all data to the table
        self.populate_table(self.projects_table)

    #######################
    ### Setup functions ###
    #######################

    def setup_widgets(self) -> None:
        '''Init all widgets.'''

        ## Instructions
        # Selection
        self.project_instructions_label = QLabel(
            "<p><strong>Select one or more project IDs.<strong></p>"
            "<p>The project decription is purely to help in the recognition "
            "of the wished after project ID. It will not be used in any "
            "calculations.</p>"
        )
        self.project_instructions_label.setTextFormat(Qt.TextFormat.RichText)

        ## Projects table
        # Start with 0 rows and 3 columns
        self.projects_table = QTableWidget(0, len(ie.ProjectIDTableHeaders))
        # self.projects_table.resizerow
        # Allow multiple row selections in the table
        self.projects_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.projects_table.setSelectionMode(QTableWidget.SelectionMode.MultiSelection)
        # Adjust the width of each column
        self.projects_table.setColumnWidth(ie.ProjectIDTableHeaders.PROJECT_ID, 250)
        self.projects_table.setColumnWidth(ie.ProjectIDTableHeaders.DESCRIPTION, 250)
        self.projects_table.setColumnWidth(ie.ProjectIDTableHeaders.FILENAME, 400)

        ## Table selection buttons
        # Add deselect all button
        self.deselect_projects_button = QPushButton(st.ButtonNames.DESELECT_ALL, self)

    def setup_layout(self) -> None:
        '''Init layout.'''

        ## Layout types
        # Main layout
        main_layout = QVBoxLayout()

        ## Add widgets
        # To main layout
        main_layout.addWidget(self.project_instructions_label)
        main_layout.addWidget(self.projects_table)
        main_layout.addWidget(self.deselect_projects_button)

        ## Display
        self.setLayout(main_layout)

    def setup_connections(self) -> None:
        '''Connect buttons to their respective actions.'''

        ## Connect functions
        # Deselect all
        self.deselect_projects_button.clicked.connect(
            self.projects_table.clearSelection)
        # Check isComplete
        self.projects_table.itemSelectionChanged.connect(
            self.completeChanged.emit)

    ############################
    ### Supporting functions ###
    ############################

    def set_each_workbooks_project_ids(self) -> None:
        '''Set project IDs for each workbook.'''

        for wb in self.managed_workbooks.workbooks:
            if isinstance(wb, wm.ManagedInputWorkbook):
                wb.managed_sheet_object.set_selectable_project_ids(
                    wb.locale_data.filter_headers.proj_id,
                    wb.locale_data.filter_headers.description,
                    wb.locale_data.filter_headers.name)

    def populate_table(self, table: QTableWidget) -> None:
        '''Populate the table with selectable project IDs.'''

        for wb in self.managed_workbooks.workbooks:
            if isinstance(wb, wm.ManagedInputWorkbook):
                # Go through each selectable ID for each row in the table
                for row_position, item in enumerate(
                    wb.managed_sheet_object.yield_from_project_id_and_desc()
                ):
                    ## Populate row. No items editable after setting
                    # Add row
                    table.insertRow(row_position)
                    # Add project ID
                    proj_id_item = QTableWidgetItem(
                        item[ie.ProjectIDTableHeaders.PROJECT_ID]
                    )
                    proj_id_item.setFlags(proj_id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(
                        row_position,
                        ie.ProjectIDTableHeaders.PROJECT_ID,
                        proj_id_item
                    )
                    # Go through each description listed for the selectable ID
                    desc_text = "\n".join(
                        list_item for list_item in item[
                            ie.ProjectIDTableHeaders.DESCRIPTION]
                    )
                    # Add project description
                    proj_desc_item = QTableWidgetItem(desc_text)
                    proj_desc_item.setFlags(proj_desc_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(
                        row_position,
                        ie.ProjectIDTableHeaders.DESCRIPTION,
                        proj_desc_item
                    )
                    # Resize the row according to the number of rows
                    # in the project description cell
                    table.resizeRowToContents(row_position)
                    # Add the file name
                    file_name = QTableWidgetItem(wb.file_name)
                    table.setItem(
                        row_position,
                        ie.ProjectIDTableHeaders.FILENAME,
                        file_name
                    )

    ##################################
    ### QWizard function overrides ###
    ##################################

    def cleanupPage(self):
        '''Override the page cleanup.
        Clear the table if the back button is pressed.'''

        self.projects_table.clear()
        self.projects_table.setRowCount(0)

    def isComplete(self) -> bool:
        '''Override the page completion.
        Check if the table has at least one project ID selected.'''

        return pu.check_completion(self.projects_table)

    def validatePage(self) -> bool:
        '''Override the page validation.
        Set the property.'''

        # Get the selected rows
        selected_rows = self.projects_table.selectionModel().selectedRows()
        # Set the selected projects for the respective workbook
        for wb in self.managed_workbooks.yield_workbooks_by_type(wm.ManagedInputWorkbook):
            pro.set_selected_project_ids(
                wb,
                self.projects_table,
                selected_rows,
                ie.ProjectIDTableHeaders
            )
        # Validation complete
        return True
