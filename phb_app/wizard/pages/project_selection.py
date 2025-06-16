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
    QHBoxLayout, QTableWidgetItem
)
import phb_app.data.workbook_management as wm
import phb_app.utils.project_utils as pro
import phb_app.utils.page_utils as pu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.io_management as io
import phb_app.data.header_management as hm

class ProjectSelectionPage(QWizardPage):
    '''Page for selecting the projects in which the hours were booked.'''
    def __init__(self, managed_workbooks: wm.WorkbookManager):
        super().__init__()
        self.managed_workbooks = managed_workbooks
        self.project_panel = None
        pu.set_titles(self, st.PROJECT_SELECTION_TITLE, st.PROJECT_SELECTION_SUBTITLE)

    def initializePage(self) -> None:
        '''Retrieve fields from other pages.'''
        self.project_panel = io.IOControls(
            page=self,
            role=st.IORole.PROJECT_TABLE,
            label=QLabel(st.PROJECT_SELECTION_INSTRUCTIONS),
            table=pu.create_table(ie.ProjectIDTableHeaders, QTableWidget.SelectionMode.MultiSelection, hm.PROJECT_COLUMN_WIDTHS),
            buttons=[QPushButton(st.ButtonNames.DESELECT_ALL, self)]
        )
        project_handler = io.EntryContext(self.project_panel, self.managed_workbooks)
        pu.setup_page(self, [pu.create_interaction_panel(self.project_panel)], QHBoxLayout())
        pu.connect_buttons(self, project_handler)
        pu.populate_selection_table(self, project_handler, wm.ManagedInputWorkbook)

    ############################
    ### Supporting functions ###
    ############################

    def set_each_workbooks_project_ids(self) -> None:
        '''Set project IDs for each workbook.'''

        for wb in self.managed_workbooks.workbooks_ctxs:
            if isinstance(wb, wm.ManagedInputWorkbook):
                wb.managed_sheet_object.set_selectable_project_ids(
                    wb.locale_data.filter_headers.proj_id,
                    wb.locale_data.filter_headers.description,
                    wb.locale_data.filter_headers.name)

    def populate_table(self, table: QTableWidget) -> None:
        '''Populate the table with selectable project IDs.'''

        for wb in self.managed_workbooks.workbooks_ctxs:
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
