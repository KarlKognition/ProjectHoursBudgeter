'''
Package
-------
PHB Wizard

Module Name
---------
Page Utilities

Author
-------
Karl Goran Antony Zuvela

Description
-----------
All functions necessary for setting up the wizard pages.
'''
#           --- Standard libraries ---
from datetime import datetime
from typing import Optional, TYPE_CHECKING
from uuid import UUID
import zipfile
#           --- Third party libraries ---
from dateutil.relativedelta import relativedelta
from openpyxl.utils.exceptions import ReadOnlyWorkbookException, InvalidFileException
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QColor
from PyQt6.QtWidgets import (
    QWizardPage, QBoxLayout, QHBoxLayout, QComboBox,
    QWidget, QLabel, QFileDialog, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QHeaderView
)
#           --- First party libraries ---
import phb_app.data.header_management as hm
import phb_app.data.location_management as loc
import phb_app.data.months_dict as md
import phb_app.logging.error_manager as em
import phb_app.logging.exceptions as ex
import phb_app.utils.date_utils as du
import phb_app.utils.file_handling_utils as fu
import phb_app.wizard.constants.integer_enums as ie
import phb_app.wizard.constants.ui_strings as st

if TYPE_CHECKING:
    import phb_app.data.io_management as io
    import phb_app.data.workbook_management as wm

#           --- Title Setup ---

def set_titles(page: QWizardPage, title: str, subtitle: str) -> None:
    '''Set the title and subtitle for the page.'''
    page.setTitle(title)
    page.setSubTitle(subtitle)

def _set_watermark(watermark: QLabel, directory: str, error: str) -> None:
    '''Set watermark.'''
    watermark_file = f"{directory}\\budget_watermark.jpg"
    watermark_pixmap = QPixmap(watermark_file)
    if watermark_pixmap.isNull():
        watermark.setStyleSheet("color: red;")
        watermark.setPixmap(QPixmap())  # Clear the pixmap
        watermark.setText(error)
        return
    watermark.setPixmap(watermark_pixmap)
    watermark.setScaledContents(False)

def create_watermark_label(directory: str, error: str) -> QLabel:
    '''Create and return the watermark label.'''
    watermark_label = QLabel()
    _set_watermark(watermark_label, directory, error)
    return watermark_label

def create_intro_message(text: str) -> QLabel:
    '''Create and return the introduction message label.'''
    intro_message = QLabel(text)
    intro_message.setWordWrap(True)
    intro_message.setAlignment(Qt.AlignmentFlag.AlignLeft)
    return intro_message

def setup_page(page: QWizardPage, widgets: list[QWidget], layout_type: QBoxLayout, spacing: int = 6, margins: tuple[int, int, int, int] = (9, 9, 9, 9)) -> None:
    '''Set up the layout for a page with the given widgets and layout type.
    The default spacing and margins from Qt are set.'''
    layout_type.setSpacing(spacing)
    layout_type.setContentsMargins(*margins)
    for widget in widgets:
        layout_type.addWidget(widget)
    page.setLayout(layout_type)

#           --- Table and Panel Setup ---

def setup_error_panel(role: st.IORole) -> QWidget:
    '''Set up the error panel for the given role with vertical layout.'''
    em.error_panels[role] = QWidget()
    em.error_panels[role].setLayout(QVBoxLayout())
    return em.error_panels[role]

def create_row_header_table(
    page: QWizardPage,
    table_headers: ie.BaseTableHeaders,
    selection_mode: QTableWidget.SelectionMode
    ) -> QTableWidget:
    '''Create a table widget with the given headers and selection mode.'''
    table = QTableWidget(len(table_headers), 0)
    table.setVerticalHeaderLabels(table_headers.cap_members_list())
    table.verticalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
    table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    table.horizontalHeader().setVisible(False)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(selection_mode)
    table.selectionModel().selectionChanged.connect(lambda selected, deselected: page.completeChanged.emit())
    return table

def create_col_header_table(
    page: QWizardPage,
    table_headers: ie.BaseTableHeaders,
    selection_mode: QTableWidget.SelectionMode,
    tab_widths: hm.TableWidths,
    invisible_headers: Optional[list[ie.BaseTableHeaders]] = None
    ) -> QTableWidget:
    '''Create a table widget with the given headers and selection mode.'''
    table = QTableWidget(0, len(table_headers))
    table.setHorizontalHeaderLabels(table_headers.cap_members_list())
    if invisible_headers:
        for header in invisible_headers:
            table.setColumnHidden(header, True)
    table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setSelectionMode(selection_mode)
    table.selectionModel().selectionChanged.connect(lambda selected, deselected: page.completeChanged.emit())
    for header, width in tab_widths.items():
        table.setColumnWidth(header, width)
    return table

def create_interaction_panel(panel: "io.IOControls", max_height: Optional[int] = None) -> QWidget:
    '''Set up the table with the given buttons and column widths.'''
    container = QWidget()
    layout = QVBoxLayout()
    panel.label.setTextFormat(Qt.TextFormat.RichText)
    layout.addWidget(panel.label)
    layout.addWidget(panel.table)
    buttons_layout = QHBoxLayout()
    # Add all assigned buttons
    if panel.buttons:
        for button in panel.buttons:
            buttons_layout.addWidget(button)
    layout.addLayout(buttons_layout)
    if panel.error_panel:
        layout.addWidget(panel.error_panel)
    container.setLayout(layout)
    if max_height is not None:
        container.setMaximumHeight(max_height)
    return container

#           --- File Handling ---

def _setup_file_dialog(page: QWizardPage, file_mode: QFileDialog.FileMode) -> QFileDialog:
    '''Set up and return a file dialog.'''
    file_dialog = QFileDialog(page)
    file_dialog.setWindowTitle(st.ADD_FILE)
    file_dialog.setNameFilter(st.EXCEL_FILE)
    file_dialog.setFileMode(file_mode)
    file_dialog.setViewMode(QFileDialog.ViewMode.Detail)
    return file_dialog

def _add_file_dialog(page: QWizardPage, wb_mngr: "wm.WorkbookManager", file_mode: QFileDialog.FileMode, file_ctx: "io.EntryContext") -> None:
    '''Add files to either the input or output selection tables.'''
    file_dialog = _setup_file_dialog(page, file_mode)
    if file_dialog.exec():
        _populate_file_table(page, wb_mngr, file_dialog.selectedFiles(), file_ctx)

def _remove_selected_file(page: QWizardPage, wb_mngr: "wm.WorkbookManager", file_ctx: "io.EntryContext") -> None:
    '''Remove the currently selected file(s) from the table.'''
    # Selected rows
    selected_items = file_ctx.panel.table.selectionModel().selectedRows()
    # Search for selected rows in reverse so that the row numbering
    # doesn't change with each deletion.
    for row in sorted([item.row() for item in selected_items], reverse=True):
        uuid = UUID(file_ctx.panel.table.item(row, ie.InputTableHeaders.UNIQUE_ID).text())
        # Remove the workbook context
        wb_mngr.remove_wb_ctx_by_uuid(file_ctx.panel.role, uuid)
        # Remove selected rows
        file_ctx.panel.table.removeRow(row)
        # Remove any associated error
        em.remove_error(uuid, file_ctx.panel.role)
        # Update UI
        page.completeChanged.emit()

#           --- Buttons ---

def connect_buttons(page: QWizardPage, ent_ctx: "io.EntryContext", wb_mngr: Optional["wm.WorkbookManager"] = None) -> None:
    '''Connect buttons to their respective actions dynamically.'''
    action_dispatch = {
        st.ButtonNames.ADD: lambda: _add_file_dialog(page, wb_mngr, QFileDialog.FileMode.ExistingFiles, ent_ctx),
        st.ButtonNames.REMOVE: lambda: _remove_selected_file(page, wb_mngr, ent_ctx),
        st.ButtonNames.SELECT_ALL: ent_ctx.panel.table.selectAll,
        st.ButtonNames.DESELECT_ALL: ent_ctx.panel.table.clearSelection
    }
    # Iterate over buttons and connect them dynamically
    for button in ent_ctx.panel.buttons:
        if action := action_dispatch.get(button.text()):
            button.clicked.connect(action)

#           --- Input Table Population ---

def _check_file_validity(file_ctx: "io.EntryContext") -> None:
    '''Check if the file is valid and highlight errors if necessary.'''
    file_exists = fu.is_file_already_in_table(file_ctx.data.file_name, ie.InputTableHeaders.FILENAME, file_ctx.panel.table)
    if file_exists:
        _highlight_bad_item(file_ctx.data.table_items.file_name)
        raise ex.FileAlreadySelected(file_ctx.data.file_name)
    if file_ctx.panel.table == st.IORole.OUTPUT and file_ctx.panel.table.rowCount() >= 2:
        _highlight_bad_item(file_ctx.data.table_items.file_name)
        raise ex.TooManyOutputFilesSelected()

def _insert_row(panel: "io.IOControls") -> int:
    '''Insert a new row in the table.'''
    row = panel.table.rowCount()
    panel.table.insertRow(row)
    return row

def _insert_col(panel: "io.IOControls") -> int:
    '''Insert a new column in the table.'''
    col = panel.table.columnCount()
    panel.table.insertColumn(col)
    return col

def insert_row_data_widget(table: QTableWidget, widget_item: QTableWidgetItem, row: int, column: ie.BaseTableHeaders) -> None:
    '''Insert given widget item into the table.'''
    widget_item.setFlags(widget_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    table.setItem(row, column, widget_item)
    table.resizeRowToContents(row)

def insert_col_data_widget(table: QTableWidget, widget_item: QTableWidgetItem, row: ie.BaseTableHeaders, column: int) -> None:
    '''Insert given widget item into the table.'''
    widget_item.setFlags(widget_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    table.setItem(row, column, widget_item)
    table.resizeRowToContents(row)

def update_handlers_country_details(data: loc.CountryData, wb_ctx: "wm.InputWorkbookContext") -> None:
    '''Update country details of the input workbook context from the IO file context.'''
    # Get country name from file name
    country_name = fu.get_origin_from_file_name(wb_ctx.mngd_wb.file_name, data, st.CountriesEnum)
    # Set the local data of the workbook
    import phb_app.data.workbook_management as wm # pylint: disable=import-outside-toplevel
    wm.set_locale_data(wb_ctx, data, country_name)

#           --- Output Table Population ---

def _create_dropdown(items: list[str], default: str) -> QComboBox:
    '''Create a dropdown with the given items and default selection.'''
    dropdown = QComboBox()
    dropdown.addItem(default)
    dropdown.addItems(items)
    dropdown.setCurrentIndex(0)
    return dropdown

def create_worksheet_dropdown(workbook_entry: "wm.OutputWorkbookContext") -> QComboBox:
    '''Create a dropdown for selecting a worksheet.'''
    return _create_dropdown(workbook_entry.managed_sheet.sheet_names, st.SpecialStrings.SELECT_WORKSHEET)

def create_year_dropdown() -> QComboBox:
    '''
    Create a dropdown for selecting a year. The default year ist he current year unless it is December,
    then due to the default month being last month, the default year is set to the previous year.
    '''
    default_year = str((datetime.now() + relativedelta(months=-ie.CONST_1)).year)
    return _create_dropdown([str(year) for year in range(2020, datetime.now().year + 3)], default_year)

def create_month_dropdown() -> QComboBox:
    '''
    Create a dropdown for selecting a month. The default month is set to the previous month.
    This is the month in which hours are yet to be analysed.
    '''
    default_month = du.abbr_month((datetime.now() + relativedelta(months=-ie.CONST_1)).month, md.LOCALIZED_MONTHS_SHORT)
    return _create_dropdown(list(md.LOCALIZED_MONTHS_SHORT.keys()), default_month)

def setup_dropdowns(table:QTableWidget, row: int, dds: "io.Dropdowns") -> None:
    '''Set up year, month, and worksheet dropdowns.'''
    table.setCellWidget(row, ie.OutputTableHeaders.WORKSHEET, dds.worksheet)
    table.setCellWidget(row, ie.OutputTableHeaders.MONTH, dds.month)
    table.setCellWidget(row, ie.OutputTableHeaders.YEAR, dds.year)

def _handle_load_file_error(row: int, file_ctx: "io.EntryContext", error: Exception) -> None:
    '''Handle errors during file selection and update the UI accordingly.'''
    _highlight_bad_item(file_ctx.panel.table.item(row, ie.OutputTableHeaders.FILENAME))
    em.add_error(file_ctx.data.uuid, file_ctx.panel.role, error)

def _handle_selection_error(row: int, file_ctx: "io.EntryContext", error: Exception) -> None:
    '''Handle errors during selection and update the UI accordingly.'''
    col: dict[st.IORole, ie.InputTableHeaders] = {
        st.IORole.INPUTS: ie.InputTableHeaders.FILENAME,
        st.IORole.OUTPUT: ie.OutputTableHeaders.FILENAME
    }
    _highlight_bad_item(file_ctx.panel.table.item(row, col[file_ctx.panel.role]))
    em.add_error(file_ctx.data.uuid, file_ctx.panel.role, error)

def handle_dropdown_selection(file_ctx: "io.EntryContext", wb_ctx: "wm.OutputWorkbookContext", row: int, dropdowns: "io.Dropdowns") -> None:
    '''Handle dropdown selection and update the selected sheet accordingly.'''
    # Remove the error if it is still there (if retrying)
    _clear_row_error_status(file_ctx, row, ie.OutputTableHeaders.FILENAME)
    import phb_app.data.io_management as io # pylint: disable=import-outside-toplevel
    io.update_current_text(dropdowns)
    try:
        wb_ctx.worksheet_service.set_selected_sheet(wb_ctx, dropdowns.current_text.worksheet)
        wb_ctx.worksheet_service.compute_employee_range()
        du.set_budgeting_date(wb_ctx, dropdowns.current_text)
    except (ex.EmployeeRowAnchorsMisalignment, ex.MissingEmployeeRow, ex.BudgetingDatesNotFound,
            KeyError) as exc:
        _handle_selection_error(row, file_ctx, exc)

#           --- Table Population ---

def _populate_file_table(page: QWizardPage, wb_mngr: "wm.WorkbookManager", files: list[str], file_ctx: "io.EntryContext") -> None:
    '''Populate the table with selected file(s).'''
    import phb_app.data.workbook_management as wm # pylint: disable=import-outside-toplevel
    for file_path in files:
        try:
            row = _insert_row(file_ctx.panel)
            wb_ctx = wm.create_wb_context_by_role(file_path, file_ctx.panel.role)
            wb_mngr.add_workbook(file_ctx.panel.role, wb_ctx)
            file_ctx.data.file_name = wb_ctx.mngd_wb.file_name
            file_ctx.data.table_items.file_name = QTableWidgetItem(file_ctx.data.file_name)
            insert_row_data_widget(file_ctx.panel.table, file_ctx.data.table_items.file_name, row, ie.InputTableHeaders.FILENAME)
            file_ctx.data.uuid = wb_ctx.mngd_wb.uuid
            file_ctx.configure_row(file_ctx, row, wb_ctx)
            _check_file_validity(file_ctx)
        except (ex.FileAlreadySelected, ex.TooManyOutputFilesSelected, ex.CountryIdentifiersNotInFilename,
                ex.IncorrectWorksheetSelected, ex.BudgetingDatesNotFound, ex.WorkbookAlreadyTracked,
                KeyError, ValueError) as exc:
            _handle_selection_error(row, file_ctx, exc)
        except (ReadOnlyWorkbookException, InvalidFileException, ex.WorkbookLoadError,
                FileNotFoundError, PermissionError, zipfile.BadZipFile,
                ) as exc:
            file_ctx.data.file_name = wm.get_file_name_from_path(file_path)
            file_ctx.data.uuid = wm.set_uuid()
            import phb_app.data.io_management as io # pylint: disable=import-outside-toplevel
            io.configure_error_row(file_ctx, row)
            _handle_load_file_error(row, file_ctx, exc)
    page.completeChanged.emit()

def populate_project_table(page: QWizardPage, proj_ctx: "io.EntryContext", wb_mngr: "wm.WorkbookManager") -> None:
    '''Populate the selection table with projects, employees or summary,'''
    for in_wb_ctx in wb_mngr.yield_workbook_ctxs_by_role(st.IORole.INPUTS): # We only need input workbooks here
        for proj_ids in in_wb_ctx.worksheet_service.yield_project_id_and_desc():
            row = _insert_row(proj_ctx.panel)
            import phb_app.data.io_management as io # pylint: disable=import-outside-toplevel
            proj_ctx.data.project_id = proj_ids[ie.ProjectIDTableHeaders.PROJECT_ID]
            proj_ctx.data.project_identifiers = io.join_str_list('\n', proj_ids[ie.ProjectIDTableHeaders.DESCRIPTION])
            proj_ctx.data.file_name = in_wb_ctx.mngd_wb.file_name
            proj_ctx.configure_row(proj_ctx, row)
    page.completeChanged.emit()

def populate_employee_table(page: QWizardPage, emp_ctx: "io.EntryContext", wb_mngr: "wm.WorkbookManager") -> None:
    '''Populate the employee table with employees.'''
    wb_ctx = wb_mngr.get_output_workbook_ctx()
    ws_ctx = wb_ctx.managed_sheet.selected_sheet.sheet_object
    start_col = wb_ctx.managed_sheet.employee_range.start_col_idx
    end_col = wb_ctx.managed_sheet.employee_range.end_col_idx + ie.CONST_1
    for col in range(start_col, end_col):
        cell = ws_ctx.cell(row=wb_ctx.managed_sheet.employee_range.start_row_idx, column=col)
        if cell.value and cell.value not in st.NON_NAMES:
            row = _insert_row(emp_ctx.panel)
            emp_ctx.data.emp_name = cell.value
            emp_ctx.data.worksheet = wb_ctx.managed_sheet.selected_sheet.sheet_name
            emp_ctx.data.coord = cell.coordinate
            emp_ctx.configure_row(emp_ctx, row)
    page.completeChanged.emit()

def populate_io_summary_table(page: QWizardPage, sum_io_ctx: "io.EntryContext", wb_mngr: "wm.WorkbookManager") -> None:
    '''Populate the summary table with IO data.'''
    col = _insert_col(sum_io_ctx.panel)
    in_wb_names = wb_mngr.get_wb_names_list_by_role(st.IORole.INPUTS)
    import phb_app.data.io_management as io # pylint: disable=import-outside-toplevel
    sum_io_ctx.data.in_file_names = io.join_str_list('\n', in_wb_names)
    out_wb_names = wb_mngr.get_wb_names_list_by_role(st.IORole.OUTPUT)
    sum_io_ctx.data.out_file_names = io.join_str_list('\n', out_wb_names)
    date = wb_mngr.get_output_workbook_ctx().managed_sheet.selected_date
    month = du.abbr_month(date.month, md.LOCALIZED_MONTHS_SHORT)
    sum_io_ctx.data.date = f"{month} {date.year}"
    sum_io_ctx.configure_row(sum_io_ctx, col)
    page.completeChanged.emit()
    

def populate_summary_data_table(page: QWizardPage, sum_data_ctx: "io.EntryContext", wb_mngr: "wm.WorkbookManager") -> None:
    '''Populate the summary data table with the collected data.'''
    out_wb_ctx = wb_mngr.get_output_workbook_ctx()
    for emp in out_wb_ctx.worksheet_service.yield_from_selected_employees():
        row = _insert_row(sum_data_ctx.panel)
        sum_data_ctx.data.emp_name = emp.name
        sum_data_ctx.data.pred_hrs = emp.hours.predicted_hours
        sum_data_ctx.data.acc_hrs = emp.hours.accumulated_hours
        sum_data_ctx.data.dev = emp.hours.deviation
        import phb_app.data.io_management as io # pylint: disable=import-outside-toplevel
        sum_data_ctx.data.proj_id = io.join_found_projects('\n', emp.found_projects.items())
        sum_data_ctx.data.out_ws_name = out_wb_ctx.managed_sheet.selected_sheet.sheet_name
        sum_data_ctx.data.coord = emp.hours.hours_coord
        sum_data_ctx.configure_row(sum_data_ctx, row, emp.hours)
    page.completeChanged.emit()


#           --- Exception Styling and Handling ---

def _highlight_bad_item(item: QTableWidgetItem) -> None:
    '''Highlight table items causing errors.'''

    item.setBackground(QColor(Qt.GlobalColor.red))
    item.setForeground(QColor(Qt.GlobalColor.white))

def _remove_highlighting(item: QTableWidgetItem) -> None:
    '''Removes the highlighting after error correction.'''

    item.setBackground(st.DEFAULT_BACKGROUND_COLOUR)
    item.setForeground(st.DEFAULT_FONT_COLOUR)

def _clear_row_error_status(file_ctx: "io.EntryContext", row: int, header: ie.BaseTableHeaders) -> None:
    '''Clear the error status of a row in the table.'''
    em.remove_error(file_ctx.data.uuid, file_ctx.panel.role)
    _remove_highlighting(file_ctx.panel.table.item(row, header))

#           --- QWizard function overrides ---

def get_combo_box(panel: "io.IOControls", row: int, col: int) -> Optional[QComboBox]:
    '''Get the combo box from the table.'''
    return panel.table.cellWidget(row, col) if panel.table.cellWidget(row, col) else None

def clean_up_table(table: QTableWidget, clear_cols: bool = False) -> None:
    '''Clean up the table.'''
    table.clearContents()
    if clear_cols:
        table.setColumnCount(0)
    else:
        table.setRowCount(0)

def _in_selection_complete(panel: "io.IOControls") -> bool:
    '''Check if both tables have at least one row selected
    and no error messages are displayed. Errors are added as QLabel widgets. No QLable, no error.'''
    return (
        panel.table.rowCount() >= ie.CONST_1
        and not panel.error_panel.findChildren(QLabel)
    )

def _out_selection_complete(panel: "io.IOControls") -> bool:
    '''Check if both tables have at least one row selected
    and no error messages are displayed. Errors are added as QLabel widgets. No QLable, no error.'''
    out_item = get_combo_box(panel, ie.OutputFile.FIRST_ENTRY, ie.OutputTableHeaders.WORKSHEET)
    return (
        panel.table.rowCount() == ie.CONST_1
        and out_item is not None
        and out_item.currentText() != st.SpecialStrings.SELECT_WORKSHEET
        and not panel.error_panel.findChildren(QLabel)
    )

def _table_selection_complete(panel: "io.IOControls") -> bool:
    '''Check if the table has at least one row selected. There is only one panel in this case.'''
    return bool(panel.table.selectionModel().selectedRows())

def _io_summary_table_complete(panel: "io.IOControls") -> bool:
    '''Check if the table has the minimum number of rows.'''
    return panel.table.rowCount() >= ie.IO_SUMMARY_ROW_COUNT

def check_completion(panel: "io.IOControls") -> bool:
    '''Return the appropriate isComplete function based on the page type.'''
    page_checker = {
        st.IORole.INPUTS: _in_selection_complete,
        st.IORole.OUTPUT: _out_selection_complete,
        st.IORole.PROJECT_TABLE: _table_selection_complete,
        st.IORole.EMPLOYEE_TABLE: _table_selection_complete,
        st.IORole.SUMMARY_IO_TABLE: _io_summary_table_complete,
        st.IORole.SUMMARY_DATA_TABLE: _table_selection_complete
    }
    return page_checker.get(panel.role, False)(panel)
