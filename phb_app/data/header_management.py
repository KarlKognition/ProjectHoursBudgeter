'''
Module Name
---------
Project Hours Budgeting Data Classes (phb_dataclasses)

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Provides several data classes for the Project Hours Budgeting Wizard.
'''

#           ---  First party library imports ---
import phb_app.wizard.constants.integer_enums as ie

type TableWidths = dict[ie.BaseTableHeaders, int]

INPUT_COLUMN_WIDTHS = {
    ie.InputTableHeaders.FILENAME: 250,
    ie.InputTableHeaders.COUNTRY: 150,
    ie.InputTableHeaders.WORKSHEET: 100

}
OUTPUT_COLUMN_WIDTHS = {
    ie.OutputTableHeaders.FILENAME: 250,
    ie.OutputTableHeaders.WORKSHEET: 150,
    ie.OutputTableHeaders.MONTH: 60,
    ie.OutputTableHeaders.YEAR: 60
}

PROJECT_COLUMN_WIDTHS = {
    ie.ProjectIDTableHeaders.PROJECT_ID: 250,
    ie.ProjectIDTableHeaders.DESCRIPTION: 250,
    ie.ProjectIDTableHeaders.FILENAME: 400
}

EMPLOYEE_COLUMN_WIDTHS = {
    ie.EmployeeTableHeaders.EMPLOYEE: 450,
    ie.EmployeeTableHeaders.WORKSHEET: 150,
    ie.EmployeeTableHeaders.COORDINATE: 80
}

SUMMARY_DATA_COLUMN_WIDTHS = {
    ie.SummaryDataTableHeaders.EMPLOYEE: 250,
    ie.SummaryDataTableHeaders.PREDICTED_HOURS: 100,
    ie.SummaryDataTableHeaders.ACCUMULATED_HOURS: 120,
    ie.SummaryDataTableHeaders.DEVIATION: 160,
    ie.SummaryDataTableHeaders.PROJECT_ID: 450,
    ie.SummaryDataTableHeaders.OUTPUT_WORKSHEET: 150,
    ie.SummaryDataTableHeaders.COORDINATE: 80
}

DEFAULT_PADDING = 5
