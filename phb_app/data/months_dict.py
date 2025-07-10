"""Localisation"""
#           --- Standard libraries ---
import locale
from datetime import date
from babel.dates import format_date

loc = locale.getdefaultlocale() # pylint: disable=deprecated-method
if loc:
    LOCALIZED_MONTHS_SHORT = {format_date(date(2025, month, 1), "MMM", locale=loc[0]).rstrip('.'): month for month in range(1, 13)}
else:
    MONTHS_SHORT_EN = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12
    }
