"""Generic types"""
#           --- Third party libraries ---
from PyQt6.QtWidgets import QPushButton

type CellCoord = str
type ButtonsList = list[QPushButton]
type StrList = list[str]
type CountryName = str
type ProjectsDict = dict[str, list[str]]
type ProjectId = str | int
type ProjectsTup = tuple[str, list[str]]
