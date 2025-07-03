"""Generic types"""
from PyQt6.QtWidgets import QPushButton

type ButtonsList = list[QPushButton]
type StrList = list[str]
type CountryName = str
type ProjectsDict = dict[str, list[str]]
type ProjectId = str | int
type ProjectsTup = tuple[str, list[str]]
