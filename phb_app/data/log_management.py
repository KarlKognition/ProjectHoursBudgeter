from dataclasses import dataclass
from datetime import datetime

@dataclass
class FileMetaData:
    '''Encapsulates all necessary log data.'''
    log_file_path: str
    selected_date: datetime
    input_workbooks: list[str]
    output_file_name: str
    output_worksheet_name: str

@dataclass
class TableStructure:
    '''Encapsulates table-related metadata.'''
    headers: list[str]
    tab_widths: list[int]
