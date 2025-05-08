# pylint: disable=missing-docstring
# pylint: disable=line-too-long
# pylint: disable=too-few-public-methods
'''
Module Name
---------
Project Hours Budgeting Protocols

Version
-------
Date-based Version: 202502010
Author: Karl Goran Antony Zuvela

Description
-----------
Provides typesetting signatures.
'''

from typing import Protocol, runtime_checkable

@runtime_checkable
class ConfigureRow(Protocol):
    '''Configure row signature'''
    def __call__(self, row_position: int) -> None: ...
