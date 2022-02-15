# -*- coding: utf-8 -*-

import re
from xlwings.conversion import Converter

# Compile regular expressions early for performance.
re_number = re.compile(r'^-?\d+(?:\.\d+)?$')


class CellAddr:
    """Class for cell coordinate"""

    def __init__(self, row=1, col=1):
        self.row = row
        self.col = col

    def increase_row(self, row=1):
        self.row += row

    def increase_col(self, col=1):
        self.col += col

    def set(self, row, col):
        self.row = row
        self.col = col

    @property
    def addr(self):
        return self.row, self.col


class MyConverter(Converter):
    """custom xlwings converter

    Override the default behavior which reads ``int`` as ``float``.

    """

    @staticmethod
    def read_value(value, options):
        if type(value) is float and value == int(value):
            ret = int(value)
        else:
            ret = value

        return ret

    @staticmethod
    def write_value(value, options):
        def is_number_string(v):
            return bool(type(v) is str and re_number.match(v.strip()))

        if type(value) is str:
            value = "'" + value.strip()
        elif type(value) is list:
            # 'value' passed as a list in this program
            for i in range(len(value)):
                # number string, either int or float
                if is_number_string(value[i]):
                    value[i] = "'" + value[i].strip()

        return value


def list_strip(seq):
    """Remove None and '' on both sides

    Args:
        seq (list or tuple): list to remove None and ''

    Returns:
        list: trimmed list

    """
    if not type(seq) in [list, tuple]:
        raise TypeError

    if type(seq) is tuple:
        temp = list(seq)
    else:
        temp = seq[0:]  # copy

    while temp:  # forward
        if temp[0] is None or temp[0] == '':
            temp.pop(0)
        else:
            break

    while temp:  # backward
        if temp[-1] is None or temp[-1] == '':
            temp.pop()
        else:
            break

    return temp


# Copied from XlsxWriter by John McNamara, jmcnamara@cpan.org
def xl_col_to_name(col, col_abs=False):
    """
    Convert a zero indexed column cell reference to a string.

    Args:
       col:     The cell column. Int.
       col_abs: Optional flag to make the column absolute. Bool.

    Returns:
        Column style string.

    """
    col_num = col
    if col_num < 0:
        return None

    col_num += 1  # Change to 1-index.
    col_str = ''
    col_abs = '$' if col_abs else ''

    while col_num:
        # Set remainder from 1 .. 26
        remainder = col_num % 26

        if remainder == 0:
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord('A') + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col_num = int((col_num - 1) / 26)

    return col_abs + col_str


def xl_rowcol_to_cell(row, col, row_abs=False, col_abs=False):
    """
    Convert a zero indexed row and column cell reference to a A1 style string.

    Args:
       row:     The cell row.    Int.
       col:     The cell column. Int.
       row_abs: Optional flag to make the row absolute.    Bool.
       col_abs: Optional flag to make the column absolute. Bool.

    Returns:
        A1 style string.

    """
    if row < 0 or col < 0:
        return None

    row += 1  # Change to 1-index.
    row_abs = '$' if row_abs else ''

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs + str(row)



