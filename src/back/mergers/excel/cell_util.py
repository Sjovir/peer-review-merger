from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

ALPHABET_LENGTH: int = 26
CAPITAL_LETTERS_START_INDEX: int = 65
COLUMN_INDEX_OFFSET: int = 1


def search_cell_by_column_and_value(sheet: Worksheet, column: int, value: [str, int, None], row_limit: int) -> Cell:
    for row_index in range(1, row_limit):
        cell_name = index_to_cell_name(column, row_index)
        cell = sheet[cell_name]
        if value is cell.value:
            return cell


def index_to_cell_name(column: int, row: int) -> str:
    return f'{index_to_column(column)}{row}'


# TODO: Make this cool and recursive
def index_to_column(index: int) -> str:
    if index <= ALPHABET_LENGTH:
        return chr(CAPITAL_LETTERS_START_INDEX + index - COLUMN_INDEX_OFFSET)

    first_char = (index // ALPHABET_LENGTH)
    second_char = index % ALPHABET_LENGTH
    if index % ALPHABET_LENGTH == 0:
        first_char = first_char - 1
        second_char = 26
    return index_to_column(first_char) + index_to_column(second_char)


def cell_above(sheet: Worksheet, cell: Cell) -> Cell:
    return sheet.cell(column=cell.column, row=cell.row-1)


class CellBuilder:

    def __init__(self):
        pass

    def build_from_index(column: int, row: int) -> Cell:
        column = index_to_column(column)
        return Cell(column, row)
