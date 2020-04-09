from back.mergers.excel.excel_merger import ExcelMerger
from back.mergers.merger import Merger
from util.logger import logger


def excel_merger() -> ExcelMerger:
    return ExcelMerger()


MERGERS = {
    'EXCEL': excel_merger
}


class Selector:
    EXCEL = 'EXCEL'
    _caller = 'Selector'

    def __init__(self, value: str):
        self.value = value

    def select(self) -> Merger:
        switch_func = MERGERS.get(self.value, lambda: logger.write_error(self._caller,
                                                                         f'Invalid Selector value: {self.value}'))
        return switch_func()
