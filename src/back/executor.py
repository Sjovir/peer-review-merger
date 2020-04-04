from util.logger import logger
from back.mergers.merger import Merger
from back.selector import Selector



class Executor:
    _caller = 'Executor.py'

    _merger: Merger = None

    def __init__(self, selector: Selector):
        self._selector = selector

    def execute(self, merger: Merger):
        self._merger = merger
        self._merger.merge()
        logger.write_log(self._caller, f'Finished merging with merger of type: {self._merger.__class__.__name__}')


