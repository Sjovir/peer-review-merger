from back.mergers.merger import Merger
from back.mergers.view_bag import MergerViewBag
from back.selector import Selector
from util.logger import logger


class Executor:
    _caller = 'Executor'

    _merger: Merger = None

    def __init__(self, selector: Selector):
        self._selector = selector

    def execute(self, view_bag: MergerViewBag):
        if not view_bag.is_valid():
            logger.write_error(self._caller, f'Tried to execute merger, but view bag is invalid')
            return

        self._merger = self._selector.select()
        self._patch_view_bag(view_bag)

        self._merger.merge()
        logger.write_log(self._caller, f'Finished merging with merger of type: {self._merger.__class__.__name__}')

    def _patch_view_bag(self, view_bag: MergerViewBag):
        for attribute_name in view_bag.get_attribute_names():
            view_bag_value = view_bag.__getattribute__(attribute_name)
            # print(f'attr: {attribute_name}, value: {view_bag_value}')
            self._merger.__setattr__(attribute_name, view_bag_value)
