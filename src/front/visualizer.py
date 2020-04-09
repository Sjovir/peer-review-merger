from back.executor import Executor
from back.mergers.view_bag import MergerViewBag
from back.selector import Selector


class Visualizer:

    _view_bag: MergerViewBag = None

    def __init__(self):
        self._selector = Selector(Selector.EXCEL)
        self._executor = Executor(self._selector)
        pass

    def start(self):
        self._show_dialog()
        self._execute()

    def _show_dialog(self):
        # TODO: Show tkinter dialog and take input
        self._view_bag: MergerViewBag = MergerViewBag()
        self._view_bag.file = 'C:\\project-merge\\template.xlsx'
        self._view_bag.input_folder = 'C:\\project-merge\\resources'
        self._view_bag.output_folder = 'C:\\project-merge'
        self._view_bag.overview_sheet_name = 'Username'
        self._view_bag.feedback_sheet_name = 'Bedømmelse'
        self._view_bag.num_tasks = 5
        self._view_bag.columns_per_task = 6
        self._view_bag.num_reviewers_per_task = 2

    def _execute(self):
        self._executor.execute(self._view_bag)
