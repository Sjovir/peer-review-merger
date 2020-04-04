from back.executor import Executor
from back.selector import Selector


class Visualizer:

    def __init__(self):
        self._selector = Selector(Selector.EXCEL)
        self._executor = Executor(self._selector)
        pass

    def start(self):
        self._show_dialog()
        self._execute()

    def _show_dialog(self):
        # TODO: Show tkinter dialog and take input
        self._merger = self._selector.select()
        self._merger.file = 'C:\\project-merge\\template.xlsx'
        self._merger.input_folder = 'C:\\project-merge\\resources'
        self._merger.output_folder = 'C:\\project-merge'
        self._merger.overview_sheet_name = 'Username'
        self._merger.feedback_sheet_name = 'Bedømmelse'
        self._merger.num_tasks = 5
        self._merger.columns_per_task = 6
        self._merger.num_reviewers_per_task = 2

    def _execute(self):
        self._executor.execute(self._merger)
