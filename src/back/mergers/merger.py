from util.logger import logger


class Merger:
    _caller = 'Merger'
    file: str = None
    input_folder: str = None
    output_folder: str = None
    overview_sheet_name: str = None
    feedback_sheet_name: str = None
    num_tasks: int = None
    columns_per_task: int = None
    num_reviewers_per_task: int = None

    def __init__(self):
        pass

    def merge(self):
        logger.write_error(self._caller, f'Merger missing implemented: merge')
