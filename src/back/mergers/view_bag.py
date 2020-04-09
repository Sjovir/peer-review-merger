from util.logger import logger


class MergerViewBag(object):
    _caller = 'MergerViewBag'

    file: str = None
    input_folder: str = None
    output_folder: str = None
    overview_sheet_name: str = None
    feedback_sheet_name: str = None
    num_tasks: int = None
    columns_per_task: int = None
    num_reviewers_per_task: int = None

    _items = [
        'file',
        'input_folder',
        'output_folder',
        'overview_sheet_name',
        'feedback_sheet_name',
        'num_tasks',
        'columns_per_task',
        'num_reviewers_per_task'
    ]

    def is_valid(self) -> bool:
        for attribute_name in self._items:
            attribute = getattr(self, attribute_name)
            if attribute is None:
                logger.write_error(self._caller, f'Attribute is not set({attribute_name})')
                return False

        return True

    def get_attribute_names(self) -> list:
        return self._items
