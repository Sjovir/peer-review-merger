from typing import List, Dict

from back.mergers.excel.cell_util import Cell


class Student:

    def __init__(self, number: int, username: str):
        self.number: int = number
        self.username: str = username
        self.students_to_review: Dict = {}
        self.feedback: List = []
        self.missing_review: List = []
        self.violating_cells: List = []

    def add_students_to_review(self, task_index: int, students: List):
        self.students_to_review[task_index] = students

    def add_feedback_from_students(self, task_index: int, feedback: List):
        self.feedback.append({'task_index': task_index, 'feedback': feedback})

    def add_cell_with_violating_input(self, cell: Cell):
        self.violating_cells.append(cell)

    def add_missing_review(self, task_index: int):
        self.missing_review.append({task_index: task_index})