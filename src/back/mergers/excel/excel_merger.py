import os
from typing import List, Dict

from openpyxl import load_workbook, Workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from back.mergers.excel.cell_util import index_to_cell_name, search_cell_by_column_and_value, cell_above
from back.mergers.excel.student import Student
from back.mergers.merger import Merger
from util.logger import logger

OUTPUT_FILE_NAME = 'Merged.xlsx'
EXCEL_EXTENSION = '.xlsx'

FIRST_STUDENT_NUMBER = 1
FIRST_CELL_SEARCH_LIMIT = 50
LAST_CELL_SEARCH_LIMIT = 150

NUM_STUDENTS_TO_REVIEW = 2
STUDENT_NUMBER_COLUMN_INDEX = 1
STUDENT_USERNAME_COLUMN_INDEX = 2
STUDENT_FIRST_OVERVIEW_COLUMN = 3


# Helper methods
def _get_first_cell_of_sheet(sheet: Worksheet) -> Cell:
    return search_cell_by_column_and_value(sheet, STUDENT_NUMBER_COLUMN_INDEX, FIRST_STUDENT_NUMBER,
                                           FIRST_CELL_SEARCH_LIMIT)


def _get_last_cell_of_sheet(sheet: Worksheet) -> Cell:
    empty_cell = search_cell_by_column_and_value(sheet, STUDENT_NUMBER_COLUMN_INDEX, None, LAST_CELL_SEARCH_LIMIT)
    return cell_above(sheet, empty_cell)


def _is_feedback_list_empty(list: List) -> bool:
    return set(list) == [None]


class ExcelMerger(Merger):

    def __init__(self):
        super().__init__()
        self._students: Dict = {}
        self._wb: Workbook
        self._overview_sheet: Worksheet
        self._feedback_sheet: Worksheet
        self._first_overview_cell: Cell
        self._last_overview_cell: Cell
        self._first_feedback_cell: Cell

    def merge(self):
        self._load_document()
        self._load_students()
        self._load_feedback()
        self._merge_feedback()
        self._save_workbook()

    def _load_document(self):
        self._wb = load_workbook(self.file, data_only=True)
        self._overview_sheet = self._wb[self.overview_sheet_name]
        self._feedback_sheet = self._wb[self.feedback_sheet_name]
        self._first_overview_cell = _get_first_cell_of_sheet(self._overview_sheet)
        self._last_overview_cell = _get_last_cell_of_sheet(self._overview_sheet)
        self._first_feedback_cell = _get_first_cell_of_sheet(self._feedback_sheet)
        self._num_students = 1 + self._last_overview_cell.row - self._first_overview_cell.row
        logger.write_log(self._caller, f'Located {self._num_students} students in the overview sheet')

    def _load_students(self):
        for row_index in range(self._first_overview_cell.row, self._num_students + self._first_overview_cell.row):
            num_cell: Cell = self._overview_sheet[index_to_cell_name(STUDENT_NUMBER_COLUMN_INDEX, row_index)]
            username_cell: Cell = self._overview_sheet[index_to_cell_name(STUDENT_USERNAME_COLUMN_INDEX, row_index)]
            student = Student(num_cell.value, username_cell.value)

            for task_index in range(self.num_tasks):
                students_to_review: List = []
                for task_to_review in range(NUM_STUDENTS_TO_REVIEW):
                    column_index: int = STUDENT_FIRST_OVERVIEW_COLUMN + task_index * NUM_STUDENTS_TO_REVIEW \
                                        + task_to_review
                    student_cell: Cell = self._overview_sheet[index_to_cell_name(column_index, row_index)]
                    students_to_review.append(student_cell.value)
                student.add_students_to_review(task_index, students_to_review)

            self._students[num_cell.value] = student

    def _load_feedback(self):
        # TODO: Refactor names to reviewer and reviewee or the like
        for student_num in range(FIRST_STUDENT_NUMBER, self._num_students + 1):
            # Load workbook
            student: Student = self._students[student_num]
            student_file: Workbook
            student_feedback_sheet: Worksheet
            try:
                student_file = self._get_student_workbook(student.username)
                student_feedback_sheet = student_file[self.feedback_sheet_name]
            except FileNotFoundError as error:
                logger.write_log(self._caller, f'Student ({student.username}) did not submit a review')
                continue

            # Load feedback and insert into student objects
            for task_index in range(self.num_tasks):
                students_to_review: List = student.students_to_review[task_index]
                for student_to_review_index in range(len(students_to_review)):
                    student_to_review: Student = self._students[students_to_review[student_to_review_index]]
                    feedback: List = self._get_feedback(student_feedback_sheet, student_to_review.number, task_index)
                    if _is_feedback_list_empty(feedback):
                        student.add_missing_review(task_index)
                    else:
                        student_to_review.add_feedback_from_students(task_index, feedback)

    def _merge_feedback(self):
        for student_num in range(FIRST_STUDENT_NUMBER, self._num_students + 1):
            student: Student = self._students[student_num]
            self._write_student_feedback(student)

    def _save_workbook(self):
        merge_file_path = self.output_folder + "\\Merged.xlsx"
        try:
            self._wb.save(merge_file_path)
        except PermissionError as error:
            logger.write_error(self._caller, f'Cannot save merged file due to permission error. Is the file open?')
            raise

    # Class level helper methods
    def _get_student_workbook(self, username: str) -> Workbook:
        student_file_folder = f'{self.input_folder}\\{username}'
        student_file_path = ''
        for file in os.listdir(student_file_folder):
            if file.endswith(EXCEL_EXTENSION):
                student_file_path = f'{student_file_folder}\\{file}'

        try:
            student_file = load_workbook(student_file_path, data_only=True)
        except PermissionError as error:
            logger.write_error(self._caller, f'Cannot open student file ({username}) due to permission error. '
                                             f'Is the file open?')
            raise

        return student_file

    def _get_feedback(self, feedback_sheet: Worksheet, student_num: int, task_index: int) -> List:
        feedback_row: int = self._get_feedback_row(student_num)
        feedback_start_column: int = STUDENT_FIRST_OVERVIEW_COLUMN + task_index * self.columns_per_task

        feedback: List = []
        for cell_column in range(self.columns_per_task):
            feedback_column: int = feedback_start_column + cell_column
            feedback_cell: Cell = feedback_sheet[index_to_cell_name(feedback_column, feedback_row)]
            feedback.append(feedback_cell.value)

        return feedback

    def _get_feedback_row(self, student_num: int) -> int:
        return self._first_feedback_cell.row + student_num - 1

    def _write_student_feedback(self, student: Student):
        """
            Write the feedback of a student into the feedback sheet of the output file

        :param student: the student object which holds the feedback
        :return:
        """
        if student.number == 2:
            print()
        row: int = self._get_feedback_row(student.number)
        for feedback_item in student.feedback:
            task_index: int = feedback_item['task_index']
            task_feedback: List = feedback_item['feedback']
            self._write_task_feedback(row, task_index, task_feedback)

    def _write_task_feedback(self, row: int, index: int, feedback: List):
        """
            Write the feedback of a task starting from the start_column

        :param row: row in sheet of output file
        :param index: task index
        :param feedback: feedback to be written
        :return:
        """
        start_column: int = STUDENT_FIRST_OVERVIEW_COLUMN + index * self.num_reviewers_per_task * self.columns_per_task
        for feedback_index in range(self.columns_per_task):
            column: int = start_column + feedback_index
            # Move to second set of columns if there exists multiple sets of feedback for the student's task
            # TODO: Make this able to take more than two sets of feedback per task into account
            if self._is_feedback_for_task_present(column, row):
                column = column + self.columns_per_task

            feedback_cell: Cell = self._feedback_sheet[index_to_cell_name(column, row)]
            feedback_cell.value = feedback[feedback_index]

    # TODO: refactor name
    def _is_feedback_for_task_present(self, column, row):
        for feedback_index in range(self.columns_per_task):
            feedback_column: int = column + feedback_index
            feedback_cell: Cell = self._feedback_sheet[index_to_cell_name(feedback_column, row)]
            if feedback_cell.value is not None:
                return True

        return False
