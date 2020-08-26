import os

from back.executor import Executor
from back.mergers.view_bag import MergerViewBag
from back.selector import Selector

from tkinter import filedialog, font, StringVar, IntVar
import tkinter as tk

from util.logger import logger

WINDOW_WIDTH = 1000
WINDOW_HEIGHT = 700

TEMPLATE_FILE_FRAME_TEXT = 'Template file'
TEMPLATE_FILE_LABEL_DEFAULT_TEXT = 'No file selected'
TEMPLATE_FILE_BUTTON_TEXT = 'Select File'
TEMPLATE_FILE_SELECT_DIRECTORY_DIALOG_TITLE = 'Choose Excel template file'

INPUT_FOLDER_FRAME_TEXT = 'Input Folder'
INPUT_FOLDER_LABEL_DEFAULT_TEXT = os.getcwd()
INPUT_FOLDER_BUTTON_TEXT = 'Select Folder'
INPUT_FOLDER_SELECT_DIRECTORY_DIALOG_TITLE = 'Choose folder with Excel files'

OUTPUT_FOLDER_FRAME_TEXT = 'Output Folder'
OUTPUT_FOLDER_LABEL_DEFAULT_TEXT = os.getcwd()
OUTPUT_FOLDER_BUTTON_TEXT = 'Select Folder'
OUTPUT_FOLDER_SELECT_DIRECTORY_DIALOG_TITLE = 'Choose folder for merged Excel file'

SETTINGS_FRAME_TEXT = 'Settings'
SETTINGS_OVERVIEW_SHEET_LABEL_TEXT = 'Overview sheet name'
SETTINGS_OVERVIEW_SHEET_DEFAULT_VALUE = 'Username'
SETTINGS_NUM_TASKS_LABEL_TEXT = 'Number of tasks'
SETTINGS_NUM_TASKS_DEFAULT_VALUE = 5
SETTINGS_NUM_REVIEWERS_PER_TASK_LABEL_TEXT = 'Number of reviewers per task'
SETTINGS_NUM_REVIEWERS_PER_TASK_DEFAULT_VALUE = 2
SETTINGS_FEEDBACK_SHEET_LABEL_TEXT = 'Feedback sheet name'
SETTINGS_FEEDBACK_SHEET_DEFAULT_VALUE = 'Bedømmelse'
SETTINGS_COLUMNS_PER_TASK_LABEL_TEXT = 'Columns per task'
SETTINGS_COLUMNS_PER_TASK_DEFAULT_VALUE = 6

MERGE_BUTTON_TEXT = 'Merge'

FRAME_FONT = None
LABEL_FONT = None
BUTTON_FONT = None


# Helper methods
def setup_fonts():
    global FRAME_FONT, LABEL_FONT, BUTTON_FONT
    FRAME_FONT = font.Font(size=15, family='Calibri')
    LABEL_FONT = font.Font(size=15, family='Calibri')
    BUTTON_FONT = font.Font(size=15, family='Calibri')


def setup_window(window: tk.Tk):
    window.update_idletasks()
    width = WINDOW_WIDTH
    height = WINDOW_HEIGHT
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))


def add_separator(window: tk.Tk):
    separator = tk.Frame(window, height=2, bd=1)
    separator.pack(fill='x', padx=5, pady=5)


def add_select_frame(window: tk.Tk, title: str, reference_variable: StringVar, button_text: str, button_cmd):
    frame = tk.LabelFrame(window, text=title, font=FRAME_FONT)
    frame.pack(fill='x', padx=5)
    label = tk.Label(frame, textvariable=reference_variable)
    label.config(font=LABEL_FONT)
    label.pack(side='left')
    button = tk.Button(frame, text=button_text, command=lambda: button_cmd())
    button['font'] = BUTTON_FONT
    button.pack(padx=5, pady=10, side='right')


def add_setting_frame(parent_frame: tk.LabelFrame, label_text: str, entry_variable):
    frame = tk.Frame(parent_frame)
    frame.pack(pady=10)
    label = tk.Label(frame, text=label_text, width=30, anchor='w')
    label.config(font=LABEL_FONT)
    label.pack(side='left')
    entry = tk.Entry(frame, textvariable=entry_variable)
    entry.config(font=LABEL_FONT)
    entry.pack(side='left')


def select_file(title: str, file_variable: StringVar):
    initial_folder = os.path.dirname(file_variable.get()) if file_variable.get().endswith('.xlsx') else os.getcwd()

    selected_file = filedialog.askopenfilename(title=title,
                                               initialdir=initial_folder,
                                               filetypes=[('Excel files', '*.xlsx')])

    if selected_file:
        file_variable.set(selected_file)

    return selected_file


def select_folder(title: str, folder_variable: StringVar):
    selected_folder = filedialog.askdirectory(title=title, initialdir=folder_variable.get())

    if selected_folder:
        folder_variable.set(selected_folder)

    return selected_folder


class Visualizer:
    _caller = 'Visualizer'

    _view_bag: MergerViewBag = None

    _template_file: StringVar = None
    _input_folder: StringVar = None
    _output_folder: StringVar = None

    _overview_sheet_name: StringVar = None
    _feedback_sheet_name: StringVar = None
    _num_tasks: IntVar = None
    _columns_per_task: IntVar = None
    _num_reviewers_per_task: IntVar = None

    _merge_button: tk.Button = None

    _merge: bool = False

    def __init__(self):
        self._selector = Selector(Selector.EXCEL)
        self._executor = Executor(self._selector)

    def start(self):
        self._show_dialog()
        if self._merge:
            self._execute()

    def _show_dialog(self):
        window = tk.Tk()

        setup_fonts()

        setup_window(window)

        add_separator(window)
        self._setup_template_frame(window)
        add_separator(window)
        self._setup_input_frame(window)
        add_separator(window)
        self._setup_output_frame(window)
        add_separator(window)
        self._setup_settings_frame(window)
        add_separator(window)
        self._setup_merge_frame(window)
        add_separator(window)

        self._validate_input()

        window.mainloop()

    def _execute(self):
        self._executor.execute(self._view_bag)

    def _setup_template_frame(self, window: tk.Tk):
        self._template_file = StringVar(value=TEMPLATE_FILE_LABEL_DEFAULT_TEXT)
        add_select_frame(window, TEMPLATE_FILE_FRAME_TEXT, self._template_file, TEMPLATE_FILE_BUTTON_TEXT,
                         self._select_template_file)

    def _setup_input_frame(self, window: tk.Tk):
        self._input_folder = StringVar(value=INPUT_FOLDER_LABEL_DEFAULT_TEXT)
        add_select_frame(window, INPUT_FOLDER_FRAME_TEXT, self._input_folder, INPUT_FOLDER_BUTTON_TEXT,
                         self._select_input_folder)

    def _setup_output_frame(self, window: tk.Tk):
        self._output_folder = StringVar(value=OUTPUT_FOLDER_LABEL_DEFAULT_TEXT)
        add_select_frame(window, OUTPUT_FOLDER_FRAME_TEXT, self._output_folder, OUTPUT_FOLDER_BUTTON_TEXT,
                         self._select_output_folder)

    def _setup_settings_frame(self, window: tk.Tk):
        settings_frame = tk.LabelFrame(window, text=SETTINGS_FRAME_TEXT, font=FRAME_FONT)
        settings_frame.pack(fill='x', padx=5)

        self._overview_sheet_name = StringVar(value=SETTINGS_OVERVIEW_SHEET_DEFAULT_VALUE)
        add_setting_frame(settings_frame, SETTINGS_OVERVIEW_SHEET_LABEL_TEXT, self._overview_sheet_name)

        self._num_tasks = IntVar(value=SETTINGS_NUM_TASKS_DEFAULT_VALUE)
        add_setting_frame(settings_frame, SETTINGS_NUM_TASKS_LABEL_TEXT, self._num_tasks)

        self._num_reviewers_per_task = IntVar(value=SETTINGS_NUM_REVIEWERS_PER_TASK_DEFAULT_VALUE)
        add_setting_frame(settings_frame, SETTINGS_NUM_REVIEWERS_PER_TASK_LABEL_TEXT, self._num_reviewers_per_task)

        self._feedback_sheet_name = StringVar(value=SETTINGS_FEEDBACK_SHEET_DEFAULT_VALUE)
        add_setting_frame(settings_frame, SETTINGS_FEEDBACK_SHEET_LABEL_TEXT, self._feedback_sheet_name)

        self._columns_per_task = IntVar(value=SETTINGS_COLUMNS_PER_TASK_DEFAULT_VALUE)
        add_setting_frame(settings_frame, SETTINGS_COLUMNS_PER_TASK_LABEL_TEXT, self._columns_per_task)

    def _setup_merge_frame(self, window: tk.Tk):
        bottom_frame = tk.Frame(window)
        bottom_frame.pack(expand='yes')
        self._merge_button = tk.Button(bottom_frame,
                                       text=MERGE_BUTTON_TEXT,
                                       command=lambda: self._save_view_bag(window),
                                       width=30,
                                       height=2)
        self._merge_button['font'] = BUTTON_FONT
        self._merge_button.pack()

    ##############################
    # Class level helper methods #
    ##############################
    def _select_template_file(self):
        template_file: str = select_file(TEMPLATE_FILE_SELECT_DIRECTORY_DIALOG_TITLE, self._template_file)
        if template_file:
            logger.write_log(self._caller, f'Selected template file: {template_file}')
            self._validate_input()

    def _select_input_folder(self):
        input_folder: str = select_folder(INPUT_FOLDER_SELECT_DIRECTORY_DIALOG_TITLE, self._input_folder)
        if input_folder:
            logger.write_log(self._caller, f'Selected input folder: {input_folder}')
            self._validate_input()

    def _select_output_folder(self):
        output_folder: str = select_folder(OUTPUT_FOLDER_SELECT_DIRECTORY_DIALOG_TITLE, self._output_folder)
        if output_folder:
            logger.write_log(self._caller, f'Selected output folder: {output_folder}')
            self._validate_input()
            logger.set_output_folder(output_folder)

    def _save_view_bag(self, window: tk.Tk):
        self._view_bag: MergerViewBag = MergerViewBag()

        self._view_bag.file = self._template_file.get()
        self._view_bag.input_folder = self._input_folder.get()
        self._view_bag.output_folder = self._output_folder.get()
        self._view_bag.overview_sheet_name = self._overview_sheet_name.get()
        self._view_bag.feedback_sheet_name = self._feedback_sheet_name.get()
        self._view_bag.num_tasks = self._num_tasks.get()
        self._view_bag.columns_per_task = self._columns_per_task.get()
        self._view_bag.num_reviewers_per_task = self._num_reviewers_per_task.get()

        self._merge = True
        window.quit()

    def _validate_input(self):
        if self._template_file.get().endswith('.xlsx'):
            self._merge_button['state'] = 'normal'
        else:
            self._merge_button['state'] = 'disabled'
