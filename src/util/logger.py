import os
import platform


class Logger:
    FILE_NAME = 'log.txt'
    HEADER = 'Log for Peer Review Merger'
    INFO_TITLE = 'Info:'
    ERROR_TITLE = 'Errors:'

    def __init__(self):
        self._info_log = []
        self._error_log = []
        self._output_folder = None

    def set_output_folder(self, output_folder):
        if os.path.isdir(output_folder):
            self._output_folder = output_folder
        else:
            if platform.system() == 'Windows':
                self._output_folder = 'C:'
            else:
                self._output_folder = '/'

    def write_log(self, caller, msg):
        log: Log = Log(caller, msg)
        print(f'INFO: {log}')
        self._info_log.append(log)

    def write_error(self, caller, msg):
        log: Log = Log(caller, msg)
        print(f'ERROR: {log}')
        self._error_log.append(log)

    def print_log(self):
        if self._output_folder is None:
            self.set_output_folder(None)

        log_path = f'{self._output_folder}\\{self.FILE_NAME}'
        with open(log_path, 'w') as logFile:
            if len(self._info_log) > 0:
                logFile.write(f'{self.INFO_TITLE}\n')
                for log in self._info_log:
                    logFile.write(f'{log.caller}: {log.msg}\n')
            if len(self._error_log) > 0:
                logFile.write(f'{self.ERROR_TITLE}\n')
                for log in self._error_log:
                    logFile.write(f'{log.caller}: {log.msg}\n')

        # for error_log in self._error_log:
        #     print(error_log)


class Log:

    def __init__(self, caller, msg):
        self.caller = caller
        self.msg = msg

    def __str__(self):
        return f'{self.caller}: {self.msg}'


logger = Logger()
