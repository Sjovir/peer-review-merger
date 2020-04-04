import os
import traceback

from util.logger import logger
from front.visualizer import Visualizer
from back.mergers.excel.cell_util import index_to_column
if __name__ == '__main__':
    print('Started Script')
    logger.set_output_folder(os.path.realpath(__file__))

    try:
        visualizer = Visualizer()
        visualizer.start()
    except Exception as error:
        print(traceback.format_exc())
        logger.write_error('Main.py', 'A vital Error has occurred:\n' + str(traceback.format_exc()))
    finally:
        logger.print_log()
    print('Finished Script')
