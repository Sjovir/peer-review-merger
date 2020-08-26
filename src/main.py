import os
import traceback

from front.visualizer import Visualizer
from util.logger import logger

if __name__ == '__main__':
    print('Started Script')

    # logger.set_output_folder(os.path.dirname(os.path.realpath(__file__)))
    logger.set_output_folder(os.getcwd())

    try:
        visualizer = Visualizer()
        visualizer.start()
    except Exception as error:
        logger.write_error('Main.py', f'A vital Error has occurred:\n{traceback.format_exc()}')
    finally:
        logger.print_log()

    print('Finished Script')
