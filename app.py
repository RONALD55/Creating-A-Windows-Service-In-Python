# -*- coding: utf-8 -*-
"""
Created on Sat Jan 22 18:40:13 2022

@author: Ronald Nyasha Kanyepi
@email : kanyepironald@gmail.com
"""

import logging
import os
import shutil
import time
import pandas as pd
from pyfiglet import Figlet

from SMWinservice import SMWinservice


def remove_folder_contents(path):
    shutil.rmtree(path)
    os.makedirs(path)


def join_path(folder_name, filename):
    return os.path.join(folder_name, os.path.basename(filename))


def copy_file(source, destination):
    for filename in os.listdir(source):
        try:
            if (os.path.splitext(filename))[1] == ".xlsx" or (os.path.splitext(filename))[1] == ".xls":
                logger.info("Converting %s to excel", filename)
                df = pd.read_excel(join_path(source, filename))
                df.to_csv(join_path(destination, (os.path.splitext(filename)[0]) + ".csv"), index=False)
                logger.info("Copying %s to destination", filename)
                os.remove(join_path(source, filename))

        except shutil.SameFileError as e:
            logger.error("Same file error has occurred : %s", e)

        except PermissionError as e:
            logger.error("Permission error %s", e)

        except OSError as e:
            logger.error("Error: %s : %s" % (filename, e.strerror))

        except Exception as e:
            logger.error("There was an exception : %s", e)


fig = Figlet(font='slant')
start_time = time.time()

print('Program now Executing   :  ')
print('-' * 100)
print(fig.renderText('WINDOWS SERVICE'))
print('-' * 100 + '\n')

# configuring the logs
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# define file handler and set formatter
file_handler = logging.FileHandler('file_path_for_logger')
formatter = logging.Formatter('%(asctime)s : %(levelname)s : %(name)s : %(message)s')
file_handler.setFormatter(formatter)

# add file handler to logger
logger.addHandler(file_handler)


class WindowsService(SMWinservice):
    _svc_name_ = "ServiceName"
    _svc_display_name_ = "ServiceName"
    _svc_description_ = "ServiceDescription"

    def start(self):
        self.isrunning = True

    def stop(self):
        self.isrunning = False

    def main(self):
        i = 0
        while self.isrunning:
            # folder path
            source = "Source_path"
            destination = "Destination_path"
            copy_file(source, destination)


if __name__ == '__main__':
    WindowsService.parse_command_line()
