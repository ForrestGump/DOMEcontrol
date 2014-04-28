# -*- coding: utf-8 -*-
"""
Created on Mon Apr 28 14:18:00 2014

@author: sborodin@rpatyphoon.ru
"""

import csv
import logging

logAAG = logging.getLogger('CloudWatcher')


def AAGCheck(str_targetFile):
    # проверка файла данных датчика облачности AAG Cloud Watcher

    # проверка - есть ли файл. Если да, то прочитать необходимое поле
    try:
        with open(str_targetFile) as f:
            file_csv = csv.DictReader(f)
            for row in file_csv:
                pass
            # end for
            SafeStatus = row['Safe Status']
        # end with
        # определение флага работы: можно или нет
        if SafeStatus == 'Safe':
            return True
        else:
            return False
        # end if
    except Exception as e:
        logAAG.exception(e)
    # end try
# end AAGCheck

if __name__ == '__main__':
    print AAGCheck("D:\Temp\CloudWatcher.csv")
