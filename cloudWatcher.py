# -*- coding: utf-8 -*-
"""
Created on Mon Apr 28 14:18:00 2014

@author: sborodin@rpatyphoon.ru
"""

import csv
import re
import logging
import serial
from math import exp

logAAG = logging.getLogger('CloudWatcher')
patternAAG = re.compile('(\![\s\w]{2})([\s\w]{11,13})(\!.{12,15})')


def AAG_Connect(SerialDevice):
    # соединение с AAG Cloud Watcher по последовательному порту
    AAG = None
    AAG = serial.Serial(SerialDevice, 9600, timeout=2)
    return AAG
# end AAG_Connect


def AAG_Disconnect(AAG):
    # отключение соединения с AAG Cloud Watcher по последовательному порту
    AAG.close()
# end AAG_Disconnect


def AAG_CheckFile(str_targetFile):
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
            logAAG.info("AAG status is Safe")
            return True
        else:
            logAAG.warning("AAG status is UnSafe")
            return False
        # end if
    except Exception as e:
        logAAG.exception(e)
    # end try
# end AAG_CheckFile


def AAG_GetSwitch(AAG):
    # проверка флага состояния непосредственно с AAG Cloud Watcher
    # Created by Josh Walawender on 2012-02-11 (c)
    AAG.write("F!")
    response = AAG.read(30)
    IsResponse = patternAAG.match(response)
    if IsResponse and re.match("\!X", IsResponse.group(1)):
        logAAG.info("AAG status is Safe")
        return True
    elif IsResponse and re.match("\!Y", IsResponse.group(1)):
        logAAG.warning("AAG status is UnSafe")
        return False
    else:
        logAAG.warning("AAG not response")
        return None
    # end if
# end AAG_GetSwitch


def AAG_GetSkyTemp(AAG):
    # Get Sky Temperature
    # Created by Josh Walawender on 2012-02-11 (c)
    AAG.write("S!")
    response = AAG.read(30)
    IsResponse = patternAAG.match(response)
    if IsResponse and re.match("\!1", IsResponse.group(1)):
        return float(IsResponse.group(2)) / 100.
    else:
        logAAG.warning("AAG not response")
        return None
    # end if
# end AAG_GetSkyTemp


def AAG_GetAmbTemp(AAG):
    # Get Ambient Temperature
    # Created by Josh Walawender on 2012-02-11 (c)
    AAG.write("T!")
    response = AAG.read(30)
    IsResponse = patternAAG.match(response)
    if IsResponse and re.match("\!2", IsResponse.group(1)):
        return float(IsResponse.group(2)) / 100.
    else:
        logAAG.warning("AAG not response")
        return None
    # end if
# end AAG_GetAmbTemp


def AAG_SkyTempCorrection(AmbTempC, SkyTempC):
    # корректировка температуры согласно аппроксимационной кривой
    K = (33., 0., 4., 100., 100.)
    CorrectionTempC = (K[0] / 100.) * (AmbTempC - K[1] / 10.) + \
        (K[2] / 100.) * (exp(K[3] / 1000. * AmbTempC) ** (K[4] / 100.))
    return SkyTempC - CorrectionTempC
# end AAG_SkyTempCorrection


def AAG_GetRainFrequency(AAG):
    # получить значение с датчика осадков
    AAG.write("E!")
    response = AAG.read(30)
    IsResponse = patternAAG.match(response)
    if IsResponse and re.match("\!R", IsResponse.group(1)):
        return int(IsResponse.group(2))
    else:
        logAAG.warning("AAG not response")
        return None
    # end if
# end AAG_GetAmbTemp


def AAG_PrintResponse(AAG):
    # вывести ответ датчика в консоль
    AAG.write("A!")
    response = AAG.read(30)
    IsResponse = patternAAG.match(response)
    if IsResponse:
        print IsResponse.group(0)
        print IsResponse.group(1)
        print IsResponse.group(2)
        print IsResponse.group(3)
    else:
        logAAG.warning("AAG not response")
    return None
    # end if
# end AAG_PrintResponse


if __name__ == '__main__':
    #print AAGCheck("D:\Temp\CloudWatcher.csv")
    AAG = AAG_Connect('COM3')
    print AAG_GetSwitch(AAG)
    print AAG_SkyTempCorrection(AAG_GetAmbTemp(AAG),
                                AAG_GetSkyTemp(AAG))
    print AAG_GetRainFrequency(AAG)
    AAG_PrintResponse(AAG)
    AAG_Disconnect(AAG)
# end __main__
