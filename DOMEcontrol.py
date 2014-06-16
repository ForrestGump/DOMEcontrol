# -*- coding: utf-8 -*-
"""
Created on Mon Mar 24 14:27:44 2014

@author: sborodin@rpatyphoon.ru
"""

import time
import datetime
import logging
import win32com.client
import weather.stations
from sunpos import time_and_location_to_sun_alt_azimuth
import cloudWatcher as cW

# TODO: загрузка информации из ini файла
UTC_DIFF = 4              # разница во времени с UTC
OBN_LAT = 55.0969400    # широта с.ш.
OBN_LONG = 36.6102800   # долгота в.д.

logDome = logging.getLogger('DOMEcontrol')


def WeatherCheck(station):
    '''Получение данных с метеостанции и определение возможности работы
    '''
    # TODO: сохранение (публикация) метеоданных
    def Fahrenheit2Celsius(fahrenheit):
        # Fahrenheit (F) to Celsius (C)
        return (fahrenheit - 32) * 5.0 / 9.0
    # end Fahrenheit2Celsius

    def mph_to_m_sec(mph):
        # Miles/hour (mph) to meters/second (m/s)
        return mph * 0.44704
    # end mph_to_m_sec

    def inches_to_mm(inches):
        # Inches (in) to millimeters (mm)
        return inches * 25.4
    # end inches_to_mm

    station.parse()
    tempOut = Fahrenheit2Celsius(station.fields['TempOut'])
    dewPoint = Fahrenheit2Celsius(station.fields['DewPoint'])
    humOut = station.fields['HumOut']
    windSpeed = mph_to_m_sec(station.fields['WindSpeed10Min'])
    rainRate = inches_to_mm(station.fields['RainRate'])

    if rainRate > 0:
        # дождь
        logDome.warning("RainRate " + str(rainRate) + "mm/hr > 0")
        return False
    elif (tempOut - dewPoint) < 1.5:
        # температура приближается к точке росы
        logDome.warning("TempOut is near DewPoint, " + str(tempOut - dewPoint))
        return False
    elif tempOut < 0:
        # температура ниже нуля
        logDome.warning("TempOut " + str(tempOut) + " < 0")
        return False
    elif humOut > 80:
        # влажность выше 80%
        logDome.warning("HumOut " + str(humOut) + "% > 80%")
        return False
    elif windSpeed > 8:
        # ветер больше 8 м/с
        logDome.warning("WindSpeed " + str(windSpeed) + "m/sec > 8 m/sec")
        return False
    else:
        logDome.info("Weather is OK")
        logDome.info(str(tempOut) + " " + str(dewPoint) + " " + str(humOut) +
                     " " + str(windSpeed) + " " + str(rainRate))
        return True
    # end if
# end WeatherCheck


def AAGCheck(AAG):
    '''Получение данных с датчика облачности
    '''
    aag_switch = cW.AAG_GetSwitch(AAG)
    aag_skyTemp = cW.AAG_SkyTempCorrection(cW.AAG_GetAmbTemp(AAG),
                                           cW.AAG_GetSkyTemp(AAG))
    aag_rainFreq = cW.AAG_GetRainFrequency(AAG)

    if not(aag_switch):
        # небезопасный режим работы
        logDome.warning("AAG_GetSwitch is " + str(aag_switch))
        return False
    elif aag_skyTemp > -2.:
        # облака отколо теплой фазы
        logDome.warning("skyTemp " + str(aag_skyTemp) + " > -2")
        return False
    elif aag_rainFreq < 2200:
        # возможен дождь
        logDome.warning("RainFrequency " + str(aag_rainFreq) + " < 2200")
        return False
    else:
        logDome.info("AAG is OK")
        logDome.info(str(aag_switch) + " " + str(aag_skyTemp)
                     + " " + str(aag_rainFreq))
        return True
    # end if
# end AAGCheck


def SAMazCheck():
    '''Проверка файла состояния АГАТ на изменение азимута с последующим чтением
    '''
    dt_begin = datetime.datetime.now()
    str_targetFile = "Z:\obni_210-" + dt_begin.strftime('%m') +\
        "-" + dt_begin.strftime('%d') + "-" + dt_begin.strftime('%Y') +\
        "\\obni_" + dt_begin.strftime('%m') + "_" +\
        dt_begin.strftime('%d') + "_" + dt_begin.strftime('%Y') + "_mt_log.dat"

    # проверка - создан ли файл
    try:
        log_move = open(str_targetFile)
        file_text = log_move.readlines()
        log_move.close()
    except Exception as e:
        logDome.exception(e)
        cord_az = -999
    # end try

    # проверка - доступно ли азимутальное положение АГАТ
    try:
        str_cord = file_text[-1]
        cord_az = float(str_cord[22:32])
    except Exception as e:
        logDome.exception(e)
        cord_az = -999
    # end try

    return cord_az
# end SAMazCheck


def WorkFlagCheck(dt_now):
    '''Расчет положения Солнца в текущий момент
    '''
    (SUN_alt, SUN_azimuth) = time_and_location_to_sun_alt_azimuth(
        dt_now.year, dt_now.month, dt_now.day,
        dt_now.hour - UTC_DIFF + dt_now.minute / 60.0 + dt_now.second / 3600.0,
        OBN_LAT, OBN_LONG)
    # если Солнце выше 15.5 градусов, то можно работать
    if SUN_alt > 15.5:
        return True
    else:
        return False
    # end if
# end WorkFlagCheck


def system_shutdown():
    try:
        import win32security
        import win32api
        import ntsecuritycon
        import os

        flags = ntsecuritycon.TOKEN_ADJUST_PRIVILEGES | \
            ntsecuritycon.TOKEN_QUERY
        htoken = win32security.OpenProcessToken(win32api.GetCurrentProcess(),
                                                flags)
        id = win32security.LookupPrivilegeValue(None,
                                                ntsecuritycon.SE_SHUTDOWN_NAME)
        newPrivileges = [(id, ntsecuritycon.SE_PRIVILEGE_ENABLED)]
        win32security.AdjustTokenPrivileges(htoken, 0, newPrivileges)
        win32api.InitiateSystemShutdown("", "", 300, 1, 0)
    finally:
        os._exit(0)
    # end try
# end system_shutdown


if __name__ == '__main__':
    # настройка логирования
    logging.basicConfig(filename='DOMEauto.log',
                        format='%(asctime)s %(name)s \
                        %(levelname)s:%(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S', level=logging.INFO)
    # вывод только сообщение нашего логера
    for handler in logging.root.handlers:
        handler.addFilter(logging.Filter('DOMEcontrol'))
    logDome.info("------------------------------")

    # установка соединения с куполом
    try:
        dome = win32com.client.Dispatch("ASCOM.DigitalDomeWorks.Dome")
        dome.Connected = True
        logDome.info("Dome is connected")
    except Exception as e:
        logDome.exception(e)
    # end try

    # установка соединения с метеостанцией
    try:
        station = weather.stations.VantagePro('COM5')
        logDome.info("Weatherstation is connected")
    except Exception as e:
        logDome.exception(e)
    # end try

    # установка соединения с датчиком облачности
    try:
        AAG = cW.AAG_Connect('COM3')
        logDome.info("AAG cloudWatcher is connected")
    except Exception as e:
        logDome.exception(e)
    # end try

    # ожидание Солнца выше 16 градусов над горизонтом и готовности АГАТа
    while not (WorkFlagCheck(datetime.datetime.today()) and SAMazCheck() > 0):
        logDome.info("Sun is below 15.5 degrees or SAM is not ready yet")
        time.sleep(60)
    # end while

    logDome.info("Getting Started")

    while WorkFlagCheck(datetime.datetime.today()):
        # NOTE: ShutterStatus 3=indeterm, 1=closed, 0=open
        if WeatherCheck(station) and AAGCheck(AAG):
            if dome.ShutterStatus != 0:
                try:
                    dome.OpenShutter()
                    logDome.info("Shutter opened")
                except Exception as e:
                    logDome.exception(e)
                # end try
                time.sleep(30)
            # end if

            TargetAz = SAMazCheck()
            try:
                dome.SlewToAzimuth(TargetAz - 20)
                logDome.info("SAM is in " + str(TargetAz))
            except Exception as e:
                logDome.exception(e)
            # end try
            time.sleep(60)
            logDome.info("Dome is in " + str(dome.Azimuth))

        else:
            if dome.ShutterStatus != 1:
                try:
                    dome.CloseShutter()
                    logDome.info("Shutter closed")
                except Exception as e:
                    logDome.exception(e)
                # end try
                time.sleep(30)
            else:
                time.sleep(30)
            # end if
        # end if
    # end while

    # завершение работы датчика облачности
    cW.AAG_Disconnect(AAG)
    logDome.info("AAG cloudWatcher is disconnected")

    # завершение работы купола - закрытие затвора
    if dome.ShutterStatus != 1:
        dome.CloseShutter()
        logDome.info("Shutter closed")
        time.sleep(30)
    # end if

    dome.Connected = False
    logDome.info("Dome is disconnected")
    logDome.info("End of work")
# end __main__
