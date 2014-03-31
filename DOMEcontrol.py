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
from sunpos import *

# TODO: загрузка информации из ini файла
UTCdif = 4              # разница во времени с UTC
OBN_lat = 55.0969400    # широта с.ш.
OBN_long = 36.6102800   # долгота в.д.


def WeatherCheck(station):
    # получение данных с метеостанции и определение возможности работы
    # TODO: написать логику реакции на состояние атмосферы
    def Fahrenheit2Celsius(Fahrenheit):
        return (Fahrenheit - 32) * 5.0 / 9.0
    # end Fahrenheit2Celsius

    station.parse()
    tempOut = Fahrenheit2Celsius(station.fields['TempOut'])
    dewPoint = Fahrenheit2Celsius(station.fields['DewPoint'])
    humOut = station.fields['HumOut']
    windSpeed = station.fields['WindSpeed10Min']
    rainDay = station.fields['RainDay']

    if rainDay > 0:
        return False
    elif (tempOut - dewPoint) < 3:
        return False
    elif tempOut < 0:
        return False
    elif humOut > 85:
        return False
    elif windSpeed > 8:
        return False
    else:
        return True
    # end if
# end WeatherCheck


def SAMazCheck():
    # проверка файла состояния АГАТ на изменение азимута с последующим чтением
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
        logging.warning(e)
        cord_az = -999
    # end try

    # проверка - доступно ли азимутальное положение АГАТ
    try:
        str_cord = file_text[-1]
        cord_az = float(str_cord[22:32])
    except Exception as e:
        logging.warning(e)
        cord_az = -999
    # end try

    return cord_az
# end SAMazCheck


def WorkFlagCheck(dt_now):
    # расчет положения Солнца в текущий момент
    (SUN_alt, SUN_azimuth) = time_and_location_to_sun_alt_azimuth(
        dt_now.year, dt_now.month, dt_now.day,
        dt_now.hour - UTCdif + dt_now.minute / 60.0 + dt_now.second / 3600.0,
        OBN_lat, OBN_long)
    # если Солнце выше 16 градусов, то можно работать
    if SUN_alt > 16:
        return True
    else:
        return False
    # end if
# end WorkFlagCheck

if __name__ == '__main__':
    # настройка логирования
    logging.basicConfig(filename='DOMEauto.log', filemode='w',
                        format='%(asctime)s %(levelname)s:%(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S', level=logging.DEBUG)

    # установка соединения с куполом
    try:
        dome = win32com.client.Dispatch("ASCOM.DigitalDomeWorks.Dome")
        dome.Connected = True
        logging.info("Dome is connected")
    except Exception as e:
        logging.warning(e)
    # end try

    # установка соединения с метеостанцией
    try:
        station = weather.stations.VantagePro('COM5')
        logging.info("Weatherstation is connected")
    except Exception as e:
        logging.warning(e)
    # end try

    # ожидание погоды и готовности АГАТа
    while not (WeatherCheck(station) and SAMazCheck() > 0):
        time.sleep(60)
    # end while

    while WorkFlagCheck(datetime.today()):
        if WeatherCheck(station):
            dome.ShutterOpen()
        else:
            dome.ShutterClose()
        # end if

        TargetAz = SAMazCheck()
        dome.SlewToAzimuth(TargetAz)
        time.sleep(30)
    # end while

    # завершение работы купола
    dome.ShutterClose()
    time.sleep(30)
    dome.FindHome()
    while dome.Slewing:
        time.sleep(15)
    # end while
    dome.Connected = False
    logging.info("Dome is disconnected")
# end __main__
