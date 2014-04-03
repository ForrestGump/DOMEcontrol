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
        logDome.info("RainRate "+str(rainRate)+" > 0")
        return False
    elif (tempOut - dewPoint) < 1.5:
        # температура приближается к точке росы
        logDome.info("TempOut is near DewPoint, "+str(tempOut - dewPoint))
        return False
    elif tempOut < 0:
        # температура ниже нуля
        logDome.info("TempOut "+str(tempOut)+" < 0")
        return False
    elif humOut > 85:
        # влажность выше 85%
        logDome.info("HumOut "+str(humOut)+"% > 85%")
        return False
    elif windSpeed > 5:
        # ветер больше 5 м/с
        logDome.info("WindSpeed "+str(windSpeed)+"m/sec > 5 m/sec")
        return False
    else:
        logDome.info("Weather is OK")
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
        logDome.warning(e)
        cord_az = -999
    # end try

    # проверка - доступно ли азимутальное положение АГАТ
    try:
        str_cord = file_text[-1]
        cord_az = float(str_cord[22:32])
    except Exception as e:
        logDome.warning(e)
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
    # если Солнце выше 15.5 градусов, то можно работать
    if SUN_alt > 15.5:
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
    logDome = logging.getLogger(__name__)

    # установка соединения с куполом
    try:
        dome = win32com.client.Dispatch("ASCOM.DigitalDomeWorks.Dome")
        dome.Connected = True
        logDome.info("Dome is connected")
    except Exception as e:
        logDome.warning(e)
    # end try

    # установка соединения с метеостанцией
    try:
        station = weather.stations.VantagePro('COM5')
        logDome.info("Weatherstation is connected")
    except Exception as e:
        logDome.warning(e)
    # end try

    # ожидание Солнца выше 16 градусов над горизонтом и готовности АГАТа
    while not (WorkFlagCheck(datetime.datetime.today()) and SAMazCheck() > 0):
        logDome.info("Sun is below 15.5 degrees or SAM is not ready yet")
        time.sleep(60)
    # end while

    logDome.info("Getting Started")

    while WorkFlagCheck(datetime.datetime.today()):
        # NOTE: ShutterStatus 3=indeterm, 1=closed, 0=open
        if WeatherCheck(station):
            if dome.ShutterStatus != 0:
                dome.OpenShutter()
                logDome.info("Shutter opened")
                time.sleep(15)
            # end if

            TargetAz = SAMazCheck()
            try:
                dome.SlewToAzimuth(TargetAz + 5)
                logDome.info("Dome is going to " + str(TargetAz))
            except Exception as e:
                logDome.warning(e)
            # end try
            time.sleep(60)
            logDome.info("Dome is in " + str(dome.Azimuth))

        else:
            if dome.ShutterStatus != 1:
                dome.CloseShutter()
                logDome.info("Shutter closed")
                time.sleep(15)
            # end if
        # end if
    # end while

    logDome.info("End of work")

    # завершение работы купола - закрытие затвора
    if dome.ShutterStatus != 1:
        dome.CloseShutter()
        logDome.info("Shutter closed")
        time.sleep(15)
    # end if

    # завершение работы купола - парковка
    try:
        dome.FindHome()
        logDome.info("DOME is in Home position")
    except Exception as e:
        logDome.warning(e)
    # end try
    while dome.Slewing:
        time.sleep(15)
    # end while
    dome.Connected = False
    logDome.info("Dome is disconnected")
# end __main__
