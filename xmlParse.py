"""
Authored by: Adam Hollow - DevOps Technician, Devices.
For: South Ayrshire Council

"""

import xml.etree.ElementTree as ET
import os
from os import path
import win32api
import win32net
import adConnect
import win32con
import datetime

filePath = r'\\PATH_TO_STORAGE_FOLDER\inf_' + adConnect.full_name + '.xml'


def initial():
    if path.exists(filePath):
        pass
    else:
        root = ET.Element("Launch")
        b = ET.SubElement(root, "Count")
        c = ET.SubElement(b, "Number")
        d = ET.SubElement(b, "Date")
        e = ET.SubElement(b, "Flag")
        e.text = "False"
        c.text = "0"
        d.text = str(datetime.datetime.now())[:10]
        tree = ET.ElementTree(root)
        tree.write(filePath)


def checkDate():
    tree = ET.parse(filePath)
    root = tree.getroot()
    recDate = root[0][1].text
    nowDate = str(datetime.datetime.now())[:10]
    Dates = [recDate, nowDate]
    x = 0
    for Date in Dates:
        dateParseDate = [Date[:4], Date[5:7], Date[8:10]]
        dateParsedDate = (dateParseDate[0] + dateParseDate[1] + dateParseDate[2])
        Dates[x] = datetime.date(int(dateParsedDate[0:4]), int(dateParsedDate[4:6]), int(dateParsedDate[6:8]))
        x += 1
    datedelta = Dates[1] - Dates[0]
    if datedelta.days < 180:
        return False
    else:
        return True


def checkFlag():
    tree = ET.parse(filePath)
    root = tree.getroot()
    flag = root[0][2].text
    if flag == 'True':
        return True
    if flag == 'False':
        return False


def checkCount():
    tree = ET.parse(filePath)
    root = tree.getroot()
    num = int(root[0][0].text)
    if num > 9:
        return True
    else:
        return False


def closedApp():
    tree = ET.parse(filePath)
    root = tree.getroot()
    root[0][0].text = str(int(root[0][0].text) + 1)
    tree.write(filePath)


def completed():
    tree = ET.parse(filePath)
    root = tree.getroot()
    root[0][0].text = 0
    tree.write(filePath)


def updateInfo():
    tree = ET.parse(filePath)
    root = tree.getroot()
    root[0][1].text = str(datetime.datetime.now())[:10]
    root[0][2].text = 'True'
    tree.write(filePath)

