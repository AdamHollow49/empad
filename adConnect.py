"""
Authored by: Adam Hollow - DevOps Technician, Devices.
For: South Ayrshire Council

"""
import pyad
from pyad import adquery, adsearch, adobject, adcomputer, aduser, adbase, adcontainer, addomain, adgroup

import win32api
import win32net
import win32con
import ctypes
import os


q = pyad.adquery.ADQuery()

def query(username):
    q.execute_query(
        attributes=["distinguishedName"],
        where_clause="SamAccountName = '{}'".format(username),
        base_dn="DC=YOURDC, DC=YOURDC"
    )
    return q.get_single_result()['distinguishedName']

def userCon():
    username = os.environ['USERNAME']
    username = username.replace("'", "''")
    user = pyad.aduser.ADUser.from_dn(query(username))
    return user


def get_display_name():
    user_info = win32net.NetUserGetInfo(win32net.NetGetAnyDCName(), win32api.GetUserName(), 2)
    fullname = user_info["full_name"]
    return fullname


full_name = get_display_name()
