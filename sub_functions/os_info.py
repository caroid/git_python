import os 
import subprocess 
import re 
import hashlib 
import getopt
from getopt import GetoptError
from bs4 import BeautifulSoup
import sys
from xlrd import open_workbook
from xlrd import open_workbook,cellname
from xlutils.copy import copy
import xlwt 

#
def get_osinfo(): 
    os_info = {} 
    i = os.uname() 
    os_info['os_type'] = i[0] 
    os_info['node_name'] = i[1] 
    os_info['kernel'] = i[2] 
    return os_info 
