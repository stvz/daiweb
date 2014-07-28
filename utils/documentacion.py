# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import xlrd, operator
from os import path
import MySQLdb as my_
from tracking.models import cat001FormatosXls, cat001CamposXls, det001FormatoCampos

