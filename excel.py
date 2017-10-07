##!/usr/bin/env python
# encoding: utf-8
#
# Created by shaohui on 2017/9/16
# Copyright © 2017 shaohui All rights reserved.
#
# *********************************************

# *********************************************
# This file provide function to get file MD5 value.
#
# *********************************************
import xlsxwriter,xlrd
import os,pandas
from hash import File_hash

debug = True  # False

def print_def(args):
	if debug:
		print(args)

file_path = "test/"

class XLSX_class(object):
	file_name = ''	#file name 
	file_path = ''	#abs file path
	file_hash = ''  #file hash
	file_snum = ''  #file sheet numbers
	def __init__(self, arg):
	    self.file_name = arg
	    print_def(self.file_name)
	    self.file_path = os.path.abspath(arg)
	    print_def(self.file_path)
	    self.file_hash = File_hash(arg).get_hash()
	    print_def(self.file_hash)
	    self.file_snum = len(xlrd.open_workbook(arg).sheet_names())
	    print_def(self.file_snum)
	    pass

	def fill_sheets(self,arg):
		pass






def creat_xls(object):
	w_b = xlsxwriter.Workbook(file_path+object)
	w_sheet = w_b.add_worksheet()
	w_sheet.write('A2','Hello world')
	w_b.close()

def open_xls(object):
    if not os.path.isfile(object):
        print("file not exist!\n")
    data = xlrd.open_workbook(object)
    data.sheet_names()
    table = data.sheet_by_index(0)
    df=pandas.read_excel(object,data.sheet_names()[0])
 #   open(table)
    print(df)
    nrows = table.nrows
    ncols = table.ncols
    print(data.sheet_names())
    print(nrows)
    print(ncols)
    #print('file ' + object + ' have ' + 'sheets')
    #print('At sheet one has' + nrows + 'rows And' + ncols + 'cols')




# if __name__ == "__main__":
#creat_xls('hello.xlsx')
pass
open_xls(file_path+'hello.xlsx')
test= XLSX_class(file_path+'hello.xlsx')