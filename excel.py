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
import hashlib

debug = True  # False

def print_def(args):
	if debug:
		print(args)

file_path = "test/"

class XLSX_sheet(object):
	sheet_index = '' #sheet index
	sheet_name = '' #sheet name
#	sheet_hash = '' #sheet content HASH
	sheet_max_col = 0  #Max col nums
	sheet_max_row = 0  #Max row nums
	sheet_row_ret = [] # only used in output file
	sheet_col_ret = [] # only used in output file
	pass
	

class XLSX_class(object):
	file_name = ''	#file name 
	file_path = ''	#abs file path
	file_hash = ''  #file hash
	file_snum = ''  #file sheet numbers
	file_sheet = [] #file sheet info
	def __init__(self, arg):
	    self.file_name = arg
	    print_def(self.file_name)
	    self.file_path = os.path.abspath(arg)
	    print_def(self.file_path)
	    self.file_hash = File_hash(arg).get_hash()
	    print_def(self.file_hash)
	    pass

	def fill_sheets(self):
			if not os.path.isfile(self.file_name):
				print("file not exist!\n")
			data = xlrd.open_workbook(self.file_name)
			self.file_snum = data.nsheets
			print_def('sheet nums:%d' %(self.file_snum))
		
			for sheet_index in range(self.file_snum):
				table = data.sheet_by_index(sheet_index)
				sheet = XLSX_sheet()
				self.file_sheet.append(sheet)
				self.file_sheet[sheet_index].sheet_index = sheet_index
				self.file_sheet[sheet_index].sheet_name = data.sheet_names()[sheet_index]
				#m = hashlib.md5()
				#df = pandas.read_excel(self.file_name,data.sheet_names()[sheet_index])
				#print_def(df)
				#m.update(df)
				#self.file_sheet[sheet_index].sheet_hash = m.hexdigest()
				self.file_sheet[sheet_index].sheet_max_col = table.ncols
				self.file_sheet[sheet_index].sheet_max_row = table.nrows

			pass


	def show_sheets(self):
			print_def('show list:')
			for s in range(self.file_snum):
				print_def('For index:%d' %(s))
				print_def(self.file_sheet[s].sheet_index)
				print_def(self.file_sheet[s].sheet_name)
				#print_def(self.file_sheet[s].sheet_hash)
				print_def(self.file_sheet[s].sheet_max_col)				
				print_def(self.file_sheet[s].sheet_max_row)		
					
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
    #print(df)
    nrows = table.nrows
    ncols = table.ncols
    print(data.sheet_names())
    print(nrows)
    print(ncols)
    print_def(table)
    print_def(table.col_values(2))
    #print('file ' + object + ' have ' + 'sheets')
    #print('At sheet one has' + nrows + 'rows And' + ncols + 'cols')




# if __name__ == "__main__":
#creat_xls('hello.xlsx')
pass
#open_xls(file_path+'hello.xlsx')
test= XLSX_class(file_path+'hello.xlsx')
test.fill_sheets()
test.show_sheets()
