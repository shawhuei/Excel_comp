##!/usr/bin/env python
# encoding: utf-8
#
# Created by shaohui on 2017/9/16
# Copyright © 2017 shaohui All rights reserved.
#
# *********************************************

# *********************************************
# This file provide function to compare excel and generate result.
#
# *********************************************
import xlsxwriter,xlrd
import os,pandas
from hash import File_hash
import hashlib

debug = True  # False
Info = True #False
Err = True

FILE_ERR = -1
FILE_SAME = 0
FILE_DIFF = 1


def print_debug(args):
	if debug:
		print(args)

def print_info(args):
	if Info:
		print(args)

def print_err(args):
	if Err:
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
	file_name = []	#file name 
	file_path_full = []	#abs file path
	file_path = [] #file path
	file_hash = []  #file hash
	file_snum = []  #file sheet numbers
	file_sheet = [[] for i in range(2)] #file sheet info
	file_output = '' #out put file name
	#private data
	__data = ''
	__data2 = ''
	__table = []
	__table2 =[]

	def __init__(self, arg, arg2):
		self.file_name.append(arg)
		self.file_name.append(arg2)
		print_debug('Input File name:%s' %self.file_name)
		self.file_path_full.append(os.path.abspath(arg))
		self.file_path_full.append(os.path.abspath(arg2))
		self.file_path.append(os.path.dirname(self.file_path_full[0]))
		self.file_path.append(os.path.dirname(self.file_path_full[1]))
		print_debug('Input File Path:%s' %self.file_path_full)
		print_debug('Input File Path:%s' %self.file_path)
		self.file_hash.append(File_hash(arg).get_hash())
		self.file_hash.append(File_hash(self.file_path_full[1]).get_hash())
		#print_debug(self.file_hash)
	pass

#private funs

	def __fill_sheets_A(self):
		#Fill in Sheet A
		if not os.path.isfile(self.file_name[0]):
			print_err("file A %s not exist!\n" %self.file_name[0])
			return FILE_ERR
		self.__data = xlrd.open_workbook(self.file_name[0])
		self.file_snum.append(self.__data.nsheets)
		print_debug('sheet nums:%d' %(self.file_snum[0]))
		for sheet_index in range(self.file_snum[0]):
			self.__table.append(self.__data.sheet_by_index(sheet_index))
			sheet = XLSX_sheet()
			self.file_sheet[0].append(sheet)
			self.file_sheet[0][sheet_index].sheet_index = sheet_index
			self.file_sheet[0][sheet_index].sheet_name = self.__data.sheet_names()[sheet_index]
			#m = hashlib.md5()
			#df = pandas.read_excel(self.file_name,__data.sheet_names()[sheet_index])
			#print_debug(df)
			#m.update(df)
			#self.file_sheet[sheet_index].sheet_hash = m.hexdigest()
			self.file_sheet[0][sheet_index].sheet_max_col = self.__table[sheet_index].ncols
			self.file_sheet[0][sheet_index].sheet_max_row = self.__table[sheet_index].nrows
		return 0

		pass
	def __fill_sheets_B(self):
		#Fill in Sheet B
		if not os.path.isfile(self.file_name[1]):
			print_err("file B %s not exist!\n" %self.file_name[1])
			return FILE_ERR
		self.__data2 = xlrd.open_workbook(self.file_name[1])
		self.file_snum.append(self.__data2.nsheets)
		print_debug('sheet nums:%d' %(self.file_snum[1]))
		for sheet_index in range(self.file_snum[1]):
			self.__table2.append(self.__data2.sheet_by_index(sheet_index))
			sheet2 = XLSX_sheet()
			self.file_sheet[1].append(sheet2)
			self.file_sheet[1][sheet_index].sheet_index = sheet_index
			self.file_sheet[1][sheet_index].sheet_name = self.__data2.sheet_names()[sheet_index]
			#m = hashlib.md5()
			#df = pandas.read_excel(self.file_name,__data.sheet_names()[sheet_index])
			#print_debug(df)
			#m.update(df)
			#self.file_sheet[sheet_index].sheet_hash = m.hexdigest()
			self.file_sheet[1][sheet_index].sheet_max_col = self.__table2[sheet_index].ncols
			self.file_sheet[1][sheet_index].sheet_max_row = self.__table2[sheet_index].nrows
		return 0

		pass
	def __compare_byrow(self,s_index,T1,T2):
		print_info('Compare By Row start!\n')
		max_row = max(self.file_sheet[1][s_index].sheet_max_row,self.file_sheet[0][s_index].sheet_max_row)
		print_debug('max row %d\n' %max_row)
		for r in range(max_row):
			if(self.__table[s_index].row_values(r) != self.__table2[s_index].row_values(r)):
				print_info('line %d is diff!\n' %r)
		print_info('Compare By Row end!\n')
		pass
	def __compare_bycol(self,s_index,T1,T2):
		print_info('Compare By Col start!\n')
		max_row = max(self.file_sheet[1][s_index].sheet_max_col,self.file_sheet[0][s_index].sheet_max_col)
		print_debug('max row %d\n' %max_row)
		for r in range(max_row):
			if(self.__table[s_index].col_values(r) != self.__table2[s_index].col_values(r)):
				print_info('col %d is diff!\n' %r)
		print_info('Compare By Col end!\n')
		pass		
	def __do_compare(self):
		#First Get all excel buffer
		####################################################################
		#TBD  Do comapre for sheet number,how to compare if del/add sheets
		#
		####################################################################
		#*******************************************************************
		####################################################################
			
		print_debug('compare start!\n')
		if(self.file_hash[0] == self.file_hash[1]):
			print_info('Same file!\n')
			return FILE_SAME
		print_info('Diff file!\n')
		
		t1 = self.__table[0]
		t2 = self.__table2[0]
		#Test for Row compare
		self.__compare_byrow(0,t1,t2)
		self.__compare_bycol(0,t1,t2)
		return FILE_DIFF
		pass
#Open funcs			
	def fill_sheets(self):
		#Open funcs to user
		if(self.__fill_sheets_A() == 0 and self.__fill_sheets_B() == 0):
			return self.__do_compare()#For test
		else:
			print_err('Fill in file Err!\n')
			return FILE_ERR
		pass

	def show_sheets(self):
		for i in range (0,2,1):
			print_info('show list:%d' %i)
			for s in range(self.file_snum[i]):
				print_info('For index:%d' %(s))
				print_info(self.file_sheet[i][s].sheet_index)
				print_info(self.file_sheet[i][s].sheet_name)
				#print_debug(self.file_sheet[0][s].sheet_hash)
				print_info(self.file_sheet[i][s].sheet_max_col)				
				print_info(self.file_sheet[i][s].sheet_max_row)		
					
			pass



#Test codes:

def creat_xls(object):
	w_b = xlsxwriter.Workbook(file_path+object)
	w_sheet = w_b.add_worksheet()
	w_sheet.write('A2','Hello world')
	w_b.close()

def open_xls(object):
    if not os.path.isfile(object):
        print_err("file not exist!\n")
    data = xlrd.open_workbook(object)
    data.sheet_names()
    table = data.sheet_by_index(0)
    df=pandas.read_excel(object,data.sheet_names()[0])
 #   open(table)
    #print_debug(df)
    nrows = table.nrows
    ncols = table.ncols
    print_debug(data.sheet_names())
    print_debug(nrows)
    print_debug(ncols)
    print_debug(table)
    print_debug(table.col_values(2))
    #print_debug('file ' + object + ' have ' + 'sheets')
    #print_debug('At sheet one has' + nrows + 'rows And' + ncols + 'cols')





pass







# Examples:

if __name__ == "__main__":
	#creat_xls('hello.xlsx')
	#open_xls(file_path+'hello.xlsx')
	test= XLSX_class(file_path+'hello.xlsx',file_path+'hello2.xlsx')
	
	ret = test.fill_sheets()
	print_debug('Debug test,ret:%d' %ret)
	#test.show_sheets()
