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
import shutil
from xlutils.copy import copy

import win32com.client

debug = True  # False
Info = True #False
Err = True

FILE_ERR = -1
FILE_SAME = 0
FILE_DIFF = 1

#Comapre flags
SAME_ROW	= 'SROW'
SAME_COL	= 'SCOL'
DIFF_CELL	= 'DCELL'
SAME_CELL	=	'SCELL'

EXTEND_FILE = '_ret'

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
	#sheet_ret = []
	pass
	

class XLSX_class(object):
	file_name = []	#file name 
	file_path_full = []	#abs file path
	file_path = [] #file path
	file_hash = []  #file hash
	file_snum = []  #file sheet numbers
	file_sheet = [[] for i in range(3)] #file sheet info
	#Note!!!!! if more sheets compared,range should enlarge
	file_output = '' #out put file name
	#private data
	__data = ''
	__data2 = ''
	__data3 = ''
	__table = []
	__table2 =[]
	__table3 = []

	def __init__(self, arg, arg2):
		self.file_name.append(os.path.basename(arg))
		self.file_name.append(os.path.basename(arg2))
		print_debug('Input File name:%s' %self.file_name)

		self.file_path_full.append(os.path.abspath(arg))
		self.file_path_full.append(os.path.abspath(arg2))
		self.file_path.append(os.path.dirname(self.file_path_full[0]))
		self.file_path.append(os.path.dirname(self.file_path_full[1]))
		print_debug('Input File Path:%s' %self.file_path_full)
		print_debug('Input File Path:%s' %self.file_path)
		self.file_hash.append(File_hash(arg).get_hash())
		self.file_hash.append(File_hash(self.file_path_full[1]).get_hash())
		self.file_output = os.path.splitext(self.file_path_full[1])[0]+EXTEND_FILE+os.path.splitext(self.file_path_full[1])[1]
		print_debug('Gen file %s ' %self.file_output)
	pass

#private funs

	def __fill_sheets_A(self):
		#Fill in Sheet A
		if not os.path.isfile(self.file_path_full[0]):
			print_err("file A %s not exist!\n" %self.file_path_full[0])
			return FILE_ERR
		self.__data = xlrd.open_workbook(self.file_path_full[0])#,formatting_info=True)
		self.file_snum.append(self.__data.nsheets)
		print_debug('sheet nums:%d' %(self.file_snum[0]))
		for sheet_index in range(self.file_snum[0]):
			self.__table.append(self.__data.sheet_by_index(sheet_index))
			sheet = XLSX_sheet()
			self.file_sheet[0].append(sheet)
			self.file_sheet[0][sheet_index].sheet_index = sheet_index
			self.file_sheet[0][sheet_index].sheet_name = self.__data.sheet_names()[sheet_index]
			#m = hashlib.md5()
			#df = pandas.read_excel(self.file_path_full,__data.sheet_names()[sheet_index])
			#print_debug(df)
			#m.update(df)
			#self.file_sheet[sheet_index].sheet_hash = m.hexdigest()
			self.file_sheet[0][sheet_index].sheet_max_col = self.__table[sheet_index].ncols
			self.file_sheet[0][sheet_index].sheet_max_row = self.__table[sheet_index].nrows
		return 0

		pass
	def __fill_sheets_B(self):
		#Fill in Sheet B
		if not os.path.isfile(self.file_path_full[1]):
			print_err("file B %s not exist!\n" %self.file_path_full[1])
			return FILE_ERR
		self.__data2 = xlrd.open_workbook(self.file_path_full[1])#,formatting_info=True)
		self.file_snum.append(self.__data2.nsheets)
		print_debug('sheet nums:%d' %(self.file_snum[1]))
		for sheet_index in range(self.file_snum[1]):
			self.__table2.append(self.__data2.sheet_by_index(sheet_index))
			sheet2 = XLSX_sheet()
			self.file_sheet[1].append(sheet2)
			self.file_sheet[1][sheet_index].sheet_index = sheet_index
			self.file_sheet[1][sheet_index].sheet_name = self.__data2.sheet_names()[sheet_index]
			#m = hashlib.md5()
			#df = pandas.read_excel(self.file_path_full,__data.sheet_names()[sheet_index])
			#print_debug(df)
			#m.update(df)
			#self.file_sheet[sheet_index].sheet_hash = m.hexdigest()
			self.file_sheet[1][sheet_index].sheet_max_col = self.__table2[sheet_index].ncols
			self.file_sheet[1][sheet_index].sheet_max_row = self.__table2[sheet_index].nrows
		return 0

		pass
		
	def __fill_sheets_C(self):
		#Fill in Sheet C
		if not os.path.isfile(self.file_output):
			print_err("file C %s not exist!\n" %self.file_output)
			return FILE_ERR
		self.__data3 = xlrd.open_workbook(self.file_output)#,formatting_info=True)
		self.file_snum.append(self.__data3.nsheets)
		print_debug('sheet nums:%d' %(self.file_snum[2]))
		for sheet_index in range(self.file_snum[2]):
			self.__table3.append(self.__data3.sheet_by_index(sheet_index))
			sheet3 = XLSX_sheet()
			self.file_sheet[2].append(sheet3)
			self.file_sheet[2][sheet_index].sheet_index = sheet_index
			self.file_sheet[2][sheet_index].sheet_name = self.__data3.sheet_names()[sheet_index]
			#m = hashlib.md5()
			#df = pandas.read_excel(self.file_path_full,__data.sheet_names()[sheet_index])
			#print_debug(df)
			#m.update(df)
			#self.file_sheet[sheet_index].sheet_hash = m.hexdigest()
			self.file_sheet[2][sheet_index].sheet_max_col = self.__table3[sheet_index].ncols
			self.file_sheet[2][sheet_index].sheet_max_row = self.__table3[sheet_index].nrows
		return 0

		pass		
	def __compare_byrow(self,s_index):
		print_info('Compare By Row start!\n')
		max_row = max(self.file_sheet[1][s_index].sheet_max_row,self.file_sheet[0][s_index].sheet_max_row)
		print_debug('max row %d\n' %max_row)
		for r in range(max_row):
			if(self.__table[s_index].row_values(r) != self.__table2[s_index].row_values(r)):
				print_info('line %d is diff!\n' %r)
				self.file_sheet[0][s_index].sheet_row_ret.append('F')
				self.file_sheet[1][s_index].sheet_row_ret.append('F')
			else:
				self.file_sheet[0][s_index].sheet_row_ret.append('S')
				self.file_sheet[1][s_index].sheet_row_ret.append('S')
								
		print_info('Compare By Row end!\n')
		pass
	def __compare_bycol(self,s_index):
		print_info('Compare By Col start!\n')
		max_col = max(self.file_sheet[1][s_index].sheet_max_col,self.file_sheet[0][s_index].sheet_max_col)
		print_debug('max row %d\n' %max_col)
		for c in range(max_col):
			if(self.__table[s_index].col_values(c) != self.__table2[s_index].col_values(c)):
				print_info('col %d is diff!\n' %c)
				self.file_sheet[0][s_index].sheet_col_ret.append('F')
				self.file_sheet[1][s_index].sheet_col_ret.append('F')
			else:
				self.file_sheet[0][s_index].sheet_col_ret.append('S')
				self.file_sheet[1][s_index].sheet_col_ret.append('S')
		print_info('Compare By Col end!\n')
		pass
	
	def __accurate_compare(self,s_index):
		print_info('Accurate Compare By Row start!\n')
		if os.path.isfile(self.file_output):
			print_info('%s is exist already!!! Will overwrite this file!!!\n' %os.path.basename(self.file_output))
		else:
			shutil.copyfile(self.file_path_full[1],self.file_output)		
		ret = self.__fill_sheets_C()
		#print_debug('ret:%d' %ret)
		max_sheet = max(self.file_snum[0],self.file_snum[1])
		print_debug('Max sheet %d' %max_sheet)
		
		for s in range(max_sheet):
			max_row = max(self.file_sheet[1][s].sheet_max_row,self.file_sheet[0][s].sheet_max_row)
			max_col = max(self.file_sheet[1][s].sheet_max_col,self.file_sheet[0][s].sheet_max_col)
			print_debug('Cur sheet %d, row %d,Max col %d' %(s,max_row,max_col))
			for r in range(max_row):	#compare by row
				if(self.__table[s].row_values(r) == self.__table2[s].row_values(r)):
					print_debug('Same Row %d' %r)
					self.file_sheet[2][s].sheet_row_ret.append(SAME_ROW)
					#self.file_sheet[2][s].sheet_ret[r].append(SAME_ROW)
					#self.file_sheet[2][s].sheet_row_ret.append(SAME_ROW)
				else:
					for c in range(max_col):
						if(self.__table[s].cell(r,c).value == self.__table2[s].cell(r,c).value):
							self.file_sheet[2][s].sheet_col_ret.append(SAME_CELL)
							#print_debug('Same cell %d %d' %(r,c))
							#self.file_sheet[2][s].sheet_ret[r].append(SAME_CELL)
						else:
							self.file_sheet[2][s].sheet_col_ret.append(DIFF_CELL)
							print_debug('Diff cell %d %d' %(r,c))
							#self.file_sheet[2][s].sheet_ret[r].append(DIFF_CELL)		
				
					self.file_sheet[2][s].sheet_row_ret.append(self.file_sheet[2][s].sheet_col_ret)
					self.file_sheet[2][s].sheet_col_ret=[]     #Clear list					
				print_debug(self.file_sheet[2][s].sheet_row_ret)
		print_info('Accurate Compare By Row end!\n')				
							
		pass


	def __gen_output(self):
		############################################################################################
		#This part run by win32
		
		winapp = win32com.client.DispatchEx('Excel.Application')
		winBook = winapp.Workbooks.Open(self.file_output)
		for s in range(self.file_snum[2]):
			winSheet = winBook.Worksheets(self.file_sheet[2][s].sheet_name)
			for r in range(self.file_sheet[2][s].sheet_max_row):
				if(self.file_sheet[2][s].sheet_row_ret[r] != SAME_ROW):
					for c in range(self.file_sheet[2][s].sheet_max_col):
						#print_debug('color cell: %d %d' %(r,c))
						#print_debug(self.file_sheet[2][s].sheet_row_ret[r])
						if(self.file_sheet[2][s].sheet_row_ret[r][c] != SAME_CELL):
							winSheet.Cells(r+1,c+1).Interior.ColorIndex = 3
							print_debug('color cell: %d %d' %(r,c))
		############################################################################################
		
		winBook.Save()
		winBook.Close()
		#w_b = xlsxwriter.Workbook(self.file_output)
		#w_b = copy(self.__data2)
	#	for sh in range(self.file_snum[1]):
			#w_sheet.append(w_b.add_worksheet(self.file_sheet[1][sh].sheet_name))
	#		w_sheet.append(w_b.get_sheet(sh))
			
		#for sh in range(self.file_snum[1]):
		#	for row in range(max(self.file_sheet[1][sh].sheet_max_col,self.file_sheet[0][sh].sheet_max_col)):
		#		if self.file_sheet[1][sh].sheet_col_ret[row] == 'S':
		#			w_sheet[sh].write(row,0,self.__table2[sh].row_values(row))
		#w_b.save(self.file_output)
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
		#self.__compare_byrow(0)
		#self.__compare_bycol(0)
		self.__accurate_compare(0)
		self.__gen_output()
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



# Examples:

if __name__ == "__main__":
	#creat_xls('hello.xlsx')
	#open_xls(file_path+'hello.xlsx')
	test= XLSX_class(file_path+'hello.xlsx',file_path+'hello2.xlsx')
	
	ret = test.fill_sheets()
	print_debug('Debug test,ret:%d' %ret)
	#test.show_sheets()
