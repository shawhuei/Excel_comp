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

import sys
import hashlib
import os

class File_hash(object):
	# we check xls and xlsx only
	supported_format = ("xls","xlsx")
	file_p = ''
	pass
	def __init__(self,file):
		self.file_p = file

	# sample check, check filename extension with supported_format
	def file_check(self,file):
		for k,v in enumerate(supported_format):
			print (k,v)
			if file_p[-4:].find(v) >= 0:
				return True
		return False

	# return hash of input file
	def get_hash(self):
		if os.path.exists(self.file_p):
			md5file = open(self.file_p,'rb')
			md5 = hashlib.md5(md5file.read()).hexdigest()
			#print(md5)
			md5file.close()
		else :
			return False
		return md5
