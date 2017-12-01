##!/usr/bin/env python
# encoding: utf-8
#
# Created by shaohui on 2017/12/1
# Copyright @ 2017 shaohui All rights reserved.
#
# *********************************************

# *********************************************
# This file provide functions for debug printing.
#
# *********************************************

#debug log level
debug = False  # False
Info = True #False
Err = True

DEBUG_MAX_LEVEL=2			#0, Err;1 Info and Err;2 ALL
debug_level = 1	#set 1 for defalut

def print_level_get():
	print("%d %d %d" %(debug,Info,Err))
	return debug_level

def print_level_set(level):
	global Err
	global Info
	global debug
	if(level>DEBUG_MAX_LEVEL):
		return -1;
	if(level == 0):
		Err = True
		Info = False
		debug = False		
	if(level == 1):
		Err = True
		Info = True
		debug = False
	if(level == 2):
		Err = True
		Info = True
		debug = True
	print_level_get()
	return 0

def print_debug(args):
	if debug:
		print(args)

def print_info(args):
	if Info:
		print(args)

def print_err(args):
	if Err:
		print(args)