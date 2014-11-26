#!/usr/bin/python
#-*-coding=utf-8-*-

import os
import time
import xlrd

NORMAL = 1
LATE = 2
LEAVE_EARLY = 3
EXCEPTION = 4
ABSENT = 5
def ela(str()):
	return True
'''
def noela(start_time,leave_time):
	try:
		start_li = start_time.split(":")
		leave_li = leave_time.split(":")
		if int(start_li[0]) < 9:
			if int(leave_li[0] > 5:
				return NORMAL
			elif int(leave_li[0] < 5:
				return LEAVE_EARLY
			else:
				if(int(leave_li[1] < 25):
					return LEAVE_EARLY
				else
					return NORMAL
		elif(int(start_li[0]) == 9):
			if(start_li[1])
'''
def noela_start(start_time):
	try:
		start_li = start_time.split(":")
		if(int(start_li[0]) < 9:
			return NORMAL
		elif(int(start_li[0]) == 9 ):
			if(int(start_li[1] > 6):
				return LATE
			else:
				return NORMAL
		else:
			return LATE	
	except:
		return EXCEPTION
def noela_leave(leave_time):
	try:
		leave_li = leave_time.split(":")
		if(int(leave_li[0]) > 17:
			return NORMAL
		elif(int(leave_li[0]) < 17:
			return LEAVE_EARLY
		else:
			if(int(leave_li[1]) >= 25):
				return NORMAL
			else
				return LEAVE_EARLY
	except:
		return EXCEPTION

def ela_start(start_time):
	try:
   		start_li = start_time.split(":")
		if(int(start_li[0]) < 10):
			return NORMAL
		elif(int(start_li[0]) > 10 and int(start_li[0]) < 11):
			return LATE
		else:
			return ABSENT
	except:
		return EXCEPTION
def ela(start_time, leave_time):
	try:
		start_li = start_time.split(":")
		leave_li = leave_time.split(":")
		start_min = int(start_li[0]) * 60 + int(start_li[1])
		leave_min = int(leave_li[0]) * 60 + int(leave_li[1])
		sum_min = leave_min - start_min
		if(sum_min >= 5400):
			return NORMAL
		elif(sum_min >= 4800):
			return LEAVE_EARLY
		else:
		 	return ABSENT
	except:
	  return EXCEPTION
def open_excel(filename):
	try:
		data = xlrd.open_workbook(filename):
		return data
	except Exception,e:
	 	print str(e)
def duty(filename):
	try:
		data = open_excel(filename)
		report_sheet = data.sheet_by_name(u"考勤日报表")
		nrows = report_sheet.nrows
		
		for rownum in range(1, nrows):
			
			
		print "seccess"
	except:
		print "fail to do work"

if __name__ == "__main__":
	if len(sys.argv) < 2:
		print "need input excel"
		return
	elif len(sys.argv) > 2:
		print "too many parameters"
		return
	else:
		duty(sys.argv[1])
