#!/usr/bin/python
#-*-coding: utf-8-*-
import os
import sys 
import time
import xlrd
import xlwt
import xlutils
from xlutils.copy import copy
import codecs
list_eal= []
list_noeal = []
NORMAL = 1
LATE = 2
LEAVE_EARLY = 3
EXCEPTION = 4
ABSENT = 5
def noela_start(start_time):
	try:
		start_li = start_time.split(":")
		if(int(start_li[0]) < 9):
			return NORMAL
		elif(int(start_li[0]) == 9 ):
			if(int(start_li[1] > 6)):
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
		if(int(leave_li[0])) > 17:
			return NORMAL
		elif(int(leave_li[0])) < 17:
			return LEAVE_EARLY
		else:
			if(int(leave_li[1]) >= 25):
				return NORMAL
			else:
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
		if(sum_min >= 540):
			return NORMAL
		elif(sum_min >= 480):
			return LEAVE_EARLY
		else:
		 	return ABSENT
	except:
	  return EXCEPTION
def open_excel(filename):
	try:
		data = xlrd.open_workbook(filename)
		return data
	except Exception,e:
	 	print str(e)


def get_list(filename):
	data = open_excel(filename)
	print"123"
	sheet1 = data.sheet_by_name("Sheet1")
	print "234"
	nrows = sheet1.nrows

	for rownum in range(1, nrows):
		list_eal.append(sheet1.row_values(rownum)[0])
		list_noeal.append(sheet1.row_values(rownum)[1])

def converse_result(value):
	if(value == NORMAL):
		return ""
	elif(value == LATE):
		return u"迟到"
	elif(value == LEAVE_EARLY):
		return u"早退"
	elif(value == ABSENT):
		return u"旷工"
	else:
	 	return u"异常"
	

def duty(filename):
	if 1:
		data = open_excel(filename)
		report_sheet = data.sheet_by_name(u"考勤日报表")
		data = xlutils.copy.copy(data)
		report_sheet_copy = data.get_sheet(0)
		nrows = report_sheet.nrows
		ncols = report_sheet.ncols
		f = codecs.open("result.csv","wb+",encoding = "gb2312")
		for rownum in range(0, nrows):	
			rowval = report_sheet.row_values(rownum)			
			write_value = ''
			value = ""
			for colnum in range(0, ncols):
				tmp_value = ""
				if(type(rowval[colnum])) is unicode :
					#tmp_value = rowval[colnum]
					tmp_value = rowval[colnum]
					
				if(type(rowval[colnum])) is str:
				 	tmp_value = rowval[colnum].decode()
				else:
					tmp_value = ""	
				#write_value =write_value + tmp_value +u","
				f.write(tmp_value)
				f.write(",")
				print rownum
			if(rowval[4] == "A01"):
				if(rowval[1] in list_noeal):
					start_value = noela_start(rowval[5])
					leave_value = noela_leave(rowval[6])
					value = converse_result(start_value) +" "+ converse_result(leave_value)
				elif(rowval[1] in list_noeal):
					start_value = ela_start(rowval[5])
					leave_value = ela(rowval[6])
					value = converse_result(start_value)+" " +converse_result( leave_value)
				else:
		 			value = u"不再考勤人员表中"		 
				#write_value = write_value+value.encode("gb2312")+","+"\n"
			f.write(value)
			f.write("\n")
#report_sheet_copy.write(rownum, 10, value)
		f.close()
		data.save("result.xlsx")			
		print "seccess"
#except Exception,e:
#		print e
#		print "fail to do work"

if __name__ == "__main__":
	if len(sys.argv) < 3:
		print "useage: duty.py 考勤人员.xlsx 考勤日报表.xlsx"
	elif len(sys.argv) > 3:
		print "too many parameters"
	else:
		get_list(sys.argv[1])
		duty(sys.argv[2])
		print list_eal
		print list_noeal
