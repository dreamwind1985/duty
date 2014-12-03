#!/usr/bin/python
#-*-coding: utf-8-*-
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import time
import xlrd
import time
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
			if(int(start_li[1]) > 6):
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
		if(int(leave_li[0]) > 17):
			return NORMAL
		elif(int(leave_li[0]) < 17):
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
		elif(int(start_li[0]) >= 10 and int(start_li[0]) <= 11):
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
			return ""
		elif(sum_min >= 480):
			leave_time = 540 - sum_min 
			return u"早退" + str(leave_time)+u"分钟"
		else:
		 	return u"旷工"
	except:
	  return u"异常"
def open_excel(filename):
	try:
		data = xlrd.open_workbook(filename)
		return data
	except Exception,e:
	 	print str(e)


def get_list(filename):
	data = open_excel(filename)
#print"123"
	sheet1 = data.sheet_by_name("Sheet1")
#	print "234"
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
		nrows = report_sheet.nrows
		ncols = report_sheet.ncols
		f = codecs.open("result.csv","wb+")
		for rownum in range(0, nrows):	
			rowval = report_sheet.row_values(rownum)			
			write_value = ''
			value = ""
			for colnum in range(0, ncols):
				if(type(rowval[colnum])) is unicode :
					tmp_value = rowval[colnum]
					try:
						tmp_value = rowval[colnum].decode("utf-8")
						tmp_value = tmp_value.encode("gb2312")
					except:
						tmp_value = ""	
					f.write(tmp_value)
				elif ((type(rowval[colnum])) is str) :
				 	tmp_value = rowval[colnum]
					f.write(tmp_value)
				else:
					tmp_value = ""	
					f.write(tmp_value)
				f.write(",")
			if(rowval[4] == "A01") :
				if(rowval[1] in list_noeal):
					start_value = noela_start(rowval[5])
					leave_value = noela_leave(rowval[6])
					value = converse_result(start_value) +" "+ converse_result(leave_value)
				elif(rowval[1] in list_eal):
					start_value = ela_start(rowval[5])
					leave_value = ela(rowval[5],rowval[6])
					value = converse_result(start_value)+" " +leave_value
				else:
		 			value = u"不在考勤人员表中"		 
			value = value.decode("utf-8").encode("gb2312")
			f.write(value)
			f.write("\n")
		f.close()
		print "seccess"
'''
if __name__ == "__main__":
	if len(sys.argv) < 3:
		print "useage: duty.py 考勤人员.xlsx 考勤日报表.xlsx"
	elif len(sys.argv) > 3:
		print "too many parameters"
	else:
		get_list(sys.argv[1])
		duty(sys.argv[2])
#print list_eal
#print list_noeal
'''
if __name__ == "__main__":
	if os.path.isfile(u"考勤人员.xlsx") and os.path.isfile(u"考勤日报表0.xlsx"):
		get_list(u"考勤人员.xlsx")
		duty(u"考勤日报表0.xlsx")
	else:
		print "考勤人员.xlsx 或者考勤日报表0.xlsx不存在"
		time.sleep(5)	
