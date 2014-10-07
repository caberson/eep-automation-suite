#!/usr/bin/python
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2011 Caber Chu

# d:\_cc\development\Python26\python.exe eep-merge-sheets-from-raw-excel.py 2011f_eep.xls 2

# cd /Users/cc/Projects/eepListFiles/
# ./eep-generate-lists.py  2011f_eep_combined.xls

# ./eep-merge-sheets-from-raw-excel.py ~/Documents/eep/2011f/2011f_eep.xls 11,12


import xlrd
import xlwt
import xlutils
from xlutils.styles import Styles
from xlutils.display import quoted_sheet_name
from xlutils.display import cell_display

import math

from xlutils.copy import copy
from xlutils.save import save

from xlwt import easyxf

import sys, os

from decimal import *

from datetime import datetime
import string

import eep_shared





# docs
# http://www.lexicon.net/sjmachin/xlrd.html
# http://groups.google.com/group/python-excel/browse_thread/thread/23a0b4d6be641755
# http://www.pythonexcels.com/2009/09/another-xlwt-example/
# http://www.python-excel.org/
# https://secure.simplistix.co.uk/svn/xlwt/trunk/xlwt/examples/xlwt_easyxf_simple_demo.py

#
# constants
#
OUTPUT_ENCODING = 'utf-8'
if sys.platform == 'win32':
	OUTPUT_ENCODING = 'big5'

#
# parameters
#
RAW_EXCEL_FILE = '2011_eep.xls'
COMBINE_SHEET_NUMBERS = 0
#CURRENT_YEAR = datetime.now().year

#
# global vars
#
ROWS_USED_BY_HEADING = 3

# 0 based, these can be moved to eep_shared.py
COL_REGION = 1
COL_LOCATION = 2
COL_SCHOOL = 3
COL_DONOR_BALANCE = 4
COL_STUDENT_NAME = 5
COL_SEX = 6
COL_GRADE = 7
COL_GRADUATION_YEAR = 8
COL_STUDENT_DONOR_ID = 9
COL_STUDENT_DONOR_NAME = 10
COL_STUDENT_DONOR_DONATION_AMOUNT_LOCAL = 11
COL_STUDENT_DONOR_DONATION_AMOUNT_US = 12
COL_COMMENT = 13	#N
COL_COMMENT_TW = 14	#O


SHEET_COLUMNS = {
	'DEFAULT': [COL_REGION, COL_LOCATION, COL_SCHOOL, COL_STUDENT_NAME, COL_SEX, COL_GRADUATION_YEAR, COL_STUDENT_DONOR_ID, COL_STUDENT_DONOR_NAME, COL_STUDENT_DONOR_DONATION_AMOUNT_LOCAL, COL_COMMENT],
	'1': [COL_REGION, COL_LOCATION, COL_SCHOOL, COL_STUDENT_NAME, COL_SEX, COL_GRADUATION_YEAR, COL_STUDENT_DONOR_ID, COL_STUDENT_DONOR_NAME, COL_STUDENT_DONOR_DONATION_AMOUNT_LOCAL, COL_COMMENT_TW],
	}


current_year = datetime.now().year
current_month = datetime.now().month



excel_row_lo = ROWS_USED_BY_HEADING
excel_row_hi = 0 # do not hard code this
#excel_row_hi = 734 # remove later

STYLES = {
	'CHINESE': xlwt.easyxf(u'font: name 宋体;'),
	'CELL_LISTING': xlwt.easyxf(u'font: name 宋体; align: wrap off, shrink_to_fit on, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
	'CELL_LISTING_WRAP': xlwt.easyxf(u'font: name 宋体; align: wrap on, shrink_to_fit off, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
	'CELL_LISTING_TITLE': xlwt.easyxf(u'font: name 宋体, bold on, height 280; align: wrap off, shrink_to_fit on; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
	'WARNING': xlwt.easyxf(u'font: name 宋体; pattern: pattern solid, fore-colour yellow;'),
	'ERROR': xlwt.easyxf(u'font: name 宋体; pattern: pattern solid, fore-colour red;'),
	}
"""
STYLES = {
	'CHINESE': xlwt.easyxf(u'font: name 新細明體;'),
	'CELL_LISTING': xlwt.easyxf(u'font: name 新細明體; align: wrap off, shrink_to_fit on, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
	'CELL_LISTING_WRAP': xlwt.easyxf(u'font: name 新細明體; align: wrap on, shrink_to_fit off, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
	'CELL_LISTING_TITLE': xlwt.easyxf(u'font: name 新細明體, bold on, height 280; align: wrap off, shrink_to_fit on; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
	'WARNING': xlwt.easyxf(u'font: name 新細明體; pattern: pattern solid, fore-colour yellow;'),
	'ERROR': xlwt.easyxf(u'font: name 新細明體; pattern: pattern solid, fore-colour red;'),
	}
"""
# determine last student row
def get_sheet_row_hi(sh_eep):
	#print excel_row_lo
	print 'Max Rows in Excel: ', sh_eep.nrows
	excel_row_hi = 0

	try:
		for rownum in range(excel_row_lo, sh_eep.nrows):#sh.nrows
			#print rownum, sh_eep.row_values(rownum)
			try:
				if (not sh_eep.cell(rownum, 0).value and not sh_eep.cell(rownum + 1, 0).value
					and not sh_eep.cell_value(rownum, 1) and not sh_eep.cell_value(rownum + 1, 1)
					and not sh_eep.cell_value(rownum, 2) and not sh_eep.cell_value(rownum + 1, 2)
					):
					excel_row_hi = rownum
					break;
			except:
				print 'Error occured trying to get sheet_row_hi'
				#excel_row_hi = rownum
				break;
	except:
		excel_row_hi = rownum
		#print rownum
	print 'Last Excel Row for sheet %s: %d' % (sh_eep.name.encode('utf-8'), excel_row_hi)
	
	
	return excel_row_hi

"""
# original.  can be deleted later on
def clean_text(val):
	#print isinstance(val)
	if type(val).__name__ in ['unicode']:
		val = val.strip()
		
	return val
"""

#
# create a new file
#
def combine_sheets():
	# open raw excel file to read
	wb_eep = xlrd.open_workbook(RAW_EXCEL_FILE, on_demand=True, formatting_info=True)

	# new workbook
	wb_new = xlwt.Workbook()
	sh_new = wb_new.add_sheet('final-data')
	sh_new.portrait = 0
	
	column_titles = ['region', 'location', 'school-na', 'student-name', 'sex', 'grad-yr', 'donor-id', 'donor-na', 'donate-amt', 'comment', 'ipt_odr_nr', 'auto-student-id', 'auto-donor-stu-cnt-id', 'scl-na-len']
	#columns = [COL_REGION, COL_LOCATION, COL_SCHOOL, COL_STUDENT_NAME, COL_SEX, COL_GRADUATION_YEAR, COL_STUDENT_DONOR_ID, COL_STUDENT_DONOR_NAME, COL_STUDENT_DONOR_DONATION_AMOUNT_LOCAL, COL_COMMENT]
	sheets = [int(x) for x in COMBINE_SHEET_NUMBERS.split(',')]
	#print sheets
	
	current_year = datetime.now().year

	"""
	style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
	style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
	ws.write(1, 0, datetime.now(), style1)
	ws.write(2, 2, xlwt.Formula("A3+B3"))
	"""
	#
	# create column titles
	#
	for cx, column_title in enumerate(column_titles):
		sh_new.write(0, cx, column_title)	
	
	i = 0
	total_rows_combined = 0
	for sheet_count, sheet_index in enumerate(sheets):
		sh_eep = wb_eep.sheet_by_index(sheet_index)
		excel_row_hi = get_sheet_row_hi(sh_eep)
		total_rows_combined += excel_row_hi - excel_row_lo
		print 'Sheet: ', sheet_index, ' Rows: ', excel_row_hi
		
		# determine which columns to use
		try:
			columns = SHEET_COLUMNS[str(sheet_count)]
		except:
			columns = SHEET_COLUMNS['DEFAULT'];

		student_names = sh_eep.col_values(COL_STUDENT_NAME, excel_row_lo, excel_row_hi)
		#for name in student_names:
			#print name

		#i = 0
		for current_sheet_row_count, rx in enumerate(range(excel_row_lo, excel_row_hi)): #sh.nrows
			current_actual_excel_row_num = i + 2 #1 for heading and 1 for actual cell number
			current_xlwt_excel_row_num = i + 1
			# row id
			#sh_new.write(i + 1, 0, i+1)
			
			# import order id
			sh_new.write(current_xlwt_excel_row_num, len(columns) + 0, i+1)
			
			# auto student id
			formula = 'COUNTIF(C$1:C%d, C%d)' % (current_actual_excel_row_num, current_actual_excel_row_num)
			sh_new.write(current_xlwt_excel_row_num, len(columns) + 1, xlwt.Formula(formula))
			
			# auto donor student count
			formula = 'COUNTIF(G$1:G%d, G%d)' % (current_actual_excel_row_num, current_actual_excel_row_num)
			sh_new.write(current_xlwt_excel_row_num, len(columns) + 2, xlwt.Formula(formula))
			
			# school name len
			formula = 'LEN(C%d)' % (current_actual_excel_row_num)
			sh_new.write(current_xlwt_excel_row_num, len(columns) + 3, xlwt.Formula(formula))
			
			for current_column, cx in enumerate(columns):
				use_warning = False
				use_error = False
				if cx > 0:
					#print rx, cx
					cell_val = sh_eep.cell_value(rx, cx)
					
					# determine styles to use
					if cell_val == "":
						use_warning = True
					else:
						cell_style = STYLES['CHINESE']
					
					# student name error checking
					if cx == COL_STUDENT_NAME:
						use_warning = cell_val in student_names[:current_sheet_row_count]
						
						# try to find if same student name exists before
						try:
							found_index = student_names[:current_sheet_row_count].index(cell_val) + ROWS_USED_BY_HEADING
							
							# check for possible duplicate
							if (sh_eep.cell_value(rx, COL_REGION) == sh_eep.cell_value(found_index, COL_REGION)
								and sh_eep.cell_value(rx, COL_LOCATION) == sh_eep.cell_value(found_index, COL_LOCATION)
								and sh_eep.cell_value(rx, COL_SCHOOL) == sh_eep.cell_value(found_index, COL_SCHOOL)):
								use_error = True
								print '\tPOSSIBLE ERROR: ', cell_val, ' CellRow:', rx, ' PreviousFoundRow:', found_index
								
						except: #success
							pass

					# check if graduation year is past current year
					#print cell_val
					if cx == COL_GRADUATION_YEAR:
						if cell_val == "":
							use_warning = True
						elif unicode(current_year) in unicode(cell_val):
							use_warning = True
						elif cell_val < current_year:
							use_error = True
					
					
					# if currentVal != previousVal, mark warning
					if cx in [COL_REGION, COL_LOCATION, COL_SCHOOL, COL_STUDENT_DONOR_NAME]:
						if current_xlwt_excel_row_num > 2:
							cell_val_prev = sh_eep.cell_value(rx - 1, cx)
							if cell_val != cell_val_prev:
								use_warning = True
					
					if use_warning:
						cell_style = STYLES['WARNING']
					if use_error:
						cell_style = STYLES['ERROR']
					
					# write value
					sh_new.write(current_xlwt_excel_row_num, current_column, eep_shared.clean_text(cell_val), cell_style)
				
			i += 1
	print 'Total Rows: ', total_rows_combined
	
	wb_new.save(eep_shared.DESTINATION_DIR + eep_shared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA + '_combined.xls')


def print_sheetnames():
	wb_eep = xlrd.open_workbook(RAW_EXCEL_FILE, on_demand=True, formatting_info=True)
	sheet_names = wb_eep.sheet_names()
	for i, name in enumerate(sheet_names):
		print i, name.strip().encode('utf-8')


# BEGIN MAIN ============================================================================================================
if __name__ == "__main__":
	print sys.platform
	try:
		raw_excel_file_na = sys.argv[1]
		RAW_EXCEL_FILE = raw_excel_file_na
	except:
		pass
	
	try:
		sheet_numbers = sys.argv[2]
		COMBINE_SHEET_NUMBERS = sheet_numbers	
	except:
		pass
	
	if len(sys.argv) < 3:
		#raw_excel_file_na_suggested = str(current_year) + ('s' if current_month < 6 else 'f') + '_eep.xls'
		#print current_year
		#print current_month
		print 'Usage:\neep-merge-sheets-from-raw-excel.py %s combine_sheet_numbers' % (eep_shared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA + '.xls')
		
		if len(sys.argv) == 2:
			print_sheetnames()
		
		sys.exit(1)
	
	# create destination folders if needed
	eep_shared.create_required_dirs()	

	#RAW_EXCEL_FILE = raw_excel_file_na
	#COMBINE_SHEET_NUMBERS = sheet_numbers
	#print raw_excel_file_na
	#print COMBINE_SHEET_NUMBERS
	
	#wb_eep = xlrd.open_workbook(raw_excel_file_na, on_demand=True, formatting_info=True) #, 
	#sh_eep = wb_eep.sheet_by_index('17')
	#sh_eep.portrait = 0
	#print "Total Rows: ", sh_eep.nrows
	
	
	combine_sheets()