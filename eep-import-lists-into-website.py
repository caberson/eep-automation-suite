#!/usr/bin/python
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2013 Caber Chu

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

import sys, os, re

from decimal import *

from datetime import datetime
import string

import mysql.connector

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



excel_row_lo = ROWS_USED_BY_HEADING
excel_row_hi = 0 # do not hard code this
#excel_row_hi = 734 # remove later

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

def print_sheetnames():
	wb_eep = xlrd.open_workbook(RAW_EXCEL_FILE, on_demand=True, formatting_info=True)
	sheet_names = wb_eep.sheet_names()
	for i, name in enumerate(sheet_names):
		print i, name.strip().encode('utf-8')

def add_record(conn, tbl):
	cursor = conn.cursor()
	qry = ("TRUNCATE j17_" + tbl)

	
	print qryInsert
	
	cursor.execute(qryInsert)
	provinces = {
		u"四川": 0,
		u"貴州": 1,
		u"陝西": 2,
		u"廣西": 3,
		u"青海": 4,
		u"遼寧": 5,
	}

def clean_text(text):
	text = text.replace(u'（', '(');
	text = text.replace(u'）', ')');
	text = re.sub(r'\([^)]*\)', '', text)
	# text = re.sub(r'\（[]*\)', '', text)
	text = text.replace('*', '').strip()
	return text

def update_health_centers(conn):
	file = "/Users/cc/Documents/eep/schoolListing/healthCentersListing.xls"

	qryInsert = (
		"INSERT INTO j17_rooms "
		"(id, ordering, state, checked_out, name, address, donation_amount, donor, description, images, province, pdf) "
		"VALUES (%(id)s, %(ordering)s, 1, 0, %(name)s, %(address)s, %(donation_amount)s,"
		"%(donor)s, %(description)s, %(images)s, %(province)s, %(pdf)s)"
	)
	
	cursor = conn.cursor()
	qry = ("TRUNCATE j17_rooms")
	cursor.execute(qry)
	# cursor.commit()

	wb_eep = xlrd.open_workbook(file, on_demand=True, formatting_info=True)
	sh_eep = wb_eep.sheet_by_index(0)
	excel_row_lo = 2
	excel_row_hi = 33
	currentOrdering = 0
	for current_sheet_row_count, rx in enumerate(range(excel_row_lo, excel_row_hi)): #sh.nrows
		currentOrdering += 1
		# current_actual_excel_row_num = i + 2 #1 for heading and 1 for actual cell number
		#current_xlwt_excel_row_num = i + 1
		student_names = sh_eep.col_values(2, excel_row_lo, excel_row_hi)
		# 1 = school id
		# 2 = name
		# 3 = province
		# 4 = location
		# 5 = donation amount
		# 6 = donor chinese name
		# 7 = donor eng name
		# print cell_val.strip().encode('utf-8')
		# print sh_eep.cell_value(rx, 6).strip().encode('utf-8')
		
		#province = sh_eep.cell_value(rx, 3)
		#province_id = provinces[province]
		province_id = 1
		if province_id is None:
			province_id = '-------'
		#elif province == u"四川":
		#	pass
		else:
			pass #continue;
		
		amount = unicode(sh_eep.cell_value(rx, 3))
		if amount[0].isdigit():
			amount = '$' + str(int(float(amount)))
		
		donor = clean_text(sh_eep.cell_value(rx, 4))
		images = 'school_img_6.jpg#|#school_img_7.jpg#|#school_img_22.jpg#|#school_img-2_10.jpg'
		images = 'room_img.jpg'
		record = {
			'id': sh_eep.cell_value(rx, 0),
			'ordering': currentOrdering,
			'name': sh_eep.cell_value(rx, 1),
			'address': sh_eep.cell_value(rx, 2),
			'donation_amount': amount, #sh_eep.cell_value(rx, 5),
			'donor': donor, #re.sub(r'\([^)]*\)', '', sh_eep.cell_value(rx, 4)),
			'description': "",
			'province': province_id, # sh_eep.cell_value(rx, 3),
			#'province_id': province_id,
			'images': images,
			'pdf': 0,
		}
		
		print "-----------\n"
		print repr(record).decode('unicode-escape')
		cursor.execute(qryInsert, record)
		conn.commit()

	cursor.close()
	

def update_schools(conn):
	file = "/Users/cc/Documents/schoolListing/schoolConstructionListingByRegion.xls"

	cursor = conn.cursor()
	qry = ("TRUNCATE j17_schools")
	# cursor.execute(qry)

	qryInsert = (
		"INSERT INTO j17_schools "
		"(id, ordering, state, checked_out, name, address, donation_amount, donor, description, images, province, pdf) "
		"VALUES (%(id)s, %(ordering)s, 1, 0, %(name)s, %(address)s, %(donation_amount)s,"
		"%(donor)s, %(description)s, %(images)s, %(province)s, %(pdf)s)"
	)
		
	# print qryInsert
	
	provinces = {
		u"四川": 0,
		u"貴州": 1,
		u"陝西": 2,
		u"廣西": 3,
		u"青海": 4,
		u"遼寧": 5,
	}

	wb_eep = xlrd.open_workbook(file, on_demand=True, formatting_info=True)
	sh_eep = wb_eep.sheet_by_index(0)
	excel_row_lo = 2
	excel_row_hi = 907
	currentOrdering = 0
	for current_sheet_row_count, rx in enumerate(range(excel_row_lo, excel_row_hi)): #sh.nrows
		currentOrdering += 1
		# current_actual_excel_row_num = i + 2 #1 for heading and 1 for actual cell number
		#current_xlwt_excel_row_num = i + 1
		student_names = sh_eep.col_values(2, excel_row_lo, excel_row_hi)
		# 1 = school id
		# 2 = name
		# 3 = province
		# 4 = location
		# 5 = donation amount
		# 6 = donor chinese name
		# 7 = donor eng name
		cell_val = sh_eep.cell_value(rx, 2)
		cell_val = sh_eep.cell_value(rx, 3)

		# print cell_val.strip().encode('utf-8')

		# print sh_eep.cell_value(rx, 6).strip().encode('utf-8')
		
		province = sh_eep.cell_value(rx, 3)
		province_id = provinces[province]

		if province_id is None:
			province_id = '-------'
		#elif province == u"四川":
		#	pass
		else:
			pass #continue;
		
		donor = clean_text(sh_eep.cell_value(rx, 6))
		school_name = clean_text(sh_eep.cell_value(rx, 2))
		amount = unicode(sh_eep.cell_value(rx, 5))
		if amount[0].isdigit():
			amount = '$' + str(int(float(amount)))

		images = 'school_img_6.jpg#|#school_img_7.jpg#|#school_img_22.jpg#|#school_img-2_10.jpg'
		images = 'school_img_6.jpg'
		record = {
			'id': sh_eep.cell_value(rx, 1),
			'ordering': currentOrdering,
			'name': school_name, # sh_eep.cell_value(rx, 2),
			'address': sh_eep.cell_value(rx, 4),
			'donation_amount': amount, #sh_eep.cell_value(rx, 5),
			'donor': donor, # sh_eep.cell_value(rx, 6),
			'description': "",
			'province': province_id, # sh_eep.cell_value(rx, 3),
			#'province_id': province_id,
			'images': images,
			'pdf': 0,
		}
		
		print "-----------\n"
		print repr(record).decode('unicode-escape')
		# cursor.execute(qryInsert, record)
		# conn.commit()

	cursor.close()
	

# BEGIN MAIN ==================================================================
if __name__ == "__main__":
	print "File: ", file
	print "Platform: ", sys.platform

	conn = mysql.connector.connect(
		host='localhost',
		user='root',
		database='eep'
	)

	update_schools(conn)
	#update_health_centers(conn)
	
	conn.close()