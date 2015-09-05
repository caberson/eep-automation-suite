#!/usr/bin/python2.7
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2011 Caber Chu

import os

import xlwt
import xlrd
import eepshared
import xlsstyles

def sort(raw_excel_file, sheet_index=0):
	try:
		wb_eep = xlrd.open_workbook(
			raw_excel_file, on_demand=True, formatting_info=True
		)
	except:
		print "Error opening file ", raw_excel_file
		return

	src_sheet = wb_eep.sheet_by_index(sheet_index)

	tmp = []
	for i in range(1, src_sheet.nrows):
		row = src_sheet.row_slice(i)

		# print row
		tmp2 = []
		for x in row:
			cell_val = x.value
			tmp2.append(cell_val)
			# print cell_val
			# print len(cell_val)

		# print tmp2
		tmp.append(tmp2)

	# print tmp

	tmp.sort(key = lambda row: ([row[13], row[2], row[6]]))

	return tmp

def get_new_wb():
	wb_new = xlwt.Workbook()
	new_sheet = wb_new.add_sheet('final-data')
	new_sheet.portrait = 0

	column_titles = [
		'region', 'location', 'school-na', 'student-name', 'sex', 'grad-yr',
		'donor-id', 'donor-na', 'donate-amt', 'comment', 'ipt_odr_nr',
		'auto-student-id', 'auto-donor-stu-cnt-id', 'scl-na-len'
    ]

	for cx, column_title in enumerate(column_titles):
		new_sheet.write(0, cx, column_title)

	return wb_new

donor_student_cnt_trackers = {}
def get_auto_donor_stu_cnt_id(donor_id):
	if not donor_id:
		return ''

	donor_student_cnt_trackers.setdefault(donor_id, 0)

	donor_student_cnt_trackers[donor_id] = (
		donor_student_cnt_trackers.get(donor_id) + 1
	)
	# print donor_student_cnt_trackers
	return donor_student_cnt_trackers.get(donor_id)


def save(data, out_file=None):
	if out_file is None:
		out_file = os.path.join(
			eepshared.DESTINATION_DIR,
            eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA + '_combined_sorted.xls'
		)

	wb_new = get_new_wb()
	new_sheet = wb_new.get_sheet(0)

	last_student_school = ''
	auto_student_id = 0
	for i, row_data in enumerate(data):
		row = i + 1
		school = row_data[2]
		donor_id = row_data[6]
		if donor_id:
			donor_id = int(donor_id)

		if school == last_student_school:
			auto_student_id += 1
		else:
			auto_student_id = 1

		# Go through columns
		for cx, value in enumerate(row_data):
			# determine auto-student-id
			if cx == 11:
				value = auto_student_id

			# determine auto-donor-stu-cnt-id
			if cx == 12:
				value = get_auto_donor_stu_cnt_id(donor_id)
			new_sheet.write(row, cx, value, xlsstyles.STYLES['CHINESE'])

		last_student_school = school

	wb_new.save(out_file)

	return out_file