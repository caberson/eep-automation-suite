#!/usr/bin/python2.7
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2011 Caber Chu

import sys
import os
import string
import re
from datetime import datetime

USER_DIR = os.path.expanduser('~')
EEP_ROOT_DIR = os.path.join(USER_DIR, 'Documents/eep')
current_year = datetime.now().year
current_month = datetime.now().month
SUGGESTED_FILE_DESTINATION_FOLDER_NAME = ('/' + str(current_year) +
	('s' if current_month <= 6 else 'f'))
SUGGESTED_RAW_EXCEL_FILE_BASE_NA = (str(current_year) +
	('s' if current_month <= 6 else 'f') + '_eep')
DESTINATION_DIR = (EEP_ROOT_DIR +
	SUGGESTED_FILE_DESTINATION_FOLDER_NAME + '/')

INSPECTION_DOCUMENTS_DESTINATION_DIR = DESTINATION_DIR + 'documents_inspection/'

STUDENT_NAME_LABELS_DIR = DESTINATION_DIR + 'student_name_labels/'

OUTPUT_ENCODING = 'utf-8'
if sys.platform == 'win32':
    OUTPUT_ENCODING = 'big5'

def clean_text(val):
	#print isinstance(val)
	if type(val).__name__ in ['unicode']:
		val = val.strip()
		val = string.replace(val, "*", "")

		val = string.replace(val, u"）", ")")
		val = string.replace(val, u'（', '(')
		
		val = string.replace(val, u'，', ',')
		val = string.replace(val, u'。', '.')
		
		# remove empty space?
		val = string.replace(val, u' ', '')
	return val

# create directory if not exists
def create_dir_if_not_exists(dir):
	if not os.path.exists(dir):
		print 'Created ' + dir
		os.makedirs(dir)

# Create required dirs
def create_required_dirs():
	print EEP_ROOT_DIR
	print INSPECTION_DOCUMENTS_DESTINATION_DIR
	print INSPECTION_DOCUMENTS_DESTINATION_DIR + 'checkinglist'
	print INSPECTION_DOCUMENTS_DESTINATION_DIR + 'receivinglist'
	print INSPECTION_DOCUMENTS_DESTINATION_DIR + 'lettersubmitlist'
	print STUDENT_NAME_LABELS_DIR

	create_dir_if_not_exists(EEP_ROOT_DIR)
	create_dir_if_not_exists(INSPECTION_DOCUMENTS_DESTINATION_DIR)
	create_dir_if_not_exists(INSPECTION_DOCUMENTS_DESTINATION_DIR + 'checkinglist')
	create_dir_if_not_exists(INSPECTION_DOCUMENTS_DESTINATION_DIR + 'receivinglist')
	create_dir_if_not_exists(INSPECTION_DOCUMENTS_DESTINATION_DIR + 'lettersubmitlist')	

	create_dir_if_not_exists(STUDENT_NAME_LABELS_DIR)

def remove_parenthesis_content(val, replaced = []):
	"""Remove parenthesis and inner strings from an unicode string."""
	if type(val).__name__ in ['unicode']:
		val = val.strip()
		val = string.replace(val, "*", "")
		val = string.replace(val, u"）", ")")
		val = string.replace(val, u'（', '(')
		
		if ')' in val:
			myRE = re.compile(r'(\(.*\))', re.U | re.I)
			m = myRE.findall(val)
		
			# print "VAL:: " + val.encode(OUTPUT_ENCODING)
			if m:
				if len(m[0]) > 4:
					replaced.append(m[0])
					#print '**', myRE.sub('', val)
					val = myRE.sub('', val)
					# print "new val: " + val.encode(OUTPUT_ENCODING)
			# print ' ', m.group(0)
	return val

def get_chinese_title_for_time(time=None):
	"""Returns Chinese title for the given time."""
	if time is None:
		time = datetime.now()

	sheet_title_base = u'{}年{}季'.format(
		time.year, (u'春' if time.month <= 6 else u'秋')
	)

	return sheet_title_base
	
if __name__ == "__main__":
	create_required_dirs()