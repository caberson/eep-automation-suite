#!/usr/bin/python2.7
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2011 Caber Chu

import sys
import os
import string
import re
from datetime import datetime

import eeputil

USER_DIR = eeputil.get_current_user_dir()
EEP_DOC_DIR = os.getenv(
    'EEP_DOC_DIR',
    os.path.join(USER_DIR, 'Documents/eep')
)
current_year = datetime.now().year
current_month = datetime.now().month
SUGGESTED_FILE_DESTINATION_FOLDER_NAME = ('/' + str(current_year) +
	('s' if current_month <= 6 else 'f'))
SUGGESTED_RAW_EXCEL_FILE_BASE_NA = (str(current_year) +
	('s' if current_month <= 6 else 'f') + '_eep')
DESTINATION_DIR = (
	EEP_DOC_DIR +
	SUGGESTED_FILE_DESTINATION_FOLDER_NAME + '/'
)

INSPECTION_DOCUMENTS_DESTINATION_DIR = DESTINATION_DIR + 'documents_inspection/'

STUDENT_NAME_LABELS_DIR = DESTINATION_DIR + 'student_name_labels/'

OUTPUT_ENCODING = 'utf-8'
if sys.platform == 'win32':
	OUTPUT_ENCODING = 'big5'

if __name__ == "__main__":
	create_required_dirs()