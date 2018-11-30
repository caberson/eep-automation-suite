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

USER_DIR = os.path.expanduser('~')
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

PHOTOS_ORIGINAL_FOLDER_NAME = 'eep_photos_original'
PHOTOS_CROPPED_FOLDER_NAME = 'eep_photos_cropped'

STUDENT_PHOTOS_ORIGINAL_DIR = DESTINATION_DIR + PHOTOS_ORIGINAL_FOLDER_NAME
STUDENT_PHOTOS_CROPPED_DIR = DESTINATION_DIR + PHOTOS_CROPPED_FOLDER_NAME

DIR_APP = os.getcwd()
DIR_DATA = os.path.join(DIR_APP, 'data')
DIR_EEP_PHOTOS_ORIGINAL_DEFAULT = os.path.join(DIR_DATA, PHOTOS_ORIGINAL_FOLDER_NAME)
DIR_EEP_PHOTOS_CROPPED_DEFAULT = os.path.join(DIR_DATA, PHOTOS_CROPPED_FOLDER_NAME)

TEMPLATES_DIR = os.path.join(DIR_APP, '..', 'templates')
DIR_ASSETS = os.path.join(DIR_APP, 'assets')
DIR_OUTPUT = os.path.join(DIR_APP, 'output')

DONOR_REPORT_FOLDER_NAME = 'donor_reports'
DONOR_REPORT_DIR = os.path.join(DIR_OUTPUT, DONOR_REPORT_FOLDER_NAME)

def build_english_year_code(year, month):
    season = 's' if month <= 6 else 'f'
    return '%s%s' % (str(year), season)

def build_chinese_year_code(year, month):
    season = u'春' if month <= 6 else u'秋'
    return '%s%s' % (str(year), season)

def get_config(config_file=None):
    import ConfigParser
    if not config_file:
        CONFIG_NAME = 'eep.cfg'
        config_file = os.path.join(os.getcwd(), 'etc', CONFIG_NAME)

    config = ConfigParser.SafeConfigParser()
    config.read(config_file)
    return config


OUTPUT_ENCODING = 'utf-8'
if sys.platform == 'win32':
    OUTPUT_ENCODING = 'big5'
