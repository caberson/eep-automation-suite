#!/usr/bin/env python
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
SUGGESTED_COMBINED_EXCEL_FILE_BASE_NA = (str(current_year) +
    ('s' if current_month <= 6 else 'f') + '_eep_combined_sorted')
DESTINATION_DIR = (
    EEP_DOC_DIR +
    SUGGESTED_FILE_DESTINATION_FOLDER_NAME + '/'
)

INSPECTION_DOC_DIR_NAME = 'documents_inspect'
STUDENT_NAME_LABELS_DIR_NAME = 'student_name_labels'

INSPECTION_DOCUMENTS_DESTINATION_DIR = os.path.join(DESTINATION_DIR, INSPECTION_DOC_DIR_NAME)
STUDENT_NAME_LABELS_DIR = os.path.join(DESTINATION_DIR, STUDENT_NAME_LABELS_DIR_NAME)

PHOTOS_ORIGINAL_FOLDER_NAME = 'eep_photos_original'
PHOTOS_CROPPED_FOLDER_NAME = 'eep_photos_cropped'

STUDENT_PHOTOS_ORIGINAL_DIR = os.path.join(DESTINATION_DIR, PHOTOS_ORIGINAL_FOLDER_NAME)
STUDENT_PHOTOS_CROPPED_DIR = os.path.join(DESTINATION_DIR, PHOTOS_CROPPED_FOLDER_NAME)

DIR_APP = os.getcwd()
DIR_DATA = os.path.join(DIR_APP, 'data')
DIR_EEP_PHOTOS_ORIGINAL_DEFAULT = os.path.join(DIR_DATA, PHOTOS_ORIGINAL_FOLDER_NAME)
DIR_EEP_PHOTOS_CROPPED_DEFAULT = os.path.join(DIR_DATA, PHOTOS_CROPPED_FOLDER_NAME)

TEMPLATES_DIR = os.path.join(DIR_APP, '..', 'templates')
DIR_ASSETS = os.path.join(DIR_APP, 'assets')
DIR_OUTPUT = os.path.join(DIR_APP, 'output')

DONOR_REPORT_FOLDER_NAME = 'donor_reports'
DONOR_REPORT_DIR = os.path.join(DESTINATION_DIR, DONOR_REPORT_FOLDER_NAME)

def get_student_photos_cropped_dir(yr_code=None):
    if yr_code is None:
        yr_code = build_english_year_code()
    return os.path.join(EEP_DOC_DIR, yr_code, PHOTOS_CROPPED_FOLDER_NAME)

def get_student_photos_original_dir(yr_code=None):
    if yr_code is None:
        yr_code = build_english_year_code()
    return os.path.join(EEP_DOC_DIR, yr_code, PHOTOS_ORIGINAL_FOLDER_NAME)

def get_donor_report_dir(yr_code=None):
    if yr_code is None:
        yr_code = build_english_year_code()
    return os.path.join(EEP_DOC_DIR, yr_code, DONOR_REPORT_FOLDER_NAME)

def get_exl_file_base_name(yr=None, mo=None):
    yr_code = build_english_year_code(yr, mo)
    return '{}_eep'.format(yr_code)

def build_english_year_code(year=None, month=None):
    if year is None:
        year = datetime.now().year
    if month is None:
        month = datetime.now().month
    
    season = 's' if month <= 6 else 'f'
    return '%s%s' % (str(year), season)

def build_chinese_year_code(year=None, month=None):
    if year is None:
        year = datetime.now().year
    if month is None:
        month = datetime.now().month

    season = u'春' if month <= 6 else u'秋'
    return u'{}{}'.format(year, season)

def build_chinese_year_code_short(year=None, month=None):
    if year is None:
        year = datetime.now().year
    
    if len(str(year)) > 2:
        year = str(year)[-2:]

    return build_chinese_year_code(year, month)

def get_config(config_file=None):
    import configparser
    if not config_file:
        CONFIG_NAME = 'eep.cfg'
        config_file = os.path.join(os.getcwd(), 'etc', CONFIG_NAME)

    config = configparser.ConfigParser()
    config.read(config_file)
    return config

OUTPUT_ENCODING = 'utf-8'
if sys.platform == 'win32':
    OUTPUT_ENCODING = 'big5'

CURRENT_SEASON_CHI = build_chinese_year_code(current_year, current_month)
CURRENT_SEASON_CHI_SHORT = build_chinese_year_code_short(current_year, current_month)
