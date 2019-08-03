#!/usr/bin/python2.7
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2011 Caber Chu

import sys
import os
import string
import re
from datetime import datetime
import eepshared

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
def create_required_dirs(yr=None, mo=None):
    join = os.path.join
    # print eepshared.EEP_DOC_DIR
    # print eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR
    # print eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR + 'checkinglist'
    # print eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR + 'receivinglist'
    # print eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR + 'lettersubmitlist'
    # print eepshared.STUDENT_NAME_LABELS_DIR
    # print eepshared.STUDENT_PHOTOS_ORIGINAL_DIR
    # print eepshared.STUDENT_PHOTOS_CROPPED_DIR

    if yr is None:
        yr = datetime.now().year
    if mo is None:
        mo = datetime.now().month
    
    season_dir = eepshared.build_english_year_code(yr, mo)
    base_dir = join(eepshared.EEP_DOC_DIR, season_dir)
    ins_doc_dst_dir = join(base_dir, eepshared.INSPECTION_DOC_DIR_NAME)
    required_dirs = [
        eepshared.EEP_DOC_DIR,
        base_dir,
        ins_doc_dst_dir,
        join(ins_doc_dst_dir, 'checkinglist'),
        join(ins_doc_dst_dir, 'receivinglist'),
        join(ins_doc_dst_dir, 'lettersubmitlist'),
        join(base_dir, eepshared.STUDENT_NAME_LABELS_DIR_NAME),
        join(base_dir, eepshared.PHOTOS_ORIGINAL_FOLDER_NAME),
        join(base_dir, eepshared.PHOTOS_CROPPED_FOLDER_NAME),
        join(base_dir, eepshared.DONOR_REPORT_FOLDER_NAME),
    ]
    for d in required_dirs:
        print(d)
        create_dir_if_not_exists(d)
    

    # create_dir_if_not_exists(eepshared.EEP_DOC_DIR)
    # create_dir_if_not_exists(eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR)
    # create_dir_if_not_exists(eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR + 'checkinglist')
    # create_dir_if_not_exists(eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR + 'receivinglist')
    # create_dir_if_not_exists(eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR + 'lettersubmitlist')

    # create_dir_if_not_exists(eepshared.STUDENT_NAME_LABELS_DIR)

    # create_dir_if_not_exists(eepshared.STUDENT_PHOTOS_ORIGINAL_DIR)
    # create_dir_if_not_exists(eepshared.STUDENT_PHOTOS_CROPPED_DIR)

    # create_dir_if_not_exists(eepshared.DONOR_REPORT_DIR)

def remove_parenthesis_content(val, replaced=[], whitelist=[], blacklist=[]):
    """Remove parenthesis and inner strings from an unicode string."""
    check_len = 4

    if type(val).__name__ in ['unicode']:
        val = val.strip()
        val = string.replace(val, "*", "")
        val = string.replace(val, u"）", ")")
        val = string.replace(val, u'（', '(')

        if ')' in val:
            blacklisted = any(x in val for x in blacklist)

            if any(x in val for x in whitelist) and not blacklisted:
                return val

            myRE = re.compile(r'(\(.*\))', re.U | re.I)
            m = myRE.findall(val)

            # print "VAL:: " + val.encode(OUTPUT_ENCODING)
            check_len = 2 if blacklisted else 4
            if m:
                if len(m[0]) > check_len:
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

def load_local_site_packages():
    # adds current site-package folder path
    current_folder = os.path.realpath(os.path.abspath(os.path.split(inspect.getfile( inspect.currentframe() ))[0]))
    local_site_packages_folder = current_folder + '/site-packages'
    if local_site_packages_folder not in sys.path and os.path.exists(local_site_packages_folder):
        sys.path.insert(0, local_site_packages_folder)