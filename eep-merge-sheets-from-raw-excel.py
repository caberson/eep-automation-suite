#!/usr/bin/python
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2011 Caber Chu
"""Merge two sheets in an EEP Excel data file.

http://www.lexicon.net/sjmachin/xlrd.html
http://groups.google.com/group/python-excel/browse_thread/thread/23a0b4d6be641755
http://www.pythonexcels.com/2009/09/another-xlwt-example/
http://www.python-excel.org/
https://secure.simplistix.co.uk/svn/xlwt/trunk/xlwt/examples/xlwt_easyxf_simple_demo.py

List out all sheets found in the Excel file.
>>> ./eep-merge-sheets-from-raw-excel.py ~/Documents/eep/2011f/20114_eep.xls

Merge sheet 11 and 12.  Sheets are 0 based.
>>> ./eep-merge_sheets-from-raw-excel.py ~/Documents/eep/2011f/20114_eep.xls --sheetnums 11 12
"""

# Standard module imports.
import sys
import os
import math

# 3rd party module imports.

# Custom module imports.
import eepshared
import eeputil
from eepsheet import EepSheet

import roster.sortsheet
import roster.mergesheet
import xlsstyles


def get_argparse():
    """Get cmd line argument parser.
    """
    import argparse
    parser = argparse.ArgumentParser(
            description='Merges Excel sheets into a new file.')
    default_excel_file_na = '{}.xls'.format(
        eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA
    )
    default_excel_file = os.path.join(
        eepshared.DESTINATION_DIR, default_excel_file_na
    )

    parser.add_argument(
        'rawexcelfile',
        nargs='?',
        default=default_excel_file,
        help='Source Excel file name (default: %(default)s)',
    )
    parser.add_argument(
        '--sheetnums',
        nargs='*',
        type=int,
        help='Sheets numbers to merge.  Sheet number starts from 0.'
    )

    return parser

# BEGIN MAIN ==================================================================
if __name__ == "__main__":
    parser = get_argparse()
    args = parser.parse_args()
    raw_excel_file = args.rawexcelfile
    # print sys.platform

    # If 'sheetnums' is not specified, print out the sheets in the src Excel file.
    if not args.sheetnums:
        parser.print_help()
        roster.mergesheet.print_sheetnames(raw_excel_file)
        sys.exit(1)

    # Create destination folders if needed
    eeputil.create_required_dirs()

    out_file = '/Users/cc/Documents/eep/2015f/2015f_eep_combined.xls'
    out_file = roster.mergesheet.combine_sheets(raw_excel_file, args.sheetnums)
    print "Out file:", out_file

    data = roster.sortsheet.sort(out_file)
    sorted_out_file = roster.sortsheet.save(data)
    print "Out file (sorted): ", sorted_out_file