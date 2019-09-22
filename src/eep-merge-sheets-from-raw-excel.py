#!/usr/bin/env python
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

# 3rd party module imports.

# Custom module imports.
import eepshared
import eeputil

import roster.sortsheet
import roster.mergesheet

from util.logger import logger

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
        eepshared.DESTINATION_DIR,
        default_excel_file_na
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

def main():
    """Main runner if invoked directly
    """
    parser = get_argparse()
    args = parser.parse_args()
    raw_excel_file = args.rawexcelfile
    print("Excel File: %s" % raw_excel_file)
    print("Destination Folder: %s" % eepshared.DESTINATION_DIR)
    # print sys.platform

    # Create destination folders if needed
    eeputil.create_required_dirs()

    # If 'sheetnums' is not specified, print out the sheets in the src Excel file.
    if not args.sheetnums:
        parser.print_help()
        roster.mergesheet.print_sheetnames(raw_excel_file)

        # Try to find latest sheet ids
        found_sheetnums = roster.mergesheet.find_latest_sheet_ids(raw_excel_file)
        use_found = raw_input('Use found sheetnums: {} (y/n)? '.format(found_sheetnums))
        if use_found.strip().lower() == 'y':
            args.sheetnums = found_sheetnums
        else:
            sys.exit(1)

    # Generate merged file
    out_file_name = os.path.join(
        eepshared.DESTINATION_DIR,
        eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA + '_combined.xls'
    )
    combined_file = roster.mergesheet.combine_sheets(
        raw_excel_file,
        args.sheetnums,
        out_file_name
    )
    logger.debug("Output file: {}".format(combined_file))

    # Generate
    data = roster.sortsheet.sort(combined_file)
    out_file_name = os.path.join(
        eepshared.DESTINATION_DIR,
        eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA + '_combined_sorted.xls'
    )
    sorted_icombined_out_file = roster.sortsheet.save(data, out_file_name)
    logger.debug("Output file (sorted): {}".format(sorted_icombined_out_file))


# BEGIN MAIN ==================================================================
if __name__ == "__main__":
    main()
