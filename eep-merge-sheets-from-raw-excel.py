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
>>> ./eep-merge_sheets-from-raw-excel.py ~/Documents/eep/2011f/20114_eep.xls 11 12
"""

# Standard module imports.
import sys
import os
import string
import math
from decimal import *
from datetime import datetime

# 3rd party module imports.
import xlrd
import xlwt
from xlwt import easyxf
import xlutils
from xlutils.styles import Styles
from xlutils.display import quoted_sheet_name
from xlutils.display import cell_display
from xlutils.copy import copy
from xlutils.save import save

# Custom module imports.
import eepshared
import eeputil
from eepsheet import EepSheet

#
# global vars
#
ROWS_USED_BY_HEADING = 3

# 0 based
colpos = {
    'region': 1,
    'location': 2,
    'school': 3,
    'donor_balance': 4,
    'student_name': 5,
    'sex': 6,
    'grade': 7,
    'graduation_year': 8,
    'student_donor_id': 9,
    'student_donor_name': 10,
    'student_donor_donation_amount_local': 11,
    'student_donor_donation_amount_us': 12,
    'comment': 13,
    'comment_tw': 14,
}

""" TODO: Obsolete code. Remove later.
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
COL_COMMENT = 13    #N
COL_COMMENT_TW = 14 #O
"""

SHEET_COLUMNS = {
    'DEFAULT': [
        colpos['region'],
        colpos['location'],
        colpos['school'],
        colpos['student_name'],
        colpos['sex'],
        colpos['graduation_year'],
        colpos['student_donor_id'],
        colpos['student_donor_name'],
        colpos['student_donor_donation_amount_local'],
        colpos['comment'],
    ],
    '1': [
        colpos['region'],
        colpos['location'],
        colpos['school'],
        colpos['student_name'],
        colpos['sex'],
        colpos['graduation_year'],
        colpos['student_donor_id'],
        colpos['student_donor_name'],
        colpos['student_donor_donation_amount_local'],
        colpos['comment_tw'],
    ],
}


STYLES = {
    'CHINESE': xlwt.easyxf(u'font: name 宋体;'),
    'CELL_LISTING': xlwt.easyxf(u'font: name 宋体; align: wrap off, shrink_to_fit on, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
    'CELL_LISTING_WRAP': xlwt.easyxf(u'font: name 宋体; align: wrap on, shrink_to_fit off, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
    'CELL_LISTING_TITLE': xlwt.easyxf(u'font: name 宋体, bold on, height 280; align: wrap off, shrink_to_fit on; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
    'WARNING': xlwt.easyxf(u'font: name 宋体; pattern: pattern solid, fore-colour yellow;'),
    'ERROR': xlwt.easyxf(u'font: name 宋体; pattern: pattern solid, fore-colour red;'),
}


class EepMergeSheets:
    # Excel cell style flags.
    STATUS_NORMAL = 0
    STATUS_WARNING = 1
    STATUS_ERROR = 2

    excel_row_lo = ROWS_USED_BY_HEADING
    excel_row_hi = 0

    #
    # create a new file
    #
    def combine_sheets(self, raw_excel_file, sheets):
        try:
            # open raw excel file to read
            wb_eep = xlrd.open_workbook(
                raw_excel_file, on_demand=True, formatting_info=True
            )
        except:
            print "Error opening file ", raw_excel_file
            return

        # New workbook
        wb_new = xlwt.Workbook()
        sh_new = wb_new.add_sheet('final-data')
        sh_new.portrait = 0

        column_titles = [
            'region', 'location', 'school-na', 'student-name', 'sex', 'grad-yr',
            'donor-id', 'donor-na', 'donate-amt', 'comment', 'ipt_odr_nr',
            'auto-student-id', 'auto-donor-stu-cnt-id', 'scl-na-len'
        ]
        # sheets = combine_sheet_numbers#  [int(x) for x in COMBINE_SHEET_NUMBERS.split(',')]
        #print sheets

        current_year = eepshared.current_year

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
        excel_row_lo = self.excel_row_lo
        total_rows_combined = 0
        for sheet_count, sheet_index in enumerate(sheets):
            sh_eep = wb_eep.sheet_by_index(sheet_index)
            sheet = EepSheet(sh_eep)
            sheet.colpos = colpos
            excel_row_hi = sheet.get_sheet_row_hi()
            total_rows_combined += excel_row_hi - excel_row_lo
            print 'Sheet: ', sheet_index, ' Rows: ', excel_row_hi

            # Determine which columns to use.
            try:
                columns = SHEET_COLUMNS[str(sheet_count)]
            except:
                columns = SHEET_COLUMNS['DEFAULT'];

            # Get all student names first.  Used to check duplicates later.
            student_names = sheet.col_values(
                sheet.colpos['student_name'], excel_row_lo, excel_row_hi
            )
            # for name in student_names:
            #    print name

            rng = range(excel_row_lo, excel_row_hi)
            section_check_cols = [
                sheet.colpos['region'],
                sheet.colpos['location'],
                sheet.colpos['school'],
                sheet.colpos['student_donor_name']
            ]
            for row_count, rx in enumerate(rng):
                # +2 because one is for heading and one is for actual cell number.
                current_actual_excel_row_num = i + 2
                current_xlwt_excel_row_num = i + 1

                # import order id
                sh_new.write(current_xlwt_excel_row_num, len(columns) + 0, i + 1)

                # auto student id
                formula = 'COUNTIF(C$1:C{}, C{})'.format(
                    current_actual_excel_row_num,
                    current_actual_excel_row_num
                )
                sh_new.write(
                    current_xlwt_excel_row_num,
                    len(columns) + 1,
                    xlwt.Formula(formula)
                )

                # auto donor student count
                formula = 'COUNTIF(G$1:G{}, G{})'.format(
                    current_actual_excel_row_num, current_actual_excel_row_num
                )
                sh_new.write(
                    current_xlwt_excel_row_num,
                    len(columns) + 2,
                    xlwt.Formula(formula)
                )

                # school name len
                formula = 'LEN(C{})'.format(current_actual_excel_row_num)
                sh_new.write(
                    current_xlwt_excel_row_num,
                    len(columns) + 3,
                    xlwt.Formula(formula)
                )

                for current_column, cx in enumerate(columns):
                    status = self.STATUS_NORMAL
                    if cx > 0:
                        #print rx, cx
                        cell_val = sheet.cell_value(rx, cx)

                        # determine styles to use
                        if cell_val == "":
                            # Cell value is empty.  Set cell style to 'warning'.
                            status = status | self.STATUS_WARNING
                        else:
                            cell_style = STYLES['CHINESE']

                        # Student name error checking
                        if cx == sheet.colpos['student_name']:
                            status = status | self.check_student_name(
                                    sheet, student_names[:row_count], rx)

                        # Check if graduation year is past current year
                        #print cell_val
                        if cx == sheet.colpos['graduation_year']:
                            status = status | self.check_graduation_year(sheet, rx)

                        # if currentVal != previousVal, mark warning
                        if cx in section_check_cols:
                            # TODO: Make sure this is correct. It was > 2 originally.
                            if current_xlwt_excel_row_num > 0:
                                status = status | self.mark_sections(sheet, rx, cx)

                        if status & self.STATUS_ERROR:
                            cell_style = STYLES['ERROR']
                        elif status & self.STATUS_WARNING:
                            cell_style = STYLES['WARNING']

                        # write value
                        sh_new.write(
                            current_xlwt_excel_row_num,
                            current_column,
                            eepshared.clean_text(cell_val),
                            cell_style
                        )

                i += 1
        print 'Total Students: ', total_rows_combined

        wb_new.save(os.path.join(
            eepshared.DESTINATION_DIR,
            eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA + '_combined.xls'
        ))

    def check_student_name(self, sheet, student_names, rownum):
        """ Checks whether a student name exists before the rownum.
        TODO: There's a bug where the first two
        """
        student_name = sheet.get_student_name(rownum)
        status = self.STATUS_NORMAL
        if student_name in student_names:
            status = self.STATUS_WARNING

        # Try to find if same student name exists before
        try:
            # Get first index of the name's appearance.
            found_index = student_names.index(student_name) + ROWS_USED_BY_HEADING

            # Check for possible duplicate within the same school.
            if (
                sheet.get_region(rownum) == sheet.get_region(found_index) and
                sheet.get_location(rownum) == sheet.get_location(found_index) and
                sheet.get_school(rownum) == sheet.get_school(found_index)
            ):
                status = self.STATUS_ERROR
                print '\tPOSSIBLE ERROR: {} Row: {} PrevRow: {}'.format(
                    student_name, rownum, found_index
                )

        except:
            #success
            pass

        return status

    def check_graduation_year(self, sheet, rownum):
        """ Check possible error in graduation year.
        """
        current_year = eepshared.current_year
        yr = sheet.get_graduation_year(rownum)
        try:
            yr = int(yr)
        except:
            return self.STATUS_WARNING

        if yr == '':
            return self.STATUS_WARNING
        elif unicode(current_year) in unicode(yr):
            return self.STATUS_WARNING
        elif yr < current_year:
            return self.STATUS_ERROR

        return self.STATUS_NORMAL

    def mark_sections(self, sheet, rownum, colnum):
        cell_val = sheet.cell_value(rownum, colnum)
        cell_val_prev = sheet.cell_value(rownum - 1, colnum)
        if cell_val != cell_val_prev:
            return self.STATUS_WARNING

        return self.STATUS_NORMAL

    def print_sheetnames(self, raw_excel_file):
        """Print sheet index number and it's name.
        """
        wb_eep = xlrd.open_workbook(
            raw_excel_file, on_demand=True, formatting_info=True
        )
        sheet_names = wb_eep.sheet_names()
        for i, name in enumerate(sheet_names):
            print i, name.strip().encode('utf-8')

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
        'sheetnums',
        nargs='*',
        type=int,
        help='Sheets numbers to merge.  Sheet number starts from 0.'
    )

    return parser

# BEGIN MAIN ==================================================================
if __name__ == "__main__":
    args = get_argparse().parse_args()
    raw_excel_file = args.rawexcelfile
    # print sys.platform

    eepms = EepMergeSheets()
    print args
    # If 'sheetnums' is not specified, print out the sheets in the src Excel file.
    if not args.sheetnums:
        eepms.print_sheetnames(raw_excel_file)
        sys.exit(1)

    # Create destination folders if needed
    eeputil.create_required_dirs()

    eepms.combine_sheets(raw_excel_file, args.sheetnums)