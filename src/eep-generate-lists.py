#!/usr/bin/env python
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2010 Caber Chu
"""Generate lists using a prepared Excel EEP data file.

http://www.lexicon.net/sjmachin/xlrd.html
http://groups.google.com/group/python-excel/browse_thread/thread/23a0b4d6be641755
http://www.pythonexcels.com/2009/09/another-xlwt-example/
http://www.python-excel.org/
https://secure.simplistix.co.uk/svn/xlwt/trunk/xlwt/examples/xlwt_easyxf_simple_demo.py
re http://gskinner.com/RegExr/

TODO: Refactor code in this script.

Usage:
usage: eep-generate-lists.py [-h] [--yr [YR]] [--mo [MO]]
                             [--destdir [DEST_DIR]] [--country {t,c}]
                             [--combinedlists]
                             [src_excel_file]

Generate various EEP Excel lists

positional arguments:
  src_excel_file        Source Excel file name (default: /Users/cc/Documents/e
                        ep/2020s/2020s_eep_combined_sorted.xls)

optional arguments:
  -h, --help            show this help message and exit
  --yr [YR]             List year (default: 2020)
  --mo [MO]             List month (default: 4)
  --destdir [DEST_DIR]  Destination dir (default: None)
  --country {t,c}       Limit to country (t = Taiwan, c = China)
  --combinedlists       Combine check and letter receiving lists
"""

# Standard module imports.
import argparse
import sys
import os
import glob
import string
import math
import shutil
from datetime import datetime
from decimal import *

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
from eepcombinedsheet import EepCombinedSheet
from util.logger import logger
from util.dir import create_directories

# TODO: Move these constants into the EepLists class.
# global vars
# Heading rows used by the data source Excel file.
SRC_HEADING_ROWS = 1
SHEET_TITLE_BASE = eeputil.get_chinese_title_for_yr_mo(None, None)

PROCESSED_EXCEL_FILE = (
    eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA + '_combined.xls'
)
TEMPLATE_DIR = (
    os.path.join(os.path.dirname(os.path.realpath(__file__)), '..', 'templates')
)


class EepLists:
    """Class to generate various Excel lists using the prepped one.
    """
    # Notes: Font height = height (x pt * 20)
    STYLES = {
        'CHINESE': xlwt.easyxf(u'font: name 宋体;'),
        'CELL_LISTING': xlwt.easyxf(
            u'font: name 宋体, height 240; align: wrap off, shrink_to_fit on, vert centre; borders: left %d, right %d, top %d, bottom %d' % (
                xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN
            )
        ),
        'CELL_LISTING_CENTER': xlwt.easyxf(
            u'font: name 宋体, height 240; align: wrap off, shrink_to_fit on, vert centre, horiz center; borders: left %d, right %d, top %d, bottom %d' % (
                xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN
            )
        ),
        'CELL_LISTING_M1': xlwt.easyxf(
            u'font: name 宋体, height 260; align: wrap off, shrink_to_fit on, vert centre; borders: left %d, right %d, top %d, bottom %d' % (
                xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN
            )
        ),
        'CELL_LISTING_WRAP': xlwt.easyxf(
            u'font: name 宋体, height 240; align: wrap on, shrink_to_fit off, vert centre; borders: left %d, right %d, top %d, bottom %d' % (
                xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN
            )
        ),
        'CELL_LISTING_TITLE': xlwt.easyxf(
            u'font: name 宋体, bold on, height 280; align: wrap off, shrink_to_fit on; borders: left %d, right %d, top %d, bottom %d' % (
                xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN
            )
        ),
        'CELL_LISTING_TITLE_SML': xlwt.easyxf(
            u'font: name 宋体, bold on, height 200; align: wrap off, shrink_to_fit on; borders: left %d, right %d, top %d, bottom %d' % (
                xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN
            )
        ),
        'WARNING': xlwt.easyxf(
            u'font: name 宋体; pattern: pattern solid, fore-colour yellow;'
        ),
    }
    CHAR_WIDTH = 256

    DEFAULT_ROW_HEIGHT = math.trunc(2.7 * 255)

    # Source Excel heading rows.
    src_heading_rows = 1

    data_sheet_row_lo = 1
    data_sheet_row_hi = 0

    templates = None
    excel_file_name = None

    raw_data_wb = None
    raw_data_sheet = None

    data_sheet = None

    def __init__(self, excel_file_name):
        self.excel_file_name = excel_file_name
        self.load_data_file(excel_file_name)
        self.data_sheet = EepCombinedSheet(self.raw_data_sheet)
        self.init_workbook_templates()

        self.filter_country = None
        self.build_combined_lists = False

    def load_data_file(self, excel_file_name):
        # open eep file
        try:
            self.raw_data_wb = xlrd.open_workbook(
                excel_file_name, on_demand=True, formatting_info=True
            )
        except:
            print('Source Excel File {} not found.'.format(excel_file_name))
            sys.exit()

        self.raw_data_sheet = self.raw_data_wb.sheet_by_index(0)
        self.data_sheet_row_hi = self.raw_data_sheet.nrows
        print("EXCEL EEP ROWS: ", self.data_sheet_row_hi)

    def open_workbook(self, file_name):
        return xlrd.open_workbook(
            os.path.join(TEMPLATE_DIR, file_name),
            formatting_info=True,
        )

    def init_workbook_templates(self):
        # create new template workbooks for copying later
        try:
            self.templates = {
                'combinedcheckletterlist': self.open_workbook('template-combined-checklist-lettersubmit-students.xls'),
                'lettersubmitlist': self.open_workbook('template-lettersubmitlist-students.xls'),
                'checklist': self.open_workbook('template-checklist-students.xls'),
                'receivinglist': self.open_workbook('template-receivinglist-students.xls'),
                # 'lettersubmitlist': xlrd.open_workbook(
                #     os.path.join(TEMPLATE_DIR, 'template-lettersubmitlist-students.xls'),
                #     formatting_info=True,
                # ),
                # 'checklist': xlrd.open_workbook(
                #     os.path.join(TEMPLATE_DIR, 'template-checklist-students.xls'),
                #     formatting_info=True,
                # ),
                # 'receivinglist': xlrd.open_workbook(
                #     os.path.join(TEMPLATE_DIR, 'template-receivinglist-students.xls'),
                #     formatting_info=True,
                # ),
            }
        except Exception as err:
            print(err)
            print('No template files found.')
            print(TEMPLATE_DIR)

    def get_new_masterlist_wb(self, column_titles):
        cw = self.CHAR_WIDTH
        wb_masterlist = xlwt.Workbook()
        sh_masterlist = wb_masterlist.add_sheet('masterlist')
        sh_masterlist.portrait = 0

        sh_masterlist.set_header_margin(0)
        sh_masterlist.set_footer_margin(0)
        sh_masterlist.set_header_str("".encode())
        sh_masterlist.set_footer_str("".encode())
        sh_masterlist.set_top_margin(0.25)
        sh_masterlist.set_left_margin(0.27)
        sh_masterlist.set_right_margin(.25)

        sh_masterlist.col(0).width = math.trunc(3 * cw)    #auto student id
        sh_masterlist.col(1).width = math.trunc(10 * cw)   #student
        sh_masterlist.col(2).width = math.trunc(3 * cw)    #sex
        sh_masterlist.col(3).width = math.trunc(4 * cw)    #graduation-yr
        sh_masterlist.col(4).width = math.trunc(4 * cw)    #donor-id
        sh_masterlist.col(5).width = math.trunc(17 * cw)   #donor-name
        sh_masterlist.col(6).width = math.trunc(4 * cw)    #donation amount
        sh_masterlist.col(7).width = math.trunc(85 * cw)   #comment

        # create column titles
        for cx, column_title in enumerate(column_titles):
            sh_masterlist.write(
                0, cx, column_title, self.STYLES['CELL_LISTING_TITLE']
            )

        return wb_masterlist

    def get_title_for_row(row):
        sheet = self.data_sheet
        region = sheet.get_region(row)
        location = sheet.get_location(row)
        school = sheet.get_school(row)
        title = region + " " + location + " " + school

        return title

    def generate_masterlist(self, for_region=None):
        column_titles = [
            '', 'student-name', 'sex', 'grad-yr', 'donor-id', 'donor-na',
            'donate-amt', 'comment'
        ]
        column_titles_len = len(column_titles)

        cps = EepCombinedSheet.colpos
        columns = [
            cps["auto_student_number"],
            cps["student_label_name"],
            cps["sex"],
            cps["graduation_year"],
            cps["student_donor_id"],
            cps["student_donor_name"],
            cps["student_donor_donation_amount_local"],
            cps["comment"],
            # COL_AUTO_STUDENT_NUMBER,
            # COL_STUDENT_LABEL_NAME,
            # COL_SEX,
            # COL_GRADUATION_YEAR, COL_STUDENT_DONOR_ID, COL_STUDENT_DONOR_NAME,
            # COL_STUDENT_DONOR_DONATION_AMOUNT_LOCAL, COL_COMMENT
        ]

        # Get new master list workbook.
        wb_masterlist = self.get_new_masterlist_wb(column_titles)
        sh_masterlist = wb_masterlist.get_sheet(0)

        last_title = ''
        total_title_rows = 0
        current_xlrd_excel_row_num = 0
        """
        current_row_count: actual count
        rx: actual row number on the Excel sheet
        """
        taiwan_name_maps = [u'臺灣', u'台灣']
        actual_row_count = 0
        sheet_range = range(self.data_sheet_row_lo, self.data_sheet_row_hi)
        sheet = self.data_sheet
        name_columns = [
            sheet.colpos['student_name']
            #sheet.colpos['student_donor_name']
        ]
        for current_row_count, rx in enumerate(sheet_range):
            #print current_row_count, ' ', rx
            current_xlrd_excel_row_num = current_row_count + self.src_heading_rows
            current_xlwt_excel_row_num = current_row_count + total_title_rows + 1
            #print "Reader Row: ", current_xlrd_excel_row_num

            region = sheet.get_region(current_xlrd_excel_row_num)
            location = sheet.get_location(current_xlrd_excel_row_num)
            school = sheet.get_school(current_xlrd_excel_row_num)
            current_title = region  + " " + location + " " + school

            # filter by region
            if for_region is not None:
                if for_region == 't' and region not in taiwan_name_maps:
                    continue
                elif for_region == 'c' and region in taiwan_name_maps:
                    continue

            current_xlwt_excel_row_num = actual_row_count + total_title_rows + 1
            actual_row_count += 1

            if last_title != current_title:
                #print current_title
                sh_masterlist.write_merge(
                    current_xlwt_excel_row_num,
                    current_xlwt_excel_row_num,
                    0,
                    column_titles_len - 1,
                    current_title,
                    self.STYLES['CELL_LISTING_TITLE_SML']
                )
                last_title = current_title
                total_title_rows += 1
                current_xlwt_excel_row_num += 1


            for current_column, cx in enumerate(columns):
                cell_val = sheet.cell_value(current_xlrd_excel_row_num, cx)

                # if cx in name_columns:
                #    cell_val = eeputil.remove_parenthesis_content(cell_val)

                if cx == sheet.colpos['comment']:
                    cell_style = self.STYLES['CELL_LISTING_WRAP']
                else:
                    cell_style = self.STYLES['CELL_LISTING']

                # Write to cell.
                sh_masterlist.write(
                    current_xlwt_excel_row_num, current_column,
                    eeputil.clean_text(cell_val), cell_style
                )

        output_file_name = u'{}_masterlist_{}.xls'.format(
            eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA,
            for_region
        )
        output_file = os.path.join(
            eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR, output_file_name
        )
        wb_masterlist.save(output_file)

    def process_schools(self, filter_country=None, build_combined_lists=False):
        excel_row_hi = self.data_sheet_row_hi
        sheet = self.data_sheet
        section_begin_row_num = self.src_heading_rows
        section_end_row_num = None
        last_school = sheet.get_school(section_begin_row_num + 1)
        school_count = 0
        current_row_count = 0
        for i in range(self.src_heading_rows, excel_row_hi):
            current_school = sheet.get_school(i)

            # If there are more rows, check for school change.
            if i + 1 < excel_row_hi:
                next_school = sheet.get_school(i + 1)
                if current_school == next_school:
                    continue

            # If filter by country, skip non-matched schools
            if filter_country is None or sheet.get_country(i) == filter_country:
                # School is going to change or this is the last row.
                school_count += 1
                print('Process school #{} {} @ {}-{}'.format(
                    school_count, current_school,
                    section_begin_row_num, i
                ))

                self.generate_school_lists(section_begin_row_num, i, build_combined_lists)
            else:
                print(u'Skipping {}'.format(current_school))

            last_school = current_school
            section_begin_row_num = i + 1
    
    def create_school_folders(self, filter_country=None):
        filter_country = 't'
        excel_row_hi = self.data_sheet_row_hi
        sheet = self.data_sheet
        section_begin_row_num = self.src_heading_rows
        section_end_row_num = None
        last_school = sheet.get_school(section_begin_row_num + 1)
        school_count = 0
        current_row_count = 0

        unique_schools = []
        for i in range(self.src_heading_rows, excel_row_hi):
            r = sheet.get_row(i)
            country = r[0].value
            region = r[1].value
            school = r[2].value
            folder = u"{}-{}-{}".format(country, region, school)
            tmp = (country, region, school, folder)

            current_school = sheet.get_school(i)

            # If there are more rows, check for school change.
            if i + 1 < excel_row_hi:
                next_school = sheet.get_school(i + 1)
                if current_school == next_school:
                    continue

            # If filter by country, skip non-matched schools
            if filter_country is None or sheet.get_country(i) == filter_country:
                # School is going to change or this is the last row.
                school_count += 1
                print('Process school #{} {} @ {}-{}'.format(
                    school_count, current_school,
                    section_begin_row_num, i
                ))

                if filter_country is None or filter_country == sheet.get_country(i):
                    unique_schools.append(tmp)
            else:
                print(u'Skipping {}'.format(current_school))

            last_school = current_school
            section_begin_row_num = i + 1
        
        folders = [os.path.join(eepshared.DESTINATION_DIR, "_schools", data[3]) for data in unique_schools]
        create_directories(folders)
        for i in unique_schools:
            to_copy = []
            (country, region, school, folder) = i

            f = os.path.join(eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR, "checkinglist")
            ff = u"{}/{} {} {}-*".format(f, country, region, school)
            files = glob.glob(ff)
            [to_copy.append(f) for f in files]

            f = os.path.join(eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR, "receivinglist")
            ff = u"{}/{} {} {}-*".format(f, country, region, school)
            files = glob.glob(ff)
            [to_copy.append(f) for f in files]

            f = os.path.join(eepshared.STUDENT_NAME_LABELS_DIR, "doc")
            ff = u"{}/{} {} {}-*".format(f, country, region, school)
            files = glob.glob(ff)
            [to_copy.append(f) for f in files]

            f = os.path.join(eepshared.STUDENT_NAME_LABELS_DIR, "pdf")
            ff = u"{}/{} {} {}-*".format(f, country, region, school)
            files = glob.glob(ff)
            [to_copy.append(f) for f in files]

            for f in to_copy:
                dest = os.path.join(eepshared.DESTINATION_DIR, "_schools", folder) 
                shutil.copy(f, dest)


    def generate_school_lists(self, beg_row, end_row, build_combined_lists=False):
        sheet = self.data_sheet
        try:
            region = sheet.get_region(beg_row)
            location = sheet.get_location(beg_row)
            school = sheet.get_school(beg_row)
        except:
            return

        # For debugging purpose
        # output = ' '.join([
        #     str(beg_row),
        #     'to',
        #     str(end_row),
        #     ' = ',
        #     str(end_row - beg_row + 1),
        #     ' create_checklist, receivinglist, lettersubmitlist',
        #     school,
        # ]).encode(eepshared.OUTPUT_ENCODING)

        if build_combined_lists:
            self.create_combined_check_letter_list(beg_row, end_row)
        else:
            self.create_checklist(beg_row, end_row)
            self.create_lettersubmitlist(beg_row, end_row)
        self.create_receivinglist(beg_row, end_row)

    def get_new_lettersubmitlist_wb(self):
        wb_new = copy(self.templates['lettersubmitlist'])

        sh_new = wb_new.get_sheet(0)#sheet_by_index(0)
        #sh_new.fit_num_pages = 1
        #sh_new.fit_height_to_pages = 0
        #sh_new.fit_width_to_pages = 1
        sh_new.portrait = 0

        # s = Styles(wb_template)
        #print wb_template.font_list[s[sh_template.cell(0, 0)].xf.font_index]
        #print wb_template.colour_map[wb_template.font_list[s[sh_template.cell(0, 0)].xf.font_index].colour_index]

        cw = self.CHAR_WIDTH
        # Modify dynamic values
        sh_new.col(0).width = math.trunc(6.15 * cw)
        # Student name.
        sh_new.col(2).width = math.trunc(10 * cw)
        # Graduation year.
        sh_new.col(4).width = math.trunc(6 * cw)
        # Donor id.
        sh_new.col(5).width = math.trunc(5.5 * cw)
        # Donor name.
        sh_new.col(6).width = math.trunc(22 * cw)
        # Comment column.  75 for office 2008 and 80 for office 2011
        sh_new.col(8).width = math.trunc(64 * cw)

        sh_new.portrait = 0
        #sh_new.show_headers = 0
        #sh_new.show_footers = 0
        #sh_new.set_show_headers ( 0 )
        #sh_new.set_print_headers( 0 )

        sh_new.set_header_margin(0)
        sh_new.set_footer_margin(0)
        sh_new.set_header_str("".encode())
        sh_new.set_footer_str("".encode())
        sh_new.set_top_margin(0.30)
        sh_new.set_left_margin(0.25)
        sh_new.set_right_margin(0.25)

        return wb_new

    def calculate_comment_col_height(self, value, min_height=None):
        """Calculate row height based on comment length."""
        if min_height is None:
            min_height = self.DEFAULT_ROW_HEIGHT

        rows = math.trunc(math.ceil(len(value) / 30.0))
        # height = rows * 300
        height = rows * 400

        # If calculated row height is less than min height, min height will be used.
        if min_height is not None and height < min_height:
            return min_height

        return height

    def create_lettersubmitlist(self, row_lo, row_hi):
        TGT_HEADING_ROWS = 2
        wb_new = self.get_new_lettersubmitlist_wb()
        sh_new = wb_new.get_sheet(0)
        sheet = self.data_sheet

        lettersubmitlist_columns = [
            0,
            0,
            sheet.colpos['student_label_name'],
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
            sheet.colpos['student_donor_id'],
            sheet.colpos['student_donor_name'],
            sheet.colpos['student_donor_donation_amount_local'],
            sheet.colpos['comment'],
        ]

        no_parenthesis_removal_columns = [
            sheet.colpos['comment'],
            sheet.colpos['student_label_name'],
        ]

        # Centered columns
        centered_columns = [
            sheet.colpos['student_donor_id'],
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
        ]

        # Title
        region = sheet.get_region(row_lo)
        location = sheet.get_location(row_lo)
        school = sheet.get_school(row_lo)
        sheet_title = region  + " " + location + " " + school
        sh_new.write_merge(
            0, 0, 0, 6, sheet_title, self.STYLES['CELL_LISTING_TITLE']
        )

        # Year
        yr_title = SHEET_TITLE_BASE + u'對口救助學生名冊'
        sh_new.write_merge(
            0, 0, 7, 8, yr_title, self.STYLES['CELL_LISTING_TITLE']
        )

        # Total rows processed so far
        i = 0
        for rx in range(row_lo, row_hi + 1):
            # 1 for heading and 1 for actual cell number
            current_actual_excel_row_num = i + 3
            current_xlrd_excel_row_num = rx
            current_xlwt_excel_row_num = i + TGT_HEADING_ROWS

            student_comment = sheet.cell_value(current_xlrd_excel_row_num, sheet.colpos['comment'])
            row_height = self.calculate_comment_col_height(student_comment)
            sh_new.row(current_xlwt_excel_row_num).height = row_height
            sh_new.row(current_xlwt_excel_row_num).height_mismatch = 1

            # Row id
            studentIDColumn = 1
            sh_new.write(
                current_xlwt_excel_row_num, studentIDColumn, i + 1,
                self.STYLES['CELL_LISTING_CENTER']
            )

            current_column = 0
            for cx in lettersubmitlist_columns:
                if cx in centered_columns:
                    cell_style = self.STYLES['CELL_LISTING_CENTER']
                else:
                    cell_style = self.STYLES['CELL_LISTING']

                if cx > 0:
                    cell_val = sheet.cell_value(current_xlrd_excel_row_num, cx)

                    if cx not in no_parenthesis_removal_columns:
                        cell_val = eeputil.remove_parenthesis_content(cell_val)
                    else:
                        cell_style = self.STYLES['CELL_LISTING_WRAP']

                    # If name column
                    if cx in [sheet.colpos['student_label_name']]:
                        dashPosition = cell_val.find(u'-')

                        if dashPosition >= 0 and len(cell_val) - dashPosition > 2:
                            cell_val = cell_val[:dashPosition] + "\n" + cell_val[dashPosition:]
                            # Also change cellstyle to wrap
                            cell_style = self.STYLES['CELL_LISTING_WRAP']

                    sh_new.write(
                        current_xlwt_excel_row_num, current_column,
                        eeputil.clean_text(cell_val), cell_style
                    )

                current_column += 1

            i += 1

        output_file_name = u'{}-lsl.xls'.format(sheet_title)
        output_file = os.path.join(
            eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR,
            'lettersubmitlist',
            output_file_name
        )
        wb_new.save(output_file)

    def get_new_checklist_wb(self):
        """Get a new checklist workbook object.
        """
        wb_new = copy(self.templates['checklist'])

        sh_new = wb_new.get_sheet(0)
        sh_new.portrait = 0
        # Modify dynamic values
        cw = self.CHAR_WIDTH
        sh_new.col(0).width = math.trunc(6.15 * cw)
        # student name
        sh_new.col(2).width = math.trunc(10 * cw)
        # graduation year
        sh_new.col(4).width = math.trunc(6 * cw)
        # Donor id
        sh_new.col(5).width = math.trunc(5.5 * cw)
        # donor name
        sh_new.col(6).width = math.trunc(22 * cw)
        # Comments
        sh_new.col(8).width = math.trunc(64 * cw)

        sh_new.portrait = 0

        sh_new.set_header_margin(0)
        sh_new.set_footer_margin(0)
        sh_new.set_header_str("".encode())
        sh_new.set_footer_str("".encode())
        sh_new.set_top_margin(0.30)
        sh_new.set_left_margin(0.25)
        sh_new.set_right_margin(0.25)

        return wb_new
 

    def create_checklist(self, row_lo, row_hi):
        TGT_HEADING_ROWS = 2
        wb_new = self.get_new_checklist_wb()
        sh_new = wb_new.get_sheet(0)
        sheet = self.data_sheet

        checklist_columns = [
            0,
            0,
            sheet.colpos['student_label_name'],
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
            sheet.colpos['student_donor_id'],
            sheet.colpos['student_donor_name'],
            sheet.colpos['student_donor_donation_amount_local'],
            sheet.colpos['comment'],
        ]

        no_parenthesis_removal_columns = [
            sheet.colpos['comment'],
            sheet.colpos['student_label_name'],
        ]

        centered_columns = [
            sheet.colpos['student_donor_id'],
            sheet.colpos['sex']
        ]

        # Title
        region = sheet.get_region(row_lo)
        location = sheet.get_location(row_lo)
        school = sheet.get_school(row_lo)
        sheet_title = region  + " " + location + " " + school
        sh_new.write_merge(0, 0, 0, 6, sheet_title, self.STYLES['CELL_LISTING_TITLE'])

        # Year
        yr_title = SHEET_TITLE_BASE + u'對口救助學生名冊'
        sh_new.write_merge(0, 0, 7, 8, yr_title, self.STYLES['CELL_LISTING_TITLE'])

        # Centered columns
        centered_columns = [
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
            sheet.colpos['student_donor_id']
        ]

        i = 0   # total rows processed so far
        for rx in range(row_lo, row_hi + 1): #sh.nrows
            current_actual_excel_row_num = i + 3 #1 for heading and 1 for actual cell number
            current_xlrd_excel_row_num = rx
            current_xlwt_excel_row_num = i + TGT_HEADING_ROWS

            student_comment = sheet.cell_value(current_xlrd_excel_row_num, sheet.colpos['comment'])
            row_height = self.calculate_comment_col_height(student_comment)
            sh_new.row(current_xlwt_excel_row_num).height = row_height
            sh_new.row(current_xlwt_excel_row_num).height_mismatch = 1

            # Row id
            studentIDColumn = 1
            sh_new.write(
                current_xlwt_excel_row_num, studentIDColumn, i + 1, self.STYLES['CELL_LISTING_CENTER']
            )

            current_column = 0
            for cx in checklist_columns:
                if cx in centered_columns:
                    cell_style = self.STYLES['CELL_LISTING_CENTER']
                else:
                    cell_style = self.STYLES['CELL_LISTING']

                if cx > 0:
                    cell_val = sheet.cell_value(current_xlrd_excel_row_num, cx)

                    if cx not in no_parenthesis_removal_columns:
                        cell_val = eeputil.remove_parenthesis_content(cell_val)
                    else:
                        cell_style = self.STYLES['CELL_LISTING_WRAP']

                    # If name column
                    if cx in [sheet.colpos['student_label_name']]:
                        dashPosition = cell_val.find(u'-')

                        if dashPosition >= 0 and len(cell_val) - dashPosition > 2:
                            cell_val = cell_val[:dashPosition] + "\n" + cell_val[dashPosition:]
                            # Also change cellstyle to wrap
                            cell_style = self.STYLES['CELL_LISTING_WRAP']

                    sh_new.write(
                        current_xlwt_excel_row_num, current_column,
                        eeputil.clean_text(cell_val), cell_style
                    )
                current_column += 1

            i += 1

        output_file_name = u'{}-cl.xls'.format(sheet_title)
        output_file = os.path.join(
            eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR,
            'checkinglist',
            output_file_name
        )
        wb_new.save(output_file)

    def get_new_receivinglist_wb(self):
        """Create a new receiving list workbook object.
        """
        wb_new = copy(self.templates['receivinglist'])

        sh_new = wb_new.get_sheet(0)
        sh_new.portrait = 0

        # Modify dynamic values
        cw = self.CHAR_WIDTH
        # checkbox
        sh_new.col(0).width = math.trunc(6.15 * cw)
        #student id
        sh_new.col(1).width = math.trunc(4.1 * cw)
        # donor name
        sh_new.col(2).width = math.trunc(10 * cw)
        # year
        sh_new.col(4).width = math.trunc(6 * cw)
        # donor id
        sh_new.col(5).width = math.trunc(4.5 * cw)
        # donor name
        sh_new.col(6).width = math.trunc(22 * cw)
        # donation amount
        # sh_new.col(7).width = math.trunc(6 * cw)
        # signature
        sh_new.col(8).width = math.trunc(36 * cw)
        # notes
        sh_new.col(9).width = math.trunc(29.05 * cw)

        sh_new.set_header_margin(0)
        sh_new.set_footer_margin(0)
        sh_new.set_header_str("".encode())
        sh_new.set_footer_str("".encode())
        sh_new.set_top_margin(0.30)
        sh_new.set_left_margin(0.25)
        sh_new.set_right_margin(0.25)

        return wb_new

    def create_receivinglist(self, row_lo, row_hi):
        TGT_HEADING_ROWS = 2

        wb_new = self.get_new_receivinglist_wb()
        sh_new = wb_new.get_sheet(0)
        sheet = self.data_sheet

        # Special column that is created for this list.
        col_reason = 9

        receivinglist_columns = [
            0,
            0,
            sheet.colpos['student_label_name'],
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
            sheet.colpos['student_donor_id'],
            sheet.colpos['student_donor_name'],
            sheet.colpos['student_donor_donation_amount_local'],
        ]

        no_parenthesis_removal_columns = [
            sheet.colpos['comment'],
            sheet.colpos['student_label_name'],
        ]

        # Centered columns
        centered_columns = [
            sheet.colpos['student_donor_id'],
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
        ]

        # Title
        region = sheet.get_region(row_lo)
        location = sheet.get_location(row_lo)
        school = sheet.get_school(row_lo)
        sheet_title = region  + " " + location + " " + school
        sh_new.write_merge(
            0, 0, 0, 6, sheet_title, self.STYLES['CELL_LISTING_TITLE']
        )

        # Year
        yr_title = SHEET_TITLE_BASE + u'對口救助學生名冊'
        sh_new.write_merge(0, 0, 7, 9, yr_title, self.STYLES['CELL_LISTING_TITLE'])


        i = 0   # total rows processed so far
        for rx in range(row_lo, row_hi + 1): #sh.nrows
            current_actual_excel_row_num = i + 3 #1 for heading and 1 for actual cell number
            current_xlrd_excel_row_num = i + self.src_heading_rows
            current_xlwt_excel_row_num = i + TGT_HEADING_ROWS

            sh_new.row(current_xlwt_excel_row_num).height = math.trunc(4 * 255)
            sh_new.row(current_xlwt_excel_row_num).height_mismatch = 1

            # row id
            sh_new.write(
                current_xlwt_excel_row_num, 1, i+1, self.STYLES['CELL_LISTING_CENTER']
            )

            current_column = 0
            for cx in receivinglist_columns:
                if cx in centered_columns:
                    cell_style = self.STYLES['CELL_LISTING_CENTER']
                else:
                    cell_style = self.STYLES['CELL_LISTING']

                if cx > 0:
                    #print i, ":", x
                    cell_val = sheet.cell_value(rx, cx)

                    if cx not in no_parenthesis_removal_columns:
                        replaced = []
                        cell_val = eeputil.remove_parenthesis_content(
                            cell_val, replaced
                        )
                        if len(replaced) > 0:
                            replacedText = replaced[0]
                    else:
                        cell_style = self.STYLES['CELL_LISTING_WRAP']

                    sh_new.write(
                        current_xlwt_excel_row_num,
                        current_column,
                        eeputil.clean_text(cell_val),
                        cell_style
                    )

                    # Special provision for student name
                    if cx == sheet.colpos['student_label_name']: # and len(replaced) > 0:
                        name_extra = sheet.cell_value(rx, sheet.colpos['student_name_extra'])
                        if len(name_extra) > 0:
                            sh_new.write(
                                current_xlwt_excel_row_num,
                                col_reason,
                                name_extra,
                                cell_style
                            )

                current_column += 1

            i += 1

        output_file_name = u'{}-rl.xls'.format(sheet_title)
        output_file = os.path.join(
            eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR,
            'receivinglist',
            output_file_name
        )
        wb_new.save(output_file)

    def get_new_combined_check_letter_list_wb(self):
        """Get a new combined checklist workbook object.
        """
        wb_new = copy(self.templates['combinedcheckletterlist'])

        sh_new = wb_new.get_sheet(0)
        sh_new.portrait = 0
        # Modify dynamic values
        cw = self.CHAR_WIDTH
        # Attendence check
        sh_new.col(0).width = math.trunc(6.15 * cw)
        # Letter submission check
        sh_new.col(1).width = math.trunc(6.15 * cw)
        # student name
        sh_new.col(3).width = math.trunc(10 * cw)
        # graduation year
        sh_new.col(5).width = math.trunc(6 * cw)
        # Donor id
        sh_new.col(6).width = math.trunc(5.5 * cw)
        # donor name
        sh_new.col(7).width = math.trunc(22 * cw)
        # Comments
        sh_new.col(9).width = math.trunc(52 * cw)

        # sh_new.portrait = 0
        sh_new.set_header_margin(0)
        sh_new.set_footer_margin(0)
        sh_new.set_header_str("".encode())
        sh_new.set_footer_str("".encode())
        sh_new.set_top_margin(0.30)
        sh_new.set_left_margin(0.25)
        sh_new.set_right_margin(0.25)

        return wb_new
 
    def create_combined_check_letter_list(self, row_lo, row_hi):
        TGT_HEADING_ROWS = 2
        wb_new = self.get_new_combined_check_letter_list_wb()
        sh_new = wb_new.get_sheet(0)
        sheet = self.data_sheet

        list_columns = [
            None,
            None,
            None,
            sheet.colpos['student_label_name'],
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
            sheet.colpos['student_donor_id'],
            sheet.colpos['student_donor_name'],
            sheet.colpos['student_donor_donation_amount_local'],
            sheet.colpos['comment'],
        ]

        no_parenthesis_removal_columns = [
            sheet.colpos['comment'],
            sheet.colpos['student_label_name'],
        ]

        centered_columns = [
            sheet.colpos['student_donor_id'],
            sheet.colpos['sex']
        ]

        # Title
        region = sheet.get_region(row_lo)
        location = sheet.get_location(row_lo)
        school = sheet.get_school(row_lo)
        sheet_title = region  + " " + location + " " + school
        sh_new.write_merge(0, 0, 0, 6, sheet_title, self.STYLES['CELL_LISTING_TITLE'])

        # Year
        yr_title = SHEET_TITLE_BASE + u'對口救助學生名冊'
        sh_new.write_merge(0, 0, 7, 8, yr_title, self.STYLES['CELL_LISTING_TITLE'])

        # Centered columns
        centered_columns = [
            sheet.colpos['sex'],
            sheet.colpos['graduation_year'],
            sheet.colpos['student_donor_id']
        ]

        i = 0   # total rows processed so far
        for rx in range(row_lo, row_hi + 1): #sh.nrows
            current_actual_excel_row_num = i + 3 #1 for heading and 1 for actual cell number
            current_xlrd_excel_row_num = rx
            current_xlwt_excel_row_num = i + TGT_HEADING_ROWS

            student_comment = sheet.cell_value(current_xlrd_excel_row_num, sheet.colpos['comment'])
            row_height = self.calculate_comment_col_height(student_comment)
            sh_new.row(current_xlwt_excel_row_num).height = row_height
            sh_new.row(current_xlwt_excel_row_num).height_mismatch = 1

            # Row id
            studentIDColumn = 2
            sh_new.write(current_xlwt_excel_row_num, studentIDColumn, i + 1, self.STYLES['CELL_LISTING_CENTER'])

            current_column = 0
            # cx = column position
            for cx in list_columns:
                if cx is not None:
                    if cx in centered_columns:
                        cell_style = self.STYLES['CELL_LISTING_CENTER']
                    else:
                        cell_style = self.STYLES['CELL_LISTING']

                    cell_val = sheet.cell_value(current_xlrd_excel_row_num, cx)

                    if cx not in no_parenthesis_removal_columns:
                        cell_val = eeputil.remove_parenthesis_content(cell_val)
                    else:
                        cell_style = self.STYLES['CELL_LISTING_WRAP']

                    # If name column
                    if cx in [sheet.colpos['student_label_name']]:
                        dashPosition = cell_val.find(u'-')

                        if dashPosition >= 0 and len(cell_val) - dashPosition > 2:
                            cell_val = cell_val[:dashPosition] + "\n" + cell_val[dashPosition:]
                            # Also change cellstyle to wrap
                            cell_style = self.STYLES['CELL_LISTING_WRAP']

                    sh_new.write(
                        current_xlwt_excel_row_num, current_column,
                        eeputil.clean_text(cell_val), cell_style
                    )
                current_column += 1

            i += 1

        output_file_name = u'{}-combined-cl.xls'.format(sheet_title)
        output_file = os.path.join(
            eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR,
            'checkinglist',
            output_file_name
        )
        wb_new.save(output_file)

# END CLASS

def generate_xmldata():
    """TODO: This method is no longer being used and needs to be updated if
    it's needed again.
    """
    from xml.dom import minidom

    roottag = "<students/>"
    root = minidom.parseString(roottag)

    fields = ['sid', 'name', 'schoolState', 'schoolCity', 'schoolName', 'sex', 'graduationYear', 'donorNumber', 'donorName', 'scholarshipAmount', 'notes','importOrder', 'autoStudentId', 'autoDonorStudentCountNumber', 'schoolNameLength']
    cps = EepCombinedSheet.colpos
    columns = [
        cps["auto_student_number"],
        cps["student_name"],
        cps["region"],
        cps["location"],
        cps["school"],
        cps["sex"],
        cps["graduation_year"],
        cps["student_donor_id"],
        cps["student_donor_name"],
        cps["student_donor_donation_amount_local"],
        cps["comment"],
        cps["import_order_number"],
        cps["auto_student_number"],
        cps["auto_donor_student_count_number"],
        cps["school_name_length"],
    ]

    for current_row_count, rx in enumerate(range(excel_row_lo, excel_row_hi)): #sh.nrows
        current_xlrd_excel_row_num = current_row_count + SRC_HEADING_ROWS
        current_xlwt_excel_row_num = current_row_count + 1

        student = root.createElement('student')
        root.documentElement.appendChild(student)

        for current_column, cx in enumerate(columns):
            field_name = fields[current_column]
            cell_val = sh_eep.cell_value(current_xlrd_excel_row_num, cx)

            #print cell_val,
            tmpNode = root.createElement(field_name)
            tmpNode.appendChild(root.createTextNode(cell_val))
            student.appendChild(tmpNode)

    # Save the file
    f = open(os.path.join(
        eepshared.INSPECTION_DOCUMENTS_DESTINATION_DIR,
        'EEPStudents.xml', 'w'
    )).write(root.toxml())

def get_argparse():
    """Get cmd line argument parser."""
    parser = argparse.ArgumentParser(
        description='Generate various EEP Excel lists')
    default_excel_file_na = '{}.xls'.format(
        eepshared.SUGGESTED_COMBINED_EXCEL_FILE_BASE_NA
    )
    default_excel_file = os.path.join(
        eepshared.DESTINATION_DIR,
        default_excel_file_na
    )

    now = datetime.now()

    parser.add_argument(
        'src_excel_file',
        nargs='?',
        default=default_excel_file,
        help='Source Excel file name (default: %(default)s)',
    )
    parser.add_argument(
        '--yr',
        nargs='?',
        default=now.year,
        type=int,
        help='List year (default: %(default)s)'
    )
    parser.add_argument(
        '--mo',
        nargs='?',
        default=now.month,
        type=int,
        help='List month (default: %(default)s)'
    )
    parser.add_argument(
        '--destdir',
        dest='dest_dir',
        nargs='?',
        # default=eepshared.DESTINATION_DIR,
        help='Destination dir (default: %(default)s)'
    )
    parser.add_argument(
        '--country',
        choices=['t', 'c'],
        help='Limit to country (t = Taiwan, c = China)'
    )
    parser.add_argument(
        '--no-masterlists',
        dest='generate_master_lists',
        action='store_false',
        help='Generate master lists?',
    )
    parser.add_argument(
        '--combinedlists',
        dest='build_combined_lists',
        action='store_true',
        help='Combine check and letter receiving lists',
    )

    return parser

def main(args):
    yr = args.yr
    mo = args.mo 

    try:
        processed_excel_file_na = args.src_excel_file
        PROCESSED_EXCEL_FILE = processed_excel_file_na
    except:
        pass

    if args.dest_dir is not None:
        destination_dir = args.dest_dir
        src_excel_file = glob.glob(destination_dir + '*_combined_sorted.xls')

        if len(src_excel_file) == 1:
            use_suggestion = raw_input('use file {}? (y/n) '.format(src_excel_file[0]))
            if use_suggestion.strip().lower() == 'y':
                PROCESSED_EXCEL_FILE = src_excel_file[0]
            else:
                return

    # Create destination folders if needed.
    eeputil.create_required_dirs(yr, mo)

    eeplists = EepLists(PROCESSED_EXCEL_FILE)
    eeplists.src_heading_rows = SRC_HEADING_ROWS
    eeplists.build_combined_lists = args.build_combined_lists

    # Generate master lists.
    # Separate China and Taiwan into different lists.
    if args.generate_master_lists:
        eeplists.generate_masterlist('c')
        eeplists.generate_masterlist('t')

    eeplists.process_schools(args.country, args.build_combined_lists)

    # TODO: Make an option to run this
    # eeplists.create_school_folders(args.country)


# BEGIN MAIN ==================================================================
if __name__ == "__main__":
    print("Platform: {}".format(sys.platform))    
    args = get_argparse().parse_args()
    print("Arguments: ", args)
    main(args)
