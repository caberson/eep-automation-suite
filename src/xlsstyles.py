#!/usr/bin/python2.7
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2011 Caber Chu

import xlwt

STYLES = {
    'CHINESE': xlwt.easyxf(u'font: name 宋体;'),
    'CELL_LISTING': xlwt.easyxf(u'font: name 宋体; align: wrap off, shrink_to_fit on, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
    'CELL_LISTING_WRAP': xlwt.easyxf(u'font: name 宋体; align: wrap on, shrink_to_fit off, vert centre; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
    'CELL_LISTING_TITLE': xlwt.easyxf(u'font: name 宋体, bold on, height 280; align: wrap off, shrink_to_fit on; borders: left %d, right %d, top %d, bottom %d' % (xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN, xlwt.Borders.THIN)),
    'WARNING': xlwt.easyxf(u'font: name 宋体; pattern: pattern solid, fore-colour yellow;'),
    'ERROR': xlwt.easyxf(u'font: name 宋体; pattern: pattern solid, fore-colour red;'),
}