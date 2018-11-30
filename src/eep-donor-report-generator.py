#-------------------------------------------------------------------------------
# Name:        eep-donorReportGenerator
# Purpose:
#
# Author:      cc
#
# Created:     07/07/2012
# Copyright:   (c) cc 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------
#!/usr/bin/env python

import sys, os, inspect, math
from datetime import datetime
from time import sleep

try:
    import win32com.client as win32
except:
    print 'win32com.client not found'

#import timeit

import xlrd

#==============================================================================
# adds current site-package folder path
current_folder = os.path.realpath(os.path.abspath(os.path.split(inspect.getfile( inspect.currentframe() ))[0]))
local_site_packages_folder = current_folder + '/site-packages'
if local_site_packages_folder not in sys.path and os.path.exists(local_site_packages_folder):
    sys.path.insert(0, local_site_packages_folder)
#==============================================================================


# from eep import common
import eepshared
import eeputil

# Word COM reference
# http://msdn.microsoft.com/en-us/library/bb244515(v=office.12).aspx
#http://webcache.googleusercontent.com/search?q=cache:kRfamjE4n6oJ:www.galalaly.me/index.php/2011/09/use-python-to-parse-microsoft-word-documents-using-pywin32-library/+python+Word+pywin32+examples&cd=3&hl=en&ct=clnk&gl=us

RANGE = range(3, 8)

R1H = 149.399993896484
R2H = 19.75
R2H = 17

# Combined sheet columns
COL_REGION = 0
COL_LOCATION = 1
COL_SCHOOL = 2
COL_STUDENT_NAME = 3
COL_SEX = 4
COL_GRADUATION_YEAR = 5
COL_STUDENT_DONOR_ID = 6
COL_STUDENT_DONOR_NAME = 7
COL_STUDENT_DONOR_DONATION_AMOUNT_LOCAL = 8
COL_COMMENT = 9
COL_IMPORT_ORDER_NUMBER = 10
COL_AUTO_STUDENT_NUMBER = 11
COL_AUTO_DONOR_STUDENT_COUNT_NUMBER = 12
COL_SCHOOL_NAME_LENGTH = 13

DEFAULT_ROWS_IN_DOC = 3 # table rows already on doc to accommodate students
STUDENTS_PER_ROW = 3
#    existingRowsOnDoc = 3   # table rows already on doc to accommodate students
#    studentsPerRow = 3

# TODO: Confirm if we can use template-donor-report-docx
FILE_DONORREPORT_TEMPLATE_FILENAME = os.path.join(
    eepshared.TEMPLATES_DIR, 'template-donor-report.doc'
)
print "Template Dir: %s" % FILE_DONORREPORT_TEMPLATE_FILENAME


def getWordHandle():
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    return word

def getEepExcelSheet(intSheetIndex=0):	#0 based
    # xls_file_na = '%s_eep_combined_sorted.xls' % (REPORT_YEAR_CODE_ENG)
    xls_file_na = '%s_combined_sorted.xls' % (eepshared.SUGGESTED_RAW_EXCEL_FILE_BASE_NA)
    #wb_eep = xlrd.open_workbook(os.getcwd() + '/assets/2011f_eep_combined.xls', on_demand=True, formatting_info=True)
    wb_eep = xlrd.open_workbook(os.path.join(eepshared.DESTINATION_DIR, xls_file_na), on_demand=True, formatting_info=True)
    sh_eep = wb_eep.sheet_by_index(intSheetIndex)
    # print xls_file_na
    return sh_eep

def getDonorList(sh):
    lastRowOfSheet = sh.nrows
    donorList = []
    for ry in range(1, lastRowOfSheet):
        try:
            donorID = sh.cell_value(ry, COL_STUDENT_DONOR_ID)
            donorName = sh.cell_value(ry, COL_STUDENT_DONOR_NAME)

            try:
                donorID = int(donorID)
            except:
                continue

            # check if donorID was already in
            try:
                donorInfo = donorID, donorName
                donorList.index(donorInfo)
                continue    #if found
            except:
                pass


            #print ry, donorID, donorName

            donorList.append(donorInfo)
        except:
             print 'Error:', sys.exc_info()

    #print donorList
    #for i, donor in enumerate(donorList):
    #    print i, donor
    #print '---------------'
    #for did, dna in donorList:
    #    print did, dna

    return donorList

def getDonorNameUsingDonorID(donorList, donorID):
    for id, name in donorList:
        if id == donorID:
            return name
    return False

def getStudentList(sh):
    lastRowOfSheet = sh.nrows
    studentList = []
    for ry in range(1, lastRowOfSheet):
        try:
            donorID = sh.cell_value(ry, COL_STUDENT_DONOR_ID)
            studentName = sh.cell_value(ry, COL_STUDENT_NAME)
            studentSchool = sh.cell_value(ry, COL_SCHOOL)

            try:
                donorID = int(donorID)
            except:
                continue

            # check if donorID was already in
            try:
                studentInfo = donorID, studentName, studentSchool
                studentList.index(donorInfo)
                continue    #if found
            except:
                pass


            #print ry, donorID, donorName

            studentList.append(studentInfo)
        except:
             print 'Error:', sys.exc_info()

    return studentList

def getDonorStudents(studentList, searchDonorID):
    newStudentList = []
    for student in studentList:
        donorID, studentName, studentSchool = student
        if donorID == searchDonorID:
            newStudentList.append(student)

    """
    for student in newStudentList:
        donorID, studentName, studentSchool = student
        print donorID, studentName, studentSchool
    """

    return newStudentList

# this is a testing method
def word():
    templateFileName = os.getcwd() + '/assets/template.doc'
    targetFileName = os.getcwd() + '/assets/test.doc'
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    doc = word.Documents.Open(templateFileName)

    activeDocument = word.ActiveDocument

    # test text
    rng = doc.Range(0,0)
    rng.InsertAfter('Hacking Word with Python\r\n\r\n')

    # get donor name

    # get student names

    # get student table
    oStudentTbl = activeDocument.Tables(1)


    # add rows if more than 9 students
    try:
        oTblRng = oStudentTbl.Range()
        #oStudentTbl.Select()
        #oStudentTbl.Select
        # print oStudentTbl.Rows.Count
        #Selection = word.Selection
        for i in range(1):
            #Selection.Rows().Add(1)
            #Selection.InsertRows(2)
            oTmpR1 = oStudentTbl.Rows.Add() #oStudentTbl.Rows(1)
            oTmpR1.SetHeight(R1H)
            oTmpR2 = oStudentTbl.Rows.Add()
            oTmpR2.SetHeight(R2H)
    except:
        pass
        # print 'Error:', sys.exc_info()

    # add students
    oStudentPhotoCell = oStudentTbl.Cell(1, 1)
    oStudentNameCell = oStudentTbl.Cell(2, 1)
    oStudentPhotoCellRange = oStudentPhotoCell.Range
    oStudentPhotoCellRange.ParagraphFormat.Alignment = 1

    word.ActiveDocument.InlineShapes.AddPicture(os.getcwd() + '/photos_cropped/400-01.jpg', LinkToFile=False, SaveWithDocument=True, Range=oStudentPhotoCellRange)
    oStudentNameCell = oStudentTbl.Cell(2, 1)
    #oStudentNameCell.FitText = True
    oStudentNameCell.Range.InsertAfter("Test Tokf fsa Ok Student Name Abc Def Gh")
    oStudentNameCell.Range.Text = "Updated"
    oStudentNameCell.Range.ParagraphFormat.Alignment = 1#wdAlignParagraphCenter

    sleep(1)
    # save
    word.ActiveDocument.SaveAs(targetFileName)#, FileFormat=win32com.client.constants.wdFormatTextLineBreaks


    doc.Close(False)
    word.Application.Quit()

def addStudentToDoc(word, studentSlot, studentInfo, studentPhotoFN):
    # get student table
    oStudentTbl = word.ActiveDocument.Tables(1)

    studentDonorID, studentName, studentSchool = studentInfo
    # table header count is a constant.  It does not increment as the number of pages increase
    #headerRowsCount = int(math.ceil(float(studentSlot + 1) / (STUDENTS_PER_ROW * DEFAULT_ROWS_IN_DOC )))
    headerRowsCount = 1
    #print 'slot, headerRowCount: ', studentSlot+1, headerRowsCount

    rx = (studentSlot % STUDENTS_PER_ROW) + 1
    ry = int(math.ceil(float(studentSlot + 1) / STUDENTS_PER_ROW))
    # print '\t', studentSlot + 1, studentName, studentSchool, ' at ', (ry, rx)

    # add students
    slotRow = ry * 2 - 1 + headerRowsCount
    oStudentPhotoCell = oStudentTbl.Cell(slotRow, rx)
    oStudentNameCell = oStudentTbl.Cell(slotRow + 1, rx)
    oStudentPhotoCellRange = oStudentPhotoCell.Range
    #oStudentPhotoCellRange.ParagraphFormat.Alignment = 1

    #oStudentNameCell.FitText = True
    #oStudentNameCell.Range.InsertAfter("Test Tokf fsa Ok Student Name Abc Def Gh")
    if studentSchool != "" and studentName != "":
        oStudentNameCell.Range.Text = studentSchool + " - " + studentName
        oStudentNameCell.FitText = True
        #oStudentNameCell.Range.Font.Name = u"æ¨™æ¥·é«”"
        oStudentNameCell.Range.ParagraphFormat.Alignment = 1#wdAlignParagraphCenter
        #oStudentNameCell.Range.ParagraphFormat.BaseLineAlignment = 2
        oStudentNameCell.VerticalAlignment = win32.constants.wdAlignVerticalCenter
        oStudentNameCell.Row.SetHeight(R2H)

    # try adding the photo
    if studentPhotoFN != '':
        try:
            #croppedStudentPhotoForDonor.remove(studentPhotoFN)
            word.ActiveDocument.InlineShapes.AddPicture(studentPhotoFN, LinkToFile=False, SaveWithDocument=True, Range=oStudentPhotoCellRange)
        except:
            #print 'Error: ', studentPhotoFN, sys.exc_info()
            pass

def getNewDonorWordDoc(donor, student_count, word):
    donorID, donorName = donor

    doc = word.Documents.Open(FILE_DONORREPORT_TEMPLATE_FILENAME)
    activeDocument = word.ActiveDocument

    # print 'Cropped: ', croppedStudentPhotoForDonor
    #croppedStudentPhotoQueue = deque(croppedStudentPhotoForDonor)
    #print croppedStudentPhotoQueue.popleft()

    tmpRange = activeDocument.Content
    # update expedition season
    docMainHeader = activeDocument.Sections(1).Headers(win32.constants.wdHeaderFooterPrimary)
    docMainHeader.Range.Find.Execute(FindText="{reportYear}", ReplaceWith=REPORT_YEAR)
    docMainHeader.Range.Find.Execute(FindText="{reportSeason}", ReplaceWith=REPORT_SEASON)

    # update donor name
    tmpRange.Find.Execute(FindText="{donorName}", ReplaceWith=donorName)
    tmpRange = activeDocument.Content
    tmpRange.Find.Execute(FindText="{donorID}", ReplaceWith=donorID)

    #activeDocument.Sections(1).Headers(1).Range.Text = "HAHAHA"

    # get student table
    oStudentTbl = activeDocument.Tables(1)

    if student_count > (DEFAULT_ROWS_IN_DOC * STUDENTS_PER_ROW):
        additionalRowsNeeded = int(math.ceil(
            float(student_count) / STUDENTS_PER_ROW
            ) - DEFAULT_ROWS_IN_DOC
        )
        # print "Need ", additionalRowsNeeded, " more rows."
        # add rows if more than 9 students
        try:
            #oTblRng = oStudentTbl.Range()
            #oStudentTbl.Select()
            #oStudentTbl.Select
            #print oStudentTbl.Rows.Count
            #Selection = word.Selection
            for i in range(additionalRowsNeeded):
                #Selection.Rows().Add(1)
                #Selection.InsertRows(2)
                oTmpR1 = oStudentTbl.Rows.Add() #oStudentTbl.Rows(1)
                oTmpR1.SetHeight(R1H)
                oTmpR2 = oStudentTbl.Rows.Add()
                oTmpR2.SetHeight(R2H)
        except:
            print 'Error:', sys.exc_info()

    return doc


def updateDonorWordDoc(donor, studentList, word=None):
    from glob import glob

    donorID, donorName = donor

    #from collections import deque
    bQuitWordApplicationAtEnd = False
    if not OPTION_CREATE_WORD_DOC and word == None:
        word = getWordHandle()
        bQuitWordApplicationAtEnd = True

    croppedStudentPhotoForDonor = glob(
        os.path.join(PHOTOS_CROPPED_DIR, str(donorID).zfill(4) + '*.*')
    )
    numberOfStudentsForDonor = len(studentList)
    totalSlotsNeeded = numberOfStudentsForDonor
    # if there are more student photos than the student list, use photos count instead
    if len(croppedStudentPhotoForDonor) > numberOfStudentsForDonor:
        totalSlotsNeeded = len(croppedStudentPhotoForDonor)

    if OPTION_CREATE_WORD_DOC:
        doc = getNewDonorWordDoc(donor, totalSlotsNeeded, word)

    # loop through students
    print 'Processing donor: ', donorID#, donorName
    for studentSlot, studentInfo in enumerate(studentList):
        studentDonorID, studentName, studentSchool = studentInfo
        studentPhotoFN = os.path.join(
            PHOTOS_CROPPED_DIR,
            str(donorID).zfill(4) + '-' + str(studentSlot + 1).zfill(2) + '.jpg'
        )
        if not os.path.isfile(studentPhotoFN):
            LOG.write(''.join(['Photo not found: ', studentPhotoFN, "\n"]))
            # print 'File not found - ', studentPhotoFN
            stuentPhotoFN = ''
        else:
            # print croppedStudentPhotoForDonor
            # print studentPhotoFN
            croppedStudentPhotoForDonor.remove(studentPhotoFN)

        # if only 1 slot is needed, center it
        if totalSlotsNeeded == 1:
            studentSlot = 1

        if OPTION_CREATE_WORD_DOC:
            addStudentToDoc(word, studentSlot, studentInfo, studentPhotoFN)

    # loop through leftover photos
    for i, studentPhotoFN in enumerate(croppedStudentPhotoForDonor):
        studentSlot = i + len(studentList)
        studentInfo = (donorID, '', '')

        LOG.write(''.join(['Extra photo: ', studentPhotoFN, "\n"]))

        if OPTION_CREATE_WORD_DOC:
            addStudentToDoc(word, studentSlot, studentInfo, studentPhotoFN)

    """
    # go over pages and add headers if possible.  This does not work
    pages = activeDocument.ActiveWindow.Panes(1).Pages
    pageCount = pages.Count
    print 'Page Count:', pageCount
    for pageNum in range(2, pageCount+1):
        print pageNum
    """

    if OPTION_CREATE_WORD_DOC:
        # save
        targetFileName = os.path.join(eepshared.DONOR_REPORT_DIR, str(donorID) + '.doc') #os.path.join(os.getcwd(), '_donor-reports-processed', str(donorID) + '.doc')
        print targetFileName
        # targetFileName = os.getcwd() + '/assets/test.doc'
        activeDocument = word.ActiveDocument
        activeDocument.SaveAs(targetFileName)#, FileFormat=win32com.client.constants.wdFormatTextLineBreaks
        doc.Close(False)

        if bQuitWordApplicationAtEnd:
            word.Application.Quit()


def getRequiredLists():
    sh_eep = getEepExcelSheet(0)    #sheet_num is 0 based
    donorList = getDonorList(sh_eep)
    studentList = getStudentList(sh_eep)

    return donorList, studentList

def processWordDocs(processDonorID=0):
    donorList, studentList = getRequiredLists()
    word = getWordHandle()

    # sort the donorList by donorID
    donorList = sorted(donorList, key=lambda donorList: donorList[0])

    # loop through all donors
    for donor in donorList:
        donorID, donorName = donor

        # Check if we are processing a specific donor
        if processDonorID > 0 and donorID != processDonorID:
            continue

        #print donorID, donorName
        studentListForDonor = getDonorStudents(studentList, donorID)
        #print len(studentListForDonor), studentListForDonor
        updateDonorWordDoc(donor, studentListForDonor, word)

    word.Application.Quit()

def setup_argparse():
    import argparse
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('--noinitdir', nargs='*',
                   help='an integer for the accumulator')
    parser.add_argument('--logonly', dest='logonly', action='store_true')
    parser.add_argument('--photodir',
        nargs = '?',
        default = '',
        help='sum the integers (default: find the max)'
    )
    parser.add_argument(
        '-y',
        '--year',
        nargs = '?',
        type = int,
        default = datetime.today().year
    )
    parser.add_argument(
        '-m',
        '--month',
        nargs = '?',
        type = int,
        default = datetime.today().month
    )

    return parser
    return parser.parse_args()

def main(args):
    print 'beg: ', datetime.now()
    processWordDocs()
    print 'end: ', datetime.now()

if __name__ == '__main__':
    #from datetime import datetime
    #word() #for testing
    # Usage: eep-donorReportGenerator.py -y2015 --month=5
    import argparse

    parser = setup_argparse()
    args = parser.parse_args()

    OPTION_INIT_DIR = 1
    if args.noinitdir:
        OPTION_INIT_DIR = 0

    OPTION_CREATE_WORD_DOC = True
    if args.logonly:
        OPTION_CREATE_WORD_DOC = False

    if args.photodir:
        photodir = args.photodir

        if os.path.isdir(photodir):
                pass
        else:
                print "Invalid photodir specified: %s" % photodir
		parser.print_help()
                sys.exit(2)

    if args.year:
        REPORT_YEAR = args.year

    if args.month:
        REPORT_MONTH = args.month
    REPORT_SEASON = u'秋' if REPORT_MONTH > 8 else u'春'
    REPORT_YEAR_CODE_ENG = eepshared.build_english_year_code(REPORT_YEAR, REPORT_MONTH)
    PHOTOS_CROPPED_DIR = eepshared.STUDENT_PHOTOS_CROPPED_DIR

    # init dir
    eeputil.create_required_dirs()

    log_file = os.path.join(eepshared.DESTINATION_DIR, 'log.txt')
    try:
        LOG = open(log_file, 'w')
    except:
        pass

    try:
        LOG.write(''.join(['Log file for ', REPORT_YEAR_CODE_ENG, "\n"]))
    except:
        print 'error'
        pass

    main(args)

    LOG.close()
