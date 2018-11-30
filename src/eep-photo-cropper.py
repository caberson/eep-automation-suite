#-------------------------------------------------------------------------------
# Name:        eep-photoCropper
# Purpose:
#
# Author:      cc
#
# Created:     07/07/2012
#-------------------------------------------------------------------------------
#:!/usr/bin/env python

import os
import sys
import inspect
import cv2
import glob
import shutil


#==============================================================================
# adds current site-package folder path
current_file = inspect.getfile(inspect.currentframe())
current_folder = os.path.realpath(os.path.abspath(
    os.path.split(current_file)[0])
)
local_site_packages_folder = current_folder + '/site-packages'
if local_site_packages_folder not in sys.path and os.path.exists(
    local_site_packages_folder
):
    sys.path.insert(0, local_site_packages_folder)
#==============================================================================

import eeputil
import eepshared

#==============================================================================
#DIRECTORY_PREFIX = 'eep'
DIR_EEP_PHOTOS_ORIGINAL_DEFAULT = eepshared.DIR_EEP_PHOTOS_ORIGINAL_DEFAULT
DIR_EEP_PHOTOS_CROPPED_DEFAULT = eepshared.DIR_EEP_PHOTOS_CROPPED_DEFAULT

DIR_CURRENT_EXECUTABLE = os.path.dirname(sys.executable)
IMAGE_MAGIC_EXE = os.path.join(DIR_CURRENT_EXECUTABLE, "..", "ImageMagick6.7.3", "convert.exe")
if os.name:
    IMAGE_MAGIC_EXE = 'convert'
CROPPER_WINDOW_NAME = 'EEP_Photo_Cropper'

IS_FILE_BROWSING_MODE = True    # if True, advances to next file upon rename

KEY_PG_UP = 2162688
KEY_UP = 2490368
KEY_PG_DOWN = 2228224
KEY_DOWN = 2621440
KEY_LEFT = 2424832
KEY_RIGHT = 2555904
KEY_ENTER = 13
KEY_SPACE = 32
KEY_BACKSPACE = 8
KEY_DELETE = 3014656
KEY_EQUAL_SIGN = 61
KEY_J = 106
KEY_K = 107
KEY_U = 117
KEY_N = 110
KEY_X = 120


MAC_KEY_DOWN = 63233
MAC_KEY_UP = 63232
MAC_KEY_LEFT = 63234
MAC_KEY_RIGHT = 63235
MAC_KEY_BACKSPACE = 127

#==============================================================================


def show_photo(photo):
    cv2.imshow(CROPPER_WINDOW_NAME, photo)

def rename_file(face_cropper, old_file_name, new_file_name, dir_photos_original, dir_photos_cropped):
    MAX_DONOR_ID_LEN = 4
    if new_file_name.strip() == '':
        return

    old_file_name_base, old_file_extension = os.path.splitext(old_file_name)

    # clean up new file name
    new_file_name_parts = new_file_name.split('.')
    if len(new_file_name_parts) == 1:
        new_file_name_parts.append('1')
    if new_file_name_parts[1] == '':
        new_file_name_parts[1] = '1'

    # greater than 3 letters, illegal name exit.
    if len(new_file_name_parts[0]) > MAX_DONOR_ID_LEN:
        return -1

    # save donor ID for possible reuse later
    donor_id = new_file_name_parts[0]

    # format file name parts
    new_file_name_parts[0] = new_file_name_parts[0].zfill(MAX_DONOR_ID_LEN)
    new_file_name_parts[1] = new_file_name_parts[1].zfill(2)
    print dir_photos_original
    # construct new file name
    new_file_name = '-'.join(new_file_name_parts) + old_file_extension.lower()
    new_file_name_count = len(glob.glob(dir_photos_original + '\\' + new_file_name))
    if new_file_name == old_file_name: #no need to rename current file to the same new file name
        return (donor_id, new_file_name)
    if new_file_name_count > 0:  # filename exists
        new_file_name = '-'.join(new_file_name_parts) + '-' + str(new_file_name_count + 1) + old_file_extension.lower()

    #rename original file
    try:
        old_file = os.path.join(dir_photos_original, old_file_name)
        new_file = os.path.join(dir_photos_original, new_file_name)

        # rename original file
        if not os.path.exists(new_file):
            shutil.move(old_file, new_file)
            face_cropper.set_current_file_name_to(new_file)
        else:
            print 'Not renamed.  File name already exists.'

        # print 'Renamed original file from ', old_file, ' to ', new_file_name
    except Exception as e:
        print e, ' Can not rename original file: ', old_file, ' to ', new_file

    # rename cropped file
    try:
        cropped_file_name = os.path.join(dir_photos_cropped, old_file_name)
        new_cropped_file_name = os.path.join(dir_photos_cropped, new_file_name)
        # rename cropped file
        shutil.move(cropped_file_name, new_cropped_file_name)

        # print 'Renaming cropped file from ', old_file_name, ' to ', new_file_name
    except Exception as e:
        print e, ' Can not rename cropped file: ', cropped_file_name, ' to ', new_cropped_file_name

    return (donor_id, new_file_name)

def start_main_loop(face_cropper):
    # setup a window for displaying stuff
    window = cv2.namedWindow(CROPPER_WINDOW_NAME)#, cv2.CV_WINDOW_AUTOSIZE
    crop_rect_position_increment = 2
    show_photo(face_cropper.get_current_photo())

    new_file_name = ""
    last_donor_id = "0"


    # loop for program
    KEY_PRESS_DEBUG = False
    while True:
        key_pressed = cv2.waitKey(0) # & 0xEFFFFF

        if KEY_PRESS_DEBUG and key_pressed > -1:
            res = key_pressed
            print 'You pressed %d (0x%x), LSB: %d (%s)' % (res, res, res % 256,
repr(chr(res%256)) if res%256 < 128 else '?')
        # pass

        #=========================================
        # Listen for important key presses
        #=========================================
        if key_pressed in (KEY_PG_UP, KEY_UP, MAC_KEY_UP, KEY_U): #page-up, up.  Previous photo
            show_photo(face_cropper.get_previous_photo())
        elif key_pressed in (KEY_PG_DOWN, KEY_DOWN, MAC_KEY_DOWN, KEY_N):
           show_photo(face_cropper.get_next_photo())

        elif key_pressed in (ord('+'), KEY_EQUAL_SIGN,): #+,=, increase crop rect size
            show_photo(face_cropper.zoom_current_rectangle_size(1))

        elif key_pressed == ord('-'): #-, decrease crop rect size
            show_photo(face_cropper.zoom_current_rectangle_size(-1))

        elif key_pressed in (KEY_LEFT, MAC_KEY_LEFT, KEY_J): #Move crop rect left
            show_photo(face_cropper.update_crop_rect_horizontal_position(-crop_rect_position_increment))

        elif key_pressed in (KEY_RIGHT, MAC_KEY_RIGHT, KEY_K): #Move crop rect right
            show_photo(face_cropper.update_crop_rect_horizontal_position(crop_rect_position_increment))

        elif key_pressed == ord('s'): #save cropped image
            face_cropper.save_cropped_image()
            # advance to next upon save
            show_photo(face_cropper.get_next_photo())

        elif key_pressed in (KEY_ENTER, KEY_SPACE) and new_file_name: #Rename file
            old_file_name = face_cropper.get_current_image_file_name()[0]
            donor_id, new_file_name = rename_file(
                face_cropper,
                old_file_name, new_file_name,
                face_cropper.dir_photos_original,
                face_cropper.dir_photos_cropped
            )
            print 'New File: %s' % new_file_name
            if int(donor_id) > 0:
                last_donor_id = donor_id
            new_file_name = ""
            print "Last donor ID: %s" % last_donor_id

            if IS_FILE_BROWSING_MODE:
                try:
                    show_photo(face_cropper.get_next_photo())
                except Exception as e:
                    print e
                    break
            else: # refresh current photo
                show_photo(face_cropper.get_current_photo())

        elif key_pressed in (KEY_BACKSPACE, MAC_KEY_BACKSPACE): #Reset new file name
            new_file_name = ""

        elif key_pressed in range(48, 58) or key_pressed == ord('.'): # number keys and period
            if key_pressed == ord('.') and len(new_file_name) == 0: #if no donor id entered, use the last one if applicable
                new_file_name = last_donor_id
            new_file_name += chr(key_pressed)

        elif key_pressed in (KEY_DELETE, KEY_X):
            #print 'Delete photo'
            face_cropper.delete_current_image()
            show_photo(face_cropper.get_current_photo())
            pass
        elif key_pressed == 27:
            #cv2.destroyWindow(CROPPER_WINDOW_NAME)
            break

# debugging method for Windows only
def compare_files():
    import os
    src_pattern = 'D:\_cc\eep_2\_renamed\*.jpg'
    src_files = glob.glob(src_pattern)
    tgt_pattern = 'F:\eep_2\*.jpg'
    tgt_files = glob.glob(tgt_pattern)

    output_dir = 'D:\_myFiles\_scripts\eep\eep_photos_original\_toCheck'

    for src_f in src_files:
        src_file_stat = os.stat(src_f)
        #print src_f, src_file_stat.st_size, src_file_stat.st_mtime
        #, file_stat

        bFound = 0
        for tgt_f in tgt_files:
            tgt_file_stat = os.stat(tgt_f)
            if (
                src_file_stat.st_size == tgt_file_stat.st_size
                and src_file_stat.st_mtime == tgt_file_stat.st_mtime
            ):
                bFound = 1
                break;
            else:
                bFound = 0

        if bFound == 0:
            file_base_name = os.path.basename(src_f)
            shutil.copyfile(src_f, output_dir + '\\' + file_base_name)
            print 'Source File Not matched:', src_f


def usage():
    print """Usage:
            --noinitdir No creating dir
            --photodir= Original photo dir
            """

def get_arguments():
    import argparse
    parser = argparse.ArgumentParser(description='Photo cropper.')
    parser.add_argument('--noinitdir', action="store_true")
    parser.add_argument('--photodir')
    parser.parse_args()

    return parser

def main(argv):
    import getopt
    import photos.cropper

    config = eepshared.get_config()
    # print config.items('path')

    OPTION_INIT_DIR = 1
    OPTION_DIR_EEP_PHOTOS_ORIGINAL = DIR_EEP_PHOTOS_ORIGINAL_DEFAULT
    OPTION_DIR_EEP_PHOTOS_CROPPED = DIR_EEP_PHOTOS_CROPPED_DEFAULT

    parser = get_arguments()
    print parser

    try:
        opts, args = getopt.getopt(argv, "h", ["help","noinitdir","photodir="])
    except getopt.GetoptError:
        print "Error"
        sys.exit(2)

    for opt, args in opts:
        #print opt, args
        if opt in ("-h"):
            usage()
            sys.exit(2)
        elif opt in ("--noinitdir",):
            OPTION_INIT_DIR = 0
        elif opt in ("--photodir",):
            if os.path.isdir(args):
                OPTION_DIR_EEP_PHOTOS_ORIGINAL = args
            else:
                print "Invalid photodir specified: %s" % args
                usage()
                sys.exit(2)

    # instantiate a cropper object
    dir_cwd_path = os.path.abspath(os.getcwd())
    face_cropper = photos.cropper.FaceCropper(
        # os.path.join(dir_cwd_path, OPTION_DIR_EEP_PHOTOS_ORIGINAL),
        # os.path.join(dir_cwd_path, OPTION_DIR_EEP_PHOTOS_CROPPED),
        eepshared.STUDENT_PHOTOS_ORIGINAL_DIR,
        eepshared.STUDENT_PHOTOS_CROPPED_DIR,
        IMAGE_MAGIC_EXE
    )

    # init this directory
    if OPTION_INIT_DIR == 1:
        # Create destination folders if needed
        eeputil.create_required_dirs()
        face_cropper.auto_save_on_view = True

    # If no files, exit program
    if len(face_cropper.original_photos) == 0:
        # print DIR_EEP_PHOTOS_ORIGINAL_DEFAULT
        print 'No photos to crop'
        sys.exit()

    start_main_loop(face_cropper)

    print 'Exiting'
    exit()


#==============================================================================
if __name__ == '__main__':
    #compare_files()

    main(sys.argv[1:])

