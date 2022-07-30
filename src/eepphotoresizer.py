#!/usr/bin/env python3
# coding=utf-8
# -*- coding: utf-8 -*-
# Copyright (C) 2012 Caber Chu

# Purpose: Resizes student photos to standard dimensions used by the generated word files.

import os, sys, inspect
import glob
import shutil
import argparse
import util.image

from datetime import datetime

#==============================================================================
# adds current site-package folder path
current_folder = os.path.realpath(os.path.abspath(os.path.split(inspect.getfile( inspect.currentframe() ))[0]))
local_site_packages_folder = current_folder + '/site-packages'
if local_site_packages_folder not in sys.path and os.path.exists(local_site_packages_folder):
    sys.path.insert(0, local_site_packages_folder)
#==============================================================================

### local site-packages includes starts
import eepshared

# Windows configs
DIR_CURRENT_EXECUTABLE = os.path.dirname(sys.executable)
IMAGE_MAGIC_EXE = os.path.join(DIR_CURRENT_EXECUTABLE, "..", "ImageMagick-6.7.3", "convert.exe")
RESIZE_RESOLUTION = '354x425'

# OSX configs
if os.name != 'nt':
    IMAGE_MAGIC_EXE = 'convert'

def resize_img(src_file, target_file=None):
    if target_file is None:
        target_file = src_file
    
    # -units is required for os x convert to work correctly
    run_cmd = IMAGE_MAGIC_EXE + ' ' + src_file + ' -units PixelsPerInch -resize ' + RESIZE_RESOLUTION + ' -density 180 ' + target_file
    print(run_cmd)
    os.system(run_cmd)

def resize_photos_for_donor_doc(photos_path, out_path):
    if not photos_path:
        dir_cwd_path = os.path.abspath(os.getcwd())
        photos_path = os.path.join(
            dir_cwd_path,
            eepshared.DIR_EEP_PHOTOS_CROPPED_DEFAULT,
        )

    pattern = os.path.join(photos_path, '*.jpg')
    print(pattern)
    files = glob.glob(pattern)

    for f in files:
        file_base_name = os.path.basename(f)
        file_name, file_extension = os.path.splitext(file_base_name)

        target_f = os.path.join(out_path, file_base_name)
        # resize_img(f, target_f)
        util.image.resize_img(f, target_f, img_magic_exe=IMAGE_MAGIC_EXE)
        print("processed {}".format(target_f))

def main():
    pass

def get_parser():
    default_yr_code = eepshared.build_english_year_code()
    yr_code = eepshared.build_english_year_code()

    if os.name != 'nt':
        # osx 
        default_base_in_dir = "/Users/cc/Documents/eep/{}".format(default_yr_code)
        default_base_out_dir = "/Users/cc/Documents/eep/{}".format(default_yr_code)
    else:
        # assume windows
        default_base_in_dir = '\\\\VBOXSVR\cc\Documents\eep',
        default_base_out_dir = r'C:\Users\cc\Documents\eep\\{}'.format(default_yr_code)


    src_photos_path = os.path.join(
        # 'C:\projects\eep-automation-suite\data\_to_resize',
        # default_base_dir,
        default_base_in_dir,
        # '\\\\VBOXSVR\cc\Documents\eep',
        'eep_photos_cropped'
    )
    output_path = os.path.join(
        default_base_out_dir,
        'eep_photos_cropped_resized'
    )
    epilog = """\
        pipenv run src/eepphotoresizer.py -p ~/Documents/eep/tmp/photos -o ~/Documents/eep/2020f/eep_photos_cropped
        """

    parser = argparse.ArgumentParser(description='Resize EEP donor photos.', epilog=epilog)
    # parser.add_argument(
    #     '-y',
    #     '--year',
    #     nargs='?',
    #     type=int,
    #     default=datetime.today().year
    # )
    # parser.add_argument(
    #     '-m',
    #     '--month',
    #     nargs='?',
    #     type=int,
    #     default=datetime.today().month
    # )
    parser.add_argument('-p', nargs='?', default=src_photos_path, help=f"sourc photos default: {src_photos_path}")
    parser.add_argument('-o', nargs='?', default=output_path, help=f"photos output folder (default: {output_path})")
    return parser

def resize_on_windows():
    # python eep-photo-resizer.py -p C:\projects\eep\data\2017f\eep_photos_cropped_original -o C:\projects\eep\data\2017f\eep_photos_cropped
    # python src/eep-photo-resizer.py -p ~/Documents/eep/2019s/eep_photos_cropped -o ~/Documents/eep/2019s/eep_photos_resized
    # TODO: Need to remove the hard defined path below.
    yr_code = eepshared.build_english_year_code()
    default_base_dir = r'C:\Users\cc\Documents\eep\\{}'.format(yr_code)

    src_photos_path = os.path.join(
        # 'C:\projects\eep-automation-suite\data\_to_resize',
        # default_base_dir,
        '\\\\VBOXSVR\cc\Documents\eep',
        yr_code,
        'eep_photos_cropped'
    )
    output_path = os.path.join(
        # 'C:\projects\eep-automation-suite\data\\2017f\eep_photos_cropped',
        default_base_dir,
        'eep_photos_cropped'
    )

    parser = get_parser()
    args, unknown = parser.parse_known_args()
    print(args)
    print(args.p, args.o, args.y, args.m)

    if args.y or args.m:
        print("Specif...")
    # resize_photos_for_donor_doc(args.p, args.o)

def resize_on_osx():
    parser = get_parser()
    args, unknown = parser.parse_known_args()
    print(args)
    resize_photos_for_donor_doc(args.p, args.o)

if __name__ == '__main__':
    if os.name != 'nt':
        resize_on_osx()
    else:
        resize_on_windows()
