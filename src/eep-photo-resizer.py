#-------------------------------------------------------------------------------
# Name: eep-photo-resizer
# Purpose: Resizes student photos to standard dimensions used by the generated word files.
#
# Author: Caber Chu
#
# Created: 09/07/2012
# Copyright: (c) Caber Chu 2012
#-------------------------------------------------------------------------------
#!/usr/bin/env python
import os, sys, inspect
import glob
import shutil
import argparse

#==============================================================================
# adds current site-package folder path
current_folder = os.path.realpath(os.path.abspath(os.path.split(inspect.getfile( inspect.currentframe() ))[0]))
local_site_packages_folder = current_folder + '/site-packages'
if local_site_packages_folder not in sys.path and os.path.exists(local_site_packages_folder):
    sys.path.insert(0, local_site_packages_folder)
#==============================================================================

### local site-packages includes starts
from eep import common
import clearcubic.utility

DIR_CURRENT_EXECUTABLE = os.path.dirname(sys.executable)
IMAGE_MAGIC_EXE = os.path.join(DIR_CURRENT_EXECUTABLE, "..", "ImageMagick-6.7.3", "convert.exe")
print IMAGE_MAGIC_EXE
DEFAULT_OUTPUT_PATH = os.path.join(
    'c:\projects\eep\data\_tmp',
    ''
)


# Windows
RESIZE_RESOLUTION = '354x425'
if os.name != 'nt':
    IMAGE_MAGIC_EXE = 'convert'
    RESIZE_RESOLUTION = '136x170'

def resize_photos_for_donor_doc(photos_path, out_path):
    if not photos_path:
        dir_cwd_path = os.path.abspath(os.getcwd())
        photos_path = os.path.join(
            dir_cwd_path,
            common.DIR_EEP_PHOTOS_CROPPED_DEFAULT,
            # '_ori'
        )

    pattern = os.path.join(photos_path, '*.jpg')
    print pattern
    files = glob.glob(pattern)

    for f in files:
        # print f
        file_base_name = os.path.basename(f)
        file_name, file_extension = os.path.splitext(file_base_name)

        target_f = os.path.join(out_path, file_base_name)
        #print target_f
        # run_cmd = IMAGE_MAGIC_EXE + ' ' + f + ' -resize 354x425 -density 180 ' + target_f
        # run_cmd = IMAGE_MAGIC_EXE + ' ' + f + ' -resize 136x170 -density 180 ' + target_f
        run_cmd = IMAGE_MAGIC_EXE + ' ' + f + ' -resize ' + RESIZE_RESOLUTION + ' -density 180 ' + target_f
        print run_cmd
        os.system(run_cmd)

def main():
    pass

if __name__ == '__main__':
    # python eep-photo-resizer.py -p C:\projects\eep\data\2017f\eep_photos_cropped_original -o C:\projects\eep\data\2017f\eep_photos_cropped
    # TODO: Need to remove the hard defined path below.
    src_photos_path = os.path.join(
        'C:\projects\eep\data\_to_resize',
        ''
    )
    output_path = os.path.join(
        'C:\projects\eep\data\\2017f\eep_photos_cropped',
        ''
    )

    parser = argparse.ArgumentParser(description='Resize EEP donor photos.')
    parser.add_argument('-p', nargs='?', default=src_photos_path)
    parser.add_argument('-o', nargs='?', default=output_path)

    args, unknown = parser.parse_known_args()
    print args.p, args.o
    resize_photos_for_donor_doc(args.p, output_path)
