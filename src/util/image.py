import os

def resize_img(src_file, target_file=None, img_magic_exe=None):
    if target_file is None:
        target_file = src_file
    
    run_cmd = img_magic_exe + ' ' + src_file + ' -units PixelsPerInch  -resize 354x425 -density 180 ' + target_file
    # print(run_cmd)
    os.system(run_cmd)
