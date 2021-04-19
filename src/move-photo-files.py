import errno
import os 
import pathlib
import shutil

def move_photo_files(src_dir, dst_dir):
    src_dir = r"/Users/cc/Documents/eep/2020s/china-photos"
    dst_dir = r"/Users/cc/Documents/eep/2020s/china-photos/flattend"

    interesting_exts = ["jpg"]
    for subdir, dirs, files in os.walk(src_dir):
        for filename in files:
            filepath = subdir + os.sep + filename

            ext = filename.split(".")[-1].lower()
            # Has capital letters, we need to rename file

            new_filepath = dst_dir + os.sep + filename.lower()
            print("-----")
            print(new_filepath)
            print(filepath)
                # print(filepath)


            
            # if fp.endswith(".JPG"):


            if ext in interesting_exts:
                shutil.copyfile(filepath, new_filepath)



def create_directories(directory_list):
    for rd in directory_list:
        try:
            print("Checking+Creating: %s" % rd)
            os.makedirs(rd)
        except OSError as e:
            if e.errno != errno.EEXIST:
                raise


move_photo_files(None, None)
