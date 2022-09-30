import os, errno

def create_directories(directory_list):
    for rd in directory_list:
        try:
            print("Checking+Creating: %s".format(rd))
            os.makedirs(rd)
        except OSError as e:
            if e.errno != errno.EEXIST:
                raise

