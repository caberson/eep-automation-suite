import os, errno

def create_directories(directory_list):
    for rd in directory_list:
        try:
            print "Checking+Creating: %s" % rd
            os.makedirs(rd)
        except OSError, e:
            if e.errno != errno.EEXIST:
                raise