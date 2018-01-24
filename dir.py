import os, sys
from os import path

#_PATH_PYTHON = os.environ['PYTHONPATH'].split(os.sep)
#_PATH_PYTHON_EXE = sys.executable

PHOTOS_ORIGINAL_FOLDER_NAME = 'eep_photos_original'
PHOTOS_CROPPED_FOLDER_NAME = 'eep_photos_cropped'

DIR_APP = os.getcwd()
DIR_DATA = os.path.join(DIR_APP, 'data')
DIR_EEP_PHOTOS_ORIGINAL_DEFAULT = path.join(DIR_DATA, PHOTOS_ORIGINAL_FOLDER_NAME)
DIR_EEP_PHOTOS_CROPPED_DEFAULT = path.join(DIR_DATA, PHOTOS_CROPPED_FOLDER_NAME)

DIR_TEMPLATES = path.join(DIR_APP, 'templates')
DIR_ASSETS = path.join(DIR_APP, 'assets')
DIR_OUTPUT = path.join(DIR_APP, 'output')
