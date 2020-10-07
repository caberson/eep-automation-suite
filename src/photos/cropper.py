import glob, os, sys, re, fnmatch
import cv2
import cv2 as cv

from commonphotocropper import draw_str
import inspect

class FaceCropper:
    """EEP Cropper"""
    # print inspect.stack()[0][1]
    DIR_CURRENT_EXECUTABLE = os.path.dirname(sys.executable)
    DIR_CURRENT_EXECUTABLE = os.path.dirname(inspect.stack()[0][1])

    dir_photos_original = ''
    dir_photos_cropped = ''
    original_photos = []

    current_view_position = 0
    auto_save_on_view = False

    # objects
    current_image = None
    current_thumbnail = None
    current_rect = None

    image_magic_exe = None

    MAX_THUMB_WIDTH = 1024.0
    MAX_THUMB_HEIGHT = 768.0
    MAX_THUMB_WIDTH = 640.0
    MAX_THUMB_HEIGHT = 480.0
    MAX_THUMB_WIDTH = 800.0
    MAX_THUMB_HEIGHT = 600.0
    #MAX_THUMB_WIDTH = 1024.0
    #MAX_THUMB_HEIGHT = 768.0

    #opencv related variables
    opencvHaarCascadePath = os.path.join(
        # DIR_CURRENT_EXECUTABLE, "..", "lib", "opencv", "data", "haarcascades"
        # '..', 'lib',  "opencv", "data", "haarcascades"
        cv2.__path__[0], "data"
    ) # 'D:\_cc\portables\PortablePython2.7/opencv/data/haarcascades/'
    print opencvHaarCascadePath
    cascade_fn = os.path.join(opencvHaarCascadePath, "haarcascade_frontalface_alt.xml")
    # cascade_fn = os.path.join("haarcascade_frontalface_alt.xml")
    print cascade_fn
    cascade = cv2.CascadeClassifier(cascade_fn)

    def __init__(self, dir_photos_original, dir_photos_cropped, image_magic_exe=None):
        self.dir_photos_original = dir_photos_original
        self.dir_photos_cropped = dir_photos_cropped
        self.image_magic_exe = image_magic_exe
        
        self.resize_image = True
        if image_magic_exe is None:
            self.resize_image = False
            print 'Warning: ImageMagicExe not set.  Thumbnail resize will not be done'

        self.get_image_list_from_directory()

    def get_image_list_from_directory(self):
        pattern = os.path.join(self.dir_photos_original, '*.jpg')
        # pattern2 = re.compile(fnmatch.translate(pattern), re.IGNORECASE)
        self.original_photos = glob.glob(pattern)
        # Hackish way to ignore case
        pattern2 = os.path.join(self.dir_photos_original, '*.JPG')
        self.original_photos = self.original_photos + glob.glob(pattern2)

        return self.original_photos.reverse()

    def get_current_image_file_name(self):
        fn = self.original_photos[self.current_view_position]
        file_base_name = os.path.basename(fn)
        file_name, file_extension = os.path.splitext(file_base_name)

        return (file_base_name, file_name, file_extension)

    #==============================================================

    def detect_small(self, img, cascade):
        #rects = cascade.detectMultiScale(img, scaleFactor=1.3, minNeighbors=4, minSize=(30, 30), flags = cv.CASCADE_SCALE_IMAGE)
        rects = cascade.detectMultiScale(
            img, scaleFactor=1.1, minNeighbors=4, minSize=(80, 80) #, flags=cv.CASCADE_SCALE_IMAGE
        )

        if len(rects) == 0:
            return []
        rects[:, 2:] += rects[:, :2]
        return rects

    def draw_rect(self, img, rect, color):
        x1, y1, x2, y2 = rect
        cv2.rectangle(img, (x1, y1), (x2, y2), color, 1)

    def get_cropped_rect(self, img, rect=None, x_conversion_factor=1.0):#rect = face rectangle
        #global currentRect
        #print 'convert face rect ', rect, ' to crop rect.'
        h, w, d = img.shape
        cropped_width = h / 5 * 4

        # if no face rect is provided, return a 4x5 rect in the center of the image
        if rect is None:
            img_center_x = w / 2
            x1 = img_center_x - (cropped_width / 2)
            x2 = img_center_x + (cropped_width / 2)
            return [x1, 0, x2, h]

        #print 'image w, h: ', w, h
        x1, y1, x2, y2 = rect
        x1 = int(x1 * x_conversion_factor)
        x2 = int(x2 * x_conversion_factor)


        #print 'CroppedWidth: ', cropped_width

        width_to_extent = (cropped_width - (x2 - x1)) / 2
        #print 'w:', w, 'x2:', x2, 'x1:',x1, 'x2-x1:', x2-x1,
        #print 'widthToExtent:', width_to_extent
        cropped_x2 = x2 + width_to_extent # + (cropped_width / 2)
        # make sure we don't exceed width
        if cropped_x2 > w:
            cropped_x2 = w

        # make sure we don't go below 0
        cropped_x1 = x1 - width_to_extent
        if cropped_x1 < 0:
            cropped_x1 = 0

        return [cropped_x1, 0, cropped_x2, h]

    def zoom_current_rectangle_size(self, zoom_increment):
        x1, y1, x2, y2 = self.current_rect
        pct = zoom_increment / 100.0
        scaleRatio = 1.0 + pct

        currentThumbnailH, currentThumbnailW, currentThumbnailD = self.current_thumbnail.shape

        xScaleIncrement = (x2 - x1) * pct
        yScaleIncrement = (y2 - y1) * pct
        #print 'xScale:', str(xScaleIncrement) + ' yScale:', str(yScaleIncrement)


        newX1 = int(x1 - xScaleIncrement)
        newX2 = int(x2 + xScaleIncrement)
        newY1 = int(y1 - yScaleIncrement)
        newY2 = int(y2 + yScaleIncrement)

        if newX1 < 0 or newX2 > currentThumbnailW or newY1 < 0 or newY2 > currentThumbnailH:
            # exceeded current view
            return self.get_image_thumbnail()
            return

        self.current_rect = [newX1, newY1, newX2, newY2]
        return self.get_image_thumbnail()

    def update_crop_rect_horizontal_position(self, position_increment):
        x1, y1, x2, y2 = self.current_rect
        self.current_rect = [x1 + position_increment, y1, x2 + position_increment, y2]

        return self.get_image_thumbnail()

    def process_image_for_detection(self, fn):
        img = cv2.imread(fn)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.equalizeHist(gray)
        return img, gray

    def get_image_thumbnail(self):
        fn = self.original_photos[self.current_view_position]
        file_base_name = os.path.basename(fn)

        vis = self.current_thumbnail.copy()
        # print 'rect:', rects.size
        self.draw_rect(vis, self.current_rect, (0, 255, 0))

        x1, y1, x2, y2 = self.current_rect
        msg = "file: {}".format(file_base_name)
        draw_str(vis, (x2 - 175, y2 - 5), msg)

        return vis
        #cv2.imshow('photosToCrop', vis)

    def save_cropped_image(self, override=True):
        fn = self.original_photos[self.current_view_position]
        file_base_name = os.path.basename(fn)
        file_name, file_extension = os.path.splitext(file_base_name)

        cropped_file = os.path.join(self.dir_photos_cropped, file_base_name)

        msg = 'Saving: %s' % os.path.join(self.dir_photos_cropped, file_base_name)
        if not override and os.path.exists(cropped_file):
            msg += '.  File exists.  Override setting: No.'
            print msg
            return
        else:
            print msg

        currentImageH, currentImageW, currentImageD = self.current_image.shape
        currentThumbImageH, currentThumbImageW, currentThumbImageD = self.current_thumbnail.shape
        up_conversion_x_ratio = float(currentImageW) / currentThumbImageW
        up_conversion_y_ratio = float(currentImageH) / currentThumbImageH

        """
        print 'Saving Cropped Image:', currentImageH, self.current_rect, currentThumbnail.shape
        up_conversion_x_ratio = float(currentImageH) / currentRect[3]
        cropped_rect = get_cropped_rect(currentImg, currentRect, up_conversion_x_ratio)
        x1, y1, x2, y2 = cropped_rect
        """

        x1 = int(self.current_rect[0] * up_conversion_x_ratio)
        y1 = int(self.current_rect[1] * up_conversion_y_ratio)
        x2 = int(self.current_rect[2] * up_conversion_x_ratio)
        y2 = int(self.current_rect[3] * up_conversion_y_ratio)
        # crop region of interest
        crop_vio_roi = self.current_image[y1:y2, x1:x2]
        #crop_vio_roi = vis_roi

        cropped_n_resized_vio_roi = cv2.resize(crop_vio_roi.copy(), (354, 425))
        #cv2.imwrite(photosCroppedDir + croppedFileName, crop_vio_roi)
        cv2.imwrite(cropped_file, cropped_n_resized_vio_roi)

        print 'Saving Cropped Image.'

        self.resize_img(cropped_file)
    
    def resize_img(self, src_file, target_file=None):
        if target_file is None:
            target_file = src_file
        
        if self.resize_image:
            run_cmd = self.image_magic_exe + ' ' + src_file + ' -units PixelsPerInch  -resize 354x425 -density 180 ' + target_file
            # print run_cmd
            os.system(run_cmd)

    def test_image(self, img):
        import numpy
        capture = cv.fromarray(img, True)
        print cv.GetDims(capture)

        #masked_image = cv.CreateImage(cv.GetSize(capture), 8, 3)
        #print img
        #cv.SetZero(masked_image)
        #cv.Not(capture, masked_image)
        #cv.Copy(capture, masked_image)
        return numpy.asarray(capture)

        grey = cv.CreateImage(cv.GetSize(capture), 8, 1)
        cv.CvtColor(capture, grey, cv.CV_BGR2GRAY)
        cv.Threshold(grey, grey, 100, 255, cv.CV_THRESH_BINARY)
        cv.Not(grey, grey)
        cv.Copy(capture, grey)
        t = numpy.asarray(grey)

    def load_image(self):
        fn = self.original_photos[self.current_view_position]
        # print 'Loading File: ', fn
        img, gray = self.process_image_for_detection(fn)

        maxThumbWidth = self.MAX_THUMB_WIDTH
        maxThumbHeight = self.MAX_THUMB_HEIGHT

        h, w, d = img.shape
        dstW = w
        dstH = h
        # resize if photo is larger than our max dimension
        if w > maxThumbWidth and h > maxThumbHeight:
            if w > h:
                ratio = maxThumbWidth / w
                dstW = int(maxThumbWidth)
                dstH = int(h * ratio)
            else:
                ratio = maxThumbHeight / h
                dstH = int(maxThumbHeight)
                dstW = int(w * ratio)

        #print dstW, dstH
        thumbnail = cv2.resize(img.copy(), (dstW, dstH))
        thumbnailGray = cv2.resize(gray.copy(), (dstW, dstH))

        #return self.test_image(thumbnail)
        rects = self.detect_small(thumbnailGray, self.cascade)

        # populate global variables
        self.current_image = img
        self.current_thumbnail = thumbnail
        #currentRect = rects[0]
        # rects[0]

        if len(rects) > 0:
            self.current_rect = self.get_cropped_rect(thumbnail, rects[0])
            #currentRect = rects[0] --face rectangle
        else:
            self.current_rect = self.get_cropped_rect(thumbnail)# get_fullimg_rect(thumbnail)

        #update_image_thumbnail()

        #save a default cropped image if non exists already so if no change is needed, no need to do anything
        if self.auto_save_on_view:
            self.save_cropped_image(override=False)

        return self.get_image_thumbnail()

    def delete_current_image(self):
        file_name = self.get_current_image_file_name()[0]
        pathToOriginalFileForDelete = os.path.join(self.dir_photos_original, file_name)
        pathToThumbnailFileForDelete = os.path.join(self.dir_photos_cropped, file_name)

        print 'Deleting\n\rOriginal > ', pathToOriginalFileForDelete, '\n\rThumbnail > ', pathToThumbnailFileForDelete

        try:
            os.unlink(pathToOriginalFileForDelete)
        except:
            print 'Error deleting original file'

        try:
            os.unlink(pathToThumbnailFileForDelete)
        except:
            print 'Error deleting cropped file'

        del self.original_photos[self.current_view_position]

    #=================================================
    #NAVIGATION functions
    #=================================================

    def get_current_photo(self):
        return self.load_image()

    def get_previous_photo(self):
        if self.current_view_position > 0:
            self.current_view_position -= 1
        return self.load_image()

    def get_next_photo(self):
        if self.current_view_position < len(self.original_photos) - 1:
            self.current_view_position += 1
        return self.load_image()

    def set_current_file_name_to(self, new_file_name):
        self.set_file_name_to(new_file_name, self.current_view_position)

    def set_file_name_to(self, new_file_name, position):
        self.original_photos[position] = new_file_name
