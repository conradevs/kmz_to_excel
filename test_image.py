from PIL import Image
from pytesseract import pytesseract

#Open image with PIL

import cv2

image = cv2.imread(r'C:\Users\Conrado\Downloads\3.jpg', 0)
inverted_image = cv2.bitwise_not(image)


#Define path to tessaract.exe
path_to_tesseract = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

#Point tessaract_cmd to tessaract.exe
pytesseract.tesseract_cmd = path_to_tesseract

min_start = 10
max_start = 255

# threshold to filter grays
thresh, im_bw = cv2.threshold(inverted_image,100, 250,cv2.THRESH_BINARY)
#Extract text from image
text = pytesseract.image_to_string(im_bw)
text_no_colon = text.replace(',', '.')
print(text_no_colon)

# write result image
cv2.imwrite(r'C:\Users\Conrado\Downloads\inverted.jpg', inverted_image)
cv2.imwrite(r'C:\Users\Conrado\Downloads\inverted-chop.jpg', im_bw)