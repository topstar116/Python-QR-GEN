import os, sys
import qrcode
from PIL import Image
import cv2 as cv

img = cv.imread('row_3_qrcode.png')
det = cv.QRCodeDetector()
retval, points, straight_qrcode = det.detectAndDecode(img)

print(retval)
print(points)
print(straight_qrcode)
