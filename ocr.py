# some code taken from askpython.com & geeksforgeeks.org

import pandas as pd
import os
import cv2
from PIL import Image
import pytesseract

def scan(image):
    pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
    print(image)
    text = ''
    image = str(image)

    if image.endswith('.jpg') or image.endswith('.jpeg') or image.endswith('.png'):
        image = cv2.imread(image)
        try:
            image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            _, threshold_image = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            image = cv2.fastNlMeansDenoising(threshold_image, None, 30, 7, 21)
        except Exception as err:
            print(f'Error: {err}')
        try:
            custom_config = r'--oem 3 --psm 3'
            text_data = pytesseract.image_to_string(image, config=custom_config)
            text += text_data
        except Exception as err:
            print(f'Error: {err}')

    return text
