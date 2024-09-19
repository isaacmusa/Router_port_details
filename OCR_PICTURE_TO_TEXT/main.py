import pytesseract as pyt
import cv2 

img = cv2.imread("C:\\Users\\ISAAC AYEGBA\\Desktop\\Isaac\\test image.jpeg") 

pyt.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

text  = pyt.image_to_string(img)

print(text)

