import pytesseract
import re
from pdf2image import convert_from_path

page = convert_from_path("./test.pdf", single_file=True)

print("- Convert the first page of scanned pdf file to image")
print(type(page))
print(page[0])

text = pytesseract.image_to_string(page[0],lang='chi_sim')
print(text)
#idcard = re.search('', text)
#print(idcard)
