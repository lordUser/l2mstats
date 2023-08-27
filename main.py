from PIL import Image
import pytesseract
import os
import pandas as pd

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
table = {'параметр': [], 'значение': []}

for filename in os.listdir(os.getcwd()+"\/l2m_stats")[0:1]:
    text = pytesseract.image_to_string(Image.open(os.path.join(os.getcwd()+"\/l2m_stats", filename)), lang="rus")
    result = text.split("\n")
    for item in result:
        value = item.split(" ")[-1]
        key = " ".join(item.split(" ")[0:-1])
        if key:
            table['параметр'].append(key)
        if value:
            table['значение'].append(value)
df = pd.DataFrame(table)
df.to_csv("table.csv", index=False)
