### Time: 20190904
### Author: YaoLing
### Des: pdf convert to jpg

# -*- coding: UTF-8 -*-

from pdf2image import convert_from_path ## pip install pdf2image  or pip install --user pdf2image

pdf_name = input('输入PDF的文件路径')
jpg_name =pdf_name[:-4]+'.jpg'

pages = convert_from_path(pdf_name, 500)
i=1
for page in pages:
    page.save(pdf_name[:-4]+f'{i}.jpg', 'JPEG')
    i=i+1