from pdf2docx import Converter
import os

#pdf_file = '/path/to/sample.pdf'
#docx_file = 'path/to/sample.docx'

pdf_file=input('输入PDF的文件路径：')


# convert pdf to docx
cv = Converter(pdf_file)
cv.convert(os.path.join(os.path.dirname(pdf_file),'test.docx')) # 默认参数start=0, end=None
cv.close()

# more samples
# cv.convert(docx_file, start=1) # 转换第2页到最后一页
# cv.convert(docx_file, pages=[1,3,5]) # 转换第2，4，6页