#!/usr/bin/env python3
# coding:utf-8
#conda install reportlab
#conda install -c conda-forge pypdf2

from PyPDF2 import PdfFileWriter, PdfFileReader
from copy import copy
def create_watermark(input_pdf, output, watermark):
    watermark_obj = PdfFileReader(watermark)
    watermark_page = watermark_obj.getPage(0)
 
    pdf_reader = PdfFileReader(input_pdf)
    pdf_writer = PdfFileWriter()
 
    # 给所有页面添加水印
    Pagelist=[i for i in range(0,pdf_reader.getNumPages())]
    # print(Pagelist)
    # while n<pdf_reader.getNumPages():
    for page in Pagelist:
        print(page)
        # print('getNumPages',pdf_reader.getNumPages())
        if page ==0 or page ==pdf_reader.getNumPages()-1:
            page = pdf_reader.getPage(page)
            pdf_writer.addPage(page)
        else:
            page = pdf_reader.getPage(page)
            new_page =copy(watermark_page)
            new_page.mergePage(page)
            pdf_writer.addPage(new_page)
            del page

    with open(output, 'wb') as out:
        pdf_writer.write(out)
 
if __name__ == '__main__':
    create_watermark(
        input_pdf='/mnt/c/Users/luping/Desktop/20211219_21T005506_黄诗雯TNP-seq_宏基因组+靶向三代纳米孔病原微生物基因检测及耐药基因鉴定DNA.pdf',
        output='/mnt/c/Users/luping/Desktop/水印2-20211219_21T005506_黄诗雯TNP-seq_宏基因组+靶向三代纳米孔病原微生物基因检测及耐药基因鉴定DNA.pdf',
        watermark='/mnt/c/Users/luping/Desktop/水印2.pdf')
