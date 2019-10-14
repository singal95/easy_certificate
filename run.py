"""a simple script for automatically generating docx file of certificates"""

# !/usr/bin/env python
import os
import time
# from math import sqrt

from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import arg_parser, tid_maker

IMAGES_PER_ROW = 2
MARGIN = 0.5


def load_images():
    """return paths of images of pwd"""
    image_files = list(f for f in os.listdir()
                       if f.endswith('.png') or f.endswith('.jpg') or f.endswith('.bmp'))
    if image_files:
        return image_files
    else:
        return None


def run(img_per_row=IMAGES_PER_ROW, margin=MARGIN):
    """main excecution func"""
    if img_per_row >= 10:
        raise ValueError("每行图片太多！请设置较小于等于10的数值")
    if margin > 10:
        raise ValueError("页边距太大！请设置小于10的页边距")

    # load images
    image_files = load_images()

    if image_files == None:
        print("此目录下无图片，请放置图片后重启此程序！")
        time.sleep(5)
        raise IOError
    image_nums = len(image_files)

    # start to write document in docx format
    document = Document()

    # controls page properties
    section = document.sections[0]
    section.page_height = Cm(21.0)  # horizontal A4 format
    section.page_width = Cm(29.7)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.top_margin = Cm(margin)
    section.bottom_margin = Cm(margin)
    section.left_margin = Cm(margin)
    section.right_margin = Cm(margin)

    # write tables
    tables = []
    for k in range(image_nums):
        table = document.add_table(img_per_row, img_per_row)
        table.style = 'Table Grid'
        tables.append(table)
        for i in range(img_per_row):
            for j in range(img_per_row):
                # set paragraph properties
                para = table.cell(i, j).paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                cell = para.add_run()

                # add image
                cell.add_picture(image_files[k],
                                 width=Cm((29.6 - 2*margin) / img_per_row),
                                 height=Cm((20.9 - 2*margin) / img_per_row))

    # finally save the .docx
    # if not os.path.isfile('my_demo.docx'):
    #     document.save('my_demo.docx')
    # else:
    #     docx_name = tid_maker() + '.docx'
    #     document.save(docx_name)
    docx_name = tid_maker() + '.docx'
    document.save(docx_name)

if __name__ == "__main__":
    args = arg_parser()
    run(img_per_row=args.number, margin=args.margin)
