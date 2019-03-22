# coding=utf-8

import os
import docx
from word_manipulation import word
from word_manipulation import docx_enhanced

# 使用set去重，不能保证顺序不变
# 源文件改名为：.bak
def txt_distinct_line(file_name):
    f = open(file_name, 'r')
    lines = set([])
    while True:
        line = f.readline()
        if not line:
            break
        else:
            lines.add(line)
    f.close()
    os.rename(file_name, file_name+".bak")
    f = open(file_name, 'w')
    f.writelines(lines)
    f.close()

# 删除word文档中重复的paragraph（不打乱原顺序）
def word_distinct_line(file_name):
    if file_name[-3:] == "doc":
        word.doc2docx(file_name, file_name+"x")
        file_name += "x"
    docx_file = docx.Document(file_name)
    dict = {}
    for paragraph in docx_file.paragraphs:
        print(paragraph.text)
        if paragraph.text not in dict:
            dict[paragraph.text] = 1
        else:
            docx_enhanced.delete_paragraph(paragraph)
    docx_file.save(file_name)
    word.docx2doc(file_name[:-1], file_name)


word_distinct_line(r"C:\Users\liuqinh2s\PycharmProjects\data_report\新建 Microsoft Word 文档.doc")