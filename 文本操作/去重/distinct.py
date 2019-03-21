# coding=utf-8

import os


# 使用set去重，不能保证顺序不变
# 源文件改名为：.bak
def distinct_line(file_name):
    f = open(file_name, 'r',encoding="utf-8")
    lines = set([])
    while True:
        line = f.readline()
        if not line:
            break
        else:
            lines.add(line)
    f.close()
    os.rename(file_name, file_name+".bak")
    f = open(file_name, 'w', encoding="utf-8")
    f.writelines(lines)
    f.close()

distinct_line('test.txt')