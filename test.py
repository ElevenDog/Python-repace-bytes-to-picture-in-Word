# coding:UTF-8
from docx import Document
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
import datetime


def readtxt():
    # 1.读文本文件，取匹配关系文本
    filepath = __file__
    realpath = os.path.realpath(filepath)
    current_path = os.path.dirname(realpath)
    txtpath = getFileName(current_path, '.txt')
    docpath = getFileName(current_path, '.docx')

    if txtpath and docpath:
        list= []
        f = open(txtpath, 'r')
        k = 0
        line =f.readline()
        while line.strip():
            arr = line.split('$$$$')
            sel_str = arr[0]
            imgpath = arr[1].strip()
            if os.path.exists(imgpath):
                txtchangepicture(docpath, sel_str, imgpath)
                k += 1
                print("替换完成次数%d" % k)
            else:
                print("文件%s 不存在" % imgpath)
            line=f.readline()
        f.close()
        #
        return "执行成功次数%d" % k
    else:
        return "缺少源文件（*.doc）或配置文件"


def getFileName(path, pattern):
    ''' 获取指定目录下的所有指定后缀的文件名 '''
    f_list = os.listdir(path)
    for i in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(i)[1] == pattern:
            return i


def txtchangepicture(docpath, sel_str, imgpath):
    # 替换文本为模板对象
    document = Document(docpath)
    for para in document.paragraphs:
        if sel_str in para.text:
            inline = para.runs
            for i in inline:
                if sel_str in i.text:
                    if "{{%s}}" % sel_str not in i.text:
                        text = i.text.replace(sel_str, "{{%s}}" % sel_str)
                        i.text = text

    document.save(docpath)

    tpl = DocxTemplate(docpath)
    context = {
        sel_str:InlineImage(tpl, imgpath, width=Mm(20)),
    }
    tpl.render(context)
    date = datetime.datetime.now()
    filepath = __file__
    realpath = os.path.realpath(filepath)
    current_path = os.path.dirname(realpath)
    path = os.path.join(current_path, "%s%s%s" % (date.year, date.month, date.day))
    if not os.path.exists(path):
        os.mkdir(path)
    filepath = os.path.join(path, "%s%s%s.docx" % (date.hour, date.minute, date.second))
    tpl.save(filepath)


info = readtxt()
print(info)
input('Press Enter to exit...')

