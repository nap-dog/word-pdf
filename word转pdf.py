#coding=utf-8

from win32com.client import gencache
from win32com.client import constants, gencache
import os
import tkinter
from tkinter import filedialog
import msvcrt
import threading
from multiprocessing.dummy import Pool as ThreadPool


def createPdf(wordPath, pdfPath):
    """
    word转pdf
    :param wordPath: word文件路径
    :param pdfPath:  生成pdf文件路径
    """
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(wordPath, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfPath,
                            constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    word.Quit(constants.wdDoNotSaveChanges)

def start(filename):
    (filepath, tempfilename) = os.path.split(filename)
    (pdfPath1, extension) = os.path.splitext(tempfilename)
    pdfPath=filepath+"\\"+pdfPath1+".pdf"
    createPdf(filename,pdfPath)
    print(filename,end="  转换为  ")
    print(pdfPath)
    print()


def main():
    root = tkinter.Tk()    # 创建一个Tkinter.Tk()实例
    root.withdraw()       # 将Tkinter.Tk()实例隐藏
    default_dir = r"文件路径"
    file_path = tkinter.filedialog.askopenfilenames(title=u'选择文件', initialdir=(os.path.expanduser(default_dir)),filetypes=[('Word', '*.doc *.docx')])
    filename=[]
    for i in range(len(file_path)):
        name=file_path[i].replace('/','\\')
        filename.append(name)
    print('开始转换,一共%s个文档' %(len(filename)))
    print()
    pool = ThreadPool()
    pool.map(start,filename)
    pool.close()
    pool.join()
    print('转换完成')
    print()
    input('按任意键退出')

if __name__ == '__main__':
    main()