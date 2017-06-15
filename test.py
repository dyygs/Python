#!usr/bin/python
# -*- coding: utf-8 -*-
import urllib2
import re
from datetime import date

import xlwt

import Tkinter
import tkFileDialog

root = Tkinter.Tk()


def openfile():
    root.geometry('%sx%s+%s+%s' % (root.winfo_width() + 100, root.winfo_height() + 100,
                                   0, 0))
    r = tkFileDialog.askopenfilename(title='打开文件', filetypes=[('Html', '*.htm'), ('All Files', '*')])
    root.destroy()
    getpage(r)
# def savefile():
#     r = tkFileDialog.asksaveasfilename(title='保存文件', initialdir='d:\mywork', initialfile='hello.py')


def showOpenFileButton():
    root.wm_attributes('-topmost', 1)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight() - 100
    root.geometry('%sx%s+%s+%s' % (root.winfo_width() + 100, root.winfo_height() + 100,
                               (screen_width - root.winfo_width())/2, (screen_height - root.winfo_height())/2))
    btn1 = Tkinter.Button(root, text='File Open', command=openfile)
    # btn2 = Tkinter.tkinter.Button(root, text='File Save', command=savefile)
    btn1.pack()
    # btn2.pack(side='left')
    root.mainloop()


headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_0) '
                          'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36',
           'Referer': 'https://mp.weixin.qq.com'}


def getpage(r):
    url = 'file://' + r
    request = urllib2.Request(url, headers=headers)
    content = urllib2.urlopen(request).read()
    # print content
    # pattern = re.compile('author clearfix.*?title="(.*?)">.*?<span>(.*?)</span>(.*?)stats-vote.*?'
    #                      '"number">(.*?)</i>.*?qiushi_comments.*?"number">(.*?)</i>',re.S)
    pattern = re.compile('.*?content":"(.*?)","date_time".*?nick_name":"(.*?)","refuse_reason', re.S)
    txt = re.findall(pattern, content)
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet("sheet1")
    row = 0
    column = 0
    dateNow = date.today()
    sheetname = dateNow.strftime('%Y-%m-%d') + ".xls"
    print sheetname
    sheet.write(column, row, '用户昵称')
    sheet.write(column, row+1, '消息内容')
    for a_txt in txt:
        column = column + 1
        sheet.write(column, row, a_txt[1])
        sheet.write(column, row + 1, a_txt[0])
    book.save(sheetname)
if __name__ == '__main__':
    showOpenFileButton()
