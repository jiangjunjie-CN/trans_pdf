# -*- coding:utf-8 -*-
# import win32com.client
# import os
# import tempfile
# import sys
# from PyPDF2.pdf import PdfFileReader, PdfFileWriter
import tkinter as tk
from tkinter import filedialog
from os.path import *
from trans_pdf import *

cover_path=None
bcover_path=None
app=tk.Tk()
app.title('Word-Convert')
label1=tk.Label(app,text='需要转换的word目录').grid(row=0,column=0)
label2=tk.Label(app,text='封面文件').grid(row=1,column=0)
label3=tk.Label(app,text='封底文件').grid(row=2,column=0)
tk.Label(app,text='未选择').grid(row=0,column=2,ipadx=150)
tk.Label(app,text='未选择').grid(row=1,column=2,padx=150)
tk.Label(app,text='未选择').grid(row=2,column=2,padx=150)
# ent1=tk.Entry(app)
# ent2=tk.Entry(app)
# ent2.grid(row=1,column=1,padx=10,pady=5)
# text_box=tk.Text(app,height=3).grid(row=0,column=3)
text_box=tk.Text(app,height=3)
text_box.insert('insert','start')
text_box.grid(row=0,column=3)
def get_info():
    global cancer
    cancer=ent2.get()
    ent2.delete(0,tk.END)

def select_cover():
    global cover_path
    cover_path = tk.filedialog.askopenfilename(title=u'选择封面', initialdir=(os.path.expanduser('default_dir')))
    tk.Label(app,text=cover_path).grid(row=1,column=2)

def select_bcover():
    global bcover_path
    bcover_path = tk.filedialog.askopenfilename(title=u'选择封底', initialdir=(os.path.expanduser('default_dir')))
    tk.Label(app,text=bcover_path).grid(row=2,column=2)

def select_dir():
    global word_path
    word_path = tk.filedialog.askdirectory(title=u'选择目录', initialdir=(os.path.expanduser('default_dir')))
    tk.Label(app,text=word_path).grid(row=0,column=2)

def run():
    # with open('trans_pdf.py','r',encoding= 'utf-8') as t:
    #     exec(t.read())
    main(word_path, cover_path, bcover_path)
    # text_box.insert('insert',log_out())
    text_box.insert('insert','done')
tk.Button(app,text='确认',width=10,command=run).grid(row=3,column=0,padx=10,pady=5,sticky='W')
tk.Button(app,text='退出',width=10,command=app.quit).grid(row=3,column=1,padx=10,pady=5,sticky='E')
tk.Button(app,text='选择文件目录',width=10,command=select_dir).grid(row=0,column=1,padx=10,pady=5,sticky='W')
tk.Button(app,text='选择封面',width=10,command=select_cover).grid(row=1,column=1,padx=10,pady=5,sticky='W')
tk.Button(app,text='选择封底',width=10,command=select_bcover).grid(row=2,column=1,padx=10,pady=5,sticky='W')


app.mainloop()


