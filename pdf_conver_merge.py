#coding:utf-8
import win32com.client
import os
import tempfile
import sys
from PyPDF2.pdf import PdfFileReader, PdfFileWriter
import tkinter as tk
from tkinter import filedialog,messagebox,Scrollbar
from os.path import *
import tkinter.font as tf

files_list=[]
def walk_dir(path):
    '''遍历目录找到 docx 文件'''
    #os.walk(),根目录下的每一个文件夹(包含它自己), 产生3-元组 (dirpath, dirnames, filenames)【文件夹路径, 文件夹名字, 文件名】
    for (root, subdirs, files) in os.walk(path):
        for fil in files:
            if str(fil).endswith('.docx') and (not str(fil).startswith('~$')) and (not str(fil).startswith('.')):
                yield os.path.join(root, fil)


def transformtion(wordapp, in_put, out_put):
    '''转 PDF'''
    wordapp.Documents.Open(in_put)  # 文件路径必须是绝对路径
    wordapp.ActiveDocument.SaveAs(out_put, FileFormat=win32com.client.constants.wdFormatPDF)
    wordapp.ActiveDocument.Close()


def merge_pdf(file_list, output_path):
    '''合并 PDF'''
    outpdf = PdfFileWriter()
    for f in file_list:
        f_pdf = PdfFileReader(open(f, 'rb'))
        for page in f_pdf.pages:
            outpdf.addPage(page)
    ous = open(output_path, 'wb')
    outpdf.write(ous)
    ous.close()

def get_files_list(inputs):
    if os.path.isfile(inputs) and str(inputs).endswith('docx'):
        files_list = [os.path.abspath(inputs)]
    elif os.path.isdir(inputs):
        files_list = list(walk_dir(inputs))
    else:
        files_list=[]
        print('输入错误！', file=sys.stderr)
        exit(1)
    for ind,each in enumerate(files_list):
        if ' 'in os.path.split(each)[1]:
            files_list.pop(ind)
            temp=os.path.split(each)[1].replace(' ','___')
            new_name=os.path.join(os.path.split(each)[0],temp)
            files_list.insert(ind,new_name)
            os.renames(each,new_name)
    return files_list

def main(inputs, cover=None, bcover=None,output_path=None):
    if output_path == None and os.path.exists(os.path.join(inputs,'PDF')) == False:
        os.mkdir(os.path.join(inputs,'PDF'))
    output_path=os.path.join(inputs,'PDF')
    wordapp = win32com.client.gencache.EnsureDispatch("Word.Application")
    tmpdir = tempfile.mkdtemp()
    for fil in files_list:
        pdf_name = os.path.splitext(os.path.basename(fil))[0] + '.pdf'
        # 转pdf
        temfile = os.path.join(tmpdir, pdf_name)
        transformtion(wordapp, fil, temfile)
        # 合并pdf(加封面、封底)
        outputfile = os.path.join(output_path, pdf_name)
        pdf_files = [cover, temfile, bcover]
        merge_pdf([p for p in pdf_files if p], outputfile)
        log='{} --> {}'.format(fil, outputfile)
        log_simple='{}:Done\n'.format(os.path.split(outputfile)[1])
        text_box.insert('end',log_simple.replace('___',' '))
        text_box.see('end')
        text_box.update()
        print(log)
    for eachw in os.listdir(inputs):
        os.chdir(inputs)
        os.renames(eachw,eachw.replace('___',' '))
    for eachp in os.listdir(os.path.join(inputs,'PDF')):
        os.chdir(os.path.join(inputs,'PDF'))
        os.renames(eachp,eachp.replace('___',' '))
    wordapp.Quit()

cover_path=None
bcover_path=None
word_path=None
app=tk.Tk()
app.title('Word-Convert')
label1=tk.Label(app,text='word目录').grid(row=0,column=0,ipadx=20)
label2=tk.Label(app,text='封面PDF').grid(row=1,column=0,ipadx=20)
label3=tk.Label(app,text='封底PDF').grid(row=2,column=0,ipadx=20)

def ini_text(row,column,master=app,height=1,width=10,ipadx=150,str='未选择'):
    text_temp=tk.Text(master,height=height,width=width,wrap='none')
    text_temp.insert('insert',str)
    text_temp.grid(row=row,column=column,ipadx=150)
    text_temp.config(state='disable')
    return text_temp

def set_text(text,str):
    text.config(state='normal')
    text.delete('1.0','end')
    text.insert('insert',str)
    text.see('end')

text_box1=ini_text(row=0,column=1)
text_box2=ini_text(row=1,column=1)
text_box3=ini_text(row=2,column=1)

app_log=tk.Tk()
app_log.title('转换进度')
ft = tf.Font(family='微软雅黑',size=8)
s1 = Scrollbar(app_log)
s1.pack(side='right', fill='y')
s2 = Scrollbar(app_log,orient='horizontal')
s2.pack(side='bottom', fill='x')
text_box=tk.Text(app_log,height=5,width=50,yscrollcommand=s1.set,xscrollcommand=s2.set,font=ft,wrap='none')
text_box.pack(expand='yes', fill='both')
app_log.withdraw()
s1.config(command=text_box.yview)
s2.config(command=text_box.xview)
def select_cover():
    global cover_path
    cover_path = tk.filedialog.askopenfilename(title=u'选择封面', initialdir=(os.path.expanduser('default_dir')))
    set_text(text_box2,cover_path)

def select_bcover():
    global bcover_path
    bcover_path = tk.filedialog.askopenfilename(title=u'选择封底', initialdir=(os.path.expanduser('default_dir')))
    set_text(text_box3,bcover_path)

def select_dir():
    global word_path,files_list
    word_path = tk.filedialog.askdirectory(title=u'选择目录', initialdir=(os.path.expanduser('default_dir')))
    set_text(text_box1,word_path)
    files_list=get_files_list(word_path)

def run():
    if word_path:
        if files_list != []:
            app_log.deiconify()
            app.withdraw()
            text_box.insert('insert','start...\n')
            text_box.update()
            main(word_path, cover_path, bcover_path)
            text_box.insert('insert','done...')
            # app.deiconify()
            tk.messagebox.showinfo(title='Done', message='已完成所有word文件的转换')
            app.quit()
        else:
            tk.messagebox.showwarning(title='Done', message='未在该目录下找到word文件，请重新选择')
            app.deiconify()
    else:
        tk.messagebox.showwarning(title='Done', message='请选择word目录')
        app.deiconify()

tk.Button(app,text='开始转换',width=10,command=run,bg='#FFFFFF').grid(row=3,column=0,columnspan=2,padx=0,pady=5)
tk.Button(app,text='退出',width=10,command=app.quit,bg='#FFFFFF').grid(row=3,column=1,columnspan=2,padx=0,pady=5)
tk.Button(app,text='更改',width=10,command=select_dir,bg='#FFFFFF').grid(row=0,column=2,padx=10,pady=5,sticky='W')
tk.Button(app,text='更改',width=10,command=select_cover,bg='#FFFFFF').grid(row=1,column=2,padx=10,pady=5,sticky='W')
tk.Button(app,text='更改',width=10,command=select_bcover,bg='#FFFFFF').grid(row=2,column=2,padx=10,pady=5,sticky='W')

app.mainloop()


