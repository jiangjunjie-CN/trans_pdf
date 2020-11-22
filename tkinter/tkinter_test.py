import tkinter as tk
from tkinter import filedialog
import os
from os.path import *
app=tk.Tk()
app.title('test_gui')
label1=tk.Label(app,text='需要解读的报告').grid(row=0,column=0)
label2=tk.Label(app,text='肿瘤类型').grid(row=1,column=0)
ent1=tk.Entry(app)
ent2=tk.Entry(app)
ent2.grid(row=1,column=1,padx=10,pady=5)


def get_info():
    global cancer
    cancer=ent2.get()
    ent2.delete(0,tk.END)

def select_file():
    global file_path
    file_path = tk.filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser('default_dir')))

tk.Button(app,text='确认',width=10,command=get_info).grid(row=3,column=0,padx=10,pady=5,sticky='W')
tk.Button(app,text='退出',width=10,command=app.quit).grid(row=3,column=1,padx=10,pady=5,sticky='E')
tk.Button(app,text='选择文件',width=10,command=select_file).grid(row=0,column=1,padx=10,pady=5,sticky='W')

app.mainloop()

print(file_path)
print(cancer)