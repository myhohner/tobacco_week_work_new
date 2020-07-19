import tkinter as tk
from tkinter import ttk
import time
import os,threading,datetime
import re
from crawl import Crawl
from excel import Excel


class Work:
    def __init__(self):
        self.c=Crawl()
        self.e=Excel()

    def thread_it(self,func):
        # 创建线程
        t = threading.Thread(target=func)
        # 守护线程
        t.setDaemon(True)
        # 启动
        t.start()

    def setUp(self):
        #pb.start()
        self.c.setUp()
        #pb.stop()

    def crawl(self):
        var.set('')
        start_row=int(start.get())
        end_row=int(end.get())
        list=self.e.get_title_list(start_row,end_row)#title_list
        print(list,flush=True)
        self.c.crawl(list)
        time.sleep(2)
        start.delete(0,tk.END)
        end.delete(0,tk.END)
        time.sleep(1)
        start.insert(0,end_row+1)
        end.insert(0,end_row+4)
        num=end_row-start_row+1
        var.set('请输入'+str(num)+'个结果 ')
        #num_list=c.insert() 
        #self.e.write_num(num_list)

    def insert(self):
        num=e.get()
        num_list=[int(i) for i in re.split('[,，]',num)]
        print(num_list,flush=True)
        self.e.write_num(num_list)
        e.delete(0,tk.END)
        var.set('数据已导入 ')

    def tearDown(self):
       self.c.tearDown()

if __name__ == '__main__':
    w=Work()
    window = tk.Tk()
    window.title('my window')
    window.geometry('400x300')



    #开始启动
    b1 = tk.Button(window, 
        text=r'启动',      # 显示在按钮上的文字
        width=15, height=2, 
        command=w.setUp)
        #command=lambda:thread_it(hit_me))     # 点击按钮式执行的命令
    b1.pack()    # 按钮位置

    start = tk.Entry(window)
    start.pack()

    end = tk.Entry(window)
    end.pack()


    #开始爬取
    b2 = tk.Button(window, 
        text=r'开始爬取',      # 显示在按钮上的文字
        width=15, height=2, 
        command=w.crawl)
        #command=lambda:thread_it(hit_me))     # 点击按钮式执行的命令
    b2.pack()    # 按钮位置

    #获取数字并输入excel
    b2 = tk.Button(window, 
        text=r'数据输入',      # 显示在按钮上的文字
        width=15, height=2, 
        command=w.insert)
        #command=lambda:thread_it(hit_me))     # 点击按钮式执行的命令
    b2.pack()    # 按钮位置

    var = tk.StringVar()    # 这时文字变量储存器

    l = tk.Label(window, 
        textvariable=var,   # 使用 textvariable 替换 text, 因为这个可以变化
        bg='white', font=('Arial', 12), width=15, height=2,fg='blue')
    l.pack() 


    #输入数字
    e = tk.Entry(window)
    e.pack()

    #结束
    b2 = tk.Button(window, 
        text=r'结束',      # 显示在按钮上的文字
        width=15, height=2, 
        command=w.tearDown)
        #command=lambda:thread_it(hit_me))     # 点击按钮式执行的命令
    b2.pack()    # 按钮位置



    window.mainloop()