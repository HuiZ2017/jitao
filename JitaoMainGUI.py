#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Y:\PycharmProjects\py3\myTest\python3\jitao\jitao\jitao_v0.3>pyinstaller -F -w -
p . -i favicon-20180415090500983.ico JitaoMainGUI.py
 @desc:
 @author: ZhangHui
 @contact: hui013@live.com
 @file: JitaoMainGUI.py
 @time: 2018/7/7 22:42
 """
import tkinter as tk
from tkinter import messagebox as msg
from tkinter import ttk
from tkinter.ttk import Notebook
import tkinter.filedialog as tkFileDialog
import xlrd,time,threading
from time import sleep
class Demo(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("技淘网后台客户管理tool v0.8 \t by 张珲")
        self.geometry("1200x600")
        self.notebook = Notebook(self,height = 400,width = 1000)
        self.searchResult = []
        self.userinfo = ['null','null']
        self.searchCard = tk.Frame(self.notebook)
        self.notebook.add(self.searchCard, text="客户信息查询")
        self.inputCard = tk.Frame(self.notebook)
        self.notebook.add(self.inputCard, text="客户信息录入")
        self.notebook.pack(side=tk.LEFT,fill=tk.Y)
        self.followCard = tk.Frame(self.notebook)
        self.notebook.add(self.followCard, text="客户跟进")
        self.notebook.pack(side=tk.LEFT, fill=tk.Y)

        self.frameUserInfo = tk.Frame(self,height = 20,width = 100)
        self.frameUserInfo.pack(side=tk.RIGHT,expand=tk.NO,fill=tk.Y)
        self.frameButton = tk.Frame(self.searchCard,height = 20,width = 100)
        self.frameButton.pack(anchor=tk.S,expand=tk.NO,fill=tk.NONE)
        ##########录入系统############
        self.frameSingleInput = tk.Frame(self.inputCard,height = 100,width = 400)
        self.frameSingleInput.pack(side=tk.TOP,anchor=tk.W)
        self.frameMultiInput = tk.Frame(self.inputCard,height = 100,width = 400,bg='blue')
        self.frameMultiInput.pack(side=tk.TOP,anchor=tk.E)
        self.frameInputOut = tk.Frame(self.inputCard,height = 500,width = 800,bg='white')
        self.frameInputOut.pack(side=tk.BOTTOM)
        tk.Button(self.frameMultiInput, text="Open", fg="black", bg="white", command=self.askopenfile).pack()
        self.clear2Botton = tk.Button(self.frameMultiInput, text="Clear", fg="green", bg="white", command=self.Clear2)
        self.clear2Botton.pack()
        self.labelCname = tk.Label(self.frameSingleInput,text='   客户企业名称：',width = 15,height = 1)
        self.labelCcontact = tk.Label(self.frameSingleInput,text='客户联系人姓名：',width = 15,height = 1)
        self.labelCMobile = tk.Label(self.frameSingleInput,text='  客户联系方式：',width = 15,height = 1)
        self.labelCtype = tk.Label(self.frameSingleInput,text='        客户类型：',width = 15,height = 1)
        self.labelSource = tk.Label(self.frameSingleInput,text='     客户信息来源：',width = 15,height = 1)
        self.buttonInput = tk.Button(self.frameSingleInput,text='单个导入',fg='blue',bg='white',command=self.singleinput)
        self.societyTag = tk.StringVar()
        self.sourceTag = tk.StringVar()
        self.societynumberChosen = ttk.Combobox(self.frameSingleInput, width=10, textvariable=self.societyTag)
        self.sourcenumberChosen = ttk.Combobox(self.frameSingleInput, width=10, textvariable=self.sourceTag)
        self.societynumberChosen['values'] = ('社会企业',
                                       '科研单位',
                                       '高等院校',
                                       '科研专家',
                                       '科技达人',
                                       '服务机构',
                                       '政府部门',
                                       '商会',
                                       '协会',
                                       '其他组织',
                                       '其他个人')     # 设置下拉列表的值
        self.sourcenumberChosen['values'] = ('搜索引擎',
                                        '广告',
                                        '社交推广',
                                        '研讨会',
                                        '客户介绍',
                                        '独立开发',
                                        '代理商',
                                        '其他')
        self.file_opt = {}
        self.file_opt['multiple'] = None
        self.file_opt['initialdir'] = '.'
        self.INIT_input()
        self.societynumberChosen.current(0)
        self.sourcenumberChosen.current(0)
        self.INIT_outbox_multi()
        ###跟进系统##
        self.followStep = 0
        tk.Label(self.followCard, width=25, height=1, text='需跟进uid粘贴至此').pack()
        self.followList = tk.Text(self.followCard,width=50,height=30)
        self.followList.pack()
        self.addFollowBotton = tk.Button(self.followCard, text="Start", fg="black", bg="white", command=self.runaddFollows)
        self.addFollowBotton.pack()
        self.labelfollowStep = tk.Label(self.followCard, width=25, height=1, text="已完成： %s" %self.followStep,
                                   justify='left')
        self.labelfollowStep.pack(anchor=tk.SE)
        #####
        self.tempVar = ''
        self.pbar = tk.Label(self.frameUserInfo,width = 25,height = 1, text=self.tempVar,justify = 'left')
        self.pbar.pack()
        ##############################e
        #置顶按钮
        self.chVarDis = tk.IntVar()
        self.topButton = tk.Checkbutton(self.frameUserInfo, text="置顶本窗口",variable=self.chVarDis,command=self.setTop)
        self.topButton.pack(anchor=tk.SE)
        #用户输入检索内容
        self.searchList = tk.Entry(self.frameButton, width=25, font=('Courier New', 15))
        self.searchList.pack(anchor=tk.SE,fill=tk.X)
        # 搜索按钮
        self.searchBotton = tk.Button(self.frameButton, text="Search", fg="blue", bg="white", command=self.Search)
        self.searchBotton.pack()
        #登录信息展示
        self.labelUname = tk.Label(self.frameUserInfo, width = 25,height = 1,text="姓名： %s" %self.userinfo[0],justify = 'left')
        self.labelID = tk.Label(self.frameUserInfo,width = 25,height = 1, text="账号： %s" %self.userinfo[1],justify = 'left')
        self.buttonoutput = tk.Button(self.frameUserInfo, text='导出私有用户', fg='blue', bg='white', command=self.outputall)
        self.buttonoutput.pack()
        self.labelUname.pack()
        self.labelID.pack()
        # 清空按钮
        self.clearBotton = tk.Button(self.frameButton, text="Clear", fg="green", bg="white", command=self.Clear)
        self.clearBotton.pack()
        #初始化下方输出框
        self.INIT_outbox()
        self.INIT_menu()
    def askopenfile(self):
        SucessFlag = 0
        self.file_opt['initialfile'] = '*.xls'
        self.file_opt['filetypes'] = [('xls', 'xls')]
        filename = tkFileDialog.askopenfilename(**self.file_opt)
        with xlrd.open_workbook(filename) as fileObj:
            worksheet1 = fileObj.sheet_by_index(0)
            num_rows = worksheet1.nrows
            num_cols = worksheet1.ncols
            flag = 0
            all = []
            for rown in range(num_rows):
                hang = []
                if worksheet1.cell_value(rown, 1):
                    for coln in range(num_cols):
                        if worksheet1.cell_type(rown, coln) == 2:
                            cell = int(worksheet1.cell_value(rown, coln))
                        else:
                            cell = worksheet1.cell_value(rown, coln)
                        hang.append(cell)
                else:
                    flag = rown
                    break
                all.append(hang)
            OkFlag = msg.askokcancel('加载成功','共加载 %s 条客户信息，点击确认开始导入' %flag)
            if OkFlag:
                for lines in all:
                    if self.singleinput(lines):
                        SucessFlag += 1
                msg.showinfo('导入完成','共导入 %s 条客户信息，成功率 %s' %(SucessFlag,SucessFlag/flag))
    def runaddFollows(self):
        t1 = threading.Thread(target=self.addFollows,args=())
        self.addFollowBotton.config(state='disabled')
        for t in [t1]:
            t.setDaemon(True)
            t.start()
    def addFollows(self):
        flag = 0
        self.followStep = 0
        followlist = self.followList.get(0.0,tk.END)
        for uid in followlist.split('\n'):
            if uid:
                if self.demo.Follow(uid):
                    self.labelfollowStep.config(text="已完成： %s" %self.followStep)
                    self.followStep += 1
                flag += 1
        msg.showinfo('跟进情况','共跟进信息%s\n跟进成功%s' %(flag,self.followStep))
            #self.demo.Follow(uid)
    def setTop(self):
        if self.chVarDis.get():
            self.wm_attributes('-topmost', 1)
        else:
            self.wm_attributes('-topmost', 0)
            # self.wm_attributes()
    def INIT_input(self):
        self.Cname = tk.Entry(self.frameSingleInput, width=20, font=('Arial', 15))
        self.Ccontacts = tk.Entry(self.frameSingleInput, width=20, font=('Arial', 15))
        self.CMobile = tk.Entry(self.frameSingleInput, width=20, font=('Arial', 15))
        self.labelCname.grid(row=1,column=1)
        self.Cname.grid(row=1,column=2)
        self.labelCcontact.grid(row=2,column=1)
        self.Ccontacts.grid(row=2,column=2)
        self.labelCMobile.grid(row=3,column=1)
        self.CMobile.grid(row=3,column=2)
        self.labelCtype.grid(row=4,column=1)
        self.labelSource.grid(row=5, column=1)
        self.societynumberChosen.grid(row=4,column=2)
        self.sourcenumberChosen.grid(row=5, column=2)
        self.buttonInput.grid(row=6,column=2)
    def singleinput(self,setinfo=[]):
        if setinfo:
            info = {}
            info['name'] = setinfo[0]
            info['contact'] = setinfo[1]
            info['cMobile'] = setinfo[2]
            info['societyTag'] = setinfo[3]
            info['sourceTag'] = setinfo[4]
            info['type'] = 1
        else:
            info = {}
            info['name'] = self.Cname.get()
            info['contact'] = self.Ccontacts.get()
            info['cMobile'] = self.CMobile.get()
            info['societyTag'] = self.societyTag.get()
            info['sourceTag'] =self.sourceTag.get()
            info['type'] = 1
        if info['name'] and info['contact'] and info['cMobile'] and info['societyTag'] and info['sourceTag'] and info['type']:
            info['societyTag'] = self.societynumberChosen['values'].index(str(info['societyTag']))
            info['sourceTag'] = self.sourcenumberChosen['values'].index(str(info['sourceTag']))
            result = self.demo.SingleCustomerInput(info)
            self.outbox2_insert(result)
            if '成功' in str(result[1]):
                return True
        else:
            msg.showwarning('警告',"%s 信息不完善!" %info['name'])

    def INIT_menu(self):
        self.menu = tk.Menu()
        self.menu.add_command(label='Login', command=self.Login)
        self['menu'] = self.menu
    def Login(self):
        #登陆方法,弹窗帐号密码
        def Submit():
            uname = new.ent1.get()
            passwd = new.ent2.get()
            from WebControl import WebControl
            self.demo = WebControl(username=uname, passwd=passwd)
            if self.demo.loginState:
                new.destroy()
                userinfo = self.demo.getUserInfo()
                msg.showinfo("提示","登陆成功 %s" %userinfo[0])
                self.userinfo = userinfo
                self.logininfoshow()
            else:
                msg.showerror("警告","账号 %s\n密码 %s\n登录失败" %(uname,passwd))
        new = tk.Toplevel(self)
        new.title('输入帐号密码')
        new.geometry("180x100")
        new.resizable(width=False, height=False)
        new.lab1 = tk.Label(new, text="账户:")
        new.lab1.grid(row=0, column=0, sticky=tk.W)
        new.ent1 = tk.Entry(new)
        new.ent1.grid(row=0, column=1, sticky=tk.W)
        new.lab2 = tk.Label(new, text="密码:")
        new.lab2.grid(row=1, column=0)
        new.ent2 = tk.Entry(new, show="*")
        new.ent2.grid(row=1, column=1, sticky=tk.W)
        new.button = tk.Button(new, text="登录", command=Submit)
        new.button.grid(row=2, column=1, sticky=tk.E)
        new.lab3 = tk.Label(new, text="")
        new.lab3.grid(row=3, column=0, sticky=tk.W)
        new.mainloop()
    def logininfoshow(self):
        self.labelUname.config(text="姓名： %s" %self.userinfo[0])
        self.labelID.config(text="账号： %s" %self.userinfo[1])

    def outputall(self):
        self.file_opt['initialfile'] = '%s导出.csv' % time.strftime("%Y-%m-%d", time.localtime())
        self.file_opt['filetypes'] = [('csv', '.csv'), ('xls', 'xls')]
        saveasname = tkFileDialog.asksaveasfile(mode='w', **self.file_opt)
        # t1 = threading.Thread(target=self.demo.listPrivateOrganizationCustomer, args=(saveasname.name,))
        t1 = threading.Thread(target=self.demo.listPrivateOrganizationCustomer, args=(saveasname.name,self.buttonoutput,))
        # t2 = threading.Thread(target=self.setButton, args=(self.buttonoutput,))
        self.buttonoutput.config(state='disabled')
        for t in [t1]:
            t.setDaemon(True)
            t.start()
        # t2.start()

    def Search(self):
        #搜索开始按钮
        searchlist = self.searchList.get()

        for sub in str(searchlist).split():
#            [('辽宁得利电子汽车衡器有限公司', '王贵强', '2018-04-13 14:18:37')]
            try:
                for res in self.demo.oneSearch(sub):
                    for subres in res:
                        self.Put_box(subres)
            except Exception as e:
                msg.showinfo("提示！", "目前该客户 %s 未录入" %sub)
    def Clear(self):
        items = self.outbox.get_children()
        [self.outbox.delete(item) for item in items]
        self.outboxFlag = 0
        msg.showinfo("提示！","已清空检索列表")
    def Clear2(self):
        items = self.outbox2.get_children()
        [self.outbox2.delete(item) for item in items]
        self.outbox2Flag = 0
        msg.showinfo("提示！","已清空检索列表")
    def Put_box(self,kw):
        self.outbox_insert(kw)
    def INIT_outbox(self):
        self.outboxFlag = 0
        self.outbox = ttk.Treeview(self.searchCard)  # 表格
        self.outbox = ttk.Treeview(self.searchCard, show="headings", height=500, columns=("a", "b", "c", "d"))
        self.outbox.column("a", width=50, anchor="center")
        self.outbox.column("b", width=500, anchor="center")
        self.outbox.column("c", width=100, anchor="center")
        self.outbox.column("d", width=300, anchor="center")
        self.outbox.heading("a", text="编号")
        self.outbox.heading("b", text="客户企业名称")
        self.outbox.heading("c", text="录入同事")
        self.outbox.heading("d", text="录入时间")
        self.outbox.pack(side=tk.BOTTOM, anchor=tk.S, fill=tk.Y)
        # self.vbar = ttk.Scrollbar(self.outbox, orient=tk.VERTICAL, command=self.outbox.yview)
        # self.outbox.configure(yscrollcommand=self.vbar.set)
        self.outbox.pack(side=tk.BOTTOM, anchor=tk.S, fill=tk.Y)
        # self.vbar.pack(side=tk.BOTTOM,anchor=tk.S)
    def INIT_outbox_multi(self):
        self.outbox2Flag = 0
        self.outbox2 = ttk.Treeview(self.frameInputOut)  # 表格
        self.outbox2 = ttk.Treeview(self.frameInputOut, show="headings",height=50, columns=("a", "b", "c", "d","e"))
        self.outbox2.column("a", width=50, anchor="center")
        self.outbox2.column("b", width=400, anchor="center")
        self.outbox2.column("c", width=200, anchor="center")
        self.outbox2.column("d", width=200, anchor="center")
        self.outbox2.column("e", width=100, anchor="center")
        self.outbox2.heading("a", text="编号")
        self.outbox2.heading("b", text="客户企业名称")
        self.outbox2.heading("c", text="录入同事")
        self.outbox2.heading("d", text="录入时间")
        self.outbox2.heading("e", text="导入状态")
        self.outbox2.pack(side=tk.BOTTOM, anchor=tk.SW, fill=tk.X)
    def outbox2_insert(self,args):
        flag = self.outbox2Flag
        self.outbox2Flag += 1
        self.outbox2.insert('', 0, text="123", values=(flag, args[0][0], args[0][1], args[0][2],args[1]))
    def outbox_insert(self,args):
        flag = self.outboxFlag
        self.outboxFlag += 1
        self.outbox.insert('',0,text="123",values=(flag,args[0],args[1],args[2]))
    def run(self):
        t = self.outbox.getvar('a')
if __name__ == "__main__":
    demo1 = Demo()
    demo1.mainloop()
    # print(demo1.run())