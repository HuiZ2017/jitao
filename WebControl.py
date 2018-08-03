#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
 @desc:
 @author: ZhangHui
 @contact: hui013@live.com
 @file: WebControl.py
 @time: 2018/6/10 18:58
 """

import requests
import json
from getUserAgent import getUserAgent
import time
from bs4 import BeautifulSoup as bs
from tkinter import messagebox as msg
import os,csv
class WebControl():
    def __init__(self,username,passwd):
        self.uname = username
        self.passwd = passwd
        self.headers = getUserAgent()
        self.MainSession = requests.Session()
        self.URL = {'login':"http://bm.jishutao.com/managesignin" ,
                    'csearch':"http://bm.jishutao.com/api/admin/customer/listAllOrganizationCustomer",
                    'addC':'http://bm.jishutao.com/api/admin/customer/addCustomer',
                    'listAllPriCu':'http://bm.jishutao.com/api/admin/customer/listPrivateOrganizationCustomer',
                    'addF':'http://bm.jishutao.com/api/admin/customer/addFollow',
                    'findOCBid':'http://bm.jishutao.com/api/admin/customer/findCustomerContacts?uid={}'}
        self.login()
        self.fileObj = None
    def login(self):
        #判断帐号密码是否为空,然后判断登陆状态
        if self.uname and self.passwd:
            LoginParams = {
                'mobile': self.uname,
                'password': self.passwd
            }
            self.loginPage = self.MainSession.post(url=self.URL['login'],data=LoginParams)
            if "404" in str(self.loginPage):
                self.loginState = False
            else:
                self.loginState = True
        pass
    def getUserInfo(self):
        if self.loginState:
            try:
                userInfo = self.loginPage.text.replace('null', '\"null\"')
                j1 = json.loads(userInfo)
                mobile = j1['data']['mobile']
                name = j1['data']['name']
            except Exception as e:
                mobile = 'Can not be found'
                name = 'Can not be found'
            return (name,mobile)

    def oneSearch(self,customer,pageNo=1):
        while True:
            result = self.MainSession.post(url=self.URL['csearch'],data={'name':customer,'pageNo':pageNo,'pageSize':100})
            if result:
                context = []
                result = json.loads(result.text)
                endpg = result['data']['totalPage']
                nextpg = result['data']['nextPage']
                alllist = result['data']['list']
                for sub in alllist:
                    context.append((sub['name'],sub['adminName'],sub['createTime']))
                pageNo += 1
                yield context
                if pageNo > endpg:
                    return None
        else:
            raise Exception
    def getTopSearch(self,name):
        result = self.MainSession.post(url=self.URL['csearch'],
                                       data={'name': name, 'pageNo': 1, 'pageSize': 1})
        if result:
            result = json.loads(result.text)
            try:
                alllist = result['data']['list'][0]
                return (alllist['name'],alllist['adminName'],alllist['createTime'])
            except Exception as e:
                return 'getTopSearch没找到'
        else:
            return 'getTopSearch没找到'
    def SingleCustomerInput(self,info):
        name = info['name']
        contact = info['contact']
        cMobile = info['cMobile']
        societyTag = info['societyTag']
        type = info['type']
        getResult = self.MainSession.post(url = self.URL['addC'],
                data={'name': name,
                      'contacts': contact,
                      'contactMobile': cMobile,
                      'societyTag':societyTag,
                      'type':type},headers=self.headers)
        try:
            if '客户已存在' in str(getResult.text):
                return (self.getTopSearch(name),'导入失败')
            else:
                return ((name,'',''),'导入成功')
        except Exception as ee:
            pass
    def listPrivateOrganizationCustomer(self,fileX,button):
        Url = self.URL['listAllPriCu']
        totalPage = 0
        Data = {
                'pageNo': 1,
                'pageSize': 100,
                'city': ''}
        filename = fileX
        with open(filename, 'w',newline='',encoding='UTF-8') as self.fileObj:
            getResult = self.MainSession.post(url=Url, data=Data)
            if getResult:
                result = json.loads(getResult.text)
                totalPage = result['data']['totalPage']
                fieldName = list(result['data']['list'][0].keys())
                writer = csv.DictWriter(self.fileObj,fieldnames=fieldName)
                writer.writeheader()
                for lines in result['data']['list']:
                    try:
                        writer.writerow(lines)
                    except Exception as e:
                        print(lines)
                Data['pageNo'] += 1
                while Data['pageNo'] <= totalPage:
                    print('#%s' % Data['pageNo'], end='')
                    getResult = self.MainSession.post(url=Url, data=Data)
                    if getResult:
                        result = json.loads(getResult.text)
                        for line in result['data']['list']:
                            try:
                                writer.writerow(line)
                                print('#', end='')
                            except Exception as e:
                                print(line)
                    Data['pageNo'] += 1
                    print()
            else:
                pass
        self.setButton(button)
    def setButton(self, button):
        button.config(state='normal')
        msg.showinfo('提示', '导出完成!')
        #（('辽宁得利电子汽车衡器有限公司', '王贵强', '2018-04-13 14:18:37'),'导入失败')
        #（('辽宁得利电子汽车衡器有限公司', '', ''),'导入成功')
    def Follow(self,uid):
        ocbid = self.findOCBid(uid)
        if ocbid:
            if self.addFollow(uid,ocbid):
                return True
        return False
    def addFollow(self,uid,ocbid):
        Url = self.URL['addF']
        Data = {
            'userBusinessList': [],
            'uid': uid,
            'ocbId': ocbid,
            'contactType': 3,
            'result':'微信沟通',
            'followTime':time.strftime("%Y-%m-%d %H:%M:%S:0", time.localtime())
        }
        getResult = self.MainSession.post(url=Url, data=Data)
        if getResult:
            result = json.loads(getResult.text)
            if not result['error']:
                return True
            else:
                return False
    def findOCBid(self,uid):
        Url = self.URL['findOCBid'].format(uid)
        getResult = self.MainSession.get(url=Url)
        if getResult:
            result = json.loads(getResult.text)
            if result['data'][0]['id']:
                return result['data'][0]['id']
            else:
                return False
        #time.strftime("%Y-%m-%d", time.localtime())
#headers = getUserAgent()
#MainSession = requests.Session()
#loginPage = MainSession.post(url=LoginURL,data=LoginParams)
#print(loginPage.cookies)

#OneSearch = MainSession.post(url=listCustomer,
#                             data=CustomerParams)
# #print(OneSearch.text)
#
# if __name__ == "__main__":
#     demo = WebControl(username='17673102113',passwd='123456')
#     print(demo.getUserInfo())
#     for i in demo.oneSearch('湖南长沙'):
#         print(i)

