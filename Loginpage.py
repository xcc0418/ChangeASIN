from tkinter.messagebox import *
from MainPage import *
import requests
import json
import os, os.path
import global_var


class LoginPage(object):
    def __init__(self, master=None):
        self.root = master  # 定义内部变量root
        self.root.geometry('%dx%d+%d+%d' % (350, 220, (960-380), (540-450)))  # 设置窗口大小
        self.username = StringVar()
        self.password = StringVar()
        self.checkbox = IntVar()
        self.checkbox.set(1)
        self.ft = tkFont.Font(family='microsoft yahei', size=11)
        self.createPage()

    def createPage(self):
        self.page = Frame(self.root)  # 创建Frame
        self.page.pack()
        Label(self.page, font=self.ft).grid(row=0, stick=W)
        Label(self.page, text='账户: ', font=self.ft).grid(row=1, stick=W, pady=8)
        Entry(self.page, textvariable=self.username, font=self.ft).grid(row=1, column=1, stick=E)
        Label(self.page, text='密码: ', font=self.ft).grid(row=2, stick=W, pady=8)
        self.e1 = Entry(self.page, textvariable=self.password, show='*', font=self.ft)
        self.e1.bind("<Return>", self.loginCheck)
        self.e1.grid(row=2, column=1, stick=E)
        Checkbutton(self.page, text="记住密码", font=self.ft, variable=self.checkbox).grid(row=3, column=0, pady=8, stick=E)
        Button(self.page, text='登陆', width=10, command=self.loginCheck, font=self.ft).grid(row=4, stick=W, pady=8)
        Button(self.page, text='退出', width=10, command=self.page.quit, font=self.ft).grid(row=4, column=1, stick=E)
        self.read_namepwd()

    def loginCheck(self, event=None):
        name = self.username.get()
        secret = self.password.get()
        self.sql_pwd(name, secret)
        # msg = self.driver_login(name, secret)
        msg = self.login_asinking(name, secret)
        if msg == '操作成功' or msg == '登录成功':
            self.page.destroy()
            MainPage(self.root)
            self.login_asinking3("IT-Test", "IT-Test")
            with open(self.filename, 'w') as fp:
                fp.write(','.join((name, secret)))
        else:
            result = self.login_secret_key(name, secret)
            if result:
                self.page.destroy()
                MainPage(self.root)
                self.login_asinking3("IT-Test", "IT-Test")
                with open(self.filename, 'w') as fp:
                    fp.write(','.join((name, secret)))
            else:
                showinfo(title='错误', message='%s' % msg)

    def read_namepwd(self):
        self.path = os.getenv('temp')
        self.filename = os.path.join(self.path, 'asingkingpwd.txt')
        # 尝试自动填写用户名和密码
        try:
            with open(self.filename) as fp:
                n, p = fp.read().strip().split(',')
                self.username.set(n)
                self.password.set(p)
        except:
            pass

    def login_asinking(self, username, password):
        # 生成Session对象，用于保存Cookie
        global_var.s = requests.Session()
        # 登录url
        login_url = 'https://erp.lingxing.com/api/passport/login'
        # 请求头
        headers = {'Host': 'erp.lingxing.com',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:95.0) Gecko/20100101 Firefox/95.0'
                   , 'Referer': 'https://erp.lingxing.com/login',
                   'Accept': 'application/json, text/plain, */*',
                   'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
                   'Accept-Encoding': 'gzip, deflate, br',
                   'Content-Type': 'application/json;charset=utf-8',
                   'X-AK-Request-Id': 'e7f7b81a-fafd-4031-8964-00376ae24d07',
                   'X-AK-Company-Id': '90136229927150080',
                   'X-AK-Request-Source': 'erp',
                   'X-AK-ENV-KEY': 'SAAS-10',
                   'X-AK-Version': '1.0.0.0.0.023',
                   'X-AK-Zid': '109810',
                   'Content-Length': '114',
                   'Origin': 'https://erp.lingxing.com',
                   'Connection': 'keep-alive'}
        # 传递用户名和密码
        data = {'account': username, 'pwd': password}
        data = json.dumps(data)
        try:
            r = global_var.s.post(login_url, headers=headers, data=data)
            r.raise_for_status()
            r1 = r.text
            r2 = json.loads(r1)
            r3 = r2['msg']
        except:
            r3 = '网络错误失败'

        return r3

    def login_asinking2(self, username, password):
        # 生成Session对象，用于保存Cookie
        global_var.s2 = requests.Session()
        # 登录url
        login_url = 'https://gw.lingxingerp.com/newadmin/api/passport/login'
        # 请求头
        headers = {'Host': 'gw.lingxingerp.com',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:95.0) Gecko/20100101 Firefox/95.0'
            , 'Referer': 'https://erp.lingxing.com/',
                   'Accept': 'application/json, text/plain, */*',
                   'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
                   'Accept-Encoding': 'gzip, deflate, br',
                   'Content-Type': 'application/json;charset=utf-8',
                   'X-AK-Request-Id': '03cabda1-739d-4cd7-bd94-6650f97775a8',
                   'X-AK-Company-Id': '90136229927150080',
                   'X-AK-Request-Source': 'erp',
                   'X-AK-Version': '1.0.0.0.0.186',
                   'X-AK-Zid': '',
                   'Content-Length': '219',
                   'Origin': 'https://erp.lingxing.com',
                   'Connection': 'keep-alive'}
        # 传递用户名和密码
        self.sql()
        sql = "select * from `flag`.`password` where `账号` = '%s' and `密码` = '%s'" % (username, password)
        self.cursor.execute(sql)
        result = self.cursor.fetchall()
        self.sql_close()
        if result:
            pwd = result[0]['密钥']
            data = {'account': username, 'pwd': pwd}
            data = json.dumps(data)
            print(data)
            # try:
            r = global_var.s2.post(login_url, headers=headers, data=data)
            print(global_var.s2.cookies.get('auth-token'))
            r.raise_for_status()
            r1 = r.text
            r2 = json.loads(r1)
            r3 = r2['msg']
            print(r2)
            return r3
        else:
            return False

    def login_asinking3(self, username, password):
        # 生成Session对象，用于保存Cookie
        global_var.s2 = requests.Session()
        # 登录url
        login_url = 'https://erp.lingxing.com/api/passport/login'
        # 请求头
        headers = {'Host': 'erp.lingxing.com',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:95.0) Gecko/20100101 Firefox/95.0'
                   , 'Referer': 'https://erp.lingxing.com/login',
                   'Accept': 'application/json, text/plain, */*',
                   'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
                   'Accept-Encoding': 'gzip, deflate, br',
                   'Content-Type': 'application/json;charset=utf-8',
                   'X-AK-Request-Id': 'e7f7b81a-fafd-4031-8964-00376ae24d07',
                   'X-AK-Company-Id': '90136229927150080',
                   'X-AK-Request-Source': 'erp',
                   'X-AK-ENV-KEY': 'SAAS-10',
                   'X-AK-Version': '1.0.0.0.0.023',
                   'X-AK-Zid': '109810',
                   'Content-Length': '114',
                   'Origin': 'https://erp.lingxing.com',
                   'Connection': 'keep-alive'}
        # 传递用户名和密码
        data = {'account': username, 'pwd': password}
        data = json.dumps(data)
        try:
            r = global_var.s2.post(login_url, headers=headers, data=data)
            print(global_var.s2.cookies.get('auth-token'))
            r.raise_for_status()
            r1 = r.text
            r2 = json.loads(r1)
            r3 = r2['msg']
        except:
            r3 = '网络错误失败'

        return r3

    def login_secret_key(self, username, password):
        # 生成Session对象，用于保存Cookie
        global_var.s = requests.Session()
        # 登录url
        login_url = 'https://gw.lingxingerp.com/newadmin/api/passport/login'
        # 请求头
        headers = {'Host': 'gw.lingxingerp.com',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:95.0) Gecko/20100101 Firefox/95.0'
                    , 'Referer': 'https://erp.lingxing.com/',
                   'Accept': 'application/json, text/plain, */*',
                   'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
                   'Accept-Encoding': 'gzip, deflate, br',
                   'Content-Type': 'application/json;charset=utf-8',
                   'X-AK-Request-Id': '03cabda1-739d-4cd7-bd94-6650f97775a8',
                   'X-AK-Company-Id': '90136229927150080',
                   'X-AK-Request-Source': 'erp',
                   'X-AK-Version': '1.0.0.0.0.186',
                   'X-AK-Zid': '',
                   'Content-Length': '219',
                   'Origin': 'https://erp.lingxing.com',
                   'Connection': 'keep-alive'}
        # 传递用户名和密码
        self.sql()
        sql = "select * from `flag`.`password` where `账号` = '%s' and `密码` = '%s'" % (username, password)
        self.cursor.execute(sql)
        result = self.cursor.fetchall()
        self.sql_close()
        if result:
            pwd = result[0]['密钥']
            data = {'account': username, 'pwd': pwd}
            data = json.dumps(data)
            print(data)
            # try:
            r = global_var.s.post(login_url, headers=headers, data=data)
            r.raise_for_status()
            r1 = r.text
            r2 = json.loads(r1)
            r3 = r2['msg']
            print(r2)
            return r3
        else:
            return False

    def sql(self):
        self.connection = pymysql.connect(host='3354n8l084.goho.co',  # 数据库地址
                                          port=24824,
                                          user='test_user',  # 数据库用户名
                                          password='a123456',  # 数据库密码
                                          db='storage',  # 数据库名称
                                          charset='utf8',
                                          cursorclass=pymysql.cursors.DictCursor)
        self.cursor = self.connection.cursor()

    def sql_close(self):
        self.cursor.close()
        self.connection.close()

    def sql_pwd(self, name, password):
        self.sql()
        sql = "select * from `flag`.`password` where `账号` = '%s' and `密码` = '%s'" % (name, password)
        self.cursor.execute(sql)
        result = self.cursor.fetchall()
        if result:
            self.sql_close()
            return
        else:
            sql1 = "insert into `flag`.`password`(`账号`, `密码`)values('%s', '%s')" % (name, password)
            self.cursor.execute(sql1)
            self.connection.commit()
            self.sql_close()