'''
软件测试小工具，实现了mysql与excel表格间的数据转换，
杨可亮 2020/9/23
使用方法：

import sqlex
host = '118.24.105.78'
user = 'root'
password = "1qaz!QAZ123***123"
db_name = "ljtestdb"
sheet_name = "t_user"
lines = 103
case = sqlex.Sqlandex(host,user,password,db_name)
case.s_to_e(sheet_name,lines)

'''
import xlwt
import xlrd
import pymysql as psql

class Sqlandex:
    def __init__(self,host,user,password,db_name):#初始化读入数据库基本的信息，包括ip地址，用户名，密码，数据库名
        self.host = host
        self.user = user
        self.password = password
        self.db_name = db_name

    def query(self,sql):#查询函数，输入查询语句
        db = psql.connect(host=self.host, user=self.user, password=self.password, db=self.db_name)#pymysqlconnect方法
        cur = db.cursor()
        cur.execute(sql)
        res = cur.fetchall()
        return res

    def get_titlelist(self,sql_title):#从数据库中获取表头，输入查询表头的sql语句
        list_title = []
        title = Sqlandex(self.host,self.user,self.password,self.db_name)
        for i in title.query(sql_title):
            list_title.append(i[0])
        return list_title

    def get_contentlist(self,sql_content):#从数据库中获取内容，输入查询内容的sql语句
        list_content = []
        content = Sqlandex(self.host,self.user,self.password,self.db_name)
        for i in content.query(sql_content):
            list_each = []
            for j in i:   
                list_each.append(j)
            list_content.append(list_each)
        return list_content

    def write_to_excel(self,list_all,sheet_name):#将获取的数据写入excel表格
        wb = xlwt.Workbook()
        ws = wb.add_sheet(sheet_name)
        for i in range(len(list_all)):
            for j in range(len(list_all[i])):
                ws.write(i,j,list_all[i][j])
        wb.save("./sqlexcel/{}.xls".format(sheet_name))

    def s_to_e(self,sheet_name,lines):#数据库数据导入表格
        sql_title = "SELECT COLUMN_NAME FROM \
            information_schema.COLUMNS\
                WHERE TABLE_SCHEMA = '{}' \
                    AND TABLE_NAME = '{}';".format(self.db_name,sheet_name)
        sql_content = "select * from {} limit {};".format(sheet_name,lines)
        list_all = []#拼接title与content
        title = Sqlandex(self.host,self.user,self.password,self.db_name)
        content = Sqlandex(self.host,self.user,self.password,self.db_name)
        list_all.append(title.get_titlelist(sql_title))
        list_all.extend(content.get_contentlist(sql_content))
        title.write_to_excel(list_all,sheet_name)

    def change(self,sqlchange):#写入数据库
        db = psql.connect(host=self.host, user=self.user, password=self.password, db=self.db_name)#查询
        cur = db.cursor()
        cur.execute(sqlchange)
        db.commit()
        
    def e_to_s(self,wb_name,sheet_name):#表格写入数据库
        wb = xlrd.open_workbook(wb_name)
        ws = wb.sheet_by_name(sheet_name)
        title_list = ws.row_values(0)
        title_sql = "CREATE TABLE {} (".format(sheet_name)
        for i in title_list:
            i = i + " varchar(255),"
            title_sql = title_sql + i
        title_sql = title_sql[:len(title_sql)-1]+")"#表头获取成功    
        con = Sqlandex(self.host,self.user,self.password,self.db_name)
        con.change(title_sql)#表头写入成功
        content_list=[]#二维数组
        for i in range(1,ws.nrows):
            content_list.append(ws.row_values(i))
        content_sql = "INSERT INTO {} (".format(sheet_name)
        for i in title_list:
            content_sql = content_sql + i +","
        content_sql = content_sql[:len(content_sql)-1]+") values ("
        content_sql_save = content_sql
        content_sql_list = []
        for i in content_list:
            for j in i:
                content_sql = content_sql+'"'+str(j)+'"'+","
            content_sql_list.append(content_sql[:len(content_sql)-1]+");")
            content_sql = content_sql_save#获取每行的sql语句成功
        for i in content_sql_list:
            con.change(i)#写入文件
   
if __name__ == "__main__":
    host = '118.24.105.78'
    user = 'root'
    password = "1qaz!QAZ123***123"
    db_name = "ljtestdb"
    con = Sqlandex(host,user,password,db_name)
    con.s_to_e("t_user",30)#测试数据库导出功能
    con.e_to_s(r"sqlexcel\2.xls","学生表")#测试数据库导入功能
    