import xlwt
import pymysql as psql

class Sqlandex:
    def __init__(self,host,user,password,db_name):#初始化读入数据库基本的信息
        self.host = host
        self.user = user
        self.password = password
        self.db_name = db_name
    def query(self,sql):#查询函数
        self.db = psql.connect(host=self.host, user=self.user, password=self.password, db=self.db_name)#查询
        self.cur = self.db.cursor()
        self.cur.execute(sql)
        self.res = self.cur.fetchall()
        return self.res
    def get_titlelist(self,sql_title):#获取表头
        self.list_title = []
        title = Sqlandex(self.host,self.user,self.password,self.db_name)
        for i in title.query(sql_title):
            self.list_title.append(i[0])
        return self.list_title
    def get_contentlist(self,sql_content):#获取表内容
        self.list_content = []
        content = Sqlandex(self.host,self.user,self.password,self.db_name)
        for i in content.query(sql_content):
            self.list_each = []
            for j in i:   
                self.list_each.append(j)
            self.list_content.append(self.list_each)
        return self.list_content
    def  write_to_excel(self,list_all,sheet_name):#写入表格
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet(sheet_name)
        for i in range(len(list_all)):
            for j in range(len(list_all[i])):
                self.ws.write(i,j,list_all[i][j])
        self.wb.save("./sqlexcel/{}.xls".format(sheet_name))
    def s_to_e(self,sheet_name,lines):
        self.sql_title = "SELECT COLUMN_NAME FROM \
            information_schema.COLUMNS\
                WHERE TABLE_SCHEMA = '{}' \
                    AND TABLE_NAME = '{}';".format(self.db_name,sheet_name)
        self.sql_content = "select * from {} limit {};".format(sheet_name,lines)
        self.list_all = []#拼接title与content
        title = Sqlandex(self.host,self.user,self.password,self.db_name)
        content = Sqlandex(self.host,self.user,self.password,self.db_name)
        self.list_all.append(title.get_titlelist(self.sql_title))
        self.list_all.extend(content.get_contentlist(self.sql_content))
        title.write_to_excel(self.list_all,sheet_name)

if __name__ == "__main__":
    host = '118.24.105.78'
    user = 'root'
    password = "1qaz!QAZ123***123"
    db_name = "ljtestdb"
    con = Sqlandex(host,user,password,db_name)
    con.s_to_e("t_user",30)