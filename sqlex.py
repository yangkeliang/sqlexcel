import xlwt
import pymysql as psql

def query(sql,db_name):
    #连接
    db = psql.connect(host='118.24.105.78', user='root', password="1qaz!QAZ123***123", db=db_name)
    #查询
    cur = db.cursor()
    cur.execute(sql)
    res = cur.fetchall()
    return res

def get_titlelist(sql_title,db_name):#获取表头
    list_title = []
    for i in query(sql_title,db_name):
        list_title.append(i[0])
    return list_title

def get_contentlist(sql_content,db_name):#获取表内容
    list_content = []
    for i in query(sql_content,db_name):
        list_each = []
        for j in i:   
            list_each.append(j)
        list_content.append(list_each)
    return list_content

def  write_to_excel(list_all,sheet_name):#写入表格
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)
    for i in range(len(list_all)):
        for j in range(len(list_all[i])):
            ws.write(i,j,list_all[i][j])
    wb.save("1.xls")

if __name__=="__main__":
    db_name = input("请输入数据库名：")
    sheet_name = input("请输入表名：")
    lines = int(input("请输入行数:"))
    sql_title = "SELECT COLUMN_NAME FROM \
        information_schema.COLUMNS\
             WHERE TABLE_SCHEMA = '{}' \
                 AND TABLE_NAME = '{}';".format(db_name,sheet_name)
    sql_content = "select * from {} limit {};".format(sheet_name,lines)
    list_all = []#拼接title与content
    list_all.append(get_titlelist(sql_title,db_name))
    list_all.extend(get_contentlist(sql_content,db_name))
    write_to_excel(list_all,sheet_name)