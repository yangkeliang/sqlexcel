# sqlexcel使用方法
#### case = sqlex.Sqlandex(host,user,password,db_name) 参数依次为ip，用户名，密码，数据库名
#### case.s_to_e(sheet_name,lines) 参数依次为表名，读取的行数，可实现数据库数据导入表格
#### case.e_to_s(wb_name,ws_name) 参数依次为工作表名，表单名，可实现表格数据导入数据库