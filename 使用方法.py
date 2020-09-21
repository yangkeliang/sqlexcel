import sqlex
host = '118.24.105.78'
user = 'root'
password = "1qaz!QAZ123***123"
db_name = "ljtestdb"
sheet_name = "t_user"
lines = 103
case = sqlex.Sqlandex(host,user,password,db_name)
case.s_to_e(sheet_name,lines)