"""
构造数据库数据
"""
import faker
import xlwt
num = int(input("请输入行数："))
f = faker.Faker("zh_CN")
wb = xlwt.Workbook()
ws = wb.add_sheet("学生表")
list1 = ["姓名","邮箱","详细地址","出生月份","身份证号","职位","感想"]
for i in range(num):
    for j in range(7):
        list2 = []
        list2.append(f.name())
        list2.append(f.ascii_email())
        list2.append(f.address())
        list2.append(f.month())
        list2.append(f.ssn())
        list2.append(f.job())
        list2.append(f.words())
        if i == 0:
            ws.write(i,j,list1[j])
        else:
            ws.write(i,j,list2[j])
wb.save("2.xls")