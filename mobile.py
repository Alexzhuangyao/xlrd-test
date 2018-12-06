import xlrd
import requests
import json

f = xlrd.open_workbook('C:\\Users\LvjingService\Documents\Tencent Files\896988640\FileRecv\公司通讯录(2018.11).xlsx')

bookSheet = f.sheet_by_name('Sheet2')

#列，行
#行总计（0，90） 列总计（0，4）（1部门/区域(改成employeeOrgid) 2姓 名 3 岗位/职级 4 手机号码）
cell_32 = bookSheet.cell_value(3,1)
print(cell_32)#姓名


# class send_post(object):
#     #发送POST请求，新增employee
#     def __init__(self,data):
#         self.url = "http://127.0.0.1:8092/employee/insertEmployee"
#         self.data = data
#         self.header = {
#              "Content-Type": "application/json"
#          }
#
#     def get_response(self):
#         try:
#             response = requests.post(url=self.url, data=self.data,headers=self.header)
#             # print(response.content)
#             print("{}:{}>>>添加EMPLOYEE表成功".format(self.url,self.data))
#         except Exception:
#             print("{}:{}>>>添加EMPLOYEE失败".format(self.url,self.data))
#
#     def run(self):
#         self.get_response()


for i in range(3,90):
    #employeeName
    employeeName = bookSheet.cell_value(i,2)
    #position
    position = bookSheet.cell_value(i,3)
    #employeeMobile
    employeeMobile = bookSheet.cell_value(i,4)
    if employeeMobile != "":
        mobileSpilt = str(employeeMobile).split("/")
        employeeMobile = str(employeeMobile)

    else:
        employeeMobile = "未知"
        mobileSpilt = "未知"
    #employeeOrgid
    employeeOrgid = bookSheet.cell_value(i,1)
    result = {"employeeId": ("E" + mobileSpilt[0]).split(".")[0],
              "employeeName": str(employeeName).replace(" ",""),
              "employeeSex": "保密",
              "employeeAge": "-1",
             "employeeMobile":employeeMobile.split(".")[0],
              "identityNumber": "未知",
              "position": str(position).replace(" ",""),
              "employeeOrgid": str(employeeOrgid).split(".")[0],
              "employeeType":str(position).replace(" ","")
              }
    print(str(result).replace('\'',"\""))
    #send_post(data=result).run()
