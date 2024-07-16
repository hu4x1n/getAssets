# _*_ coding:utf-8 _*_
# @Time : 2023/6/3 20:57
# @Author: zhxknb1


import requests
import openpyxl
import time

header = {
    "Host": "nsiscp.jscert.cn:8078",
    "Connection": "close",
    "sec-ch-ua": "Not?A_Brand;v=8, Chromium;v=108, Google Chrome;v=108",
    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOlwvXC8xMjcuMC4wLjE6MTAwMDFcL2FwaVwvYXV0aFwvbG9naW4iLCJpYXQiOjE3MjEwOTI3NDksImV4cCI6MTcyMTE3OTE0OSwibmJmIjoxNzIxMDkyNzQ5LCJqdGkiOiJHQ2NBU2I2TE96RHNDeTFUIiwic3ViIjoiRkU0QUNGNzUzMUY5MDRDQjZGOUE1OTIzNjIyODAwMjEiLCJwcnYiOiIyM2JkNWM4OTQ5ZjYwMGFkYjM5ZTcwMWM0MDA4NzJkYjdhNTk3NmY3In0.L0Bgy75j_3vWuhA9V8F9MCiVtjf6kdE9TNsRe5CEew0",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
    "sec-ch-ua-platform": "Windows",
    "Cookie": "assetreport_session=eyJpdiI6ImxkUWRHS3BpTHQwWmVXanMzL0NubUE9PSIsInZhbHVlIjoiYlJCcGlFN255NTFmN1V0clI1QmlVNlE1OTh5eHZzT3dtVE1EY0VMZm9UcS83VFk4MDYrYjlNeHJTZ09vekF5UWlaampvNG0xbUNOMWRWZnk0OTIrMHhSMEJZZENyaEgwN3BncmRZNEpsNElOenYycW5OanBUVnFlKzNkZ2x5dHAiLCJtYWMiOiJmMmQ5NDI5OGFmNTg3MTk1YjI1NDk3NzMxN2QzNTJmYTdmNzEwMTQxNTgyNTA2OGYyYTUxYWJjOTgxNzhiMTEwIiwidGFnIjoiIn0%3D",
}
page = 1
url = "https://nsiscp.jscert.cn:8078/api/activity/B531B7A7A12BE88C77D94527C8CFB0EB/asset/white_list"

# 在某个文件目录下创建一个excel文件，文件名为TZHWdata.xlsx
workbook = openpyxl.load_workbook("/Users/zhxknb1/Desktop/TZHWdata.xlsx")
worksheet = workbook["sheet1"]
worksheet["A1"] = "assetUrl"
worksheet["B1"] = "realIp"
worksheet["C1"] = "eventNum"
worksheet["D1"] = "noCheckEventNum"
worksheet["E1"] = "findType"
worksheet["F1"] = "createdAt"
worksheet["G1"] = "ingorevulNum"
worksheet["H1"] = "submitRestrict"
worksheet["I1"] = "riskTypeId"
worksheet["J1"] = "riskTypeName"
row = 2
for n in range(1, 30):
    postdata = {
        "page": n,
        "pageSize": "50",
        "assetUrl": "",
        "findType": "",
        "startTime": "",
        "endTime": "",
    }
    response = requests.post(url=url, headers=header, data=postdata)

    data_list = response.json()["data"]["list"]
    for dic in data_list:
        assetUrl = dic["assetUrl"]
        realIp = dic["realIp"]
        eventNum = dic["eventNum"]
        noCheckEventNum = dic["noCheckEventNum"]
        findType = dic["findType"]
        createdAt = dic["createdAt"]
        ingorevulNum = dic["ingorevulNum"]
        submitRestrict = dic["submitRestrict"]
        riskTypeId = dic["riskTypeId"]
        riskTypeName = dic["riskTypeName"]
        print(f"assetUrl:{assetUrl},createdAt:{createdAt}")

        worksheet.cell(row=row, column=1).value = assetUrl
        worksheet.cell(row=row, column=2).value = realIp
        worksheet.cell(row=row, column=3).value = eventNum
        worksheet.cell(row=row, column=4).value = noCheckEventNum
        worksheet.cell(row=row, column=5).value = findType
        worksheet.cell(row=row, column=6).value = createdAt
        worksheet.cell(row=row, column=7).value = ingorevulNum
        worksheet.cell(row=row, column=8).value = submitRestrict
        worksheet.cell(row=row, column=9).value = riskTypeId
        worksheet.cell(row=row, column=10).value = riskTypeName
        row = row + 1
        time.sleep(0.1)

    print(f"第{n}页爬取完成")
    # 修改保存路径
    workbook.save("/Users/zhxknb1/Desktop/TZHWdata.xlsx")

workbook.close()
