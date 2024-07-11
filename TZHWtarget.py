# _*_ coding:utf-8 _*_
# @Time : 2023/5/29 20:57
# @Author: 为赋新词强说愁


import requests
import openpyxl
import time

header = {
    "Host": "nsiscp.jscert.cn:8078",
    "Connection": "close",
    "Content-Length": "58",
    "sec-ch-ua": "Not?A_Brand;v=8, Chromium;v=108, Google Chrome;v=108",
    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOlwvXC8xMjcuMC4wLjE6MTAwMDFcL2FwaVwvYXV0aFwvbG9naW4iLCJpYXQiOjE3MTc0MDMzMDUsImV4cCI6MTcxNzQ4OTcwNSwibmJmIjoxNzE3NDAzMzA1LCJqdGkiOiJMM2V4MWFVOFBtZDhyOUhhIiwic3ViIjoiRkU0QUNGNzUzMUY5MDRDQjZGOUE1OTIzNjIyODAwMjEiLCJwcnYiOiIyM2JkNWM4OTQ5ZjYwMGFkYjM5ZTcwMWM0MDA4NzJkYjdhNTk3NmY3In0.GYj0XZ7-OTyePp_lwmfI1cDUw9kI6sXDOwdiYKoIxT4",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
    "sec-ch-ua-platform": "Windows",
    "Origin": "https://nsiscp.jscert.cn:8078",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Dest": "empty",
    "Referer": "https://nsiscp.jscert.cn:8078/",
    "Cookie": "assetreport_session=eyJpdiI6IlNtd2l6V1d3OXFqd1VOejJLbndQNUE9PSIsInZhbHVlIjoiMnVGQ1pzQlRPNWoxZXhYSTBndTRmTk84aFJpYzhIeXdXTEg0c0hkOXJRUTQrNHFaL3piMW03TW1CME1LN2xWeWhJdjE4bGt2Vm1BVndpNUMyYkprY2FzVDdmZDZZbUl5RHBCRXRzNGd0TVdXR0xJcGJ2cG96SnQwNEQ2MXI1YXEiLCJtYWMiOiJmZTA4NjUzNGJjYjZmMDRjZWQyNDA0NzZiM2FjN2Q3ZDc5YzAxNDBjNzFmYTI1MzdhOGE5ZmUxZDNiMTMyYjhkIiwidGFnIjoiIn0%3D",
}
page = 1
url = "https://nsiscp.jscert.cn:8078/api/activity/C3462D409C57497E5F054FF9C11F4366/asset/white_list"

workbook = openpyxl.load_workbook("TZHWdata.xlsx")
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
for n in range(1, 280):
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
    workbook.save("TZHWdata.xlsx")

workbook.close()
