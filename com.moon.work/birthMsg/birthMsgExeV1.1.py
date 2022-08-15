import time
import xlrd
from borax.calendars.lunardate import LunarDate
import requests
import json
import hmac
import hashlib
import base64
import urllib.parse
import datetime


# 读取excel文件，生成生日消息
def readExcel(date):
    date = str(date)
    if date.find(" ") > -1:
        date = date[:date.find(" ")]
    # 读取excel文件
    excelFile = xlrd.open_workbook_xls("birthF.xls")
    excelSheet = excelFile.sheet_by_index(0)
    rows = excelSheet.nrows
    # 解析农历和阳历日期
    anaDate = getLunarDate(date)
    norYear = anaDate[0]
    # 获取农历日期
    chnDateNum = anaDate[1]
    chnDateChn = anaDate[2]

    # 循环excel文件，判断是否有满足条件的日期
    perBirthChnStr = ""
    if rows < 3:
        return "birth.xls文件中没有生日信息"

    # excel中存的是农历日期
    for rowNum in range(2, rows):
        row = excelSheet.row_values(rowNum)
        # 拼接当年农历生日日期yyyy/mm/dd
        perBirth = norYear + "/" + row[1]
        remaindDays = row[2]
        if remaindDays == "":
            remaindDays = 0
        remaindDays = int(remaindDays)

        num = getDaysFromTwoDate(perBirth,chnDateNum)
        # 小于0说明当年生日已过，计算来年生日
        if num < 0:
            nextYear = int(norYear) + 1
            perBirth = str(nextYear) + "/" + row[1]
            num = getDaysFromTwoDate(perBirth,chnDateNum)

        if num <= remaindDays:
            perBirthChnStr = perBirthChnStr + row[0] + "生日差" + str(num) + "天.\n"

    perBirthStr = ""
    if perBirthChnStr != "":
        perBirthStr = "今天农历" + chnDateChn + ":" + perBirthChnStr

    return perBirthStr


# 获取两个日期之间查多少天
def getDaysFromTwoDate(date1, date2):
    date1 = datetime.datetime.strptime(date1[0:10], "%Y/%m/%d")
    date2 = datetime.datetime.strptime(date2[0:10], "%Y/%m/%d")
    num = (date1 - date2).days
    return num


# 获取农历日期
def getLunarDate(norDate):
    # 先把空格之后的截掉
    norDate = str(norDate)
    if norDate.find(" ") > -1:
        norDate = norDate[:norDate.find(" ")]
    norDateArr = norDate.split("/")
    # 获取阳历年份
    norYear = str(norDateArr[0])
    # 转换成农历日期
    formatDate = LunarDate.from_solar_date(int(norDateArr[0]), int(norDateArr[1]), int(norDateArr[2]))
    # 农历日期yyyy/mm/dd版本
    chnDateNum = str(formatDate.year) + "/" + str(formatDate.month).zfill(2) + "/" + str(formatDate.day).zfill(2)
    # 农历日期中文版本
    chnDateChn = formatDate.cn_month + "月" + formatDate.cn_day
    return norYear, chnDateNum, chnDateChn


# 向钉钉发送消息
def sendMsg(msg):
    msg = str(msg)
    headers = {"Content-Type": "application/json"}
    data = {
        "msgtype": "text",
        "text": {
            "content": msg
        }
    }
    timestamp = str(round(time.time() * 1000))
    secret = 'SEC97cf7cba9c67f2bb6c74214777c3439272acf7c4e1ee70ead6411e13fa0cfeb5'
    secret_enc = secret.encode('utf-8')
    string_to_sign = '{}\n{}'.format(timestamp, secret)
    string_to_sign_enc = string_to_sign.encode('utf-8')
    hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))

    webhook = "https://oapi.dingtalk.com/robot/send?access_token" \
              "=da2be94addd339065881ccdb9e73a9f62d2c5fba63392b6cfa5c080f6a566763&timestamp=" + timestamp + "&sign=" + sign
    requests.post(webhook, data=json.dumps(data), headers=headers)


now = time.strftime("%Y/%m/%d", time.localtime())
birthMsg = readExcel(now)
if birthMsg != "" and birthMsg != "birth.xls文件中没有生日信息":
    sendMsg(birthMsg)
