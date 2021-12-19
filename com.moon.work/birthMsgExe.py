import time
import xlrd
from borax.calendars.lunardate import LunarDate
import requests
import json
import hmac
import hashlib
import base64
import urllib.parse


# 读取excel文件，生成生日消息
def readExcel(date):
    date = str(date)
    if date.find(" ") > -1:
        date = date[:date.find(" ")]
    # 读取excel文件
    excelFile = xlrd.open_workbook_xls("birth.xls")
    excelSheet = excelFile.sheet_by_index(0)
    rows = excelSheet.nrows
    # 解析农历和阳历日期
    anaDate = getLunarDate(date)
    norDate = anaDate[0]
    # 获取农历日期
    chnDateNum = anaDate[1]
    chnDateChn = anaDate[2]

    # 循环excel文件，判断是否有满足条件的日期
    perBirthNorStr = ""
    perBirthChnStr = ""
    if rows < 3:
        return "birth.xls文件中没有生日信息"

    for rowNum in range(2, rows):
        row = excelSheet.row_values(rowNum)
        perBirth = getLunarDate(row[1])
        if norDate == perBirth[0]:
            perBirthNorStr = perBirthNorStr + row[0] + "、"
        elif chnDateNum == perBirth[1]:
            perBirthChnStr = perBirthChnStr + row[0] + "、"

    perBirthStr = ""
    if perBirthNorStr != "":
        perBirthNorStr = perBirthNorStr[:perBirthNorStr.rfind("、")]
        perBirthStr = "今日" + date + "生日：" + perBirthNorStr + "。\n"
    if perBirthChnStr != "":
        perBirthChnStr = perBirthChnStr[:perBirthChnStr.rfind("、")]
        perBirthStr = perBirthStr + "农历" + chnDateChn + "生日：" + perBirthChnStr + "。"
    return perBirthStr


# 获取农历日期
def getLunarDate(norDate):
    # 先把空格之后的截掉
    norDate = str(norDate)
    if norDate.find(" ") > -1:
        norDate = norDate[:norDate.find(" ")]
    norDateArr = norDate.split("/")
    # 获取阳历日期mm/dd版本
    norDate = str(norDateArr[1]).zfill(2) + "/" + str(norDateArr[2]).zfill(2)
    # 转换成农历日期
    formatDate = LunarDate.from_solar_date(int(norDateArr[0]), int(norDateArr[1]), int(norDateArr[2]))
    # 农历日期mm/dd版本
    chnDateNum = str(formatDate.month).zfill(2) + "/" + str(formatDate.day).zfill(2)
    # 农历日期中文版本
    chnDateChn = formatDate.cn_month + "月" + formatDate.cn_day
    return norDate, chnDateNum, chnDateChn


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
    secret = 'SEC07640358776a56534d8ea02d4ca4913f1c2e7d688a082ef306d5e78efa4cb232'
    secret_enc = secret.encode('utf-8')
    string_to_sign = '{}\n{}'.format(timestamp, secret)
    string_to_sign_enc = string_to_sign.encode('utf-8')
    hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))

    webhook = "https://oapi.dingtalk.com/robot/send?access_token" \
              "=9d53fb7f9e571a154fbaaf1943957601f2d209dea3a8c806f3d99a4120d68059&timestamp=" + timestamp + "&sign=" +\
              sign
    requests.post(webhook, data=json.dumps(data), headers=headers)


now = time.strftime("%Y/%m/%d", time.localtime())
birthMsg = readExcel(now)
if birthMsg != "" and birthMsg != "birth.xls文件中没有生日信息":
    sendMsg(birthMsg)
