import xlrd
from borax.calendars.lunardate import LunarDate
import wx
import wx.adv
import requests
import json
import hmac
import hashlib
import base64
import urllib.parse
import time


class MyFrame(wx.Frame):
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, 'birthToolPanel', size=(500, 500))

        # 创建面板
        panel = wx.Panel(self)

        self.date = wx.adv.CalendarCtrl(panel, pos=(100, 10), size=(300, 200))

        wx.StaticText(panel, label="农历:", pos=(50, 220))
        self.chnDate = wx.TextCtrl(panel, pos=(100, 220), style=wx.TE_READONLY)
        dateArr = self.getLunarDate(self.date.GetDate())
        self.chnDate.AppendText(dateArr[2])
        # 生日消息
        wx.StaticText(panel, label="生日消息:", pos=(30, 260))
        self.birthMsg = wx.TextCtrl(panel, pos=(100, 260), size=(300, 150), style=wx.TE_MULTILINE | wx.TE_RICH2)

        # 在Panel上添加Button
        getMsgButton = wx.Button(panel, label='获取生日消息', pos=(120, 420))
        sendMsgButton = wx.Button(panel, label='发送消息', pos=(240, 420))

        # 绑定单击事件
        self.Bind(wx.EVT_BUTTON, self.getBirthMsg, getMsgButton)
        self.Bind(wx.EVT_BUTTON, self.sendMsg, sendMsgButton)
        self.date.Bind(wx.adv.EVT_CALENDAR_SEL_CHANGED, self.chgDate)

    def chgDate(self, event):
        self.chnDate.Clear()
        dateArr = self.getLunarDate(self.date.GetDate())
        self.chnDate.AppendText(dateArr[2])

    def getBirthMsg(self, event):
        birthMsg = self.readExcel(self.date.GetDate())
        self.birthMsg.Clear()
        self.birthMsg.AppendText(birthMsg)

    def sendMsg(self, event):
        headers = {"Content-Type": "application/json"}
        data = {
            "msgtype": "text",
            "text": {
                "content": self.birthMsg.GetValue()
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

    @staticmethod
    def readExcel(date):
        date = str(date)
        if date.find(" ") > -1:
            date = date[:date.find(" ")]
        # 读取excel文件
        excelFile = xlrd.open_workbook_xls("birth.xls")
        excelSheet = excelFile.sheet_by_index(0)
        rows = excelSheet.nrows
        # 解析农历和阳历日期
        anaDate = MyFrame.getLunarDate(date)
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
            perBirth = MyFrame.getLunarDate(row[1])
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

    @staticmethod
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


if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = MyFrame(parent=None, id=-1)
    frame.Show()
    app.MainLoop()
