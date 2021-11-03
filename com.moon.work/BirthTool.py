import xlrd
import xlwt
import time
from borax.calendars.lunardate import LunarDate
import wx


class MyFrame(wx.Frame):
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, 'birthToolPanel', size=(800, 600))

        # 创建面板
        panel = wx.Panel(self)

        # 在Panel上添加Button
        getMsgButton = wx.Button(panel, label='获取生日消息', pos=(300, 450))
        sendMsgButton = wx.Button(panel, label='发送消息', pos=(450, 450))

        # 绑定单击事件
        self.Bind(wx.EVT_BUTTON, self.getBirthMsg, getMsgButton)
        self.Bind(wx.EVT_BUTTON, self.sendMsg, sendMsgButton)

    def getBirthMsg(self, event):
        date = "2021/11/3"
        birthMsg = self.readExcel(date)
        print(birthMsg)

    def sendMsg(self, event):
        pass

    @staticmethod
    def readExcel(date):
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
            perBirthStr = "今日" + date + "生日：" + perBirthNorStr + "。"
        if perBirthChnStr != "":
            perBirthChnStr = perBirthChnStr[:perBirthChnStr.rfind("、")]
            perBirthStr = perBirthStr + "农历" + chnDateChn + "生日：" + perBirthChnStr + "。"
        return perBirthStr

    @staticmethod
    def getLunarDate(norDate):
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
