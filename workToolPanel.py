import wx
# import wx.stc as stc


class MyFrame(wx.Frame):

    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, 'workToolPanel', size=(800, 600))

        # 创建面板
        panel = wx.Panel(self)

        # 在Panel上添加Button
        processButton = wx.Button(panel, label='处理', pos=(300, 450))
        resetButton = wx.Button(panel, label='重置', pos=(450, 450))

        wx.StaticText(panel, label="待处理sql:", pos=(40, 10))
        self.sqlTempStr = wx.TextCtrl(panel, pos=(110, 10), size=(650, 100), style=wx.TE_MULTILINE | wx.TE_RICH2)
        # self.sqlTempStr.SetInsertionPoint(1)
        # self.sqlTempStr.SetWindowStyleFlag(self.sqlTempStr.GetWindowStyleFlag() & ~wx.TE_BESTWRAP | wx.HSCROLL)

        wx.StaticText(panel, label="填充数据:", pos=(45, 120))
        self.insertData = wx.TextCtrl(panel, pos=(110, 120), size=(650, 100), style=wx.TE_MULTILINE | wx.TE_RICH2)

        wx.StaticText(panel, label="填充数据类型:", pos=(15, 230))
        self.dataType = wx.TextCtrl(panel, pos=(110, 230), size=(650, 100), style=wx.TE_MULTILINE | wx.TE_RICH2)

        wx.StaticText(panel, label="结果:", pos=(75, 340))
        self.result = wx.TextCtrl(panel, pos=(110, 340), size=(650, 100), style=wx.TE_MULTILINE | wx.TE_RICH2)

        # 绑定单击事件
        self.Bind(wx.EVT_BUTTON, self.OnClickProcess, processButton)
        self.Bind(wx.EVT_BUTTON, self.OnClickReset, resetButton)

    def OnClickProcess(self, event):
        # 获取录入数据
        sqlTempStr = self.sqlTempStr.GetValue()
        insertData = self.insertData.GetValue()
        insertData = insertData[insertData.rfind("["):].replace("[", "").replace("]", "").replace(" ","")

        dataType = self.dataType.GetValue()
        dataType = dataType[dataType.rfind("["):].replace("[", "").replace("]", "")
        # 处理数据
        result = self.getSqlResult(sqlTempStr, insertData, dataType)
        # 数据赋值
        self.result.Clear()
        self.result.AppendText(result)

    @staticmethod
    def getSqlResult(sqlTempStr, insertData, dataType):
        result = sqlTempStr
        # 转换成list
        insertData = insertData.split(",")
        dataType = dataType.split(",")
        # 循环处理将sqlTempStr 中的?替换成insertData
        index = 0
        while result.find("?") > -1:
            if dataType[index].find("String") > -1:
                data = "\'" + insertData[index] + "\'"
            elif dataType[index].find("Timestamp") > -1:
                data = "to_timestamp(\'" + insertData[index] + "\','yyyy-mm-dd hh24:mi:ss.ff')"
            elif dataType[index].find("Date") > -1:
                data = "to_date(\'" + insertData[index] + "\','yyyy-mm-dd')"
            else:
                data = insertData[index]
            result = result.replace("?", data, 1)
            index = index + 1

        # 去掉特殊字符
        result = result.replace("\t", "").replace("\r", "").replace("\n", "")
        # 去除result多余的空格
        while result.find("  ") > -1:
            result = result.replace("  ", " ")

        return result + ";"

    def OnClickReset(self, event):
        self.sqlTempStr.Clear()
        self.insertData.Clear()
        self.dataType.Clear()
        self.result.Clear()


if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = MyFrame(parent=None, id=-1)
    frame.Show()
    app.MainLoop()
