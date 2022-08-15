import xlrd
import io
from docxtpl import DocxTemplate
import json
import os


def readExcel():
    # 读取excel文件
    excelFile = xlrd.open_workbook("data.xlsx")
    # 获取第一个sheet中的数据
    excelSheet = excelFile.sheet_by_index(0)
    rows = excelSheet.nrows
    # 循环excel文件行
    if rows < 3:
        return "birth.xls文件中没有生日信息"
    contexts = []
    # 循环每一列，填充数据到word文件中
    for rowNum in range(2, rows):
        # 创建填充字符
        content = io.StringIO()
        content.write("{")
        # 获取每一行的最大列数
        row = excelSheet.row_values(rowNum)
        itemNums = len(row)
        for itemNum in range(0, itemNums - 1):
            data = row[itemNum]
            if data == "":
                data = "/"
            content.write("\"d" + str(itemNum) + "\": \"" + str(data) + "\",")
        # 去掉最后一个逗号
        context = content.getvalue().strip(",") + "}"
        contexts.append(json.loads(context))
    # 创建要保存的文件
    if not os.path.exists("./result"):
        os.mkdir("./result")
    for content in contexts:
        tpl = DocxTemplate("model.docx")
        tpl.render(content)
        tpl.save("./result/{}.docx".format(content["d0"]))


readExcel()
