import sys, os
import xlwt, xlrd
from xlwt.Workbook import Workbook
from datetime import datetime


def excelToTable(xlsfile):
    # 打开文件
    xls = xlrd.open_workbook(xlsfile)
    sheet = xls.sheet_by_index(0)

    maxRow = sheet.nrows
    maxCol = sheet.ncols

    # 空文件直接退出
    if maxRow == 0: 
        return False

    # 注册表头
    table = []
    header = {}
    for iCol in range(maxCol):
        value = sheet.cell(0, iCol).value
        header[iCol] = value

    # 读取数据
    for iRow in range(1, maxRow):
        item = {}
        for iCol in range(maxCol):
            value = sheet.cell(iRow, iCol).value
            item[header[iCol]] = value
        table.append(item)

    return table

def getStyle(bold = False, border = False):
    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment

    font = xlwt.Font()
    font.name = "宋体"
    font.height = 256
    font.bold = bold
    style.font = font

    if border:
        borders = xlwt.Borders() # Create Borders
        borders.left = xlwt.Borders.THIN 
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        style.borders = borders

    return style

def main(xlsfile):
    table = excelToTable(xlsfile)

    # 新建文件
    workbook = xlwt.Workbook(encoding = 'utf-8')
    sheet = workbook.add_sheet("sheet1", cell_overwrite_ok = True)

    # 写公司名称
    xlsfile = os.path.basename(xlsfile)
    coName = xlsfile.split("-")[0]
    idxRow = 0
    sheet.write_merge(idxRow, idxRow, 0, 9, coName, getStyle(bold = True))
    # 写日期
    startDate = table[0]["税款所属期起"]
    startDate = datetime.strptime(startDate, r'%Y-%m-%d')
    endDate = table[0]["税款所属期止"]
    endDate = datetime.strptime(endDate, r'%Y-%m-%d')
    date = datetime.strftime(startDate, r'%Y年%m月%d日')
    date += datetime.strftime(endDate, r'-%d日')
    idxRow = idxRow + 1
    sheet.write_merge(idxRow, idxRow, 0, 9, date, getStyle(bold = True))
    # 写表头
    idxRow = idxRow + 1
    sheet.write_merge(idxRow, idxRow + 1, 0, 0, "序号", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow + 1, 1, 1, "姓名", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow + 1, 2, 2, "税前工资", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow, 3, 7, "应扣部分", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow + 1, 8, 8, "应发工资", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow + 1, 9, 9, "签字", getStyle(bold = True, border = True))
    idxRow = idxRow + 1
    sheet.write_merge(idxRow, idxRow, 3, 3, "本期基本养老保险费", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow, 4, 4, "本期基本医疗保险费", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow, 5, 5, "本期失业保险费", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow, 6, 6, "公积金", getStyle(bold = True, border = True))
    sheet.write_merge(idxRow, idxRow, 7, 7, "个税", getStyle(bold = True, border = True))

    # 写数据
    sumData = [0] * 9
    headCount = len(table)
    for index in range(headCount): #采用计数，方便之后统计人数
        idxRow = idxRow + 1
        sheet.write(idxRow, 0, index + 1, getStyle(border = True))
        sheet.write(idxRow, 1, table[index]["姓名"], getStyle(border = True))

        value = table[index]["本期收入"]
        value = float(value)
        sumData[2] += value
        salary = value
        sheet.write(idxRow, 2, '%.2f' % value, getStyle(border = True))

        value = table[index]["本期基本养老保险费"]
        value = float(value)
        sumData[3] += value
        salary -= value
        sheet.write(idxRow, 3, '%.2f' % value, getStyle(border = True))

        value = table[index]["本期基本医疗保险费"]
        value = float(value)
        sumData[4] += value
        salary -= value
        sheet.write(idxRow, 4, '%.2f' % value, getStyle(border = True))

        value = table[index]["本期失业保险费"]
        value = float(value)
        sumData[5] += value
        salary -= value
        sheet.write(idxRow, 5, '%.2f' % value, getStyle(border = True))

        value = table[index]["本期住房公积金"]
        value = float(value)
        sumData[6] += value
        salary -= value
        sheet.write(idxRow, 6, '%.2f' % value, getStyle(border = True))

        value = table[index]["累计应补(退)税额"]
        value = float(value)
        sumData[7] += value
        salary -= value
        sheet.write(idxRow, 7, '%.2f' % value, getStyle(border = True))

        value = salary
        sumData[8] += value
        sheet.write(idxRow, 8, '%.2f' % value, getStyle(border = True))
        sheet.write(idxRow, 9, "", getStyle(border = True))

    # 写表尾
    idxRow = idxRow + 1
    sheet.write_merge(idxRow, idxRow, 0, 1, "合计", getStyle(bold = True, border = True))
    for index in range(2, 9):
        value = sumData[index]
        sheet.write(idxRow, index, '%.2f' % value, getStyle(bold = True, border = True))
    sheet.write(idxRow, 9, "", getStyle(border = True))

    # 设置表格大小
    fontSize = 1024
    sheet.col(0).width = fontSize * 2
    sheet.col(1).width = fontSize * 4
    sheet.col(2).width = fontSize * 4
    sheet.col(3).width = fontSize * 8
    sheet.col(4).width = fontSize * 8
    sheet.col(5).width = fontSize * 8
    sheet.col(6).width = fontSize * 4
    sheet.col(7).width = fontSize * 4
    sheet.col(8).width = fontSize * 4
    sheet.col(9).width = fontSize * 4

    heightStyle = xlwt.easyxf('font:height 360') # 18pt
    idxRow = idxRow + 1
    for index in range(idxRow):
        sheet.row(index).set_style(heightStyle)

    # 生成
    filepath = "工资单_" + coName + "_" + date + ".xls"
    if os.path.exists(filepath):
        os.remove(filepath)
    workbook.save(filepath)

if __name__ == "__main__":
    print(sys.argv)
    if len(sys.argv) == 0:
        print("[错误] 未导入文件。")
    
    xlsfile = sys.argv[1]
    if not(os.path.exists(xlsfile)):
        print("[错误] 文件不存在。")
    
    main(xlsfile)