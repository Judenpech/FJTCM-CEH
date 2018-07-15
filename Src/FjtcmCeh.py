import os, xlrd, xlwt
import sys
import traceback
import winreg
from xlutils.copy import copy


def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, \
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', )
    return winreg.QueryValueEx(key, "Desktop")[0]


def open_excel(file):
    try:
        wb = xlrd.open_workbook(file, formatting_info=True)
        return wb
    except Exception as e:
        print(str(e))


def get_files_name(file_dir):
    """
    获取要处理的文件夹目录下的所有文件名
    :param file_dir:
    :return:
    """
    L = []
    for roots, dirs, files in os.walk(file_dir):
        # files是一个list，内容是该文件夹中所有的文件
        if files:
            for f in files:
                L.append(f)
    return L


def set_formula(dirPath, filesList):
    for f in filesList:
        fPath = dirPath + "\\" + f
        print("正在处理" + f)
        wb = copy(open_excel(fPath))
        ws = wb.get_sheet(0)
        try:
            ws.write(5, 10, xlwt.Formula('SUM(D6:D10,F6:F10,H6:H10)+SUM(J6:J10)'))
            ws.write(13, 10, xlwt.Formula('SUM(D14:D20,F14:F20,H14:H20)+SUM(J14:J20)'))
            ws.write(23, 10, xlwt.Formula('SUM(D24:D29,F24:F29,H24:H29,J24:J29)'))
            ws.write(29, 10, xlwt.Formula('SUM(K6,K14,K24)'))  # bug：求和等于零的话单元格值不为0
            newDir = dirPath + "\\处理后"
            if not os.path.exists(newDir):
                os.makedirs(newDir)
            wb.save(newDir + "\\" + f)
        except:
            traceback.print_exc()
        print("处理成功！")


def get_moral(dirPath, filesList):  # 德育
    print("正在计算德育加减分...")
    moralAward = {}
    moralPenalty = {}
    moralTotal = {}
    for f in filesList:
        fPath = dirPath + "\\" + f
        wb = open_excel(fPath)
        ws = wb.sheet_by_index(0)
        cntAward = 0.0
        cntPenalty = 0.0
        for i in range(5, 10):
            for j in range(3, 8, 2):
                value1 = ws.cell(i, j).value
                if not value1 == "":
                    cntAward += value1
            value2 = ws.cell(i, 9).value
            if not value2 == "":
                cntPenalty += value2
        for i in range(23, 29):
            for j in range(3, 10, 2):
                value3 = ws.cell(i, j).value
                if not value3 == "":
                    cntAward += value3
        cntTotal = cntAward + cntPenalty
        moralAward[f] = round(cntAward, 1)
        moralPenalty[f] = round(abs(cntPenalty), 1)
        moralTotal[f] = round(cntTotal, 1)
    print("计算德育加减分成功！")
    return [moralAward, moralPenalty, moralTotal]


def get_discipline(dirPath, filesList):  # 智育
    print("正在计算智育加减分...")
    disciplineAward = {}
    disciplinePenalty = {}
    disciplineTotal = {}
    for f in filesList:
        fPath = dirPath + "\\" + f
        wb = open_excel(fPath)
        ws = wb.sheet_by_index(0)
        cntAward = 0.0
        cntPenalty = 0.0
        for i in range(13, 20):
            for j in range(3, 8, 2):
                value1 = ws.cell(i, j).value
                if not value1 == "":
                    cntAward += value1
            value2 = ws.cell(i, 9).value
            if not value2 == "":
                cntPenalty += value2
        cntTotal = cntAward + cntPenalty
        disciplineAward[f] = round(cntAward, 1)
        disciplinePenalty[f] = round(abs(cntPenalty), 1)
        disciplineTotal[f] = round(cntTotal, 1)
    print("计算智育加减分成功！")
    return [disciplineAward, disciplinePenalty, disciplineTotal]


def save_moral(moral, filesList):
    print("正在保存德育加减分...")
    wb = xlwt.Workbook()
    ws = wb.add_sheet('德育加减分')
    header = ["文件名", "加奖分", "减罚分", "总分"]
    col = 0
    for h in header:
        ws.write(0, col, h)
        col += 1
    row = 1
    for f in filesList:
        ws.write(row, 0, f)
        for j in range(1, col):
            ws.write(row, j, moral[j - 1][f])
        row += 1
    wb.save(get_desktop() + "\\德育加减分.xls")
    print("德育加减分已保存至桌面[德育加减分.xls]")


def save_discipline(discipline, filesList):
    print("正在保存智育加减分...")
    wb = xlwt.Workbook()
    ws = wb.add_sheet('智育加减分')
    header = ["文件名", "加奖分", "减罚分", "总分"]
    col = 0
    for h in header:
        ws.write(0, col, h)
        col += 1
    row = 1
    for f in filesList:
        ws.write(row, 0, f)
        for j in range(1, col):
            ws.write(row, j, discipline[j - 1][f])
        row += 1
    wb.save(get_desktop() + "\\智育加减分.xls")
    print("智育加减分已保存至桌面[智育加减分.xls]")


def main():
    print("欢迎使用福建中医药大学综测助手，本程序可对加减分申报表进行批量处理和统计。")
    dirPath = input("请输入完整的文件夹路径(文件夹中所有的文件都必须以[.xls]为扩展名)：")
    filesList = get_files_name(dirPath)
    set_formula(dirPath, filesList)
    dirPath = dirPath + "\\处理后"
    n = 0
    while n != 3:
        print("\n请选择您要执行的操作：\n1. 获取德育加奖分、减罚分及总分\n2. 获取智育加奖分、减罚分及总分\n3. 退出\n")
        n = eval(input())
        while n not in [1, 2, 3]:
            print("输入有误，请重新输入！")
            print("\n请选择您要执行的操作：\n1. 获取德育加奖分、减罚分及总分\n2. 获取智育加奖分、减罚分及总分\n3. 退出\n")
            n = eval(input())
        if n == 1:
            moral = get_moral(dirPath, filesList)
            save_moral(moral, filesList)
        elif n == 2:
            discipline = get_discipline(dirPath, filesList)
            save_discipline(discipline, filesList)
        elif n == 3:
            print("感谢使用，再见！")
            sys.exit(0)


if __name__ == '__main__':
    main()
