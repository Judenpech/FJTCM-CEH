import datetime
import os
import sys
import time
import traceback
import winreg
import xlrd
import xlwt

from xlutils.copy import copy


def get_desktop():
    """
    获取当前用户桌面路径
    :return: 当前用户桌面路径
    """
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, \
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', )
    return winreg.QueryValueEx(key, "Desktop")[0]


def open_excel(file):
    """
    获取workbook
    :param file: excel文件
    :return: workbook
    """
    try:
        wb = xlrd.open_workbook(file, formatting_info=True)
        return wb
    except Exception as e:
        print(str(e))


def get_files_name(file_dir):
    """
    获取要处理的文件夹目录下的所有文件名
    :param file_dir:文件夹路径
    :return:文件名列表
    """
    filesList = []
    for roots, dirs, files in os.walk(file_dir):
        # # files是一个list，内容是该文件夹中所有的文件
            if files:
                for f in files:
                    filesList.append(f)
    return filesList


def set_formula(dirPath, filesList):
    """
    批量为加减分申报表设定公式
    :param dirPath: 文件夹路径
    :param filesList: 文件名列表
    :return: 成功：1  失败：0
    """
    print("即将进行处理加减分申报表...*´∀`)´∀`)*´∀`)*´∀`)\n")
    try:
        for f in filesList:
            fPath = dirPath + "\\" + f
            print("正在处理" + f)
            wb = copy(open_excel(fPath))
            ws = wb.get_sheet(0)
            ws.write(5, 10, xlwt.Formula('SUM(D6:D10,F6:F10,H6:H10)+SUM(J6:J10)'))
            ws.write(13, 10, xlwt.Formula('SUM(D14:D20,F14:F20,H14:H20)+SUM(J14:J20)'))
            ws.write(23, 10, xlwt.Formula('SUM(D24:D29,F24:F29,H24:H29,J24:J29)'))
            ws.write(29, 10, xlwt.Formula('SUM(K6,K14,K24)'))  # bug：求和等于零的话单元格值不为0
            newDir = dirPath + "\\处理后"
            if not os.path.exists(newDir):
                os.makedirs(newDir)
            wb.save(newDir + "\\" + f)
            print("处理成功！(๑•̀ㅂ•́)و✧")
    except:
        print("\nOops!w(ﾟДﾟ)w\n程序出错啦，可以访问作者的GitHub页面上报bug哦~\n"
              + "访问链接：https://github.com/jl223vy/FJTCM-CEH")
        print("\n出错信息:\n" + traceback.format_exc())
        return 0
    return 1


def get_moral(dirPath, filesList):
    """
    批量计算德育加奖分、减罚分和总分
    :param dirPath: 文件夹路径
    :param filesList: 文件名列表
    :return: flag和一个List，包括德育加奖分、减罚分和总分三个Dict
    """
    print("即将进行计算德育加减分...*´∀`)´∀`)*´∀`)*´∀`)\n")
    moralAward = {}
    moralPenalty = {}
    moralTotal = {}
    try:
        for f in filesList:
            print("正在计算" + f)
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
        print("\n计算德育加减分成功！(๑•̀ㅂ•́)و✧")
        return 1, [moralAward, moralPenalty, moralTotal]
    except:
        print("\nOops!w(ﾟДﾟ)w\n程序出错啦，可以访问作者的GitHub页面上报bug哦~\n"
              + "访问链接：https://github.com/jl223vy/FJTCM-CEH")
        print("\n出错信息:\n" + traceback.format_exc())


def get_discipline(dirPath, filesList):
    """
    批量计算智育加奖分、减罚分和总分
    :param dirPath: 文件夹路径
    :param filesList: 文件名列表
    :return: flag和一个List，包括智育加奖分、减罚分和总分三个Dict
    """
    print("即将进行计算智育加减分...*´∀`)´∀`)*´∀`)*´∀`)\n")
    disciplineAward = {}
    disciplinePenalty = {}
    disciplineTotal = {}
    try:
        for f in filesList:
            print("正在计算" + f)
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
        print("\n计算智育加减分成功！(๑•̀ㅂ•́)و✧")
        return 1, [disciplineAward, disciplinePenalty, disciplineTotal]
    except:
        print("\nOops!w(ﾟДﾟ)w\n程序出错啦，可以访问作者的GitHub页面上报bug哦~\n"
              + "访问链接：https://github.com/jl223vy/FJTCM-CEH")
        print("\n出错信息:\n" + traceback.format_exc())


def save_moral(moral, filesList):
    """
    保存德育加减分表至桌面
    :param moral: 包含德育加奖分、减罚分和总分三个Dict的List
    :param filesList: 文件名列表
    :return: none
    """
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
    wb.save(get_desktop() + "\\德育加减分" + datetime.datetime.now().strftime('%Y%m%d%H%M%S') + ".xls")
    print("文件已保存至桌面！文件名为【德育加减分(保存时间).xls】")


def save_discipline(discipline, filesList):
    """
        保存智育加减分表至桌面
        :param moral: 包含智育加奖分、减罚分和总分三个Dict的List
        :param filesList: 文件名列表
        :return: none
        """
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
    wb.save(get_desktop() + "\\智育加减分" + datetime.datetime.now().strftime('%Y%m%d%H%M%S') + ".xls")
    print("文件已保存至桌面！文件名为【智育加减分(保存时间).xls】")


def main():
    print("(｡･∀･)ﾉﾞ欢迎使用福建中医药大学综测助手!\n\n【程序功能说明】：\n"
          + "1、对所有加减分申报表批量设定公式\n"
          + "2、统计汇总德、智育加奖分、减罚分和总分至Excel表格\n"
          + "\n【使用须知】:\n"
          + "1、请提供完整的文件夹路径，如：C:\\Users\\Lenovo\\Desktop\\加减分申报表汇总\n"
          + "2、所有的加减分申报表都必须以[.xls]为扩展名\n"
          )
    dirPath = input("请输入完整的文件夹路径：")
    filesList = get_files_name(dirPath)
    setFormulaSucceed = set_formula(dirPath, filesList)
    if setFormulaSucceed == 1:
        dirPath = dirPath + "\\处理后"
        n = 0
        while n != 3:
            print("\n请选择您要执行的操作：\n[1] 获取德育加奖分、减罚分和总分\n[2] 获取智育加奖分、减罚分及总分\n[3] 退出\n")
            n = eval(input())
            while n not in [1, 2, 3]:
                print("输入有误，请重新输入！")
                print("\n请选择您要执行的操作：\n[1] 获取德育加奖分、减罚分和总分\n[2] 获取智育加奖分、减罚分及总分\n[3] 退出\n")
                n = eval(input())
            if n == 1:
                flag, moral = get_moral(dirPath, filesList)
                if flag == 1:
                    save_moral(moral, filesList)
            elif n == 2:
                flag, discipline = get_discipline(dirPath, filesList)
                if flag == 1:
                    save_discipline(discipline, filesList)
            elif n == 3:
                print("感谢使用，再见！(*´∀`)~♥")
                time.sleep(3)
                sys.exit(0)


if __name__ == '__main__':
    main()
