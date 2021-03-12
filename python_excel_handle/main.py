# -*- coding: UTF-8 -*-
import xlrd
import xlwt
import sys
import logging
from PIL import ImageTk, Image

import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
############ 配置 ################

my_price_select = ""
############ 配置 ################

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', filename="log.txt")
logger = logging.getLogger(__name__)


def get_input_wuliaohao(filename):
    """
    得到输入的物料号列表
    :param filename:
    :return:
    """
    input_excel = xlrd.open_workbook(filename)
    input_sheets = input_excel.sheets()
    input_sheet_num = input_excel.nsheets
    if input_sheet_num is not 1:
        logger.error("input.xlsx 的sheet数目错误！")
        exit(-1)
    input_wuliaohao = []
    for i in range(input_sheets[input_sheet_num - 1].nrows):
        wuliao = str(input_sheets[input_sheet_num - 1].cell(i, 0).value)
        if wuliao != "物料号" and i == 0:
            logger.error("input 第一行必须是物料号")
        if wuliao == "物料号":
            continue
        input_wuliaohao.append(wuliao)
    return input_wuliaohao


def get_price_list_by_wuliaohao(filename, wuliaohaos, which_price):
    """
    得到字典类型的结果 {物料号，价钱}
    :param filename: 存储价钱的excel文件
    :param wuliaohaos: 物料号列表
    :param which_price: 价钱类型
    :return: {物料号，价钱}
    """
    price_dict = {}
    excel = xlrd.open_workbook(filename)
    '''	打印文件信息 '''
    sheet_num = excel.nsheets
    logger.info("表格sheet的数量为：{}".format(sheet_num))
    all_sheet = excel.sheets()

    for wuliao_str in wuliaohaos:
        for sigle_sheets in all_sheet:
            for row in range(sigle_sheets.nrows):
                for col in range(sigle_sheets.ncols):

                    ceil_str = str(sigle_sheets.cell(row, col).value)
                    # 找到物料号的字符串
                    if ceil_str.strip() == wuliao_str:
                        # 找价钱在第几列
                        res_price = find_price_loc(sigle_sheets, which_price, row)
                        if res_price is None:
                            logger.error("工作表中有物料号，但是没有：{},返回空".format(which_price))
                            price_dict[ceil_str] = ""
                            return price_dict
                        if price_dict.__contains__(ceil_str):
                            logger.error("包含多个{}对应的{},可能input中输入多次".format(wuliao_str, which_price))
                            continue
                        price_dict[ceil_str] = res_price
    return price_dict


def find_price_loc(sheet, price_str, price_str_row):
    """
    找价钱在当前表的第列
    :param sheet: 工作表
    :param price_str: 价格类型的字符串
    :param price_str_row: 要多少行的价格
    :return:  价格
    """
    is_find = 0
    res_col = -1
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            # 找到列表中的一个物料号
            if price_str in str(sheet.cell(row, col).value):
                logger.info("找到了 {} 的位置: sheet:{}，行:{},：列{}".format(price_str, sheet.name, row, col))
                is_find += 1
                if is_find > 1:
                    logger.error("错误，工作表：{} 中包含超过一个字符串{}".format(sheet.name, price_str))
                    return ""
                res_col = col
        # 找价钱在第几列
    if is_find != 1:
        logger.error("错误，工作表：{} 中未找到字符串{}".format(sheet.name, price_str))
        return ""
    return str(sheet.cell(price_str_row, res_col).value)


def output_res(input_list, out_file, outdict, price_select):
    """
    按照输入的格式、顺序输出结果
    :param input_file:
    :param out_fil:
    :return:
    """
    save_excel = xlwt.Workbook()
    out_sheet = save_excel.add_sheet(price_select, cell_overwrite_ok=True)  # 创建sheet
    i = 1
    out_sheet.write(0, 0, "物料号")
    out_sheet.write(0, 1, price_select)

    for wuliaohao in input_list:
        out_sheet.write(i, 0, wuliaohao)
        if outdict.__contains__(wuliaohao):
            out_sheet.write(i, 1, outdict.get(wuliaohao))
            logger.info("保存{}，key:{} value:{}".format(out_file, wuliaohao, outdict.get(wuliaohao)))
        else:
            logger.error("物料号：{} ,价格未找到".format(wuliaohao))
        i += 1
    save_excel.save(out_file)


def main():
    my_price_file = "我的价格.xlsx"
    inpu_file_name = u'input.xlsx'
    if my_price_select == "" or my_price_file is None:
        return
    output = "{}-结果输出.xls".format(my_price_select)
    intput_wuliaos = get_input_wuliaohao(inpu_file_name)
    logger.info(intput_wuliaos)
    res_dict = get_price_list_by_wuliaohao(my_price_file, intput_wuliaos, my_price_select)
    logger.info(res_dict)
    output_res(intput_wuliaos, output, res_dict, my_price_select)


def find_price_res():
    logger.info("find_price_res")
    main()
    tk.messagebox.showinfo(title='查找结果', message='查找成功')

selected = None


def combo_chose(*args):
    global my_price_select
    logger.info("下拉选择：{}".format(comboxlist.get()))
    my_price_select = str(comboxlist.get())


def get_image(filename, width, height):
    im = Image.open(filename).resize((width, height))
    return ImageTk.PhotoImage(im)

window = tk.Tk()  # 主窗口
window.title('赵楠的自动报价工具')  # 窗口标题
window.geometry('450x600')  # 窗口尺寸
canvas_root = tk.Canvas(window, width=450, height=600)
im_root = get_image("a.jpg", 450, 600)
canvas_root.create_image(225, 300, image=im_root)
canvas_root.pack()

b = tk.Button(window, text='查找价格', command=find_price_res).place(x=300, y=90)  # 点击按钮执行的命令

var2 = tk.StringVar()
var2.set(("行业价", "区域价", "产品线底价"))  # 为变量设置值

""" 下拉框"""
comvalue = tk.StringVar()  # 窗体自带的文本，新建一个值
comboxlist = ttk.Combobox(window, textvariable=comvalue)  # 初始化
comboxlist["values"] = ("标准价", "总部价", "行业价", "产品线低价", "成本", "区域经理价")
#comboxlist.current(0)  # 选择第一个
comboxlist.bind("<<ComboboxSelected>>", combo_chose)  # 绑定事件,(下拉列表框被选中时，绑定go()函数)
comboxlist.place(x=100, y=100)
#w = tk.Label(window, text="请选择查询价格类型", font=("华文行楷", 10), fg="black")
w = tk.Label(window, text="请选择查询价格类型", fg="black")

w.place(x=120, y=75)
w = tk.Label(window, text="nice day !", font=("黑体", 20), fg="green")
w.place(x=140, y=10)


window.mainloop()
