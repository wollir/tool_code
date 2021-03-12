# -*- coding: UTF-8 -*-
import xlrd
import xlwt
import sys
import logging

import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
############ 配置 ################

my_price_select = ""
############ 配置 ################

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


if __name__ == "__main__":
    filename = "a.xlsx"
    input_excel = xlrd.open_workbook(filename)
    input_sheets = input_excel.sheets()
    # input_sheet_num = input_excel.nsheets
    input_sheet_num = 1
    logger.info(input_sheet_num)
    input_wuliaohao = []
    bu_han_shui_shouru_2 = []
    AR_num_3 = []
    dan_jia_4 = []
    nei_bu_xing_hao_5 = []
    for i in range(1, input_sheets[0].nrows):
        logger.info("i: {}".format(i))
        temp_buhanshui = str(input_sheets[0].cell(i, 2).value)
        if temp_buhanshui == "":
            bu_han_shui_shouru_2.append(0)
        else:
            bu_han_shui_shouru_2.append(float(input_sheets[0].cell(i, 2).value))
        AR_num_3.append(float(input_sheets[0].cell(i, 3).value))
        dan_jia_4.append(float(input_sheets[0].cell(i, 4).value))
        nei_bu_xing_hao_5.append(str(input_sheets[0].cell(i, 5).value))

    nei_bu_xing_hao_set = list(set(nei_bu_xing_hao_5))
     # 四个列表的大小
    list_size = len(nei_bu_xing_hao_5)
    # 型号数量
    version_num = len(nei_bu_xing_hao_set)
    # todo 这个列表要根据查询数据填好
    mimi_out_price = dict()
    for version in nei_bu_xing_hao_set:
        mimi_out_price[version] = 1000
    mimi_out_price["DH-S3000C-16GT"] = 227
    mimi_out_price["DH-S3000C-8GT"] = 98
    # shou ru he
    all_income = dict()
    all_AR_sum = dict()
    for version in nei_bu_xing_hao_set:
        all_income[version] = 0
        for l_i in range(list_size):
            if nei_bu_xing_hao_5[l_i] == version and dan_jia_4[l_i] < mimi_out_price[version]:
                all_income[version] += bu_han_shui_shouru_2[l_i]

        all_AR_sum[version] = 0
        for l_i in range(list_size):
            if nei_bu_xing_hao_5[l_i] == version:
                all_AR_sum[version] += AR_num_3[l_i]

    logger.info(all_income[nei_bu_xing_hao_set[0]])
    logger.info(all_AR_sum[nei_bu_xing_hao_set[0]])

    logger.info(all_AR_sum)
    logger.info(all_income)
