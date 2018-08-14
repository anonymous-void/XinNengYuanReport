# -*- coding: utf-8 -*-

import xlrd
import time
import pandas as pd
import numpy as np


def GetItemLoc(items, file_name):
# items = the name or list of items of what you want
# ret = a list of the location, each location is also a list, including x and y
    file_handler = xlrd.open_workbook(file_name, on_demand=True)
    excel_book = file_handler.sheet_by_index(0)
    ret = list()
    
    for row in range(excel_book.nrows):
        for col in range(len(excel_book.row_values(row))):
            if type(items) == type(list()):
                for idx, item in enumerate(items):
                    if item == excel_book.row_values(row)[col]:
                        ret.append([row, col])
            if type(items) == type(str()):
                if items == excel_book.row_values(row)[col]:
                    ret.append([row, col])
    file_handler.release_resources()
    del file_handler
    return ret   


def GetItemByLoc(loc, file_name):
    file_handler = xlrd.open_workbook(file_name, on_demand=True)
    excel_book = file_handler.sheet_by_index(0)
    
    ret = list()
    # if loc is a series of location pairs
    if type(loc[0]) == type(list()):
        for each_loc in loc:
            ret.append(excel_book.row_values(each_loc[0])[each_loc[1]])
    else:
        ret.append(excel_book.row_values(loc[0])[loc[1]])
    
    file_handler.release_resources()
    del file_handler
    return ret


def GetTable(sites, columns, file_name):
    row_loc_pair = GetItemLoc(sites, file_name)
    col_loc_pair = GetItemLoc(columns, file_name)
    
    df = pd.DataFrame(data=None, index=None, columns=columns)
    
    for row_loc in row_loc_pair:
        row_cache = list()
        for col_loc in col_loc_pair:
            row_cache.append( GetItemByLoc([row_loc[0], col_loc[1]], file_name)[0] )
        df_tmp = pd.DataFrame(data=[row_cache], index=GetItemByLoc(row_loc, file_name),\
                              columns=columns)
        df = pd.concat([df, df_tmp])
    return df



if __name__ == '__main__':
    file_name = '2017年10月新能源利用小时.xlsx'
    row_name = ['红松风电场', '吉风风电场', '聚风风电场', '承德平均', '尚义平均', '全网']
    col_name = ['装机容量\n（兆瓦）', '10月利用小时数', '年累计利用小时数']
    mydf = GetTable(row_name, col_name, file_name)
    print(mydf)
    