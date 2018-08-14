# -*- coding: utf-8 -*-

import xlrd

file = xlrd.open_workbook('2017年10月新能源利用小时.xlsx')
book = file.sheet_by_index(0)

def GetNum(row_name, col_name, excel_book):
# row_name = 德润风电场, col_name = 10月利用小时数
# excel_book = file.sheet_by_index(0)
# return = 德润风电场 在 10月的利用小时数
    row_name_loc = [0, 0]
    col_name_loc = [0, 0]
    target_loc = [0, 0]
    for row in range(excel_book.nrows):
        for col in range(len(excel_book.row_values(row))):
            if row_name == excel_book.row_values(row)[col]:
                row_name_loc = [row, col]
            if col_name == excel_book.row_values(row)[col]:
                col_name_loc = [row, col]
    
    target_loc = [row_name_loc[0], col_name_loc[1]]
    return excel_book.row_values(target_loc[0])[ target_loc[1] ]


if __name__ == '__main__':
    print( GetNum('红松风电场', '装机容量\n（兆瓦）', book) )

