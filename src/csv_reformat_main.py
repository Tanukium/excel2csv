#!/anaconda3/envs/msemi/bin/python
# -*- coding: UTF-8 -*-

import csv
import os
import sys
from csv_reformat import DataLength

table = []
with open(os.path.abspath(sys.argv[1]),
          'r', encoding='cp932') as csv_file:
    reader = csv.reader(csv_file)
    for row in reader:
        table.append(row)
data_length = DataLength(sys.argv[1])


def reformat_comma_int(data):
    try:
        return int(data.replace(',', ''))
    except ValueError:
        return data


def reformat_int_in_data(arr):
    count = 0
    while count < len(arr):
        poped = arr.pop(0)
        arr.append(reformat_comma_int(poped))
        count += 1


def del_repeat(arr):
    for item in arr:
        while arr.count(item) > 1:
            del arr[arr.index(item)]
    return arr


def del_brank_str(arr):
    for item in arr:
        if item == '':
            del arr[arr.index(item)]
    return arr


def pick_index_row_up(arr):
    global row_num
    rows = []
    for row in arr:
        if type(row[0]) == int:
            row_num = arr.index(row)
            break
        index = del_brank_str(del_repeat(row))
        if index:
            rows.append(index)
    return rows, row_num


if __name__ == '__main__':
    datas = []
    for row in table:
        data_length_in_row = len(row) - data_length.data_length
        datas.append(row[data_length_in_row:])
    for row in datas:
        reformat_int_in_data(row)
    index_rows, index_num = pick_index_row_up(datas)
    for row in index_rows:
        print(row)
    datas = datas[index_num:]
    for row in datas:
        print(row)
