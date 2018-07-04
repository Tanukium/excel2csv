#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import pandas as pd
import sys
import re
import os


class Excel2csv(object):

    def __init__(self, path, file_name):
        if len(sys.argv) > 1 and sys.argv[1]:
            self.fileName = sys.argv[1]
            self.path = os.path.dirname(os.path.abspath(sys.argv[1])) + '\\'
        elif file_name:
            self.fileName = file_name
            self.path = os.path.dirname(os.path.abspath(file_name)) + '\\'
        else:
            raise RuntimeError('No path or filename')

    def convert(self):
        # xls book Open (xls, xlsxのどちらでも可能)
        input_book = pd.ExcelFile(self.fileName)
        # sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
        input_sheet_name = input_book.sheet_names
        for sheet_name in input_sheet_name:
            data_xls = pd.read_excel(self.fileName, sheet_name=sheet_name)
            csv_name = self.path + sheet_name + ".csv"
            data_xls.to_csv(csv_name, encoding='mskanji')
            with open(csv_name, 'r') as f:
                data = f.read()
                data = re.sub(r'Unnamed: .*', '', data, count=0)
            with open(csv_name, 'w') as f:
                f.write(data)


if __name__ == '__main__':
    excel2csv = Excel2csv(sys.argv[1], None)
    excel2csv.convert()
