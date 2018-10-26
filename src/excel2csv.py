#!/anaconda3/envs/msemi/bin/python
# -*- coding: UTF-8 -*-

import sys
import os
import csv
import xlrd


class Excel2csv(object):

    def __init__(self, file_name):
        if len(sys.argv) > 1 and sys.argv[1]:
            self.file_path = os.path.abspath(sys.argv[1])
        elif file_name:
            self.file_path = os.path.abspath(file_name)
        else:
            raise RuntimeError('No path or filename')
        self.file_name = os.path.basename(self.file_path)

    def make_csv_path(self):
        file_dirname = os.path.dirname(self.file_path) + os.sep
        csv_path = file_dirname + os.path.splitext(self.file_name)[0] + os.sep
        if not os.path.exists(csv_path):
            os.mkdir(csv_path)
        return csv_path

    def csv_from_sheet(self, book_name, sheet_name):
        sheet = book_name.sheet_by_name(sheet_name)
        csv_name = sheet_name + '.csv'
        with open(self.make_csv_path() + csv_name, 'w',
                  newline='', encoding='cp932') as csv_file:
            writer = csv.writer(csv_file, delimiter=',', 
                                quotechar='|', quoting=csv.QUOTE_MINIMAL)
            for row_num in range(sheet.nrows):
                writer.writerow(sheet.row_values(row_num))

    def csv_from_excel(self):
        book = xlrd.open_workbook(self.file_path)
        sheet_names = book.sheet_names()
        for sheet_name in sheet_names:
            self.csv_from_sheet(book, sheet_name)


if __name__ == '__main__':
    excel2csv = Excel2csv(None)
    excel2csv.csv_from_excel()
