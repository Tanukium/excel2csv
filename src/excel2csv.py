# -*- coding: UTF-8 -*-

import sys
import os
import csv
import xlrd
from collections import Counter as ct


def _get_data_from_sheet(book_name, sheet_name,
                         switch, container):
    sheet = book_name.sheet_by_name(sheet_name)
    data_range = sheet.nrows if switch == "row" else sheet.ncols
    for index in range(data_range):
        if switch == "row":
            new_data = sheet.row_values(index)
        else:
            new_data = sheet.col_values(index)
        container.append(new_data)
    return container


def _count_aline_data_length(data, index):
    pop_num = 0
    data_length = len(data[index])
    while type(data[index][0]) is not float:
        data[index].pop(0)
        pop_num += 1
        if not data[index]:
            break
    return data_length - pop_num


class Excel2csv(object):
    """Exchanging .xls/.xlsx files to .csv file.

    The object of this class will contain the name and absolute
    path of a excel(.xls/.xlsx) file, and it has some methods to
    read the excel file by line, then save it with csv format.

    Attributes:
        file_path: the path of .xls/.xlsx file.
        file_name: the name of .xls/.xlsx file.

    """

    def __init__(self, file_name):
        """Reading the name and absolute path of excel file.

        For testing, the __init__ method also provide a command-
        line usage like `$ python excel2csv.py hogehoge.xls`.

        Args:
            file_name (str, optional): The name of excel file.
                Should not write it in the *command-line usage*.

        """
        if len(sys.argv) > 1 and sys.argv[1]:
            self.file_path = os.path.abspath(sys.argv[1])
        elif file_name:
            self.file_path = os.path.abspath(file_name)
        else:
            raise RuntimeError('No path or filename')

        self.file_name = os.path.basename(self.file_path)
        self.row_data, self.col_data = self._get_data_from_excel()
        self.row_length = len(self.col_data)
        self.col_length = len(self.row_data)
        self.row_index_length, self.col_index_length = None, None

        (self.row_index_length,
         self.col_index_length) = (self.row_length -
                                   self._count_row_data_length(),
                                   self.col_length -
                                   self._count_col_data_length())

    def _make_csv_path(self):
        """Making the path for saving csv file and return it.

        For saving csv files, this method will make a folder named
        with the name of excel file under that file's folder, and
        return the absolute path of the new folder.

        Returns:
            csv_path (str): the absolute path which csv file
                would be saved in.

        """
        file_dirname = os.path.dirname(self.file_path) + os.sep
        csv_path = (file_dirname +
                    os.path.splitext(self.file_name)[0] + os.sep)
        if not os.path.exists(csv_path):
            os.mkdir(csv_path)
        return csv_path

    def _get_data_from_excel(self):
        row_data, col_data = [], []
        book = xlrd.open_workbook(self.file_path)
        sheet_names = book.sheet_names()
        for sheet_name in sheet_names:
            _get_data_from_sheet(book, sheet_name,
                                 "row", row_data)
            _get_data_from_sheet(book, sheet_name,
                                 "col", col_data)
        return row_data, col_data

    def _count_row_data_length(self):
        if not self.row_index_length:
            row_data_lengths = []
            if self.row_data and self.col_length:
                data = self.row_data[:]
                for index in range(self.col_length):
                    row_data_lengths.append(
                        _count_aline_data_length(data, index))
                self.row_data = self._get_data_from_excel()[0]
                return ct(row_data_lengths).most_common()[0][0]
            else:
                print("No row data or row length in the obj!")

    def _count_col_data_length(self):
        if not self.col_index_length:
            col_data_lengths = []
            if self.col_data and self.row_length:
                data = self.col_data[:]
                for index in range(self.row_length):
                    col_data_lengths.append(
                        _count_aline_data_length(data, index))
                self.col_data = self._get_data_from_excel()[1]
                return ct(col_data_lengths).most_common()[0][0]
            else:
                print("No col data or col length in the obj!")

    def _get_index_rows(self):
        rows = []
        for index in range(self.row_index_length):
            rows.append(self.row_data[index])
        return rows

    def _csv_from_sheet(self, book_name, sheet_name):
        """Writing contents in the new csv file and save it.

        This method will read contents from a sheet in the excel file,
        then write it in a new csv file by line and save it.

        Params:
            book_name (str): The workbook's name from a excel file.
            sheet_name (str): The name of a sheet in a workbook.

        """
        sheet = book_name.sheet_by_name(sheet_name)
        csv_name = sheet_name + '.csv'
        with open(self._make_csv_path() + csv_name, 'w',
                  newline='', encoding='cp932') as csv_file:
            writer = csv.writer(csv_file, delimiter=',',
                                quotechar='|',
                                quoting=csv.QUOTE_MINIMAL)
            for row_num in range(sheet.nrows):
                writer.writerow(sheet.row_values(row_num))

    def csv_from_excel(self):
        book = xlrd.open_workbook(self.file_path)
        sheet_names = book.sheet_names()
        for sheet_name in sheet_names:
            self._csv_from_sheet(book, sheet_name)


def main():
    e2c = Excel2csv(None)
    print(e2c.row_length, e2c.col_length)
    print(e2c.row_index_length, e2c.col_index_length)
    print(e2c._get_index_rows())


if __name__ == '__main__':
    main()
