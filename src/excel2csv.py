# -*- coding: UTF-8 -*-

import sys
import os
import csv
import xlrd
from collections import Counter as CoT


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


def _del_repeat(lst):
    for item in lst:
        while lst.count(item) > 1:
            del lst[lst.index(item)]
    return lst


def _del_blank_mbr(lst):
    for item in lst:
        try:
            while not item:
                del lst[lst.index(item)]
        except ValueError:
            break
    return lst


class Excel2csv(object):

    def __init__(self, file_name):
        if len(sys.argv) > 1 and sys.argv[1]:
            self.file_path = os.path.abspath(sys.argv[1])
        elif file_name:
            self.file_path = os.path.abspath(file_name)
        else:
            raise RuntimeError('No path or filename')

        self.file_name = os.path.basename(self.file_path)

    def _get_data_from_excel(self, index):
        row_data, col_data = [], []
        book = xlrd.open_workbook(self.file_path)
        sheet_names = book.sheet_names()
        _get_data_from_sheet(book, sheet_names[index],
                             "row", row_data)
        _get_data_from_sheet(book, sheet_names[index],
                             "col", col_data)
        return row_data, col_data

    def _count_row_data_length(self, col_index):
        if not self.index_row_amount:
            row_data_lengths = []
            if self.col_list and self.col_amount:
                data = self.col_list[:]
                for index in range(self.col_amount):
                    row_data_lengths.append(
                        _count_aline_data_length(data, index))
                self.col_list = self._get_data_from_excel(col_index)[1]
                result = CoT(row_data_lengths).most_common()
                if not result[0][0]:
                    return result[1][0]
                else:
                    return result[0][0]
            else:
                print("No col data or col length in the obj!")

    def _count_col_data_length(self, row_index):
        if not self.index_col_amount:
            col_data_lengths = []
            if self.row_list and self.row_amount:
                data = self.row_list[:]
                for index in range(self.row_amount):
                    col_data_lengths.append(
                        _count_aline_data_length(data, index))
                self.row_list = self._get_data_from_excel(row_index)[0]
                result = CoT(col_data_lengths).most_common()
                if not result[0][0]:
                    return result[1][0]
                else:
                    return result[0][0]
            else:
                print("No row data or row length in the obj!")

    def _check_title_cell(self):
        if self.col_list:
            _del_blank_mbr(_del_repeat(self.col_list[0]))
            title = ""
            if not self.col_list[0]:
                _del_blank_mbr(self.row_list[0])
                if len(self.row_list[0]) == 1 and self.row_list[0][0] != '':
                    title = self.row_list[0][0]
            elif self.col_list[0][0] == self.row_list[0][0]:
                title = self.row_list[0][0]
            else:
                title = "No title"
            self.title = title
        else:
            title = "No title"
            self.title = title

    def _get_index_rows(self):
        cols, blank_col_amount_list = [], []
        for index in range(self.index_col_amount, self.col_amount):
            cols.append(self.col_list[index][0:self.index_row_amount])
        for col in cols:
            for item in col:
                if item == self.title:
                    col[col.index(item)] = ''
        for col in cols:
            blank_col_count = 0
            for item in col:
                if not item:
                    blank_col_count += 1
                else:
                    blank_col_amount_list.append(blank_col_count)
                    break
        blank_col_amount_list = _del_repeat(blank_col_amount_list)
        blank_col_amount_list.sort()
        for col in cols:
            if min(blank_col_amount_list) != 0:
                del col[:min(blank_col_amount_list)]
            else:
                for index in range(blank_col_amount_list[1]):
                    if not col[index]:
                        col.pop(index)
            while not col[-1]:
                col.pop()
                if not col:
                    break
        cols = _del_blank_mbr(cols)
        for col in cols:
            for index in range(len(col)):
                if not col[index]:
                    col[index] = cols[cols.index(col) - 1][index]
                else:
                    break
        lst = []
        for col in cols:
            col_new = []
            for item in col:
                if type(item) is float:
                    col_new.append(str(item))
                else:
                    col_new.append(item)
            s = "_".join(col_new)
            lst.append(s)
            del s, col_new
        return lst

    def _get_index_cols(self):
        rows, blank_row_amount_list = [], []
        for row in self.row_list:
            rows.append(row[:self.index_col_amount])
        for row in rows:
            for item in row:
                if item == self.title:
                    row[row.index(item)] = ''
        for row in rows:
            blank_row_count = 0
            for item in row:
                if not item:
                    blank_row_count += 1
                else:
                    blank_row_amount_list.append(blank_row_count)
                    break
        blank_row_amount_list = _del_repeat(blank_row_amount_list)
        blank_row_amount_list.sort()
        for row in rows:
            if min(blank_row_amount_list) != 0:
                del row[:min(blank_row_amount_list)]
            else:
                for index in range(blank_row_amount_list[1]):
                    if not row[index]:
                        row.pop(index)
            while not row[-1]:
                row.pop()
                if not row:
                    break
        rows = _del_blank_mbr(rows)
        for row in rows:
            for index in range(len(row)):
                if not row[index]:
                    row[index] = rows[rows.index(row) - 1][index]
                else:
                    break
        for row in rows:
            _del_blank_mbr(row)
        lst = []
        for row in rows:
            row_new = []
            for item in row:
                if type(item) is float:
                    row_new.append(str(item))
                else:
                    row_new.append(item)
            s = "_".join(row_new)
            lst.append(s)
            del s, row_new
        # diff = self.row_amount - self.index_row_amount - len(lst)
        # if diff != 0:
        #     for i in range(diff):
        #        lst = [""] + lst
        return lst

    def _get_data_rows(self):
        data_rows = []
        for index in range(self.index_row_amount, self.row_amount):
            data_rows.append(self.row_list[index][self.index_col_amount:])
        return data_rows

    def _make_btf_sheet(self):
        index_row = self._get_index_rows()  # [str * n]
        index_col = self._get_index_cols()  # [str * n]
        data_rows = [index_row] + self._get_data_rows()  # [list * n]
        data_container = []
        for row in data_rows:
            if index_col:
                for item in index_col:
                    row = [item] + row
                    data_container.append(row)
                    index_col.pop(0)
                    break
            elif data_rows:
                row = [""] + row
                data_container.append(row)
        return data_container

    def _re_init(self, index):
        self.row_list, self.col_list = self._get_data_from_excel(index)
        self._check_title_cell()
        self.row_list, self.col_list = self._get_data_from_excel(index)

        self.row_amount = len(self.row_list)
        self.col_amount = len(self.col_list)

        self.index_row_amount, self.index_col_amount = None, None
        (self.index_row_amount,
         self.index_col_amount) = (self.row_amount -
                                   self._count_row_data_length(index),
                                   self.col_amount -
                                   self._count_col_data_length(index))

    def _make_csv_path(self):
        file_dirname = os.path.dirname(self.file_path) + os.sep
        csv_path = (file_dirname +
                    os.path.splitext(self.file_name)[0] + os.sep)
        if not os.path.exists(csv_path):
            os.mkdir(csv_path)
        return csv_path

    def _csv_from_sheet(self, book_name, sheet_name, index):
        self.row_list, self.col_list = self._get_data_from_excel(index)
        self._check_title_cell()
        csv_name = self.title + "_" + sheet_name + ".csv"
        sheet = book_name.sheet_by_name(sheet_name)
        with open(self._make_csv_path() + csv_name, 'w',
                  newline='', encoding='cp932') as csv_file:
            writer = csv.writer(csv_file, delimiter=',',
                                quotechar='|',
                                quoting=csv.QUOTE_MINIMAL)
            if sheet.nrows > 1 and sheet.ncols > 1:
                self._re_init(index)
                rows = self._make_btf_sheet()
                for i in range(len(rows)):
                    writer.writerow(rows[i])
            elif sheet.nrows > 1 and sheet.ncols == 1:
                for i in range(sheet.nrows):
                    writer.writerow(sheet.row_values(i))
            elif sheet.nrows == 1 and sheet.ncols > 1:
                writer.writerow(sheet.row_values(0))
            elif not sheet.nrows:
                writer.writerow([])

    def csv_from_excel(self):
        book = xlrd.open_workbook(self.file_path)
        sheet_names = book.sheet_names()
        for sheet_name in sheet_names:
            self._csv_from_sheet(book, sheet_name, sheet_names.index(sheet_name))


def main():
    e2c = Excel2csv(None)

    """
    e2c._re_init(0)

    print("このテーブルには", e2c.row_amount, "行があり、",
          e2c.col_amount, "列がある.")
    print("上から", e2c.index_row_amount, "番目までの行と、",
          e2c.index_col_amount, "番目までの列はインデックス.")

    print("インデックス行の内容は、", e2c._get_index_rows())
    print("インデックス行には、", len(e2c._get_index_rows()), "個セルがある.")

    print("インデックス列の内容は、", e2c._get_index_cols())
    print("インデックス列には、", len(e2c._get_index_cols()), "個セルがある.")
    print("このテーブルのタイトルは、", e2c.title, "です.")

    print(e2c._make_btf_sheet())
    """

    e2c.csv_from_excel()


if __name__ == '__main__':
    main()
