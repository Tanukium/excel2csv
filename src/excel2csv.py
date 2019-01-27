# -*- coding: UTF-8 -*-

import sys
import os
import csv
import xlrd
import copy
import re
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
    num = 0
    data_len = len(data[index])
    data_copy = copy.deepcopy(data[index])
    for item in data_copy:
        if type(item) is not float:
            num += 1
        else:
            if type(data_copy[data_copy.index(item) + 1]) is float:
                break
            else:
                num += 1
    return data_len - num


def _del_repeat(lst):
    for item in lst:
        while lst.count(item) > 1:
            del lst[lst.index(item)]
    return lst


def _u3000_killer(lst):
    while '\u3000' in lst:
        lst[lst.index('\u3000')] = ''
    return lst


def _del_blank_cell(lst):
    while '' in lst:
        lst.remove('')
    return lst


def _del_blank_list(lst):
    while [] in lst:
        lst.remove([])
    return lst


def _del_blank_cell_at_start(lst):
    lst_copy = copy.deepcopy(lst)
    result = []
    for item in lst_copy:
        num = 0
        while not item[0]:
            item.remove('')
            num += 1
        result.append(num)
    result.sort()
    if not min(result):
        result = result[1]
    else:
        result = min(result)
    for item in lst:
        del item[:result]
    return lst


class Excel2csv(object):

    def __init__(self, file_name):
        if file_name:
            self.file_path = os.path.abspath(file_name)
        elif len(sys.argv) > 1 and sys.argv[1]:
            self.file_path = os.path.abspath(sys.argv[1])
        else:
            raise RuntimeError('No path or filename')
        self.file_name = os.path.basename(self.file_path)

        self.result = self._get_data_from_excel()
        self.comment, self.data = {}, {}
        self.index_row, self.index_col = {}, {}

    def _get_data_from_excel(self):
        result = {}
        book = xlrd.open_workbook(self.file_path)
        sheet_names = book.sheet_names()
        for index in range(len(sheet_names)):
            rows, cols = [], []
            _get_data_from_sheet(book, sheet_names[index],
                                 "row", rows)
            _get_data_from_sheet(book, sheet_names[index],
                                 "col", cols)
            result[sheet_names[index]] = [rows, cols]
        return result

    def _get_title_from_sheet(self, sheet_name):
        if self.result[sheet_name]:
            row_result, col_result = [], []

            rows = self.result[sheet_name][0]
            for row in rows:
                _u3000_killer(row)
            for row in rows:
                if not any(row):
                    rows.remove(row)
            rows_copy = copy.deepcopy(rows)
            for item in rows_copy:
                _del_blank_cell(item)
            for row in rows_copy:
                if len(row) == 1:
                    row_result.append(row[0])
                elif len(row) == 2:
                    for item in row:
                        row_result.append(item)
                else:
                    break
            for item in row_result:
                for row in rows:
                    if item in row:
                        rows.remove(row)

            cols = self.result[sheet_name][1]
            for col in cols:
                _u3000_killer(col)
            for col in cols:
                if not any(col):
                    cols.remove(col)
            cols_copy = copy.deepcopy(cols)
            for item in cols_copy:
                _del_blank_cell(item)
            for col in cols_copy:
                if len(col) == 1:
                    col_result.append(col[0])
                else:
                    break
            for item in col_result:
                for col in cols:
                    if item in col:
                        cols.remove(col)

            if sheet_name not in self.comment:
                self.comment[sheet_name] = []
                for item in row_result:
                    self.comment[sheet_name].append(item)
                for item in col_result:
                    self.comment[sheet_name].append(item)

            for col in cols:
                for member in self.comment[sheet_name]:
                    if member in col:
                        col[col.index(member)] = ''

            _del_blank_cell_at_start(rows)
            _del_blank_cell_at_start(cols)

            return _del_repeat(row_result + col_result)

    def _get_comment_from_end_of_sheet(self, sheet_name):
        if self.result[sheet_name]:
            row_result = []
            rows = self.result[sheet_name][0]
            for row in rows:
                if not any(row):
                    rows.remove(row)
            rows_copy = copy.deepcopy(rows)
            for item in rows_copy:
                _del_blank_cell(item)
            for row in rows_copy:
                if (len(row) == 1 and
                        len(rows_copy[rows_copy.index(row)]) > 1):
                    row_result.append(row[0])
            for item in row_result:
                for row in rows:
                    if item in row:
                        rows.remove(row)
            if sheet_name in self.comment:
                for item in row_result:
                    self.comment[sheet_name].append(item)
                _del_repeat(self.comment[sheet_name])
            return row_result

    def _count_row_data_length(self, sheet_name):
        result = []
        rows = self.result[sheet_name][0]
        rows_copy = copy.deepcopy(rows)
        if rows_copy:
            for index in range(len(rows_copy)):
                result.append(_count_aline_data_length(rows_copy,
                                                       index))
            result = CoT(result).most_common()
            if not result[0][0] and result[0][1] == result[1][1]:
                return result[1][0]
            else:
                return result[0][0]

    def _count_col_data_length(self, sheet_name):
        result = []
        cols = self.result[sheet_name][1]
        cols_copy = copy.deepcopy(cols)
        if cols_copy:
            for index in range(len(cols_copy)):
                result.append(_count_aline_data_length(cols_copy,
                                                       index))
            result = CoT(result).most_common()
            if not result[0][0] and result[0][1] == result[1][1]:
                return result[1][0]
            else:
                return result[0][0]

    def _make_index_rows(self, sheet_name):
        container = []
        rows = self.result[sheet_name][0]
        rows_copy = copy.deepcopy(rows)
        rows_len = len(rows_copy[0])
        row_data_len = self._count_row_data_length(sheet_name)
        row_index_len = rows_len - row_data_len
        for row in rows_copy:
            container.append(row[:row_index_len])
        container = list(filter(any, container))
        _del_blank_cell_at_start(container)
        for row in container:
            for index in range(len(row)):
                if not row[index]:
                    row[index] = container[container.index(row) - 1][index]
                else:
                    break
        con_copy = copy.deepcopy(container)
        del container
        for row in con_copy:
            if '\u3000' in row:
                row.remove('\u3000')
        container = list(filter(any, con_copy))
        tmp2, result = [], []
        for row in container:
            tmp = []
            for item in row:
                if type(item) is float:
                    tmp.append(str(int(item)))
                else:
                    tmp.append(item)
            s = "_".join(tmp)
            tmp2.append(s)
        for item in tmp2:
            if item.find("\n") != -1:
                item = re.sub(r"\n", " ", item)
            if item.endswith("_"):
                result.append(item.rstrip("_"))
            elif item.startswith("_"):
                result.append(item.lstrip("_"))
            else:
                result.append(item)
        if sheet_name not in self.index_row:
            self.index_row[sheet_name] = result
        return result

    def _make_index_cols(self, sheet_name):
        container = []
        cols = self.result[sheet_name][1]
        cols_copy = copy.deepcopy(cols)
        cols_len = len(cols_copy[0])
        col_data_len = self._count_col_data_length(sheet_name)
        col_index_len = cols_len - col_data_len
        for col in cols_copy:
            container.append(col[:col_index_len])
        container = list(filter(any, container))
        for col in container:
            for member in self.comment[sheet_name]:
                if member in col:
                    col[col.index(member)] = ''
        _del_blank_cell_at_start(container)
        for col in container:
            for index in range(len(col)):
                if not col[index]:
                    col[index] = container[container.index(col) - 1][index]
                else:
                    break
        con_copy = copy.deepcopy(container)
        del container
        for col in con_copy:
            if '\u3000' in col:
                col.remove('\u3000')
        container = list(filter(any, con_copy))
        tmp2, result = [], []
        for col in container:
            tmp = []
            for item in col:
                if type(item) is float:
                    tmp.append(str(int(item)))
                else:
                    tmp.append(item)
            s = "_".join(tmp)
            tmp2.append(s)
        for item in tmp2:
            if item.find("\n") != -1:
                item = re.sub(r"\n", " ", item)
            if item.endswith("_"):
                result.append(item.rstrip("_"))
            elif item.startswith("_"):
                result.append(item.lstrip("_"))
            else:
                result.append(item)
        if sheet_name not in self.index_col:
            self.index_col[sheet_name] = result
        return result

    def _make_data_rows(self, sheet_name):
        container = []
        rows = self.result[sheet_name][0]
        cols = self.result[sheet_name][1]
        rows_len = len(rows[0])
        cols_len = len(cols[0])
        for row in rows[(cols_len -
                         self._count_col_data_length(sheet_name)):]:
            container.append(row[(rows_len -
                                  self._count_row_data_length(sheet_name)):])
        if sheet_name not in self.data:
            self.data[sheet_name] = container
        return container

    def _make_csv_path(self):
        file_dir_name = os.path.dirname(self.file_path) + os.sep
        csv_path = (file_dir_name +
                    os.path.splitext(self.file_name)[0] + os.sep)
        if not os.path.exists(csv_path):
            os.mkdir(csv_path)
        return csv_path

    def _make_uncover_list(self, item):
        with open(self._make_csv_path() + "uncovered.txt", 'a+',
                  newline='', encoding='utf-8') as uncover_list:
            uncover_list.write(item + "\n")

    def _make_btf_sheet(self):
        self._get_data_from_excel()
        container, zero_group, one_group = {}, [], []
        for key in self.result:
            container[key] = []
            for item in self.result[key]:
                container[key].append(len(item))
        print(container)
        for key in container:
            if not container[key][0] or not container[key][1]:
                zero_group.append(key)
            elif container[key][0] == 1 or container[key][1] == 1:
                one_group.append(key)
        for key in self.result:
            if key in zero_group or key in one_group:
                self._make_uncover_list(key)
            else:
                self._get_title_from_sheet(key)
                self._get_comment_from_end_of_sheet(key)
                self._make_index_rows(key)
                self._make_index_cols(key)
                self._make_data_rows(key)
                print(key,
                      self.comment[key],
                      self.index_row[key],
                      self.index_col[key],
                      self.data[key], sep="\n")

    def _csv_from_sheet(self, sheet_name):
        csv_name = sheet_name + ".csv"
        with open(self._make_csv_path() + csv_name, 'w',
                  newline='', encoding='cp932') as csv_file:
            writer = csv.writer(csv_file, delimiter=',',
                                quotechar='|',
                                quoting=csv.QUOTE_MINIMAL)
            writer.writerow(self.comment[sheet_name][0:1])
            writer.writerow(self.comment[sheet_name][1:])
            writer.writerow(self.index_col[sheet_name])
            print(len(self.index_row[sheet_name]),
                  len(self.data[sheet_name]))
            if (len(self.index_row[sheet_name])
                    > len(self.data[sheet_name])):
                for index in range(len(self.data[sheet_name])):
                    self.data[sheet_name][index].insert(0,
                                                        self.index_row[sheet_name][index + 1])
                for item in self.data[sheet_name]:
                    writer.writerow(item)
            elif (len(self.index_row[sheet_name])
                  == len(self.data[sheet_name])):
                if self.index_row[sheet_name][0] != self.index_col[sheet_name][0]:
                    for index in range(len(self.data[sheet_name])):
                        self.data[sheet_name][index].insert(0,
                                                            self.index_row[sheet_name][index])
                else:
                    for index in range(len(self.data[sheet_name])):
                        if index != len(self.data[sheet_name]) - 1:
                            self.data[sheet_name][index].insert(0,
                                                                self.index_row[sheet_name][index + 1])
                        else:
                            break
                for item in self.data[sheet_name]:
                    writer.writerow(item)

    def csv_from_excel(self):
        self._make_btf_sheet()
        for key in self.comment:
            self._csv_from_sheet(key)


def main():
    e2c = Excel2csv(None)
    e2c.csv_from_excel()


if __name__ == '__main__':
    main()
