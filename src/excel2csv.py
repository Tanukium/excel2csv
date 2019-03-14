import xlrd
import copy
import re
import os
import csv
import sys


def remove_space_strings(row):
    """
    とあるlistにおける全部のスペースと'\u3000'要素を空白の文字列に入れ替えるメソッド.
    :param row: (list) メソッドが応用するlist.
    :return: (list) 全部の'\u3000'要素が削除されたlist.
    """
    new_row = []
    for string in row:
        if string == '\u3000' or string == ' ':
            new_row.append('')
        else:
            new_row.append(string)
    return new_row


def remove_space_strings_in_elements(sheet):
    new_sheet = []
    for row in sheet:
        new_sheet.append(remove_space_strings(row))
    return new_sheet


def remove_blank_strings(row):
    """
    とあるlistにおける全部の空白の文字列（''）要素を削除するメソッド.
    :param row: (list) メソッドが応用するlist.
    :return: (list) 全部の空白文字列要素が削除されたlist.
    """
    new_row = []
    for string in row:
        if string == '':
            pass
        else:
            new_row.append(string)
    return new_row


def remove_blank_lists(sheet):
    """
    とあるlistにおける全部の空白のlist要素（[]）を削除するメソッド.
    :param sheet: (list) メソッドが応用するlist.
    :return: (list) 全部の空白list要素が削除されたlist.
    """
    new_sheet = []
    for row in sheet:
        if row == []:
            pass
        else:
            new_sheet.append(row)
    return new_sheet


def remove_lists_fulled_with_blank_strings(sheet):
    new_sheet = []
    for row in sheet:
        if any(row):
            new_sheet.append(row)
        else:
            pass
    return new_sheet


def remove_blank_strings_at_start_of_lists(sheet):
    result = []
    for row in sheet:
        num = 0
        for string in row:
            if string:
                break
            else:
                num += 1
        result.append(num)
    result.sort()
    for num in result:
        if num:
            result = num
            break
    new_sheet = []
    for row in sheet:
        new_sheet.append(row[result:])
    return new_sheet


def remove_repeat_elements(rows):
    """
    とあるlistにおける重複（2個以上存在する）の要素を, 一個だけにするメソッド.
    :param rows: (list) メソッドが応用するlist.
    :return: (list) 全部の重複の要素が削除されたlist.
    """
    new_rows = []
    for string in rows:
        if string in new_rows:
            pass
        else:
            new_rows.append(string)
    return new_rows


def data_pretreatment(sheet):
    sheet = remove_space_strings_in_elements(sheet)
    sheet = remove_blank_strings_at_start_of_lists(sheet)
    sheet = remove_lists_fulled_with_blank_strings(sheet)
    return sheet


def get_size_of_index_area(sheet):
    for r_index, row in enumerate(sheet):
        num = 0
        for cell in row:
            if isinstance(cell, float):
                num += 1
            if num >= 2:
                return r_index


def get_title_and_comment(sheet):
    sheet_copy = copy.deepcopy(sheet)
    new_sheet = []
    for row in sheet_copy:
        new_sheet.append(remove_blank_strings(row))
    result = []
    for row in new_sheet:
        if len(row) == 1:
            result.append(row[0])
    return result


def remove_rows_contained_title_and_comment(rows, titles):
    new_rows = []
    for row in rows:
        new_row = []
        for string in row:
            if string in titles:
                new_row.append("")
            else:
                new_row.append(string)
        new_rows.append(new_row)
    new_rows = remove_lists_fulled_with_blank_strings(new_rows)
    return new_rows


def make_list_vertical(rows, index):
    new_row = []
    for row in rows:
        new_row.append(row[index])
    return new_row


def get_index_lists(sheet, index_length):
    source = sheet[:index_length]
    index_list = []
    for i in range(len(source[0])):
        new_row = make_list_vertical(source, i)
        new_row = remove_repeat_elements(new_row)
        index_list.append(new_row)
    return index_list


def get_index_row_contained_strings(index_list):
    source = index_list
    index_row = []
    for cell in source:
        new_cell = []
        for element in cell:
            new_cell.append(str(element))
        s = '_'.join(new_cell)
        index_row.append(s)
        del s
    return index_row


def reformat_index_row(index_row):
    new_index_row = []
    for string in index_row:
        if string.find("\n") != -1:
            string = re.sub(r"\n", "_", string)
        if string.startswith("_"):
            string = string.lstrip("_")
        if string.endswith("_"):
            string = string.rstrip("_")
        new_index_row.append(string)
    return new_index_row


def get_content_lists(sheet, index_row, index_length):
    content_lists = [index_row]
    for row in sheet[index_length:]:
        content_lists.append(row)
    return content_lists


def make_uncover_csv_file(path, sheet_name):
    with open(path + "{}.csv".format(sheet_name), 'w',
              newline='', encoding='cp932', errors='ignore') as uncover_list:
        uncover_list.write(sheet_name)
    return None


def print_warning(sheet_name):
    print("-" * 8)
    print("警告： ワークシート {} は整形しませんでした。".format(sheet_name))
    print("　　　 データ構造か何かに原因があります。")
    return None


def get_merged_cells(sheet):
    return sheet.merged_cells


def get_merged_cells_value(sheet, row_index, col_index):
    merged = get_merged_cells(sheet)
    for (rlow, rhigh, clow, chigh) in merged:
        if row_index >= rlow and row_index < rhigh:
            if col_index >= clow and col_index < chigh:
                cell_value = sheet.cell_value(rlow, clow)
                return cell_value
    return None


def get_rows_from_sheet(book_name, sheet_name):
    book = xlrd.open_workbook(book_name, formatting_info=True)
    sheet = book.sheet_by_name(sheet_name)
    rows_num = sheet.nrows
    cols_num = sheet.ncols
    new_sheet = []
    if rows_num > 1 and cols_num > 1:
        for r in range(rows_num):
            new_row = []
            for c in range(cols_num):
                cell_value = sheet.row_values(r)[c]
                if cell_value is None or cell_value == '':
                    cell_value = (get_merged_cells_value(sheet, r, c))
                if cell_value is None:
                    cell_value = ""
                new_row.append(cell_value)
            new_sheet.append(new_row)
        return new_sheet
    else:
        return None


def get_sheet_names_from_book(book_name):
    book = xlrd.open_workbook(book_name, formatting_info=True)
    sheet_names = book.sheet_names()
    return sheet_names


class Excel2csv(object):
    def __init__(self, file_name):
        if file_name:
            self.book_name = os.path.abspath(file_name)
        else:
            raise RuntimeError('ファイル名はありません')

        self.file_name = os.path.basename(self.book_name)
        self.sheet_names = get_sheet_names_from_book(self.book_name)

    def get_csv_path_and_make_folder(self):
        folder_name = os.path.dirname(self.book_name) + os.sep
        csv_path = (folder_name +
                    os.path.splitext(self.file_name)[0] + os.sep)
        if os.path.exists(csv_path):
            pass
        else:
            os.mkdir(csv_path)
        return csv_path

    def get_content_lists_and_titles_from_book(self):
        csv_path = self.get_csv_path_and_make_folder()
        result = {}
        uncover_sheets = []
        for sheet_name in self.sheet_names:
            sheet = get_rows_from_sheet(self.book_name, sheet_name)
            try:
                sheet = data_pretreatment(sheet)
                titles = get_title_and_comment(sheet)
                sheet = remove_rows_contained_title_and_comment(sheet, titles)
                index_length = get_size_of_index_area(sheet)
                index_area = get_index_lists(sheet, index_length)
                index_row = get_index_row_contained_strings(index_area)
                index_row = reformat_index_row(index_row)
                content_lists = get_content_lists(sheet, index_row, index_length)
                result[sheet_name] = [content_lists, titles]
            except:
                print_warning(sheet_name)
                make_uncover_csv_file(csv_path, sheet_name)
                result[sheet_name] = None
                uncover_sheets.append(sheet_name)
            else:
                pass
        return result, uncover_sheets

    def output_csv_files(self):
        csv_path = self.get_csv_path_and_make_folder()
        csv_source, uncover_sheets = self.get_content_lists_and_titles_from_book()
        for sheet_name in self.sheet_names:
            csv_name = sheet_name + ".csv"
            if csv_source[sheet_name]:
                with open(csv_path + csv_name, 'w', newline='',
                          encoding='cp932', errors='ignore') as csv_file:
                    writer = csv.writer(csv_file, delimiter=',',
                                        quotechar='|', quoting=csv.QUOTE_MINIMAL)
                    output = csv_source[sheet_name][1]
                    if output:
                        writer.writerow(output[:1])
                        writer.writerow(output[1:])
                    else:
                        pass
                    output = csv_source[sheet_name][0]
                    for row in output:
                        writer.writerow(row)
            else:
                pass
        return uncover_sheets


def main():
    e2c = Excel2csv(sys.argv[1])
    # con = e2c.get_content_lists_and_titles_from_book()
    # print(con)
    uncover_sheet = e2c.output_csv_files()
    return None


if __name__ == "__main__":
    main()
