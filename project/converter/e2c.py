import xlrd
import os


"以下は, book(xlrd.book.Book)及びsheet(xlrd.sheet.Sheet)をロードするためのメソッド."


def open_book(book_name):
    """
    book_name(str: ファイル名 -> XLSファイル)をxlrd.book.Bookとして開く.
    """
    book = xlrd.open_workbook(book_name, formatting_info=True)
    return book


def get_sheet_names_from_book(book):
    """
    Bookの中の全てのsheetの名前(str)を返す.
    """
    sheet_names = book.sheet_names()
    return sheet_names


"以下は, sheetの結合セルを除去するための下処理メソッド."


def get_merged_cells(sheet):
    """
    Sheet(xlrd.sheet.Sheet)に対し, その中の*結合セル*の集合を返す.
    """
    return sheet.merged_cells


def get_merged_cells_value(sheet, row_index, col_index):
    """
    Sheetの特定の結合セルのインデックス(何行何列)に対し, その値を返す.
    """
    merged = get_merged_cells(sheet)
    for (rlow, rhigh, clow, chigh) in merged:
        if rlow <= row_index < rhigh:
            if clow <= col_index < chigh:
                cell_value = sheet.cell_value(rlow, clow)
                return cell_value
    return None


def get_unmerged_sheet(book_name, sheet_name):
    """
    Book(xlrd.book.Book)の特定sheetに対し, 全てのセルの結合を解除した上、
    結合が解除された空白セルに元の結合セルの値を充填してから, new_sheetとして返す.
    """
    book = open_book(book_name)
    sheet = book.sheet_by_name(sheet_name)
    rows_num = sheet.nrows
    cols_num = sheet.ncols
    new_sheet = []
    if rows_num > 1 and cols_num > 1:
        for r in range(rows_num):
            new_row = []
            for c in range(cols_num):
                cell_value = sheet.row_values(r)[c]
                if cell_value is None or cell_value == "":
                    cell_value = (get_merged_cells_value(sheet, r, c))
                if cell_value is None:
                    cell_value = ""
                new_row.append(cell_value)
            new_sheet.append(new_row)
        return new_sheet
    else:
        return None


"以下は, sheetの目視的可読性を上げるための下処理メソッド."


def replace_space_strings(row):
    """
    Row(list)の要素の中, 全ての*スペース*と*'\u3000'*を,
    *''*に入れ替えてからnew_row(list)を返す.
    """
    new_row = []
    for string in row:
        if string == '\u3000' or string == ' ':
            new_row.append('')
        else:
            new_row.append(string)
    return new_row


def get_no_space_cell_sheet(sheet):
    """
    Sheetの中, *スペース*と*'\u3000'*に充填されたセルを,
    *''*に入れ替えてからnew_sheet(xlrd.sheet.Sheet)を返す.
    """
    new_sheet = []
    for row in sheet:
        new_sheet.append(replace_space_strings(row))
    return new_sheet


def remove_blank_cells(row):
    """
    Row(list)の中, 全ての*''*を削除してから, new_row(list)として返す.
    """
    new_row = []
    for string in row:
        if string == '':
            pass
        else:
            new_row.append(string)
    return new_row


def get_no_blank_cell_sheet(sheet):
    """
    Sheetの中, 全ての空っぽcell(str: '')を削除してから, new_sheetとして返す.
    """
    new_sheet = []
    for row in sheet:
        new_sheet.append(remove_blank_cells(row))
    return new_sheet


def get_no_blank_row_sheet(sheet):
    """
    Sheetの中, 全ての空っぽrow(list: [])を削除してから, new_sheetとして返す.
    """
    new_sheet = []
    for row in sheet:
        if row == []:
            pass
        else:
            new_sheet.append(row)
    return new_sheet


"以下は, 変換後出力したcsvファイルのフォルダを作成するための下処理メソッド."


def get_csv_path_and_make_folder(abs_file_name):
    path = os.path.dirname(abs_file_name)
    file_name = os.path.basename(abs_file_name)
    csv_path = os.path.join(path, os.path.splitext(file_name)[0])
    if os.path.exists(csv_path):
        pass
    else:
        os.mkdir(csv_path)
    return csv_path


if __name__ == "__main__":
    test_sheet = get_unmerged_sheet("02-2.xls", "1")
    # test_sheet = get_unmerged_sheet("11509.xls", "8")
    test_sheet = get_no_space_cell_sheet(test_sheet)
    test_sheet = get_no_blank_cell_sheet(test_sheet)
    test_sheet = get_no_blank_row_sheet(test_sheet)
    for row in test_sheet:
        print(row)
