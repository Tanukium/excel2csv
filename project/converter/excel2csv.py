import xlrd
import os
import csv
import shutil
import sys


"以下は, book(xlrd.book.Book)及びsheet(xlrd.sheet.Sheet)をロードするためのメソッド."


def open_book(book_name):
    """
    book_name(str: ファイル名 -> XLSファイル)をxlrd.book.Bookとして開く.
    """
    book = xlrd.open_workbook(book_name, formatting_info=True)
    return book


def get_sheet_names_from_book(book_name):
    """
    book(xlrd.book.Book)の中の全ての,
    sheet(xlrd.book.Book)の名前(str)の集合をsheet_names(list)として返す.
    """
    book = open_book(book_name)
    sheet_names = book.sheet_names()
    return sheet_names


"以下は, sheetの結合セルを除去するための下処理メソッド."


def get_merged_cells(sheet):
    """
    sheetに対し, その中の*結合セル*の集合を返す.
    """
    return sheet.merged_cells


def get_merged_cells_value(sheet, row_index, col_index):
    """
    sheetの特定の結合セルのインデックス(何行何列)に対し, その値を返す.
    結合セルでない, 値がない場合, Noneを返す.
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
    bookの特定sheetに対し, 全てのセルの結合を解除した上、
    結合が解除された空白セルに元の結合セルの値を充填してから, new_sheet(list: -> list)として返す.
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
        raise ValueError("シート {} がCSVに変換できませんでした.".format(sheet_name))


"以下は, 出力するCSVファイルの目視的可読性を上げるための下処理メソッド."


def replace_space_strings(row):
    """
    row(list)の要素の中, 全ての*スペース*と*'\u3000'*を,
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
    sheet(list: -> list)の中, *スペース*と*'\u3000'*に充填されたセルを,
    *''*に入れ替えてからnew_sheetを返す.
    """
    new_sheet = []
    for row in sheet:
        new_sheet.append(replace_space_strings(row))
    return new_sheet


def remove_blank_cells(row):
    """
    row(list)の全ての*''*を削除してから, new_row(list)として返す.
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
    sheetの全てのrow(list)要素に対して, 以下の走査を行う.
    1. もし中の全てのcell(str)要素が空っぽ文字列(str: -> '')であれば, 全ての空っぽ文字列を削除してから,
    要素なしになったrowをnew_sheet(list)に入れる;
    2. もし''でないcell要素があれば, そのrowをそのままnew_sheetに入れる.
    走査完了後, new_sheetを返す.
    """
    new_sheet = []
    for row in sheet:
        is_blank = True
        for cell in row:
            if cell != "":
                is_blank = False
                break
        if is_blank:
            new_sheet.append(remove_blank_cells(row))
        else:
            new_sheet.append(row)
    return new_sheet


def get_no_blank_row_sheet(sheet):
    """
    sheetの全ての空っぽrow(list: [])を削除してから, new_sheetとして返す.
    """
    new_sheet = []
    for row in sheet:
        if not row:
            pass
        else:
            new_sheet.append(row)
    return new_sheet


def sheet_pretreatment(book_name, sheet_name):
    """
    sheet_name(str: sheet(xlrd.sheet.Sheet)の名前)に対し, その名前のsheetを
    結合セル解除 -> スペース入替 -> 空っぽセル削除 -> 空っぽ行削除をしてから, sheetを返す.
    上のメソッドを順番通り呼び出すだけ.
    """
    sheet = get_unmerged_sheet(book_name, sheet_name)
    sheet = get_no_space_cell_sheet(sheet)
    sheet = get_no_blank_cell_sheet(sheet)
    sheet = get_no_blank_row_sheet(sheet)
    return sheet


"以下は, 変換後出力したCSVファイルを入れるフォルダを作成するための下処理メソッド."


def get_csv_path_and_make_csv_folder(abs_file_name):
    """
    abs_file_name(str: 絶対パスを含めるXLSファイル名)に対し,
    XLSファイルの所在フォルダに, 拡張子が含まないXLSファイル名でフォルダを作成し,
    そのパスをcsv_path(str)として返す.
    """
    path = os.path.dirname(abs_file_name)
    file_name = os.path.basename(abs_file_name)
    csv_path = os.path.join(path, os.path.splitext(file_name)[0])
    if os.path.exists(csv_path):
        pass
    else:
        os.mkdir(csv_path)
    return csv_path


"""
以下は, 写真かグラフが入ったためCSVに変換できないシートを例外処理したり,
シェルで変換するとき警告をstdoutに出力したりするためのメソッド.
"""


def output_unconverted_csv_file(path, sheet_name):
    """
    sheet_name(str: CSVに変換できないsheetの名前)に対し, sheet_name.csvを作成し,
    その中にsheet_nameを書き込んでから保存するメソッド.
    """
    with open(os.path.join(path, "{}.csv".format(sheet_name)), 'w',
              newline='', encoding='cp932', errors='ignore') as unconverted_csv_file:
        unconverted_csv_file.write(sheet_name)
    return None


def print_unconverted_warning(sheet_name):
    """
    sheet_name(str: CSVに変換できないsheetの名前)に対し,
    変換できない警告文をstdoutに出力するメソッド.
    """
    print("-" * 8)
    print("警告：")
    print("シート {} はCSVに変換できませんでした。".format(sheet_name))
    print("シートの中に写真・グラフが入ったかもしれません。")
    return None


class Converter(object):
    """
    ExcelファイルをCSVに変換するための容器クラス.
    param(str): "foo.xls"のようなXLSファイル名, もしくは"../../foo.xls"のようなXLSファイルの相対パス.
    self.abs_file_name(str): "/../../foo.xls"のようなXLSファイルの絶対パス.
    self.file_name(str): "foo.xls"のように, *クリーン*なXLSファイル名.
    self.sheet_names(list -> str): sheet(xlrd.sheet.Sheet)の名前(str)のリスト.
    """
    def __init__(self, file_name):
        """
        クラス初期化メソッド.
        もしインスランスを定義するときparamがなければ, RuntimeErrorを挙げる.
        """
        if file_name:
            self.abs_file_name = os.path.abspath(file_name)
            self.file_name = os.path.basename(self.abs_file_name)
        else:
            raise RuntimeError('ファイル名はありません')

        self.sheet_names = get_sheet_names_from_book(self.abs_file_name)

    def sheet_to_csv(self):
        """
        Convertインスタンスのself.file_nameに対し, それが指しているXLSファイルの
        全てのsheet(xlrd.sheet.Sheet)をresult(dict: -> list/None -> list)に変換し,
        変換不能のsheetの名前(str: sheet_name)をunconverted_sheets(list: -> str)に入れて,
        resultとunconverted_sheetを返す.
        """
        csv_path = get_csv_path_and_make_csv_folder(self.abs_file_name)
        result = {}
        unconverted_sheets = []
        for sheet_name in self.sheet_names:
            try:
                sheet = sheet_pretreatment(self.abs_file_name, sheet_name)
                result[sheet_name] = sheet
                print("シート {} をCSVに変換しました.".format(sheet_name))
            except ValueError:
                print_unconverted_warning(sheet_name)
                output_unconverted_csv_file(csv_path, sheet_name)
                result[sheet_name] = None
                unconverted_sheets.append(sheet_name)
            else:
                pass
        return result, unconverted_sheets

    def output_csv_files(self):
        """
        Converter.sheet_to_csvで変換したsheet(list: -> list)を, CSVファイルに書き込む.
        """
        csv_path = get_csv_path_and_make_csv_folder(self.abs_file_name)
        csv_source, unconverted_sheets = self.sheet_to_csv()
        for sheet_name in self.sheet_names:
            csv_name = sheet_name + ".csv"
            if csv_source[sheet_name]:
                with open(os.path.join(csv_path, csv_name), 'w', newline='',
                          encoding='cp932', errors='ignore') as csv_file:
                    writer = csv.writer(csv_file, delimiter=',',
                                        quotechar='|', quoting=csv.QUOTE_MINIMAL)
                    for row in csv_source[sheet_name]:
                        writer.writerow(row)
            else:
                pass
        return unconverted_sheets

    def pack_csv_files(self):
        """
        Converter.output_csv_fileで出力したCSVファイルを,
        フォルダ丸ごとZIPファイルとして圧縮する.
        """
        zip_name = self.file_name
        zip_name = os.path.splitext(zip_name)[0]
        path = os.path.dirname(self.abs_file_name)
        path = os.path.join(path, zip_name)
        pack = shutil.make_archive(path, format='zip',
                                   root_dir=path, base_dir='.')
        shutil.rmtree(path)
        return pack


if __name__ == "__main__":
    e2c = Converter(sys.argv[1])
    e2c.output_csv_files()
    e2c.pack_csv_files()
