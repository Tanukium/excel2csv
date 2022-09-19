import zipfile
import xlrd
import csv
import io

"以下は, book(xlrd.book.Book)及びsheet(xlrd.sheet.Sheet)をロードするためのメソッド."


def open_book(xls_file_obj):
    """
    xls_name(bytes: XLSファイル)をxlrd.book.Bookとして開く.
    """
    book = xlrd.open_workbook(file_contents=xls_file_obj, formatting_info=True)
    return book


def get_sheet_names_from_book(xls_file_obj):
    """
    book(xlrd.book.Book)の中の全ての,
    sheet(xlrd.book.Book)の名前(str)の集合をsheet_names(list)として返す.
    """
    book = open_book(xls_file_obj)
    sheet_names = book.sheet_names()
    return sheet_names


def get_row_col_num_from_sheet(sheet_name):
    """
    sheet_name(str: -> xlrd.sheet.Sheet)という名のシートに対し,
    それの行数, 列数を返す.
    """
    row_num = sheet_name.nrows
    col_num = sheet_name.ncols
    return row_num, col_num


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


def get_unmerged_sheet(xls_file_obj, sheet_name):
    """
    bookの特定sheetに対し, 全てのセルの結合を解除した上、
    結合が解除された空白セルに元の結合セルの値を充填してから, new_sheet(list: -> list)として返す.
    """
    book = open_book(xls_file_obj)
    sheet = book.sheet_by_name(sheet_name)
    rows_num, cols_num = get_row_col_num_from_sheet(sheet)
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


def replace_line_break(row):
    """
    row(list)の要素の中に存在する, 全ての*\n*を削除してからnew_row(list)として返す.
    """
    new_row = []
    for string in row:
        if type(string) == str and '\n' in string:
            new_row.append(string.replace('\n', ''))
        else:
            new_row.append(string)
    return new_row


def get_no_space_cell_sheet(sheet):
    """
    sheet(list: -> list)の中, *スペース*と*'\u3000'*に充填されたセルを,
    *''*に入れ替えてからnew_sheetを返す.
    また, セルに*\n*があれば, それを削除する.
    """
    new_sheet = []
    for row in sheet:
        new_sheet.append(replace_line_break(replace_space_strings(row)))
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


def transpose_sheet(sheet):
    """
    x行y列のsheetを, y行x列のsheetに変換してtransposed_sheetとして返す.
    """
    row_num, col_num = len(sheet), len(sheet[0])
    transposed_sheet = []
    for i in range(col_num):
        temp = []
        for row in sheet:
            temp.append(row[i])
        transposed_sheet.append(temp)
    return transposed_sheet


def count_blank_cell_at_row_start(transposed_sheet):
    """
    transposed_sheet(list: -> list -> str)の要素(list)に対し,
    連続何行が中の要素が全て*''*になる(つまり, 列がからっぽ)ことを計算し, その結果をintとして返す.
    結果がtransposed_sheetの長さ(つまり, sheetの列数)と一致すれば,
    sheetがからっぽに判定し, -1を返す.
    """
    count = 0
    for col in transposed_sheet:
        if any(col):
            break
        count += 1
    if count == len(transposed_sheet):
        return -1
    else:
        return count


def remove_blank_cell_at_row_start(sheet):
    """
    sheetのrowの始まりに存在する空っぽセル('')を削除して,
    new_sheetとして返す.
    """
    new_sheet = []
    transposed_sheet = transpose_sheet(sheet)
    count = count_blank_cell_at_row_start(transposed_sheet)
    if count != -1:
        for row in sheet:
            new_sheet.append(row[count:])
    else:
        for row in sheet:
            new_sheet.append(row)
    return new_sheet


def sheet_pretreatment(xls_file_obj, sheet_name):
    """
    sheet_name(str: sheet(xlrd.sheet.Sheet)の名前)に対し, その名前のsheetを
    結合セル解除 -> スペース入替 -> 空っぽセル削除 -> 空っぽ行削除 -> 行の始まりの空っぽセル削除をしてから,
    sheetを返す. つまり, 上のメソッドを順番通り呼び出すだけ.
    """
    sheet = get_unmerged_sheet(xls_file_obj, sheet_name)
    sheet = get_no_space_cell_sheet(sheet)
    sheet = get_no_blank_cell_sheet(sheet)
    sheet = get_no_blank_row_sheet(sheet)
    sheet = remove_blank_cell_at_row_start(sheet)
    return sheet


"""
以下は, 整形済みのsheetをcsvフォーマットのオンメモリfile-like objectとして
出力するメソッド.
"""


def output_converted_csv_file_on_memory(csv_source, sheet_name):
    """
    csv_source(dict: k->sheet_name, v->sheet)の中に入っている整形済みの特定sheetを、
    csvフォーマットのオンメモリfile-like objectに変換して値で返すメソッド.
    """
    csv_on_memory = io.StringIO()
    writer = csv.writer(csv_on_memory, delimiter=',',
                        quotechar='|', quoting=csv.QUOTE_MINIMAL)
    for row in csv_source[sheet_name]:
        writer.writerow(row)
    csv_data = csv_on_memory.getvalue()
    csv_on_memory.close()
    csv_data = csv_data.encode('utf-8')
    return csv_data


"""
以下は, 写真かグラフが入ったためCSVに変換できないシートを例外処理したり,
シェルで変換するとき警告をstdoutに出力したりするためのメソッド.
"""


def output_unconverted_csv_file_on_memory(sheet_name):
    """
    sheet_name(str: CSVに変換できないsheetの名前)に対し,
    csvフォーマットのオンメモリfile-like objectを作成して値で返すメソッド.
    """
    csv_on_memory = io.StringIO()
    csv_on_memory.write(("シート {} はCSVに変換できませんでした。\r\n".format(sheet_name)))
    csv_on_memory.write("シートの中に写真やグラフが入ったり、"
                        "シートに入っている有効データが少なかったりするかもしれません。")
    unconverted_csv_data = csv_on_memory.getvalue()
    csv_on_memory.close()
    unconverted_csv_data = unconverted_csv_data.encode('utf-8')
    return unconverted_csv_data


def print_unconverted_warning(sheet_name):
    """
    sheet_name(str: CSVに変換できないsheetの名前)に対し,
    変換できない警告文をstdoutに出力するメソッド.
    """
    print("-" * 8)
    print("警告：")
    print("シート {} はCSVに変換できませんでした。".format(sheet_name))
    print("シートの中に写真・グラフが入ったり、"
          "シートに入ってる有効データが少なかったりするかもしれません。")
    return None


class Converter(object):
    """
    ExcelファイルをCSVに変換するための容器クラス.
    param(str): "foo.xls"のようなXLSファイル名, もしくは"../../foo.xls"のようなXLSファイルの相対パス.
    # self.abs_file_name(str): "/../../foo.xls"のようなXLSファイルの絶対パス.
    # self.file_name(str): "foo.xls"のように, *クリーン*なXLSファイル名.
    self.file_like_xls(bytes): S3からダウンロードしたXLSファイルをオンメモリにしたfile-like object.
    self.sheet_names(list -> str): sheet(xlrd.sheet.Sheet)の名前(str)のリスト.
    """

    def __init__(self, xls_file_obj, bucket_name):
        """
        クラスのイニシャルメソッド.
        インスランスを定義するときparamがなければ, RuntimeErrorを挙げる.
        """
        if xls_file_obj and bucket_name:
            self.xls_file_obj = xls_file_obj
            self.bucket_name = bucket_name
        else:
            raise RuntimeError('処理させる.xlsのバイナリファイルは渡されていない！')

        self.sheet_names = get_sheet_names_from_book(self.xls_file_obj)

    def sheet_to_csv(self):
        """
        Convertインスタンスのself.file_nameに対し, それが指しているXLSファイルの
        全てのsheet(xlrd.sheet.Sheet)をresult(dict: -> list/None -> list)に変換し,
        resultを返す.
        """
        result = {}
        for sheet_name in self.sheet_names:
            try:
                sheet = sheet_pretreatment(self.xls_file_obj, sheet_name)
                result[sheet_name] = sheet
                print("シート {} をCSVに変換しました.".format(sheet_name))
            except ValueError:
                print_unconverted_warning(sheet_name)
                result[sheet_name] = None
            else:
                pass
        return result

    def output_csv_files_to_memory(self):
        """
        Converter.sheet_to_csvで変換したsheet(list: -> list)を, CSVファイルに書き込む.
        """
        result = {}
        csv_source = self.sheet_to_csv()
        for sheet_name in self.sheet_names:
            csv_name = sheet_name + ".csv"
            if csv_source[sheet_name]:
                csv_file_obj = output_converted_csv_file_on_memory(csv_source, sheet_name)
                result[csv_name] = csv_file_obj
            else:
                csv_file_obj = output_unconverted_csv_file_on_memory(sheet_name)
                result[csv_name] = csv_file_obj
        return result

    def pack_csv_files(self):
        """
        Converter.output_csv_fileで出力したCSVファイルを,
        フォルダ丸ごとオンメモリのZIPファイルとして圧縮する.
        """
        zip_stream = io.BytesIO()
        result = self.output_csv_files_to_memory()

        with zipfile.ZipFile(zip_stream, 'w', compression=zipfile.ZIP_DEFLATED) as writer:
            for sheet_name in self.sheet_names:
                writer.writestr(sheet_name + '.csv', result[sheet_name + '.csv'])
        zip_upload = zip_stream.getvalue()
        zip_stream.close()
        return zip_upload


if __name__ == "__main__":
    # xls_file_name = '国勢調査_上田.xls'
    pass
