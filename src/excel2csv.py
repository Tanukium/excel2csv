# -*- coding: UTF-8 -*-

"""
【プログラム名】Excel→CSV直行便
【モジュール名】excel2csv.py
【機能】Excel統計ファイルに収録された表データ(ワークシート)を
       機械判読可能なCSV形式に変換処理する
【処理概要】
    入力 統計Excelファイル名 (例)02_009_danjo_nenrei.xls
    出力 Excelファイルに対応したCSV格納フォルダ (例)02_009_danjo_nenrei
        そのフォルダ内にワークシートごとにCSVファイルを書き出す。
        未処理のワークシート名はuncovered.txtを作成しその中に出力する。　　
    出力例 フォルダ名 02_009_danjo_nenrei
        出力CSVファイル　9（旧丸子町）.csv, 9（旧上田市）.csv, …
    未処理ファイル名リスト uncovered.txt
【作成日】2018/07/04
【最終更新】2019/02/05 AM03:00:00 -> 作成日と最終更新は, Gitの履歴に於いて閲覧できる.
【作成者】シュクメイスイ（長野大学前川ゼミ）
"""

import sys
import os
import csv
import xlrd
# Excelファイルを読み取り,
# Workbookオブジェクトとして返すライブラリ.
# リスト走査の繰り返しをすることで,
# Workbook>Worksheet>List（行または列のセルら）
# の順番で値を得られる.
import copy
# Pythonではメモリでのリストやタプルなどの
# オブジェクトのコピーができないため,
# このライブラリのdeepcopy()メソッドを使用し
# リストオブジェクトのコピー機能を実現した.
import re
# 正規表現機能を提供するライブラリ.
# sub()メソッドを用いて, 特定の文字列をマッチして書き替えられる.
from collections import Counter as CoT


# リストにおける要素の出現個数を統計するライブラリ.
# most_common()メソッドは出現個数の最頻値を返す.


def _get_data_from_sheet(book_name, sheet_name,
                         switch, container):
    """
    Excelファイルのワークシートから行または列ごとに一行分か一列分のセルらを読み取り,
    containerに入れて値として返すメソッド.
    containerは次のような形のものである:
    [[cell1, cell2, cell3, ...], [cell1, cell2, cell3, ...], ...]
    :param book_name: (Workbook.workbook)
    xlrd.open_workbook()で開いたExcelファイル.
    :param sheet_name: (str) Excelファイル
    (Workbook.workbookオブジェクト)におけるWorksheetの名前.
    :param switch: (str: "row" or "col")
    行モードか列モードでExcelファイルの内容を読み取るのをコントロールするパラメータ.
    :param container: (list) 値として返すリスト.
    :return: (list) 行か列のlistを含めているlist. 行か列のlistには、セルを含める.
    """
    sheet = book_name.sheet_by_name(sheet_name)
    # Workbook.worksheetオブジェクトを開く.
    data_range = sheet.nrows if switch == "row" else sheet.ncols
    # 行の数か列の数でsheetを走査するのを決める.
    for index in range(data_range):
        if switch == "row":
            new_data = sheet.row_values(index)
            # new_dataはlistである（一行分か一列分のセルら）.
        else:
            new_data = sheet.col_values(index)
        container.append(new_data)
    return container


def _count_aline_data_length(data, index):
    """
    一行分か一列分のセルを含めるリストのサイズ（要素の個数）を計算し,
    （安全のために）そのコピーを作成する.
    コピーにおける始めの要素からfloat型の要素までのサイズをさらに計算し,
    最初計算したリストのサイズでそれを引き、結果を値として返す.
    :param data: (list) [[cell1, cell2, cell3, ...],
    [cell1, cell2, cell3, ...], ...]の形の, 行か列のlistを含めているlist.
    :param index: (int) 行か列のlistは要素としてdataにおける順番.
    :return: (int) data[index]におけるfloat形の要素の個数.
    いわゆる行か列におけるデータの長さ.
    """
    num = 0
    data_len = len(data[index])
    # data[index]（一行分か一列分のセルらのlist）のサイズを計算
    data_copy = copy.deepcopy(data[index])  # data[index]のコピーを作成
    for item in data_copy:  # data_copyを走査する
        if type(item) is not float:
            # type()はオブジェクトの型を値として返すメソッド.
            num += 1
        else:
            if type(data_copy[data_copy.index(item) + 1]) is float:
                break
                # データの連続性を考え,
                # float型の要素が2個連続して出現する場合だけ走査を停止させ
            else:
                num += 1
    return data_len - num


def _del_repeat(lst):
    """
    とあるlistにおける重複（2個以上存在する）の要素を,
    一個だけにするメソッド.
    :param lst: (list) メソッドが応用するlist.
    :return: (list) 全部の重複の要素が削除されたlist.
    """
    for item in lst:
        while lst.count(item) > 1:
            del lst[lst.index(item)]
    return lst


def _u3000_killer(lst):
    """
    とあるlistにおける全部の'\u3000'
    （UTF-8における特殊なスペース）要素を空白の文字列に
    入れ替えるメソッド.
    :param lst: (list) メソッドが応用するlist.
    :return: (list) 全部の'\u3000'要素が削除されたlist.
    """
    while '\u3000' in lst:
        lst[lst.index('\u3000')] = ''
    return lst


def _del_blank_cell(lst):
    """
    とあるlistにおける全部の空白の文字列（''）要素を削除するメソッド.
    :param lst: (list) メソッドが応用するlist.
    :return: (list) 全部の空白文字列要素が削除されたlist.
    """
    while '' in lst:
        lst.remove('')
        # del list[index] か list.remove(member)は
        # listにおける要素を削除するメソッド.
    return lst


def _del_blank_list(lst):
    """
    とあるlistにおける全部の空白のlist要素（[]）を削除するメソッド.
    :param lst: (list) メソッドが応用するlist.
    :return: (list) 全部の空白list要素が削除されたlist.
    """
    while [] in lst:
        lst.remove([])
    return lst


def _del_blank_cell_at_start(lst):
    """
    空白の文字列要素を含めるlist要素を含めるlistにおける,
    list要素のはじめにある余った空白の文字列要素を削除するメソッド.
    このメソッドはlst_copyというlistのコピーを作成し,
    その中のlist要素ごとにlist要素の始めから連続の空白文字列要素を削除し,
    その個数を計算し, resultというlistに入れる.
    result.sort()でresultにおける連続の空白文字列要素の個数要素を
    一番小さいから並び直し, その一番小さい個数要素を取り出す.
    もし一番小さい個数要素は0の場合, ２番目小さい個数要素を取り出す.
    listにおける各list要素の始めから,
    先ほど取り出した個数要素分の空白文字列要素を削除してから,
    各list要素を含めるlistを値として返す.
    例えば, [["", 1, 2], ["", "", 5], [6, 7, 8]]の場合,
    このメソッドを使用し返された値は
    [["", 1, 2], ["", 5], [6, 7, 8]]のはずである.
    :param lst: (list) メソッドが応用するlist.
    :return: (list) 空白の文字列要素を含めるlist要素を含めるlistが,
    list要素のはじめにある余った空白の文字列要素を削除するメソッド.
    """
    lst_copy = copy.deepcopy(lst)  # パラメータのlstのコピーを作成する.
    result = []  # 空白文字列要素の個数を入れる容器(list)を作成.
    for item in lst_copy:
        # lstにおける各list要素の,
        # 始めからの連続の空白文字列要素の個数を計算する.
        num = 0
        while not item[0]:
            # Pythonでは, とあるオブジェクトは0であれば,
            # そのBooleanの値がfalseになる.
            item.remove('')
            num += 1
        result.append(num)
    result.sort()
    if not min(result):  # if min(result) == 0と同じ意味
        result = result[1]  # ここからresultはlistからintになった
    else:
        result = min(result)
    for item in lst:
        del item[:result]
        # lstにおける各list要素の,
        # 始め（item[0]）からitem[result]までの要素らを削除
    return lst


class Excel2csv(object):
    """
    Excelファイルを整形されたCSVファイルに転換する.
    :arg file_path: (str) Excelファイルのパス（ファイル名含め）.
    :arg file_name: (str) Excelファイルの名前（拡張子なし）.
    :arg result: (dict) "{sheetname: [sheet], ...}"のような形の容器.
    Excelから読み取ったworksheetごとのデータを収納する.
    :arg comment: (dict) "{sheetname: [title, comment, ...], ...}"
    のような形の容器.
    _get_title_from_sheet()と_get_comment_from_end_of_sheet()で
    抜かれたタイトル文字列とコメント文字列はこの中に入れる.
    :arg data: (dict) "{sheetname: [[data_of_row], ...], ...}"
    のような形の容器. 行ごとのデータlistらを含めるシートlistを収納する.
    :arg index_row: (dict) "{sheetname: [index_of_row], ...}"
    のような形の容器. シートごとの見出し行listを収納する.
    :arg index_col: (dict) 同index_row.
    """

    def __init__(self, file_name):
        """
        新たなExcel2csvオブジェクトを初期化するメソッド.
        新たなExcel2csvオブジェクトを作成するとき自動で稼働するので,
        外部でコールする必要はない.
        :param file_name: (str) Excelファイルの名前（拡張子ある）.
        外部ファイルでこのファイルをコールしてExcel2csvオブジェクトを初期化する
        場合, このパラメータが必要である.
        シェルでこのファイルを実行する場合, このパラメータを書く必要はない.
        :raise RuntimeError: 外部ファイルでのコールの場合file_nameを書かないか
        シェルでファイル自身を実行する場合パラメータを書かないと,
        このエラーがraiseされる.
        """
        if file_name:  # file_nameは空白でない（存在する）限り
            self.file_path = os.path.abspath(file_name)
        elif len(sys.argv) > 1 and sys.argv[1]:
            # （シェルで実行する場合）パラメータが存在する限り
            self.file_path = os.path.abspath(sys.argv[1])
        else:
            raise RuntimeError('No path or filename')
        self.file_name = os.path.basename(self.file_path)

        self.result = self._get_data_from_excel()
        self.comment, self.data = {}, {}
        self.index_row, self.index_col = {}, {}

    def _get_data_from_excel(self):
        """
        _get_data_from_sheet()を使用し,
        Excelファイルからシートごとにセル内容を読み取り,
        値として返すメソッド.
        :return: (dict) {sheetname: [[[row], [row], ...],
        [[col], [col], ...]], ...}の形のdictオブジェクト.
        """
        result = {}
        book = xlrd.open_workbook(self.file_path)  # Excelファイルを開く.
        sheet_names = book.sheet_names()  # sheet_names()でworksheetらを含めるlistオブジェクトを取得する.
        for index in range(len(sheet_names)):
            # 行と列ごとにbookをget_data_from_sheet()で走査し,
            # 結果にシートの名前というキーをつけてresultに入れる.
            rows, cols = [], []
            _get_data_from_sheet(book, sheet_names[index],
                                 "row", rows)
            _get_data_from_sheet(book, sheet_names[index],
                                 "col", cols)
            result[sheet_names[index]] = [rows, cols]
        return result

    def _get_title_from_sheet(self, sheet_name):
        """
        Worksheet（のはじめ）からタイトルとコメントを取り出し,
        Excel2csvオブジェクトのcommentというattributeに入れるメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (list) １番目の文字列要素はタイトルであり,
        ２番目からの文字列要素はコメントであるlistを返す.
        """
        if self.result[sheet_name]:
            # result[sheet_name]というattrが存在する限り
            # （そうでないとエラーがraiseされる）
            row_result, col_result = [], []
            # list容器2個を作る
            rows = self.result[sheet_name][0]
            # sheet_nameという名前のシートにおける
            # 行のlistらを含めているlistとrowsという変数をバインド

            for row in rows:
                _u3000_killer(row)
                # 行のlistらにおける\u3000文字列を削除
            for row in rows:
                if not any(row):
                    rows.remove(row)
                    # 行のlistらには全ての要素が空白文字列の
                    # listがあれば, それを削除
            rows_copy = copy.deepcopy(rows)
            # rowsのコピーを作成
            for item in rows_copy:
                _del_blank_cell(item)
            # rows_copyにおける行のlistらにおける
            # 空白文字列要素を削除
            for row in rows_copy:
                if len(row) == 1:
                    row_result.append(row[0])
                elif len(row) == 2:
                    for item in row:
                        row_result.append(item)
                else:
                    break
                # rows_copyにおけるサイズは1か2の行のlist
                # における要素を全部コメントと見なし,
                # row_resultに入れる（１番目の要素はタイトル）.
            for item in row_result:
                for row in rows:
                    if item in row:
                        rows.remove(row)
                    # row_resultの要素を含める
                    # rowsにおける行のlistを削除.

            # 以下の段落は, 上の段落と同じ論理.
            # ただし, 行ではなく列を処理する.
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

            # comment[sheet_name]というattrが存在しない限り
            # row_result, col_resultの要素を
            # comment[sheet_name]に入れる.
            if sheet_name not in self.comment:
                self.comment[sheet_name] = []
                for item in row_result:
                    self.comment[sheet_name].append(item)
                for item in col_result:
                    self.comment[sheet_name].append(item)

            # comment[sheet_name]の要素は列のlistのなかに存在すれば
            # 列のlistのなかのその要素を空白文字列にする.
            for col in cols:
                for member in self.comment[sheet_name]:
                    if member in col:
                        col[col.index(member)] = ''

            # rowsとcolsのはじめにある
            # 余った空白の文字列要素を削除する.
            _del_blank_cell_at_start(rows)
            _del_blank_cell_at_start(cols)

            return _del_repeat(row_result + col_result)

    def _get_comment_from_end_of_sheet(self, sheet_name):
        """
        Worksheetの全体からコメントを取り出し,
        Excel2csvオブジェクトのcommentというattributeに入れるメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (list) 文字列要素（コメント）によって構成されるlistを返す.
        """
        if self.result[sheet_name]:
            # result[sheet_name]というattrが存在する限り
            row_result = []
            # list容器一個を作る
            rows = self.result[sheet_name][0]
            # sheet_nameという名前のシートにおける
            # 行のlistらを含めているlistとrowsという変数をバインド

            for row in rows:
                if not any(row):
                    rows.remove(row)
            # 行のlistらには全ての要素が空白文字列の
            # listがあれば, それを削除
            rows_copy = copy.deepcopy(rows)
            # rowsのコピーを作成
            for item in rows_copy:
                _del_blank_cell(item)
            # rows_copyにおける行のlistらにおける
            # 空白文字列要素を削除
            for row in rows:
                if (len(row) == 1 and
                        len(rows_copy[rows_copy.index(row)]) > 1):
                    row_result.append(row[0])
                # もしとある行のlistは上記の操作を行う前のサイズが1より大きい,
                # 操作を行った後のサイズがちょうど1（つまり, [hoge, "", "", ...]
                # のようなlist）であれば, その中の唯一の要素をrow_resultに入れる.
            for item in row_result:
                for row in rows:
                    if item in row:
                        rows.remove(row)
                # row_resultの要素を含める
                # rowsにおける行のlistを削除.

            # comment[sheet_name]というattrが存在する限り
            # row_resultの要素をcomment[sheet_name]に入れて,
            # 重複分の要素を削除する.
            if sheet_name in self.comment:
                for item in row_result:
                    self.comment[sheet_name].append(item)
                _del_repeat(self.comment[sheet_name])
            return row_result

    def _count_row_data_length(self, sheet_name):
        """
        Worksheetにおける行における連続するデータセルの個数
        （データエリアの長さ）を計算し, その結果を値として返すメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (int) 行における連続するデータセルの個数
        （データエリアの長さ）.
        """
        result = []
        rows = self.result[sheet_name][0]
        rows_copy = copy.deepcopy(rows)
        # 容器を作り, rowsとシートの行のlistらをバインドし,
        # rows_copyというrowsのコピーを作る.
        if rows_copy:  # rows_copyに要素がある限り
            for index in range(len(rows_copy)):
                result.append(_count_aline_data_length(rows_copy, index))
            # 行ごとのデータの長さを計算し, resultに入れる.
            result = CoT(result).most_common()
            # ここからresultは[int, int, ...]から
            # [(int, int), (int, int), ...]の形のlistになる.
            # most_common()メソッドは出現個数の最頻値を返す.
            # 返された結果は, (個数, 出現回数)の形のタプル要素で構成された
            # listになる.
            if not result[0][0] and result[0][1] == result[1][1]:
                # 最頻値の個数は0, しかも最頻値の出現回数と
                # 最頻値の次の個数の出現回数が同じ場合
                return result[1][0]
                # 最頻値の次の個数を返す.
            else:
                return result[0][0]
                # そうでない場合, 最頻値の個数の返す.

    def _count_col_data_length(self, sheet_name):
        """
        Worksheetにおける列における連続するデータセルの個数
        （データエリアの長さ）を計算し, その結果を値として返すメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (int) 列における連続するデータセルの個数
        （データエリアの長さ）.
        """
        # このメソッドの論理は,
        # _count_row_data_length()と全く同じものである.
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
        """
        Worksheetの見出し行を取り出し, 整形して,
        listとして返すメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (list) ["a_b_c", "a_b_d", ...]の形のlist.
        """
        container = []
        rows = self.result[sheet_name][0]
        rows_copy = copy.deepcopy(rows)
        # 容器を作り, rowsとシートの行のlistらをバインドし,
        # rows_copyというrowsのコピーを作る.

        rows_len = len(rows_copy[0])
        row_data_len = self._count_row_data_length(sheet_name)
        row_index_len = rows_len - row_data_len
        # rowsのサイズを計算し,
        # rowsにおける行の連続するデータセルの個数を計算し,
        # 両者の引き算は見出し行の長さになる.

        for row in rows_copy:
            container.append(row[:row_index_len])
        # rowsのコピーから行ごとに見出し行の部分を割り出し,
        # listとしてcontainerに入れる.
        container = list(filter(any, container))
        # containerにおける空白list要素([])を全部削除する.
        _del_blank_cell_at_start(container)
        # containerにおけるlist要素のはじめにある
        # 余った空白の文字列要素を削除する.

        for row in container:
            for index in range(len(row)):
                if not row[index]:
                    row[index] = container[container.index(row) - 1][index]
                else:
                    break
        # [["上田", "上田", "東部"], ["", "上田", "南部"]]という形のlistを
        # [["上田", "上田", ""], ["上田", "上田", "南部"]]という形にする.
        con_copy = copy.deepcopy(container)
        # container容器のコピーを作る.
        del container
        # container自体を削除する.
        for row in con_copy:
            if '\u3000' in row:
                row.remove('\u3000')
        # container容器のコピーにおける行における
        # 文字列要素の中に, "\u3000"があれば
        # "\u3000"を削除する.
        container = list(filter(any, con_copy))
        # container容器のコピーにおける空白のlist要素([])を
        # 削除し, またcontainerという変数とバインドする.
        tmp2, result = [], []
        # 容器二つを作る.
        for row in container:
            tmp = []
            for item in row:
                if type(item) is float:
                    tmp.append(str(int(item)))
                else:
                    tmp.append(item)
            s = "_".join(tmp)
            tmp2.append(s)
            # tmpという容器に
            # containerにおける数字要素を文字列化したものと,
            # もともと文字列であった要素を入れて,
            # "_"を要素の区切りにしてtmp自体を文字列化して,
            # tmp2に入れる.
        for item in tmp2:
            if item.find("\n") != -1:
                item = re.sub(r"\n", " ", item)
            if item.endswith("_"):
                result.append(item.rstrip("_"))
            elif item.startswith("_"):
                result.append(item.lstrip("_"))
            else:
                result.append(item)
            # tmp2における要素は,
            # "\n"（改行）を含める場合改行をスペースにし,
            # 末尾かはじめに"_"がある場合"_"を削除し,
            # 以上の処理をしてからresultに入れる.
        if sheet_name not in self.index_row:
            # index_rowというattributeには
            # sheet_nameというキーが存在しない限り
            self.index_row[sheet_name] = result
            # sheet_nameというキーの値はresultにする.
        return result

    def _make_index_cols(self, sheet_name):
        """
        Worksheetの見出し列を取り出し, 整形して,
        listとして返すメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (list) ["a_b_c", "a_b_d", ...]の形のlist.
        """
        # このメソッドの論理は, _make_index_rows()と同じである.
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
        """
        行ごとにWorksheetにおける行のlistにおけるデータセルらを割り出し,
        値として返すメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (list) [["data", "data", ...], ["data", ...], ...]
        の形のlist.
        """
        container = []
        rows = self.result[sheet_name][0]
        cols = self.result[sheet_name][1]
        rows_len = len(rows[0])
        cols_len = len(cols[0])
        # 容器を作り, rowsとcolsをそれぞれ
        # worksheetの行・列のlistらとバインドし,
        # 一行分・一列分のセル個数を計算する.
        for row in rows[(cols_len -
                         self._count_col_data_length(sheet_name)):]:
            container.append(row[(rows_len -
                                  self._count_row_data_length(sheet_name)):])
            # 連続するデータセルの個数（データエリアの長さ）に応じて,
            # 行ごとに行のlistにおけるデータセルらを割り出し,
            # containerという容器に入れる.
        if sheet_name not in self.data:
            # dataというattributeには
            # sheet_nameというキーが存在しない限り
            self.data[sheet_name] = container
            # sheet_nameというキーの値はcontainerにする.
        return container

    def _make_csv_path(self):
        """
        出力されるCSVファイルを収納するフォルダを作成するメソッド.
        値として, そのフォルダのパスを返す.
        :return: (str) 出力されるCSVファイルを収納するフォルダのパス.
        """
        file_dir_name = os.path.dirname(self.file_path) + os.sep
        csv_path = (file_dir_name +
                    os.path.splitext(self.file_name)[0] + os.sep)
        # 出力されるCSVファイルのパス（str）を作成する.
        # os.path.dirname("/dev/sda1/file")は"/dev/sda1/"のようなパスを返し,
        # os.path.splitext("file.ext")は"file"を返す.
        # os.sepは, *nixにおいては"/"であり,
        # windowsでは"\"である.
        if not os.path.exists(csv_path):
            os.mkdir(csv_path)
        # csv_path自体はシステムに存在しない限り,
        # それを作成する.
        return csv_path

    def _make_uncover_list(self, item):
        """
        Workbookにおける転換しないworksheetの名前を含める
        テキストファイルを作成するメソッド.
        :param item: (str) 転換しないworksheetの名前.
        :return: (none) なし.
        """
        with open(self._make_csv_path() + "uncovered.txt", 'a+',
                  newline='', encoding='utf-8') as uncover_list:
            uncover_list.write(item + "\n")

    def _make_btf_sheet(self):
        """
        Workbookにおける処理できるworksheetと
        処理できないworksheetを判断し, 処理できるworksheetを
        整形し, できないworksheetの名前をuncovered.txtに書き込む
        メソッド.
        :return: (none) なし.
        """
        self._get_data_from_excel()
        container, zero_group, one_group = {}, [], []
        # dict, list, listという容器を3個作る.
        for key in self.result:
            container[key] = []
            for item in self.result[key]:
                container[key].append(len(item))
        # containerを{sheetname: [行の数, 列の数], ...}のようにする.
        print(container)
        for key in container:
            if not container[key][0] or not container[key][1]:
                zero_group.append(key)
            elif container[key][0] == 1 or container[key][1] == 1:
                one_group.append(key)
        # containerに於いて, 要素の行の数か列の数は0か1の場合,
        # そのキー（sheetname）をzero_groupかone_groupに入れる.
        for key in self.result:
            if key in zero_group or key in one_group:
                self._make_uncover_list(key)
            # zero_groupとone_groupにおける要素（sheetname）を
            # uncovered.txtに書き込む.
            else:
                self._get_title_from_sheet(key)
                self._get_comment_from_end_of_sheet(key)
                self._make_index_rows(key)
                self._make_index_cols(key)
                self._make_data_rows(key)
                # 正常のシートのタイトル・コメントを取り出し,
                # 見出し行・列を整形し, データセルを取り出す.
                print(key,
                      self.comment[key],
                      self.index_row[key],
                      self.index_col[key],
                      self.data[key], sep="\n")

    def _csv_from_sheet(self, sheet_name):
        """
        処理されたattributeらからデータを読み取り, 結合し,
        CSVファイルに書き込むメソッド.
        :param sheet_name: (str) worksheetの名前.
        :return: (none) なし.
        """
        csv_name = sheet_name + ".csv"
        with open(self._make_csv_path() + csv_name, 'w',
                  newline='', encoding='cp932') as csv_file:
            writer = csv.writer(csv_file, delimiter=',',
                                quotechar='|',
                                quoting=csv.QUOTE_MINIMAL)
            # CSVファイルに書き込む用意をする.
            writer.writerow(self.comment[sheet_name][0:1])
            writer.writerow(self.comment[sheet_name][1:])
            writer.writerow(self.index_col[sheet_name])
            # タイトル, コメントら, 見出し行を一行ずつCSVに書き込む.
            print(len(self.index_row[sheet_name]),
                  len(self.data[sheet_name]))
            if (len(self.index_row[sheet_name])
                    > len(self.data[sheet_name])):
                # 見出し列の長さとdataというattributeのサイズ
                # （データエリアの長さではない, 列におけるセルの数とほぼ同じ）
                # を比較する. 見出し列の方が長かったら,
                # 見出し列の２番目の要素から一個ずつ,
                # データセルで構成された行のlistのはじめにくっつける.
                for index in range(len(self.data[sheet_name])):
                    self.data[sheet_name][index].insert(0,
                                                        self.index_row[sheet_name][index + 1])
                for item in self.data[sheet_name]:
                    writer.writerow(item)
                # データセルで構成された行のlistらを一個ずつ
                # CSVに書き込む. 一個の行のlistはCSVファイルの
                # 一行を占める.
            elif (len(self.index_row[sheet_name])
                  == len(self.data[sheet_name])):
                if self.index_row[sheet_name][0] != self.index_col[sheet_name][0]:
                    # 見出し列の長さとdataというattributeのサイズは同じだが,
                    # 見出し行の１番目の要素(str)と見出し列の１番目の要素は違う場合,
                    # 見出し列の１番目の要素から一個ずつ,
                    # データセルで構成された行のlistのはじめにくっつける.
                    for index in range(len(self.data[sheet_name])):
                        self.data[sheet_name][index].insert(0,
                                                            self.index_row[sheet_name][index])
                else:
                    # 見出し列の長さとdataというattributeのサイズは同じ,
                    # しかも見出し行の１番目の要素(str)と見出し列の１番目の要素も同じ場合,
                    # 見出し列の２番目の要素から一個ずつ,
                    # データセルで構成された行のlistのはじめにくっつける.
                    for index in range(len(self.data[sheet_name])):
                        if index != len(self.data[sheet_name]) - 1:
                            self.data[sheet_name][index].insert(0,
                                                                self.index_row[sheet_name][index + 1])
                        else:
                            break
                for item in self.data[sheet_name]:
                    writer.writerow(item)
                    # データセルで構成された行のlistらを一個ずつ
                    # CSVに書き込む. 一個の行のlistはCSVファイルの
                    # 一行を占める.

    def csv_from_excel(self):
        """
        外部から唯一見られるメソッド.
        このメソッドをコールすることで, 整形されたCSVファイルが出力される.
        :return: (none) なし.
        """
        self._make_btf_sheet()
        for key in self.comment:
            self._csv_from_sheet(key)


def main():
    """
    モジュールではなく, 単独のプログラムとして
    このファイルを実行するとき稼働するメソッド.
    "python excel2csv.py filename.xls"のように
    コールされる時, シェルから伝達されるシェルのパラメータ
    （sys.argv[1]）が必要である.
    :return: (none) なし.
    """
    e2c = Excel2csv(None)
    e2c.csv_from_excel()


if __name__ == '__main__':
    main()
