from excel2csv import Excel2csv
import os
import sys


def names2list(list_file):
    container = []
    path = os.path.dirname(os.path.abspath(list_file))
    with open(list_file, mode="r", encoding="UTF-8-sig") as file:
        for line in file:
            line = line.rstrip()
            line = path + os.sep + line
            container.append(line)
        return container, path


def make_log(name, path, switch, uncover_sheet):
    log = path + os.sep + "result.log"
    basename = os.path.basename(name)
    with open(log, mode="a+", encoding="UTF-8") as file:
        if switch == "all":
            print("-" * 8)
            print("{} は転換されました。".format(name))
            file.write("{}, 問題なし\n".format(basename))
        elif switch == "some":
            print("-" * 8)
            print("警告： {} は転換されましたが、".format(name))
            print("　　　 一部のシートは転換が飛ばされました。")
            file.write("{}, 一部のシートは問題あり\n".format(basename))
            for sheet in uncover_sheet:
                file.write("-- {}\n".format(sheet))
        elif switch == "none":
            print("-" * 8)
            print("警告： {} は転換されませんでした。".format(name))
            print("　　　 リストファイルか、このプログラムの論理をチェックしてください。")
            file.write("{}, 全体は問題あり\n".format(basename))
    return None


def main():
    names, path = names2list(sys.argv[1])
    print("今度転換するファイルは以下のものです：")
    for name in names:
        print("{},".format(name))
    for name in names:
        try:
            e2c = Excel2csv(name)
            uncovered_sheet = e2c.output_csv_files()
            if uncovered_sheet:
                make_log(name, path, "some", uncovered_sheet)
            else:
                make_log(name, path, "all", uncovered_sheet)
        except:
            temp_sheet = []
            make_log(name, path, "none", temp_sheet)
    print("-" * 8)
    print("リストファイルにおける全ての Excel ファイルは転換されました。")
    print("転換の過程と結果は、{} にて確認できます。".format("result.log"))
    return None


if __name__ == '__main__':
    main()
