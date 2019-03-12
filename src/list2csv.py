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
        if switch:
            print("{} was reformatted without error.".format(name))
            file.write("{}, no error\n".format(basename))
        else:
            print("{} was reformatted with some error.".format(name))
            file.write("{}, with error\n".format(basename))
            for sheet in uncover_sheet:
                file.write("-- {}\n".format(sheet))
    return None


def main():
    names, path = names2list(sys.argv[1])
    print(names)
    for name in names:
        e2c = Excel2csv(name)
        uncovered_sheet = e2c.output_csv_files()
        if uncovered_sheet:
            make_log(name, path, False, uncovered_sheet)
        else:
            make_log(name, path, True, uncovered_sheet)
    print("All files in list was reformatted.")
    print("You can check names of error sheets in result.log.")
    return None


if __name__ == '__main__':
    main()
