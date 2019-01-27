from excel2csv import Excel2csv
import os
import sys


def names2list(name):
    container = []
    with open(name, mode="r", encoding="UTF-8-sig") as file:
        for line in file:
            line = line.rstrip()
            line = os.path.dirname(os.path.abspath(name)) + os.sep + line
            container.append(line)
        return container


def main():
    names = names2list(sys.argv[1])
    print(names)
    for name in names:
        e2c = Excel2csv(name)
        e2c.csv_from_excel()


if __name__ == '__main__':
    main()
