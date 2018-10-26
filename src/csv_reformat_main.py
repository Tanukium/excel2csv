#!/anaconda3/envs/msemi/bin/python
# -*- coding: UTF-8 -*-

import csv
import os
import sys
from csv_reformat import DataLength

table = []
with open(os.path.abspath(sys.argv[1]), 'r', encoding='cp932') as csv_file:
    reader = csv.reader(csv_file)
    for row in reader:
        table.append(row)
data_length = DataLength(sys.argv[1])

if __name__ == '__main__':
    for row in table:
        data_length_in_row = len(row) - data_length.data_length
        print(row[data_length_in_row:])
