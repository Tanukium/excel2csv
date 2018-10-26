#!/anaconda3/envs/msemi/bin/python
# -*- coding: UTF-8 -*-

import csv
import os
from collections import Counter


class DataLength(object):

    def __init__(self, file_name):
        self.table = []
        self.data_length = None
        with open(os.path.abspath(file_name),
                  'r', encoding='cp932') as csv_file:
            reader = csv.reader(csv_file)
            for row in reader:
                self.table.append(row)
        self.count_data_length()
    
    def receive_data_length(self, index):
        table_length = len(self.table[index])
        pop_num = 0
        while self.table[index][0] == '':
            self.table[index].pop(0)
            pop_num += 1
            if not self.table[index]:
                break
        return table_length - pop_num - 1
    
    def count_data_length(self):
        if not self.data_length:
            table_lengths = []
            for index in range(len(self.table)-1):
                table_lengths.append(self.receive_data_length(index))
            self.data_length = Counter(table_lengths).most_common()[0][0]

