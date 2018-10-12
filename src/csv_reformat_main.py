import csv
import os
import sys
from csv_reformat import Data_length


table = []
with open(os.path.abspath(sys.argv[1]),'r') as csv_file:
    reader = csv.reader(csv_file)
    for row in reader:
         table.append(row)
data_length = Data_length(sys.argv[1])


if __name__ == '__main__':
    for row in table:
        data_length_in_row = len(row)-data_length.data_length
        print(row[data_length_in_row:])

