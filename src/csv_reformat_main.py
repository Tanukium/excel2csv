import csv
import os
from csv_reformat import Data_length


table = []
with open(os.path.abspath('sample.csv'),'r') as csv_file:
    reader = csv.reader(csv_file)
    for row in reader:
         table.append(row)
data_length = Data_length('sample.csv')


if __name__ == '__main__':
    for row in table:
        print(row[(len(row)-data_length.data_length()):])

