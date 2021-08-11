from csv import *
import os
import glob
import csv
from xlsxwriter.workbook import Workbook
import XLSX


def Get_titles(read_file, index):
    set_val = set()
    f = open(read_file, 'r')
    csv_reader = reader(f)
    header = next(csv_reader)
    if header != None:
        for line in csv_reader:
            if line[index] == "N-A":
                continue
            else:
                set_val.add(line[index].strip())
    f.close()
    return set_val


def Make_Files(values, read_file, index, path):
    path = "ProgramFiles/Csv_files/style_data_csv/" + path + "/"
    for val in values:
        #need to replace the / in val

        f = open(read_file, 'r')
        csv_reader = reader(f)
        # print(val)
        csv_file = open(path + val + ".csv", "w", newline='')
        csv_writer = writer(csv_file)
        for line in csv_reader:
            # print(line[4])
            if val == line[index]:
                csv_writer.writerow(line)



def Sort_Data(values, index, path):
    largest_vals = []
    smallest_vals = []
    for val in values:
        file_path = "ProgramFiles/Csv_files/style_data_csv/"+path+"/" + val + ".csv"
        f = open(file_path, 'r')
        csv_reader = reader(f)
        # print(csv_reader)
        set_val = True
        search_val_large = ''
        search_val_small = ''
        for line in csv_reader:
            # print(line[index])
            if set_val:
                search_val_large = line
                search_val_small = line
            if line[index] == "NA":
                continue
            if Get_larger(float(search_val_large[index].rstrip("%")), float(line[index].rstrip("%"))):
                search_val_large = line

            if Get_Smallest(float(search_val_small[index].rstrip("%")), float(line[index].rstrip("%"))):
                search_val_small = line

            set_val = False
        if search_val_small != search_val_large:
            largest_vals.append(search_val_large)
            smallest_vals.append(search_val_small)
        f.close()
    return largest_vals, smallest_vals


def Get_larger(val1, val2):
    # print(val1)
    if val1 < val2:
        # print(val1)
        return True
    return False


def Get_Smallest(val1, val2):
    # print(val1)
    if val1 > val2:
        # print(val1)
        return True
    return False


def write_to_file(Best, Worst, name, path):
    path_best = 'ProgramFiles/Csv_files/style_data_csv/' + path + '/Best_' + name + ".csv"
    csv_file = open(path_best, "w", newline='')
    csv_writer = writer(csv_file)
    for element in Best:
        csv_writer.writerow(element)
    path_worst = 'ProgramFiles/Csv_files/style_data_csv/' + path + '/Worst_' + name + ".csv"
    csv_file_worst = open(path_worst, "w", newline='')
    csv_writer = writer(csv_file_worst)
    for element in Worst:
        csv_writer.writerow(element)
    csv_file.close()
    csv_file_worst.close()


def Convert_Csv_XLSX():
    for csvfile in glob.glob(os.path.join('ProgramFiles', '*.csv')):
        workbook = Workbook(csvfile[13:-4] + '.xlsx')
        worksheet = workbook.add_worksheet()
        with open(csvfile, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
        workbook.close()
        f.close()


def RunFile(Names, file1, indexes, reverse_indexes, path):
    values = Get_titles(file1, Names)
    Make_Files(values, file1, Names,path)
    for index in indexes:
        largest, smallest = Sort_Data(values, index,path)
        write_to_file(largest, smallest, str(index), (path+"/sorted"))
    for index in reverse_indexes:
        largest, smallest = Sort_Data(values, index,path)
        write_to_file(smallest, largest, str(index),(path+"/sorted"))


def Stylize():
    Names = 3
    TR1 = 8
    TR3 = 9
    TR5 = 10
    TR10 = 11
    YTD = 23
    NET = 24
    Gross = 25
    STD3 = 26
    STD5 = 27
    STD10 = 28
    SHR3 = 29
    SHR5 = 30
    SHR10 = 31
    indexes = [TR1, TR3, TR5, TR10, YTD, SHR3, SHR5, SHR10]
    reverse_indexes = [NET, Gross, STD3, STD5, STD10]
    file1 = "ProgramFiles/Stock_Mutual_funds.csv"
    file2 = "ProgramFiles/Bond_Mutual_funds.csv"
    RunFile(Names, file1, indexes, reverse_indexes, 'stocks')
    RunFile(Names, file2, indexes, reverse_indexes, 'bonds')
    Convert_Csv_XLSX()
    XLSX.Stylize_Data_Main()
