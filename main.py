import os
import threading
from csv import writer
import requests
from lxml import html
import csv
import time
from tkinter import *
import sys

from xlsxwriter.exceptions import FileCreateError

import Get_Style_Data

from tkinter.ttk import *



def Get_roots(name):
    roots = []
    url1 = "https://finance.yahoo.com/quote/" + name
    url2 = url1 + "/profile"
    url3 = url1 + "/performance"
    url4 = url1 + "/risk"
    url5 = "https://www.zacks.com/funds/etf/" + name
    url6 = url1 + "/holdings"
    url7 = 'https://markets.ft.com/data/etfs/tearsheet/holdings?s=' + name
    url8 = 'https://research.tdameritrade.com/grid/public/etfs/profile/profile.asp?symbol=' + name
    s = requests.session()
    try:

        page = s.get(url1, headers={'User-Agent': 'Custom'})
        root1 = html.fromstring(page.text)
        roots.append(root1)
    except requests.exceptions.ConnectionError:
        root1 = False
        roots.append(root1)
    try:
        page = s.get(url2, headers={'User-Agent': 'Custom'})
        root2 = html.fromstring(page.text)
        roots.append(root2)
    except requests.exceptions.ConnectionError:
        root2 = False
        roots.append(root2)

    try:
        page = s.get(url3, headers={'User-Agent': 'Custom'})
        root3 = html.fromstring(page.text)
        roots.append(root3)
    except requests.exceptions.ConnectionError:
        root3 = False
        roots.append(root3)

    try:
        page = s.get(url4, headers={'User-Agent': 'Custom'})
        time.sleep(0.5)
        root4 = html.fromstring(page.text)
        roots.append(root4)
    except requests.exceptions.ConnectionError:
        root4 = False
        roots.append(root4)

    try:
        page = s.get(url5, headers={'User-Agent': 'Custom'})
        root5 = html.fromstring(page.text)
        roots.append(root5)

    except requests.exceptions.ConnectionError:
        root5 = False
        roots.append(root5)

    try:
        page = s.get(url6, headers={'User-Agent': 'Custom'})
        root6 = html.fromstring(page.text)
        roots.append(root6)
    except requests.exceptions.ConnectionError:
        root6 = False
        roots.append(root6)

    try:
        page = s.get(url7, headers={'User-Agent': 'Custom'})
        root7 = html.fromstring(page.text)
        roots.append(root7)
    except requests.exceptions.ConnectionError:
        root7 = False
        roots.append(root7)
        # print(page.text)
        # quit()

        s.close()

    return roots


def Get_Assets(url_num, root):
    result = []
    search_val = ''
    if url_num == 12:
        search_val = "Cash"
    elif url_num == 13:
        search_val = "US stock"
    elif url_num == 14:
        search_val = "Non-US stock"
    elif url_num == 15:
        search_val = "US bond"
    elif url_num == 16:
        search_val = "Non-US bond"
    elif url_num == 17:
        search_val = "Other"
    # print(url_num)
    for val in range(1, 7):
        val = str(val)
        # print(val)
        try:
            x_path = '/html/body/div[3]/div[3]/section/div[1]/div/div/div[2]/table/tbody/tr[' + val + ']/td[1]/span/text()'

            tree = root.getroottree()
            result = tree.xpath(x_path)
            # print(result)

            if result[0].strip() == search_val:
                tree = root.getroottree()
                x_path = '/html/body/div[3]/div[3]/section/div[1]/div/div/div[2]/table/tbody/tr[' + val + ']/td[2]/text()'
                result = tree.xpath(x_path)
                # print(result)
                break

        except IndexError:
            continue
    try:
        val = result[0]
        val = str(val).rstrip("%")
        float(val)
    except (ValueError, IndexError):
        result = ["NA"]
        #print(11)
    return result


def Get_Criteria(name, X_path, url_num, roots, criteria_name):
    root = None
    if url_num == 1:
        root = roots[0]
    elif url_num == 2:
        root = roots[1]
    elif url_num == 3:
        root = roots[2]
    elif url_num == 4:
        root = roots[3]
    elif url_num == 5:
        root = roots[4]
    elif url_num == 6:
        return name
    elif url_num == 7:
        root = roots[5]

    if root == False:
        return "NA"

    if url_num >= 12:
        result = Get_Assets(url_num, roots[6])
    else:
        tree = root.getroottree()
        result = tree.xpath(X_path)
    if len(result) == 0:
        if url_num == 4:
            result = Etf_Data(criteria_name, result, root)
            if len(result) == 0:
                result = "NA"
            else:
                try:
                    result = result[0]
                    result = result.strip()
                except IndexError:
                    x = 0
        else:
            result = "NA"
    else:
        try:
            result = result[0]
            result = result.strip()
        except (IndexError):
            x = 0

    if url_num == 5:
        result = result.replace(" Research Report", "")
        result = result.replace("Risk", "")
    if '★' in result:
        result = str(len(result))

    if result == "0.00%" or result == "0":
        result = "NA"

    if result.strip() == "N/A":
        result = "NA"
    result = result.replace('/', '-')
    return result


def Etf_Data(criteria_name, result, root):
    if criteria_name.strip() == "Std Dev: 3 years":
        tree = root.getroottree()
        result = tree.xpath('//*[@id="Col1-0-Risk-Proxy"]/section/div/div/div[6]/div[2]/span[1]/text()')
    elif criteria_name.strip() == "Std Dev: 5 years":
        tree = root.getroottree()
        result = tree.xpath('//*[@id="Col1-0-Risk-Proxy"]/section/div/div/div[6]/div[3]/span[1]/text()')
    elif criteria_name.strip() == "Std Dev: 10 years":
        tree = root.getroottree()
        result = tree.xpath('//*[@id="Col1-0-Risk-Proxy"]/section/div/div/div[6]/div[4]/span[1]/text()')
    elif criteria_name.strip() == "Sharpe Ratio: 3 years":
        tree = root.getroottree()
        result = tree.xpath('//*[@id="Col1-0-Risk-Proxy"]/section/div/div/div[7]/div[2]/span[1]/text()')
    elif criteria_name.strip() == "Sharpe Ratio: 5 years":
        tree = root.getroottree()
        result = tree.xpath('//*[@id="Col1-0-Risk-Proxy"]/section/div/div/div[7]/div[3]/span[1]/text()')
    elif criteria_name.strip() == "Sharpe Ratio: 10 years":
        tree = root.getroottree()
        result = tree.xpath('//*[@id="Col1-0-Risk-Proxy"]/section/div/div/div[7]/div[4]/span[1]/text()')
    return result


def check_slash(result):
    if "/" in result:
        print("hi")


def append_list_as_row(file_name, list_of_elem):
    with open(file_name, 'a', newline='') as write_obj:
        # Create a writer object from csv module
        csv_writer = writer(write_obj)
        # Add contents of list as last row in the csv file
        csv_writer.writerow(list_of_elem)


def Set_ETF_MF(result, symbol):
    star_list = []
    if result.strip() == "MF":
        star_list = Get_Stars_MF(symbol)
    elif result.strip() == "ETF":
        star_list = Get_Stars_ETF(symbol)
    elif symbol == "NA":
        print("MF")
        star_list = Get_Stars_MF("MF")
        if len(star_list) == 0:
            print("ETF")
            star_list = Get_Stars_ETF("ETF")
    else:
        return

    return star_list


def Get_Stars_MF(symbol):
    s = requests.session()
    url = "https://research.tdameritrade.com/grid/public/mutualfunds/profile/ratingsandriskBuffer.asp?symbol=" + symbol
    page = s.get(
        url,
        headers={'User-Agent': 'Custom'})
    root = html.fromstring(page.text)
    tree = root.getroottree()
    starlist = []
    val1 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[5]')
    if val1 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val1))

    val2 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[5]')
    if val2 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val2))

    val3 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[5]')
    if val3 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val3))
    val4 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[5]')
    if val4 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val4))

    return starlist


def Get_Stars_ETF(symbol):
    s = requests.session()
    url = "https://research.tdameritrade.com/grid/public/etfs/profile/ratingsandriskBuffer.asp?symbol=" + symbol
    page = s.get(
        url,
        headers={'User-Agent': 'Custom'})

    root = html.fromstring(page.text)
    tree = root.getroottree()
    starlist = []
    val1 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[2]/span/img[5]')
    if val1 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val1))

    val2 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[3]/span/img[5]')
    if val2 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val2))

    val3 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[4]/span/img[5]')
    if val3 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val3))

    val4 = Get_Stars_From_Xpath(tree,
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[1]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[2]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[3]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[4]',
                                '//*[@id="table-morningstarRatingsTable"]/tbody/tr[1]/td[5]/span/img[5]')
    if val4 == 0:
        starlist.append("N/A")
    else:
        starlist.append(str(val4))

    return starlist


def Get_Stars_From_Xpath(tree, x_path_one, x_path_two, x_path_three, x_path_four, x_path_five):
    result = tree.xpath(x_path_one)
    result += tree.xpath(x_path_two)
    result += tree.xpath(x_path_three)
    result += tree.xpath(x_path_four)
    result += tree.xpath(x_path_five)

    return len(result)


def Update_bar():
    bar.start(70)
    bar.update()


def run_program(output_file, input_file, fund_type):
    #output_file = output_file
    open(output_file, "w").close()
    lists_from_csv = []
    f = open("ProgramFiles/XPaths/Yahoo_X_paths.csv", 'r')
    csv_reader = csv.reader(f)
    header = False
    header_list = []
    for row in csv_reader:
        lists_from_csv.append(row)
    bonds = open(input_file, 'r')
    lines = bonds.readlines()
    results_list = []
    count = 0
    Update_bar()
    for line in lines:
        Window_Update(count, fund_type, line, lines)
        star_list = []
        line = line.strip()
        count += 1
        print("Gathering Data on: " + line)
        print(str(count) + "/" + str(len(lines)))

        page_roots = Get_roots(line)
        if page_roots != False:
            for item in lists_from_csv:
                if header == False:
                    header_list.append(item[0])
                val = str(item[1] + "/text()")
                if len(val) > 2:
                    if int(item[2]) == 8:
                        try:
                            result = star_list[0]
                        except (TypeError, IndexError):
                            result = "N/A"

                    elif int(item[2]) == 9:
                        try:
                            result = star_list[1]
                        except (TypeError, IndexError):
                            result = "N/A"
                    elif int(item[2]) == 10:
                        try:
                            result = star_list[2]
                        except (TypeError, IndexError):
                            result = "N/A"
                    elif int(item[2]) == 11:
                        try:
                            result = star_list[3]
                        except (TypeError, IndexError):
                            result = "N/A"

                    else:
                        result = Get_Criteria(line, str(val), int(item[2]), page_roots, item[0])
                        if result == 'Foreign Small/Mid Growth':
                            print(11)
                        try:
                            star_list_hold = Set_ETF_MF(result.strip(), line)
                            if star_list_hold != None:
                                star_list = star_list_hold
                        except AttributeError:
                            x = 0
                    try:
                        result = result.rstrip("%")
                        result = float(result)
                    except ValueError:
                        x = 0
                    results_list.append(result)
                else:
                    continue
            if header == False:
                append_list_as_row(output_file, header_list)
            append_list_as_row(output_file, results_list)
            results_list.clear()
            header = True
    f.close()
    bonds.close()


def Window_Update(count, fund_type, line, lines):
    title_text.config(text="Running...")
    label_text.config(text="Gathering data on: " + line)
    data_process = (str(count) + "/" + str(len(lines)) + " " + fund_type + " Completed")
    progress_text.config(text=data_process)
    window.update()


def Done():
    title_text.config(text="Done")
    label_text.config(text="Data compiled check spreadsheets")
    window.update()


def funs_main():
    label_text.config(text="Loading Stock Funds....")
    window.update()
    time.sleep(2)
    run_program("ProgramFiles/Stock_mutual_funds.csv", "Stock_Funds.txt", "Stock Funds")
    print("Gathering Bond Funds Data")
    label_text.config(text="Loading Bond Funds....")
    window.update()
    time.sleep(5)
    run_program("ProgramFiles/Bond_Mutual_funds.csv", "Bond_funds.txt", "Bond Funds")

    label_text.config(text="Processing Data")
    progress_text.config(text="")
    window.update()
    time.sleep(5)
    go = True
    while go:
        try:
            Get_Style_Data.Stylize()
            go = False
        except FileCreateError:
            label_text.config(text="Please close both excel files to continue")
            window.update()

    Done()
    Stop()


def Stop():
    label_text.config(text="Program Ending....")

    window.update()
    window.quit()
    sys.exit()
    t1.join()


window = Tk()

window.title('FundBot')
window.iconbitmap("ProgramFiles/icon.ico")
window.geometry('500x300')
title_text = Label(window, text="Starting...", font=("Arial", 15))
title_text.place(
    rely=0.05, relx=0.4)
bar = Progressbar(window, orient=HORIZONTAL, length=400, mode='indeterminate')

bar.pack(pady=70)

label_text = Label(window, text="", font=("Arial", 15))
label_text.place(relx=0.28,
                 rely=0.4, )

progress_text = Label(window, text="", font=("Arial", 10))
progress_text.place(relx=0.32,
                    rely=0.5, )
t1 = threading.Thread(target=funs_main)
button_two = Button(window, text="End Program", command=Stop).pack(pady=20, )
button = Button(window, text="Start Getting Data", command=t1.start())

window.mainloop()
