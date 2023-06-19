import openpyxl
from pybliometrics.scopus import ScopusSearch
import pandas as pd
import requests
import os

def q_filter(path, sheet_name, title_column, q_column, q_factor):
    """
        :param path: takes the pass to input exel file
        :param sheet_name: takes the name of sheet for analysis
        :param title_column: letter of  column with title of journals
        :param q_column: letter of column with Q factor
        :param q_factor: required Q factor
    """

    print('Searching has been started...')
    book = openpyxl.load_workbook(path)
    sheet = book[sheet_name]
    title_q = {}
    for i in range(1, sheet.max_row+1):
        title_q[sheet[title_column + str(i)].value] = sheet[q_column + str(i)].value
    q_required = []
    for key, val in title_q.items():
        if val == q_factor:
            q_required.append(f'SRCID ({key}) OR')
    print('List for request has been formed')
    if purpose == 'data':
        scopus_request_data(q_required)
    elif purpose == 'count':
        scopus_request_count(q_required)
    return


def scopus_request_data(q_required):
    """
    :param q_required: list of all journals with required quartile
    :return: exel sheets with information about publication in required year
    """
    print('data forming...')
    with open(r'key.txt',
              'r') as f:  # path to the file with your api key
        key = f.read()
    output_data = []
    stuck = 80
    t = 0
    n = len(q_required) // stuck + 1
    print('Request has been sent')
    for i in range(n):
        output_data.append(q_required[i * stuck:i * stuck + stuck])
    print('sheet creation...')
    for elem in output_data:
        query_list = ''
        for srcid in elem:
            t += 1
            query_list += srcid + ' '
        query = f'{query_list[0:len(query_list)-3]} AND pubyear is {pab_year} AND doctype(ar) OR doctype(re)'
        print(query)
        fullquery = r'https://api.elsevier.com/content/search/scopus?start=0&count=1&query= ' + str(query) + '&apiKey=' \
                    + str(key)
        export = rf'export{pab_year}{output_data.index(elem)+1}.xlsx'

        df = pd.DataFrame(pd.DataFrame(ScopusSearch(fullquery, subscriber=False,
                                                    verbose=True).results))

        df.to_excel(export, index=False)
    return


def scopus_request_count(q_required, key_path):
    """
    :param q_required: list of all journals with required quartile
    :param key_path: path to the key to scopus
    :return: amount of founded articles
    """
    with open(r'key.txt',
              'r') as f:  # path to the file with your api key
        key = f.read()
    output_data = []
    rj_count = 0
    stuck = 80
    n = len(q_required) // stuck + 1
    print('Counting has been started')
    for i in range(n):
        output_data.append(q_required[i * stuck:i * stuck + stuck])
    for elem in output_data:
        query_list = ''
        for srcid in elem:
            query_list += srcid + ' '
        query = f'{query_list[0:len(query_list) - 3]} AND pubyear is {pab_year} AND doctype(ar) OR doctype(re)'
        fullquery = r'https://api.elsevier.com/content/search/scopus?start=0&count=1&query=' + str(query) + '&apiKey=' \
                    + str(key)
        r = requests.get(fullquery)
        rj = r.json()
        rj_count += int(rj['search-results']['opensearch:totalResults'])
    return print(f'Full amount of publacation = {rj_count}')


purpose = input('What do you need to find? (please, print data, count or both): ')
path_to_file = input('Print name of initial file: ')
sheet_name_in_file = input('Print sheet name: ')
title_column_in_file = input('Print column with journal titles (A, B, C itc.): ')
Q_column_in_file = input('Print column with Q factor (A, B, C itc.): ')
Q_factor_in_file = input('Print required Q factor (Q1, Q2 itc.): ')
#path_output_to_file = os.getcwd()
pab_year = input('Print publication year: ')
q_filter(path_to_file, sheet_name_in_file, title_column_in_file, Q_column_in_file, Q_factor_in_file)
