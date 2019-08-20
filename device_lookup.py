import requests
import json
import subprocess
import sys
import os
import xlsxwriter
import webbrowser
import certifi
import urllib3
from datetime import datetime
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


authorize_url = <<authorizationurl>>
token_url = <<<Token_URL>>>
callback_uri = "http://www.example.org"
client_id = input("What is your client ID: ")
secret_key = input("What is your secret key: ")
company = input('\nCompany ID: ')
base_url = <<<base URL>>>
basic_search = <<<Base URL & parameters to begin search>>>
auth_redirect_url = (authorize_url + '?client_id=' + client_id +
                     '&response_type=code&redirect_uri=' + callback_uri)
item1_req_all = <<<portion of URL that goes behind basic_search>>>
item1_sort = <<<portion of URL that goes behind basic_search>>>
item1_sort2 = <<<portion of URL that goes behind basic_search>>>
item1_sort3 = <<<portion of URL that goes behind basic_search>>>
item2_req_all = <<<portion of URL that goes behind basic_search>>>
item2_sort = <<<portion of URL that goes behind basic_search>>>
item2_sort2 = <<<portion of URL that goes behind basic_search>>>
item2_sort3 = <<<portion of URL that goes behind basic_search>>>
item3_req_all = <<<portion of URL that goes behind basic_search>>>
item3_sort = <<<portion of URL that goes behind basic_search>>>
item3_sort2 = <<<portion of URL that goes behind basic_search>>>

print("Device Systems Lookup\n\n"
      "This Application will allow for quick lookups"
      "and bulk requests that can be exported using excel.\n\n")
print("You will now be redirected to your web "
      "browser for your authorization code.\n")
webbrowser.open(auth_redirect_url)
print("---  " + auth_redirect_url + "  ---\n")

authorization_code = input('code: \n')
data = {'grant_type': 'authorization_code',
        'code': authorization_code,
        'redirect_uri': callback_uri
        }
print("...Requesting access token\n\n")
access_token_response = requests.post(token_url, data=data,
                                      verify=False,
                                      allow_redirects=False,
                                      auth=(client_id, secret_key))

print("response below:\n")
print(access_token_response.headers)
print('body: ' + access_token_response.text)
tokens = json.loads(access_token_response.text)
access_token = tokens['access_token']
print("access token: " + access_token)


def item1_bulk(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item1_req_all +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item1_rv(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item1_sort +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item1_mp(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item1_sort2 +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item1_gx(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item1_sort3 +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item2_bulk(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item2_req_all +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item2_rv(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item2_sort +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item2_mp(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item2_sort2 +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item2_gx(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item2_sort3 +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item3_bulk(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item3_req_all +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item3_rv(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item3_sort + company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def item3_mp(search):
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(base_url +
                                     item3_sort2 +
                                     company,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def ID_search(query):
    company_end = '&company=' + company
    api_call_headers = {'Authorization': 'Bearer ' + access_token}
    api_call_response = requests.get(basic_search +
                                     query +
                                     company_end,
                                     headers=api_call_headers,
                                     verify=False)
    print(api_call_response)
    return(api_call_response).json()


def request_data(): # for individual searches(outputs data on screen)
    try:
        ID = ''
        while not ID:
            ID = input("Please provide ID:\t")
        ID_info = ID_search(imei)
        if len(ID) == 0:
            print("Please provide a proper ID:\t")
        else:
            search = ID_info
            display_data(search)
    except requests.exceptions.ConnectionError:
        print("Couldn't connect to server! Is the network up?")
        request_data()


def bulk_item3():
    print_bulk = ''
    bulk_request = item3_mp(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item3_trv():
    print_bulk = ''
    bulk_request = item3_rv(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item3_tmo():
    print_bulk = ''
    bulk_request = item3_bulk(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def bulk_item1():
    print_bulk = ''
    bulk_request = item1_bulk(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item1_vmp():
    print_bulk = ''
    bulk_request = item1_mp(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item1_vrv():
    print_bulk = ''
    bulk_request = item1_rv(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item1_vgx():
    print_bulk = ''
    bulk_request = item1_gx(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def bulk_item2():
    print_bulk = ''
    bulk_request = item2_bulk(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item2_amp():
    print_bulk = ''
    bulk_request = item2_mp(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item2_arv():
    print_bulk = ''
    bulk_request = item2_rv(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


def item2_agx():
    print_bulk = ''
    bulk_request = item2_gx(print_bulk)
    if print_bulk == '':
        print("Information sent to Excel, file location below:\n")
        display_wb(bulk_request)


# Diplays data output to your current screen
def display_data(search):
    print(f"Information for your query below:\n")
    for entry in search["items"]:
        example = entry['data']['example']
        example2 = entry['data']["example2"]
        example3 = entry['data']["example3"]
        example4 = entry['data']["example4"]
        example5 = entry['data']["example5"]
        example6 = entry['data']["example6"]
        example7 = entry['data']["example7"]
        device = entry['type']
        state = entry['subscription']["state"]
        example9 = entry['subscription']['example9']
        example10 = entry['subscription']["example10"]
        example11 = entry['subscription']["example11"]
        example12 = entry['subscription']["example12"]
        example13 = entry["example13"]
        example14 = entry['gateway']["example14"]
        example15 = entry['gateway']["example15"]
        example16 = entry['gateway']["example16"]
        example17 = entry['gateway']["example17"]
        example18 = entry['name']
        print(f"Key: {siteName}\n\nKey: {example}\tKey: {example3} \t"
              f"Key: {example2}\n\nKey: {example4}\t"
              f"Key: {example6}\tKey: {example5}\n"
              f"Key: {example7}\nKey: {example}\nKey: {example}\n"
              f"Key: {example}\tKey: {example}\nKey: {example}\n"
              f"Key: {example}\nKey: {example}\n"
              f"Key: {example}\n"
              f"Key: {example}\n"
              f"Key: {example}\nKey: {example}\n")


# Sends data to a named workbook, used for large amounts of data
def display_wb(search):
    row = 0
    col = 0
    wbname = input('Name your workbook: ')
    avwb = xlsxwriter.Workbook(wbname + '.xlsx')
    worksheet = avwb.add_worksheet("Bulk list")
    cell_format = avwb.add_format()
    cell_format.set_bg_color('cyan')
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    body_format = avwb.add_format()
    body_format.set_align('center')
    body_format.set_align('vcenter')
    body_format.set_bg_color('gray')
    worksheet.write('A1', 'Cell_header', cell_format)
    worksheet.write('B1', 'Cell_header', cell_format)
    worksheet.write('C1', 'Cell_header', cell_format)
    worksheet.write('D1', 'Cell_header', cell_format)
    worksheet.write('E1', 'Cell_header', cell_format)
    worksheet.write('F1', 'Cell_header', cell_format)
    worksheet.write('G1', 'Cell_header', cell_format)
    worksheet.write('H1', 'Cell_header', cell_format)
    worksheet.write('I1', 'Cell_header', cell_format)
    worksheet.write('J1', 'Cell_header', cell_format)
    worksheet.write('K1', 'Cell_header', cell_format)
    worksheet.write('L1', 'Cell_header', cell_format)
    worksheet.write('M1', 'Cell_header', cell_format)
    worksheet.write('N1', 'Cell_header', cell_format)
    worksheet.write('O1', 'Cell_header', cell_format)
    worksheet.write('P1', 'Cell_header', cell_format)
    worksheet.write('Q1', 'Cell_header', cell_format)
    worksheet.write('R1', 'Cell_header', cell_format)
    worksheet.write('S1', 'Cell_header', cell_format)
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 14)
    worksheet.set_column('E:E', 24)
    worksheet.set_column('F:F', 10)
    worksheet.set_column('G:G', 15)
    worksheet.set_column('H:H', 15)
    worksheet.set_column('I:I', 14)
    worksheet.set_column('J:J', 24)
    worksheet.set_column('K:K', 10)
    worksheet.set_column('L:L', 15)
    worksheet.set_column('M:M', 15)
    worksheet.set_column('N:N', 14)
    worksheet.set_column('O:O', 24)
    worksheet.set_column('P:P', 10)
    worksheet.set_column('Q:Q', 15)
    worksheet.set_column('R:R', 15)
    worksheet.set_column('S:S', 14)
    for entry in search["items"]:
        example = entry['data']['example']
        example2 = entry['data']["example2"]
        example3 = entry['data']["example3"]
        example4 = entry['data']["example4"]
        example5 = entry['data']["example5"]
        example6 = entry['data']["example6"]
        example7 = entry['data']["example7"]
        device = entry['type']
        state = entry['subscription']["state"]
        example9 = entry['subscription']['example9']
        example10 = entry['subscription']["example10"]
        example11 = entry['subscription']["example11"]
        example12 = entry['subscription']["example12"]
        example13 = entry["example13"]
        example14 = entry['gateway']["example14"]
        example15 = entry['gateway']["example15"]
        example16 = entry['gateway']["example16"]
        example17 = entry['gateway']["example17"]
        example18 = entry['name']
        row += 1
        worksheet.write(row, col, example, body_format)
        worksheet.write(row, col + 1, example2, body_format)
        worksheet.write(row, col + 2, example3, body_format)
        worksheet.write(row, col + 3, example4, body_format)
        worksheet.write(row, col + 4, example5, body_format)
        worksheet.write(row, col + 5, example6, body_format)
        worksheet.write(row, col + 6, example7, body_format)
        worksheet.write(row, col + 7, example8, body_format)
        worksheet.write(row, col + 8, example9, body_format)
        worksheet.write(row, col + 9, example10, body_format)
        worksheet.write(row, col + 10, example11, body_format)
        worksheet.write(row, col + 11, example12, body_format)
        worksheet.write(row, col + 12, example13, body_format)
        worksheet.write(row, col + 13, example14, body_format)
        worksheet.write(row, col + 14, example15, body_format)
        worksheet.write(row, col + 15, example16, body_format)
        worksheet.write(row, col + 16, example17, body_format)
        worksheet.write(row, col + 17, example18, body_format)
        worksheet.write(row, col + 18, example19, body_format)
    avwb.close()
    print(os.getcwd())


# Menu items
def all_item2():
    search = input('\nPress 1: All item2 Devices\n'
                   'Press 2: item2 Devices\n'
                   'Press 3: item2 Devices\n'
                   'Press 4: item2 Devices\n'
                   'Press 5: Main Menu: ')
    if search == "1":
        bulk_item2()
    elif search == "2":
        item2_arv()
    elif search == "3":
        item2_amp()
    elif search == "4":
        item2_agx()
    elif search == "5":
        menu()
    else:
        print('Please choose a valid option')
        menu()


def all_item1():
    search = input('\nPress 1: All item1 Devices\n'
                   'Press 2: item1 Devices\n'
                   'Press 3: item1 Devices\n'
                   'Press 4: item1 Devices\n'
                   'Press 5: Main Menu: ')
    if search == "1":
        bulk_item1()
    elif search == "2":
        item1_vrv()
    elif search == "3":
        item1_vmp()
    elif search == "4":
        item1_vgx()
    elif search == "5":
        menu()
    else:
        print('Please choose a valid option: ')
        menu()


def all_item3():
    search = input('\nPress 1: All item3 Devices\n'
                   'Press 2: item3 Devices\n'
                   'Press 3: item3 Devices\n'
                   'Press 4: Main Menu: ')
    if search == "1":
        bulk_item3()
    elif search == "2":
        item3_trv()
    elif search == "3":
        item3_tmp()
    elif search == "4":
        menu()
    else:
        print('Please choose a valid option: ')
        menu()


def menu():
    lookup = input("Please choose an option\n\n"
                   "Press 1: Single ID Search\n"
                   "Press 2: item1 Bulk Search\n"
                   "Press 3: item2 Bulk Search\n"
                   "Press 4: item3 Bulk Search\n"
                   "Press 5: Quit Program"
                   ": ")
    if lookup == "1":
        request_data()
    elif lookup == "2":
        all_item1()
    elif lookup == "3":
        all_item2()
    elif lookup == "4":
        all_item3()
    elif lookup == "5":
        print('Thank you. Logging out!')
        quit()
    else:
        print("Not a valid choice, try again: \n")
        menu()


if __name__ == '__main__':
    while True:
        menu()
