import time
import requests
import json
import xlwings as xw
from datetime import date


# Shift data typr and convert it into list of list
def shifter(list, title):
    output = [[title]]
    for item in list:
        temp = []
        temp.append(item)
        output.append(temp)
        output.append([''])
        output.append([''])
        output.append([''])
        output.append([''])
        output.append([''])
        output.append([''])
        output.append([''])

    return output


def rename_date(prices):
    classes = ['C', 'P']
    for i in classes:
        for j in prices[i]:
            for k in prices[i][j]:
                year = '20' + k[4] + k[5]
                month = k[2] + k[3]
                day = k[0] + k[1]
                diff = date(int(year), int(month), int(day)) - date.today()
                prices[i][j][str(diff.days)] = prices[i][j][k]
                del prices[i][j][k]

    return prices


def sort(file, prices,sheet):

        wb = xw.Book(file)
        sheet = wb.sheets[sheet]
        strikes = []
        sheet.range('K102').value = 'latest update'
        sheet.range('K103').value = date.today()
        sheet.range('N103').value = time.localtime().tm_min
        sheet.range('L103').value = time.localtime().tm_hour

        for i in prices['P'].keys():
            strikes.append(int(i))
        strikes.sort()


        index = shifter(strikes, "Strikes")
        sheet.range('K1').value = index
        l1 = ['L', 'M', 'N', "O", 'P','Q','R','S','T','U','V','W']
        k = 2
        for i in strikes:

            mat = []

            for j in prices['P'][str(i)].keys():


                year = '20' + j[4] + j[5]
                month = j[2] + j[3]
                day = j[0] + j[1]
                diff = date(int(year), int(month), int(day)) - date.today()
                l = [str(j) , diff.days]
                mat.append(l)


            mat.sort(key = lambda x: x[1])
            for f in range(len(mat)):
                for j in range(len(l1)):
                    sheet.range((l1[f] + str(k))).value = mat[f][1]
                    sheet.range((l1[f] + str(k + 1))).value = prices['P'][str(i)][mat[f][0]]
                print(f)


            k = k + 2
        strikes = []
        for i in prices['C'].keys():
            strikes.append(int(i))
        strikes.sort()

        index = shifter(strikes, "Strikes")
        sheet.range('K1').value = index
        l1 = ['J','I','H','G','F', 'E', 'D', 'C', 'B', 'A']
        k = 2
        for i in strikes:

            mat = []
            for j in prices['C'][str(i)].keys():
                year = '20' + j[4] + j[5]
                month = j[2] + j[3]
                day = j[0] + j[1]
                diff = date(int(year), int(month), int(day)) - date.today()
                l = [str(j), diff.days]
                mat.append(l)

            mat.sort(key = lambda x: x[1])
            for f in range(len(mat)):
                for j in range(len(l1)):
                    sheet.range((l1[f] + str(k))).value = mat[f][1]
                    sheet.range((l1[f] + str(k + 1))).value = prices['C'][str(i)][mat[f]]
                print(f)

            k = k + 2
        wb.save()


def sort_full(file, prices, sheet):
    wb = xw.Book(file)
    sheet = wb.sheets[sheet]
    strikes = []
    sheet.range('K102').value = 'latest update'
    sheet.range('K103').value = date.today()
    sheet.range('N103').value = time.localtime().tm_min
    sheet.range('L103').value = time.localtime().tm_hour

    for i in prices['P'].keys():
        strikes.append(int(i))
    strikes.sort()

    index = shifter(strikes, "Strikes")
    sheet.range('K1').value = index
    l1 = ['L', 'M', 'N', "O", 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W']
    k = 2
    for i in strikes:

        mat = []

        for j in prices['P'][str(i)].keys():
            year = '20' + j[4] + j[5]
            month = j[2] + j[3]
            day = j[0] + j[1]
            diff = date(int(year), int(month), int(day)) - date.today()
            l = [str(j), diff.days]
            mat.append(l)

        mat.sort(key=lambda x: x[1])
        for f in range(len(mat)):
            for j in range(len(l1)):
                sheet.range((l1[f] + str(k))).value = mat[f][1]
                sheet.range((l1[f] + str(k + 1))).value = prices['P'][str(i)][mat[f][0]][0]
            print(f)

        k = k + 8
    strikes = []
    for i in prices['C'].keys():
        strikes.append(int(i))
    strikes.sort()

    index = shifter(strikes, "Strikes")
    sheet.range('K1').value = index
    l1 = ['J', 'I', 'H', 'G', 'F', 'E', 'D', 'C', 'B', 'A']
    k = 2
    for i in strikes:

        mat = []
        for j in prices['C'][str(i)].keys():
            year = '20' + j[4] + j[5]
            month = j[2] + j[3]
            day = j[0] + j[1]
            diff = date(int(year), int(month), int(day)) - date.today()
            l = [str(j), diff.days]
            mat.append(l)

        mat.sort(key=lambda x: x[1])
        for f in range(len(mat)):
            for j in range(len(l1)):
                sheet.range((l1[f] + str(k))).value = mat[f][1]
                sheet.range((l1[f] + str(k + 1))).value = prices['C'][str(i)][mat[f][0]][0]
            print(f)

        k = k + 8
    wb.save()


# Get option and future products details
def get_products_list(req_link):
    print('getting products list.')
    headers = {
        'Accept': 'application/json'
    }
    r = requests.get(req_link, params={}, headers=headers)
    Product_id = []
    Product_ticker = []
    Product_type = []
    Respond_to_text = r.text
    prod = []
    Text_loaded_json = json.loads(Respond_to_text)
    for i in (Text_loaded_json.get("result")):
        Product_ticker.append(i.get("symbol"))
        Product_id.append(i.get("id"))
        Product_type.append(i.get("contract_type"))
        prod.append([i.get("symbol"), i.get("contract_type")])

    return Product_ticker, prod


# Extract option products out of all products
def get_options_list(List_of_products):
    List_of_options = []

    for i in List_of_products:
        if len(i) >= 15:
            List_of_options.append(i)

    Extracted_options = []

    for i in List_of_options:
        Extracted_options.append(i.split('-'))

    return Extracted_options


def split_option_assets(options_list):
    name_list = []
    name_list.append(options_list[0][1])
    for i in range(len(options_list)):
        for j in name_list:
            if options_list[i][1] == j:
                flag = 0
                break
            else:
                flag = 1

        if flag == 1:
            name_list.append(options_list[i][1])
    splited = {}
    for i in name_list:
        splited[i] = []
    for i in range(len(options_list)):

        for j in name_list:
            if options_list[i][1] == j:
                splited[j].append([options_list[i][0], options_list[i][2], options_list[i][3]])

    return splited



def get_options_classes(options_list):
    name_list = []
    name_list.append(options_list[0][0])
    for i in range(len(options_list)):
        for j in name_list:
            if options_list[i][0] == j:
                flag = 0
                break
            else:
                flag = 1

        if flag == 1:
            name_list.append(options_list[i][0])

    splited = {}
    for i in name_list:
        splited[i] = []
    for i in range(len(options_list)):

        for j in name_list:
            if options_list[i][0] == j:
                splited[j].append([options_list[i][1], options_list[i][2]])

    return splited


def get_options_strikes(options_classes, options_classes2):
    options_calls = options_classes['C']
    options_puts = options_classes['P']

    name_list = []
    name_list.append(options_calls[0][0])

    for i in range(len(options_calls)):
        for j in name_list:
            if options_calls[i][0] == j:
                flag = 0
                break
            else:
                flag = 1

        if flag == 1:
            name_list.append(str(options_calls[i][0]))


    splited = {}
    for i in name_list:
        splited[i] = {}

    for i in range(len(options_calls)):

        for j in name_list:
            if options_calls[i][0] == j:


                splited[j][str(options_calls[i][1])] = []

    options_classes['C'] = splited

    name_list = []
    name_list.append(options_puts[0][0])

    for i in range(len(options_puts)):
        for j in name_list:
            if options_puts[i][0] == j:
                flag = 0
                break
            else:
                flag = 1

        if flag == 1:
            name_list.append(str(options_puts[i][0]))

    splited = {}
    for i in name_list:
        splited[i] = {}

    for i in range(len(options_puts)):

        for j in name_list:
            if options_puts[i][0] == j:

                splited[j][str(options_puts[i][1])] = []


    options_classes['P'] = splited

    return options_classes



def get_price(option_splited, asset, option_splited2):
    class_list = ['C', 'P']
    strike_list = []
    maturity_list = []
    b = 0
    for i in class_list:
        for j in option_splited[i]:
            for k in option_splited[i][j]:
                try:
                    code = i + '-' + asset + '-' + j + '-' + k
                    address = 'https://api.delta.exchange/v2/l2orderbook/' + code

                    headers = {
                        'Accept': 'application/json'
                    }
                    price = []
                    r = requests.get(address, params={}, headers=headers)
                    s = r.text
                    z = json.loads(s)
                    price.append((z.get("result")).get("sell"))
                    price = price[0].get("price")
                    b = b + 1
                    print(code)
                except:
                    print('rid')
                    continue

                option_splited[i][j][k].append(price)

    return option_splited

def get_price_full(option_splited, asset, option_splited2):
    class_list = ['C', 'P']
    strike_list = []
    maturity_list = []
    b = 0
    for i in class_list:
        for j in option_splited[i]:
            for k in option_splited[i][j]:
                try:
                    code = i + '-' + asset + '-' + j + '-' + k
                    address = 'https://api.delta.exchange/v2/l2orderbook/' + code

                    headers = {
                        'Accept': 'application/json'
                    }
                    price = []
                    r = requests.get(address, params={}, headers=headers)
                    s = r.text
                    z = json.loads(s)
                    ps = (z.get("result")).get("sell")
                    ps = ps[0].get("price")
                    price.append([ps])
                    ps = (z.get("result")).get("buy")
                    ps = ps[0].get("price")
                    price.append([ps])
                    b = b + 1
                    print(code)
                except:
                    print('rid')
                    continue

                option_splited[i][j][k].append(price)

    return option_splited

def get_futures_price(symbols):
    pric_list = []
    for i in symbols:
        address = 'https://api.delta.exchange/v2/l2orderbook/' + i[0]
        headers = {
            'Accept': 'application/json'
        }

        r = requests.get(address, params={}, headers=headers)
        s = r.text
        z = json.loads(s)
        price = (z.get("result")).get("sell")
        price = price[0].get("price")
        pric_list.append([i[0],price])


    return  pric_list


