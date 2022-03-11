from options import sort, split_option_assets, get_options_list, get_products_list,\
    get_options_classes, get_options_strikes, get_price , get_futures_price
import xlwings as xw
import http

headers = {
    'Accept': 'application/json'
}

while True:


        symbol_list , r= get_products_list('https://api.delta.exchange/v2/products')

        list_of_futures = []
        perp_futures = []
        for i in r:
            if i[1] == 'futures':
                list_of_futures.append(i)



        for i in r:
            if i[1] == 'perpetual_futures':
                for j in list_of_futures:
                    if j[0].split('_')[0] == i[0]:

                         perp_futures.append(i)


        list_of_futures.sort(key = lambda x: x[0])
        perp_futures.sort(key = lambda x: x[0])

        futures = get_futures_price(list_of_futures)
        future_perp = get_futures_price(perp_futures)
        wb = xw.Book('b1.xlsx')
        sheet = wb.sheets[2]
        sheet.range('A3').value = futures
        sheet.range('C3').value = future_perp
