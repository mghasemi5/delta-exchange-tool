
from options import sort, split_option_assets, get_options_list, get_products_list,\
    get_options_classes, get_options_strikes, get_price

while True:


   try:

        symbol_list = get_products_list('https://api.delta.exchange/v2/products')
        options_list = get_options_list(symbol_list)
        option_assets = split_option_assets(options_list)
        eth = option_assets['BTC']
        eth_classes = get_options_classes(eth)
        options_splitted = get_options_strikes(eth_classes, eth_classes)
        print(options_splitted)
        prices = get_price(options_splitted, 'BTC', options_splitted)
        sort('btc.xlsx', prices)
   except:
           continue



