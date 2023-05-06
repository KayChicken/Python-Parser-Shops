import requests
import win32com.client as win32
import pandas as pd
import os
from bs4 import BeautifulSoup
import random 
import re
import tkinter as tk
import tkinter.messagebox as mb


list_functions = []

def get_data_pyaterochka(search,list_items):
    cookies = {
        'location_id': '12122',
        'TS01658276': '01a93f75478d4fc3708cc238ec2d3f25a23db06d24709c45a22d714dbeaeb8110f02a065152906e9c726f0ed0faf6540f883ff90f5333257f1ae663ab5efcd59b3fa3f7fda45cce42666f75b82472e85b6cc2f2ee8',
        '_ym_uid': '1671550434868939606',
        '_ym_d': '1681744133',
        'TS010a09ac': '01a93f7547a60300ca5dca0908cb0e2921c2878f547932440029fb091326a0ffdcf6141850a4e19f6fa678006396f8d4a35ce1f571',
        '_ym_isad': '1',
        '_ym_visorc': 'b',
        '_gcl_au': '1.1.261062118.1681744133',
        '_gid': 'GA1.2.1209124289.1681744134',
        'location': '%7B%22id%22%3A12122%2C%22name%22%3A%22%D0%B3.%D0%9C%D0%BE%D1%81%D0%BA%D0%B2%D0%B0%22%2C%22type%22%3A%22city%22%2C%22new_loyalty_program%22%3Atrue%2C%22site_shops_count%22%3A1223%2C%22region%22%3A%7B%22id%22%3A14%2C%22name%22%3A%22%D0%9C%D0%BE%D1%81%D0%BA%D0%B2%D0%B0%22%7D%2C%22isDefault%22%3Atrue%2C%22isConfirmed%22%3Atrue%7D',
        '_dc_gtm_UA-77780429-1': '1',
        '_ga': 'GA1.2.1209410130.1681744134',
        '_gali': '__layout',
        '_ga_0BDW7QLPEC': 'GS1.1.1681744133.1.1.1681744400.0.0.0',
    }

    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Referer': 'https://5ka.ru/special_offers',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    params = {
        'records_per_page': '18',
        'page': '1',
        'store': '31Z6',
        'ordering': '',
        'price_promo__gte': '',
        'price_promo__lte': '',
        'categories': '',
        'search': search,
    }

    response = requests.get('https://5ka.ru/api/v2/special_offers/', params=params, cookies=cookies, headers=headers).json()
    items = response.get('results')
    for item in items:      
        item_id = item.get('id')
        item_name = item.get('name')
        item_basePrice = item.get('current_prices').get('price_reg__min')
        item_salePrice = item.get('current_prices').get('price_promo__min')
        list_items.append({
             'item_id' : item_id,
            'item_name' : item_name,
            'item_basePrice' : item_basePrice,
            'item_salePrice' : item_salePrice,
            'shop' : 'Пятёрочка'
        })

def get_data_magnit(search,list_items):
    cookies = {
        'PHPSESSID': '1v3n7rgppu5gkp9o80admncbv7',
        '_gid': 'GA1.2.1411163097.1681744879',
        '_ym_uid': '1681744879734484319',
        '_ym_d': '1681744879',
        'tmr_lvid': 'f08698db43cd3cb805f4a0e260a694ba',
        'tmr_lvidTS': '1681744878587',
        '_ym_isad': '1',
        '_ym_visorc': 'w',
        '_clck': '5dcf8h|1|fau|0',
        '_clsk': 'ci4j03|1681745363624|8|1|n.clarity.ms/collect',
        '_gaexp': 'GAX1.2.GPmbGdBVQ0GDPJ9fobr4wg.19539.0',
        '_ga_MEJT7QPK6J': 'GS1.1.1681745456.1.1.1681745477.39.0.0',
        '_ga': 'GA1.1.698379430.1681744879',
        'MGNT_CLIENT_TIMEZONE': 'Europe%2FMoscow',
        'tmr_detect': '1%7C1681745485538',
        'APPID': 'app04',
        '_ga_72BLMGVWY6': 'GS1.1.1681745484.1.1.1681745888.0.0.0',
    }

    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Authorization': 'null',
        'Connection': 'keep-alive',
        'MAGNIT-APP-DEVICE-ID': '25146d116c692ed1b96ae04fd2f700d0',
        'MAGNIT-APP-OS': 'Web',
        'MAGNIT-APP-VERSION': '1',
        'Referer': 'https://dostavka.magnit.ru/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
        'X-App-Version': '1.0.0',
        'X-Device-ID': '25146d116c692ed1b96ae04fd2f700d0',
        'X-Device-Platform': 'Web',
        'X-Device-Tag': 'disabled',
        'X-Platform-Version': '537.36',
        'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    params = {
        'q': search,
        'page': '1',
        'sort': 'price_asc',
        'is_count': 'true',
        'shop_code': '230026',
        'useSiteXmlIdService': 'express',
    }

    response = requests.get(
        'https://dostavka.magnit.ru/api/catalog/product-list/search',
        params=params,
        cookies=cookies,
        headers=headers,
    ).json()
    
    items = response.get('items')
    for item in items:
        item_id = item.get('id')
        item_name = item.get('name')
        item_basePrice = item.get('offers')[0].get('previousPrice')
        item_salePrice = item.get('offers')[0].get('currentPrice')
        list_items.append({
             'item_id' : item_id,
            'item_name' : item_name,
            'item_basePrice' : item_basePrice,
            'item_salePrice' : item_salePrice,
            'shop' : 'Магнит'
        })
    

def get_data_perekrestok(search,list_items):
    rs = requests.get("https://www.perekrestok.ru/")
    cookies = rs.cookies['session']
    index = cookies.find('accessToken') + 20
    index_lastWord = cookies.find('refreshToken') - 9
    token = rs.cookies['session'][index : index_lastWord]

    headers = {
    'authority': 'www.perekrestok.ru',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'en-US,en;q=0.9',
    'auth': f'Bearer {token}',
    'content-type': 'application/json;charset=UTF-8',
    # 'cookie': '_gcl_au=1.1.1694392960.1681811579; _ym_uid=1676709163984046375; _ym_d=1681811580; tmr_lvid=95dc69a85fac712c9c1088d101024da4; tmr_lvidTS=1676709162559; flocktory-uuid=9429ee99-0ac4-472e-b987-80cf92fd4cda-2; _ymab_param=BBLL2HyWaaoigFFSmFzCYTYDO6FZHq8kxotvT7wZGDHkMcxABFih7NcFdiF33GQJZlwYNx_bTWvsS9YH0stK1DcMZMg; _gpVisits={"isFirstVisitDomain":true,"idContainer":"100024EE"}; agreements=j:{"isCookieAccepted":true,"isAdultContentEnabled":true,"isAppAppInstallPromptClosed":true}; _ym_isad=1; _gid=GA1.2.234065395.1682157008; _ym_visorc=w; httpsReferer=https%3A%2F%2Fwww.google.com%2F; tmr_detect=1%7C1682159702391; _ga=GA1.2.1773586271.1681811580; _gp100024EE={"hits":5,"vc":1,"ac":1,"a6":1}; _dc_gtm_UA-189134493-1=1; TS015bfe9d=01b7bf369096c73cfb49d1447e53a09e67979279917d22b9be238bb5bae178b2d73cc827c348e2a7ab7ecffeca69a8dd5c28f1249deb0136a8baa56f02db22906330f4b908ee442e69e6a5d8c88bd14aa3bbc38b80; _ga_5K49P5RFR8=GS1.1.1682157006.6.1.1682161444.57.0.0; session=j:{"accessToken":"eyJ0eXAiOiJKV1QiLCJhbGciOiJFUzUxMiJ9.eyJqdGkiOiJmZThhY2IyZC00OTY1LTRkNTQtYTQzYy1iOGYyNjU2MzY4ZGUiLCJpYXQiOjE2ODIxNTcwMTcsImV4cCI6MTY4MjE4NTgxNywiZCI6ImYxM2YzYmVmLTAzNzMtNDNjMy05YWRiLTMzZGEwNmFjMGUzZiIsImFwaSI6IjEuNC4xLjAiLCJpcCI6IjE5NC4xMDYuMTk2LjQyIiwidSI6IjZlY2ZlODEyLTJhODQtNGZkYS1iZDY1LTQ2NWY3Y2JlMGIyOCIsInQiOjF9.ABK3IJX5Qlvlad0jjn7U_XLGNQwHsAayFyIdMnPlQSuo2KzohnxD2LMitYiC9Q1g2fT4W2IlqIw2VSrFrFddW8LvAAp-zwMfWpnuPXiS_RHXgRV7rl2GqB3X4gJYwC_IKie8h5y43sYv8YfT9HNxFlEfLQLT_VER2piaFDWH2zPLOWM5","refreshToken":"eyJ0eXAiOiJKV1QiLCJhbGciOiJFUzUxMiJ9.eyJqdGkiOiJhYzNkNDZhMS0xNWI3LTQ1MzQtYmZlZi03M2IyZTE1ODM5YjciLCJpYXQiOjE2ODIxNTcwMTcsImV4cCI6MTY5NzcwOTAxNywiZCI6ImYxM2YzYmVmLTAzNzMtNDNjMy05YWRiLTMzZGEwNmFjMGUzZiIsImFwaSI6IjEuNC4xLjAiLCJpcCI6IjE5NC4xMDYuMTk2LjQyIiwidSI6IjZlY2ZlODEyLTJhODQtNGZkYS1iZDY1LTQ2NWY3Y2JlMGIyOCIsInQiOjJ9.AARV0eRsZXmEKJCeuFUjNuyk7qzozDNNJEcvNpT1To1cDBE2IS7lmuZ6e8GoDxNyZNfeUxt6rX5V3Mq2qSMNUe9iAHRUX-yQg_Mbq4VtD-DXhzzLjC4zk7khGTH2vVHLfzMms81UqfcLR7T2owmFYjIbaBZYgjQ0ymvGpT4DsMw4GDKp","accessTokenExpiredAt":1682185817263,"refreshTokenExpiredAt":1697709017263,"device":{"uuid":"f13f3bef-0373-43c3-9adb-33da06ac0e3f"}}',
    'origin': 'https://www.perekrestok.ru',
    'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    }

    json_data = {
    'page': 1,
    'perPage': 48,
    'filter': {
        'textQuery': search,
    },
    'withBestProductReviews': False,
    }

    response = requests.post(
    'https://www.perekrestok.ru/api/customer/1.4.1.0/catalog/product/feed',
    headers=headers,
    json=json_data
    ).json()

    items = response['content']['items']
    for item in items:
        item_id = item.get('id')
        item_name = item.get('title')
        item_salePrice = item.get('priceTag').get('price')
        item_basePrice = item.get('priceTag').get('grossPrice') if item.get('priceTag').get('grossPrice') != None else item_salePrice
        list_items.append({
            'item_id' : item_id,
            'item_name' : item_name,
            'item_basePrice' : item_basePrice / 100,
            'item_salePrice' : item_salePrice / 100,
            'shop' : 'Перекрёсток'
        })

def get_data_spar(search,list_items):
    headers = {
        'Accept': '*/*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Origin': 'https://myspar.ru',
        'Referer': 'https://myspar.ru/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-site',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    params = {
        'query': search,
        'location': '1508',
        'isWeb': 'true',
        'limit': '15',
    }

    response = requests.get('https://ysearch.myspar.ru/web', params=params, headers=headers).json()
    items = response['products']
    for item in items:
        item_id = item.get('id')
        item_name = item.get('name')
        item_salePrice = item.get('price')
        item_basePrice = item.get('price_old') if item.get('price_old') != "" else item_salePrice
        list_items.append({
            'item_id' : int(item_id),
            'item_name' : item_name,
            'item_basePrice' : float(item_basePrice.replace("₽", "")) ,
            'item_salePrice' : float(item_salePrice.replace("₽", "")),
            'shop' : 'Спар'
        })

    
        


def get_data_globus(search,list_items): 
    cookies = {
    'PHPSESSID': 'kDP1BrBvSx7K4BtD3OkFnLCHh90W0wud',
    'BX_USER_ID': 'f00e62c40bceffdf87880152f3403cc1',
    '_ym_uid': '167670916789754217',
    '_ym_d': '1682176051',
    '_gid': 'GA1.2.1380009944.1682176051',
    '_ym_isad': '1',
    '_ym_visorc': 'w',
    'globus_hyper_show_select': '1',
    'rrpvid': '384382780933773',
    'advcake_session_id': '207045a7-3517-e37d-8bcf-97003e5efc33',
    'rcuid': '63c3ebb4f3258d66d28cca75',
    '_gcl_au': '1.1.1489734571.1682176314',
    'tmr_lvid': '75584938c40540e709500b33e34d2efd',
    'tmr_lvidTS': '1682176314327',
    'globus_hyper_id': '71',
    'globus_hyper_name': '%D0%9A%D0%BB%D0%B8%D0%BC%D0%BE%D0%B2%D1%81%D0%BA',
    'url_hyper_id': '5002',
    'online_hyper_id': '5002',
    'advcake_track_url': 'https%3A%2F%2Fonline.globus.ru%2F%3Futm_source%3Dglobus.ru%26utm_medium%3Dpopup%26utm_campaign%3Donline.globus.ru%26utm_content%3Dvybor-gm',
    'advcake_utm_partner': 'online.globus.ru',
    'advcake_utm_webmaster': 'vybor-gm',
    'advcake_click_id': '',
    'token': 'eyJhbGciOiJSUzI1NiIsImtpZCI6IjBFOEYxREM5QjhCMEI5OEMzRDRFQUExRjU3ODc2MjU3OUM3OUEwOUNSUzI1NiIsInR5cCI6IkpXVCIsIng1dCI6IkRvOGR5Yml3dVl3OVRxb2ZWNGRpVjV4NW9KdyJ9.eyJuYmYiOjE2ODIxNzY5NTQsImV4cCI6MTY4MjE4MDU1NCwiaXNzIjoiaHR0cHM6Ly9pZGVudGl0eS5nbG9idXMucnUiLCJhdWQiOiJPbmxpbmVTaG9wV2ViIiwiaWF0IjoxNjgyMTc2OTU0LCJhdF9oYXNoIjoiX1NiQjdJZHpYRkgwdWlWTUdzMlJWdyIsInNpZCI6IjY0N0M5NUE1NkUyMjUzRTMxN0QwRkM4RDMyOEFFNDYzIiwic3ViIjoiNDc2RDc3RTkzRDIyNDY1NTg2MjRBQjhERDNGRTBDMTUiLCJhdXRoX3RpbWUiOjE2ODIxNzY5NTQsImlkcCI6ImxvY2FsIiwiZ2l2ZW5fbmFtZSI6ItCU0LzQuNGC0YDQuNC5IiwiZmFtaWx5X25hbWUiOiLQmtCw0LnQvdC-0LIiLCJmYXZvcml0ZV9yZXRhaWxfc3RvcmUiOjUwMDIsImN1c3RvbWVyX2RhdGFfZmlsbGVkIjp0cnVlLCJwZXJtaXNzaW9uLnBsX2NvbW11bmljYXRpb24iOnRydWUsInBlcm1pc3Npb24ucGVyc29uYWxfZGF0YSI6dHJ1ZSwiYWdyZWVtZW50LnBsIjp0cnVlLCJtZW1iZXJzaGlwX3N0YXR1cyI6IkUwMDAyIiwiaXNfcmVnaXN0ZXJlZCI6dHJ1ZSwidGVtcF9lbWFpbCI6ImthaW5vdi5kbWl0cnkzNzBAeWFuZGV4LnJ1IiwicHJlZmVycmVkX3VzZXJuYW1lIjoiNzkyNjE4OTcwNjAiLCJuYW1lIjoiNzkyNjE4OTcwNjAiLCJwaG9uZV9udW1iZXIiOiI3OTI2MTg5NzA2MCIsImFtciI6WyJvdHBfY29kZSJdfQ.t98PLTEtTaK6A73owvtPK7Z3ZQhNpA2bGlC3wsMvz1zikQEAeWt4Eoofu3yzB-0wRtlTMCPatDkt_v77ax9sqjCY0uMsolXaxW5ebGABjVU0GZOEaO9ie2bjvlx5hvUkdmgNves1JCIVuxBRklrqVz5KERVWBj6efmh0N2b-oHMEdVkLwKM-XzsNDJkZYuAuH0vL7n85PJ7_XDDdRfJh9rpaO5WO_JED0932cBMN5MnQjc2IvI9MEVr1rUIJvgmD9HGW15bWHaMlSkmIm41Mo4RYyWSnu8k-E_LlmJoftMtJ3fNFIUmL_jyIOIEKuzlLsX0R88md8CupUbjJSnt65A',
    'advcake_track_id': 'c1bb0a11-61ba-643b-1eb9-866e0ec4630a',
    '_ga_WYXVN1FFMV': 'GS1.1.1682176314.1.1.1682177113.12.0.0',
    '_ga': 'GA1.2.1456204225.1682176051',
    }

    headers = {
    'authority': 'www.globus.ru',
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-language': 'en-US,en;q=0.9',
    'referer': 'https://www.globus.ru',
    'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'same-origin',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    }

    req = requests.get(f"https://www.globus.ru/search/?q={search}",cookies=cookies,headers=headers).text
    soup = BeautifulSoup(req,'lxml')
    items = soup.find('div',class_='search-page').find('ul').find_all('a') if soup.find('div',class_='search-page').find('ul') != None else []
    for item in items:
        link = f'https://www.globus.ru{item.get("href")}'
        req = requests.get(link, headers=headers).text
        link_parser = BeautifulSoup(req, 'lxml')
        if (link_parser.find("div" , "page404") or (link_parser.find("div", "no-offer-page"))):
            continue
        item_id = int(link_parser.find("div", class_="catalog-detail__header-desc-small--smaller").text.replace("Арт. ", ""))
        item_name = link_parser.find("h1", class_="catalog-detail__title-h1").text
        item_salePrice = link_parser.find("span", class_="catalog-detail__item-price-actual-main").text
        item_basePrice = link_parser.find("span", class_="catalog-detail__item-price-old-main").text if link_parser.find("span", class_="catalog-detail__item-price-old-main") != None else item_salePrice
        list_items.append({
            'item_id' : int(item_id),
            'item_name' : item_name,
            'item_basePrice' : int(item_basePrice),
            'item_salePrice' : int(item_salePrice),
            'shop' : 'Глобус'
        })
        
def get_data_diksi(search,list_items):
    req = requests.get(f"https://dixy.ru/catalog/search.php?q={search}").text
    soup = BeautifulSoup(req, 'lxml')
    items = soup.find_all('div', class_ = 'dixyCatalogItem')
    for item in items:
        item_id = random.randint(65030,100000)
        item_name = item.find('div', class_ = 'dixyCatalogItem__hover').text.strip() if item.find('div', class_ = 'dixyCatalogItem__hover') != None else item.find('div', class_='dixyCatalogItem__title').text.strip()
        item_salePrice = f'{item.find(attrs={"itemprop": "price"}).text}.{item.find("div", class_="dixyCatalogItemPrice__kopeck").text}'
        item_basePrice = item.find("div", class_="dixyCatalogItemPrice__oldprice").text.strip() if item.find("div", class_="dixyCatalogItemPrice__oldprice") != None else item_salePrice
        list_items.append({
            'item_id' : int(item_id),
            'item_name' : item_name,
            'item_basePrice' : float(re.sub(r"[\n\t\s]*", "", item_basePrice)),
            'item_salePrice' : float(re.sub(r"[\n\t\s]*", "", item_salePrice)),
            'shop' : 'Дикси'
        })


def get_data_ashan(search,list_items):
    headers = {
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.9',
    'Connection': 'keep-alive',
    'Origin': 'https://www.auchan.ru',
    'Referer': 'https://www.auchan.ru/',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'cross-site',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    }

    response = requests.get(
        f'https://sort.diginetica.net/search?st={search}&apiKey=06U4652632&strategy=vectors_extended,zero_queries&fullData=true&withCorrection=true&withFacets=true&treeFacets=true&regionId=1&useCategoryPrediction=false&size=20&offset=0&showUnavailable=false&unavailableMultiplier=0.2&preview=false&withSku=false&sort=DEFAULT',
        headers=headers,
    ).json()
    items = response['products']
    for item in items:
        item_id = item.get('id')
        item_name = item.get('name')
        item_basePrice = item.get('oldPrice') if item.get('oldPrice') != None else item.get('price')
        item_salePrice = item.get('price')
        list_items.append({
            'item_id' : int(item_id),
            'item_name' : item_name,
            'item_basePrice' : float(item_basePrice),
            'item_salePrice' : float(item_salePrice),
            'shop' : 'Ашан'
        })



def get_data_yarche(search,list_items):
    headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'en-US,en;q=0.9',
    'Connection': 'keep-alive',
    'Content-Type': 'application/json',
    'Origin': 'https://yarcheplus.ru',
    'Referer': 'https://yarcheplus.ru/',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'cross-site',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'token': '',
    }

    json = {"search":search,"page":{"page":1,"limit":48},"filter":{"byCategoryIds":[]},"productFilterTimeRangeUnit":"month"}
    json_data = {
    'query': '\n  query ($search: String!, $sort: SortInput, $page: PageInput, $filter: ProductsFilterRequest, $productFilterTimeRangeUnit: PeriodEnum) {\n    search (search: $search, sort: $sort, page: $page, filter: $filter) {\n      \n  products {\n  id\n  code\n  name\n  description\n  amount\n  image {\n  id\n  title\n  alt\n}\n  price\n  previousPrice\n  itemSum\n  isNew\n  isHit\n  isVkusvill\n  isFavorite\n  isAvailable\n  isSubscribed\n  isVeterinaryControl\n  quant {\n  code\n  fullName\n  shortName\n  multiple\n  pricePerUnit\n  previousPricePerUnit\n  unit\n  type\n  minAmount\n  stepAmount\n  amountPerQuant\n  currency\n  unitCode\n}\n  categories {\n  id\n  name\n  code\n  treeId\n  parentTreeId\n}\n  rating\n  numberOfRatings\n  brand\n}\n  productCategories {\n  \n  id\n  name\n  code\n  treeId\n  parentTreeId\n\n  description\n  amount\n  image {\n  id\n  title\n  alt\n}\n  metaTags {\n  key\n  value\n}\n  isBold\n}\n  taps {\n  key\n  value\n}\n  filters {\n  type\n  name\n  title\n  ... on FilterList {\n    options {\n  id\n  value\n  label\n  count\n  isApplied\n}\n  }\n  ... on FilterTimeRange {\n    minValue (unit: $productFilterTimeRangeUnit)\n    maxValue (unit: $productFilterTimeRangeUnit)\n  }\n}\n  sort {\n  param\n  title\n  direct\n  isApplied\n}\n  page {\n  total\n  limit\n  page\n}\n\n    }\n  }\n',
    'variables': json,
    }

    get_token = requests.get("https://yarcheplus.ru/",headers=headers)
    headers['token'] = get_token.text[get_token.text.find('token') : get_token.text.find('token') + 100].split('"')[2]
    req = requests.post('https://api.magonline.ru/api/graphql', headers=headers, json=json_data).json()
    items = req['data']['search']['products']
    for item in items:
        print(item)
        item_id = item.get('id')
        item_name = item.get('name')
        item_salePrice = item.get('price')
        item_basePrice = item.get('previousPrice') if item.get('previousPrice') != None else item_salePrice
        list_items.append({
            'item_id' : int(item_id),
            'item_name' : item_name,
            'item_basePrice' : float(item_basePrice),
            'item_salePrice' : float(item_salePrice),
            'shop' : 'Ярче'
        })

    



class App(tk.Tk):
    def __init__(self):
            super().__init__()
            btn_info = tk.Button(self, text="Информационное окно",
                                command=self.show_info)
            btn_warn = tk.Button(self, text="Окно с предупреждением",
                                command=self.show_warning)
            btn_error = tk.Button(self, text="Окно с ошибкой",
                                command=self.show_error)

            opts = {'padx': 40, 'pady': 5, 'expand': True, 'fill': tk.BOTH}
            btn_info.pack(**opts)
            btn_warn.pack(**opts)
            btn_error.pack(**opts)

    def show_info(self):
            msg = "Отчёт был успешно составлен"
            mb.showinfo("Информация", msg)

    def show_warning(self):
            msg = "Временные файлы удалены не правильно"
            mb.showwarning("Предупреждение", msg)

    def show_error(self):
            msg = "Приложение обнаружило неизвестную ошибку"
            mb.showerror("Ошибка", msg)


    
def convertToExcel(list_items, file_name = "my_file"):
        print(list_items)
        df = pd.json_normalize(list_items)
        df.to_csv('file.csv')
        if (len(list_items) != 0):
             df = df.sort_values('item_salePrice')
        df.to_excel(f'{file_name}.xlsx', index=None, header=True)
        xlApp = win32.Dispatch('Excel.Application')
        wb = xlApp.Workbooks.Open(os.path.abspath(f'{file_name}.xlsx'))
        ws = wb.Worksheets['Sheet1']
        ws.Columns.AutoFit()
        wb.Save()
        xlApp.Application.Quit()
        App.show_info(tk.Tk)
        
        

def main():
    window = tk.Tk()
    window.title("Парсер магазинов")
    window.geometry('300x400')
    window.config(bg="#f7f6f2") #F2994A #F2C94C #B06AB3 #fad0c4

    def fun_button():
        list_items = []
        check_box(item_entry.get(),list_items)
        for func in list_functions:
            func
            
        file_name = file_entry.get() if len(file_entry.get()) != 0 else "my_file"
        convertToExcel(list_items,file_name)

    def check_box(item,list_items):
        if (pyterochka.get() == 1):list_functions.append(get_data_pyaterochka(item,list_items))
        if (magnit.get() == 1):list_functions.append(get_data_magnit(item,list_items))
        if (diksi.get() == 1):list_functions.append(get_data_diksi(item,list_items))
        if (perekrestok.get() == 1):list_functions.append(get_data_perekrestok(item,list_items))
        if (spar.get() == 1):list_functions.append(get_data_spar(item,list_items))
        if (ashan.get() == 1):list_functions.append(get_data_ashan(item,list_items))
        if (globus.get() == 1):list_functions.append(get_data_globus(item,list_items))
        if (yarche.get() == 1):list_functions.append(get_data_yarche(item,list_items))
        
            

    label = tk.Label(window, text="Поиск продуктов в магазинах" , background="#f7f6f2" , font=("Arial", 14))
    label.pack()

    item_label = tk.Label(window, text="Название товара" , background="#f7f6f2")
    item_label.pack()
    item_entry = tk.Entry(window, width=40)
    item_entry.pack()

    shops_label = tk.Label(window, text="Магазины" , background="#f7f6f2")
    shops_label.pack()


    pyterochka = tk.IntVar()
    checkButton_pyterochka = tk.Checkbutton(window , text="Пятёрочка" , background="#f7f6f2", variable=pyterochka )
    checkButton_pyterochka.pack()

    magnit = tk.IntVar()
    checkButton_magnit = tk.Checkbutton(window , text="Магнит" ,background="#f7f6f2", variable=magnit )
    checkButton_magnit.pack()

    perekrestok = tk.IntVar()
    checkButton_perekrestok = tk.Checkbutton(window , text="Перекрёсток" , background="#f7f6f2" , variable=perekrestok)
    checkButton_perekrestok.pack()

    diksi = tk.IntVar()
    checkButton_diksi = tk.Checkbutton(window , text="Дикси" , background="#f7f6f2", variable=diksi)
    checkButton_diksi.pack()

    spar = tk.IntVar()
    checkButton_spar = tk.Checkbutton(window , text="Спар" , background="#f7f6f2",variable=spar)
    checkButton_spar.pack()

    ashan = tk.IntVar()
    checkButton_ashan = tk.Checkbutton(window , text="Ашан" , background="#f7f6f2",variable=ashan)
    checkButton_ashan.pack()

    globus = tk.IntVar()
    checkButton_globus = tk.Checkbutton(window , text="Глобус" , background="#f7f6f2",variable=globus)
    checkButton_globus.pack()

    yarche = tk.IntVar()
    checkButton_yarche = tk.Checkbutton(window , text="Ярче" , background="#f7f6f2",variable=yarche)
    checkButton_yarche.pack()

    file_frame = tk.Frame(window , background="#f7f6f2")
    file_frame.pack(pady=10)
    file_label = tk.Label(file_frame,text="Введите название файла",  background="#f7f6f2")
    file_label.pack()
    file_entry = tk.Entry(file_frame , width=40)
    file_entry.pack()

    button = tk.Button(window , text="Сгенерировать информацию", command=fun_button , background="#f7f6f2")
    button.pack(pady=10)

    window.mainloop()

main()