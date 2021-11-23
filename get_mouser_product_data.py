# -*- coding: utf-8 -*-
"""
Created on Sat Mar  6 10:19:53 2021

Получение массива данных с MOUSER по артикулу 
"""

import requests,json
from bs4 import BeautifulSoup


class GetMouserProductData():
    
    #HEADERS = {'user-agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36 OPR/74.0.3911.107',
             #'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
             #'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7'}
    
    HEADERS = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36',
               'accept': '*/*', 'accept-language':'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7','content-type': 'text/plain','origin': 'https://ru.mouser.com','referer': 'https://ru.mouser.com/'}

    '''Инициализация переменной-индивидуального ключа для Mouser API'''
    def __init__(self, api_key):
        self.API_KEY = api_key
    
    def SearchPart(self, partname):
        '''получение по поиску SearchByKeywordRequest'''
        SearchApi = 'https://api.mouser.com/api/v1/search/keyword'
        
        models = '{ SearchByKeywordRequest: { "keyword": "%s", "records": 1, "startingRecord": 1, "searchWithYourSignUpLanguage":true }}' %  str(partname)
        # важно перекодировать в ютф 8 чтобы русские символы не глючили
        models = models.encode('utf-8')
        headers = {
            'accept': 'application/json', #application/json
            'Content-Type': 'application/json; charset=UTF-8',#application/json text/xml
            
        }
        params = (
            ('apiKey', self.API_KEY), 
        )
        response = requests.post( SearchApi, headers=headers, params=params, data=models)
        binary = response.content
        output = json.loads(binary)
        if output['Errors'] == []:
            return output
        else:
            print('Ошибка ответа API: ',output)
            return None
    
    def productURL(self, datadict):
        '''Получение прямой ссылки на продукт для парсера'''
        try:
            return datadict['SearchResults']['Parts'][0]['ProductDetailUrl']
        except KeyError:
            print('Не удалось получить URL детали из массива')
            return None
        except IndexError:
            print('Деталь не найдена по API')
            return None
    #___________________________________________________________________parse
    def get_html(self, url):
        '''ПАРСЕР: Получение по прямой ссылке на товар кода страницы'''
        if url == None:
            print('Нет URL детали')
            return None
        else:
            r = requests.get(url, headers=self.HEADERS)
            if r.status_code == 200:
                print('Status code 200\n')
                #print("Первая цифра статуса",str(r.status_code)[0])
                return r
                
            else:
                return None
                raise Exception('ParseError, статус ответа: ', r.status_code)
                #print('ParseError, статус ответа: ', r.status_code)
                #return None
        
    def get_params(self, html):
        '''ПАРСЕР: отдает словарь с параметрами продукта'''
        if html == None:
            print('Нет html кода целевой страницы')
            return None
        else:
            soup = BeautifulSoup(html.text, 'html.parser')
            data={}
            try:
                datasheetURL = soup.find('a', id='pdp-datasheet_0').get('href')
                data['datasheetURL'] = datasheetURL
            except AttributeError:
                data['datasheetURL'] = None
            table_items = soup.find('table', class_='specs-table')
            items = table_items.find_all('tr')
            for item in items[1:]:
                data[item.find('td',class_='attr-col').get_text(strip=True).replace(':','')] = item.find('td', class_='attr-value-col').get_text(strip=True)
            
            return data
    #_____________________________________________________________________process
    def all_data(self,partname):
        APIdata = self.SearchPart(partname)
        #print(APIdata)
        if APIdata == None:
            return None
        elif APIdata['SearchResults']['NumberOfResult'] == 0:
            return None
        else:
            URL = self.productURL(APIdata)
            if URL == None:
                return {'APIdata': APIdata,'ParserData':None}
            else:
                html = self.get_html(URL)
                ParserData = self.get_params(html)
                return {'APIdata': APIdata,'ParserData':ParserData}
        

if __name__ == "__main__":

    API_KEY = "4aaaca99-d258-45ba-ad1d-aca303dd0a8a"# "97463f63-1d87-497e-8067-1524f6cc016b"# "23fd309c-4214-48da-82bc-9f217e4acc08"
    part="TPS76618DR"#"усилитель".encode('utf-8') "TPS65994ADRSLR"
    #part="ABRACADABRA"
    #part='FMMT2222A'
    #print(part)
    #part="ECAP 400V 100uF"
    #part='абракадабра'
    GMPD = GetMouserProductData(API_KEY)
    data = GMPD.all_data(part)
    #print(data['ParserData']['datasheetURL'])
    print(data)