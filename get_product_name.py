# -*- coding: utf-8 -*-
"""
Created on Sun Feb  2 22:26:59 2020
регуляркb для отделения. Требуется для создания заголовков для коллекции типов продукции сайта ТМЕ
"""
import re
'''регулярка для отделения по точке-запятой первых слов из декрипшина (столбец С) 
Далее это будет называться ТИПОМ продукта'''
# НЕИСПОЛЬЗУЕТСЯ, использовалось еслибы детали классифицировались по дескрипшену
def product_name(descr):
    product = re.split(r';', descr)
    return(product[0])

"""Отделеяем типы от параметров. Новая классификация по столбцу параметров G"""
# создает из строки вида {'Тип аксессуаров для диодов LED': 'светодовод для LED', 'Форма': 'круглая', 'Размер ... параметр "Тип аксессуаров для диодов LED светодовод для LED"
def product_type(parameters):
    try:
        product = re.search('Категория.+?,', parameters).group()
        product = product.replace(":","")
        product = product.replace("'","")
        product = product.replace(",","")
        #product = product.replace("Тип","")
        #product = product.replace(":","")
        product  = product.strip()
        #print(product)
        return(product)
    except TypeError:
        return None
    except AttributeError:
        return None


if __name__ == '__main__':
    test_descr = "Вентилятор: DC; осевой; 12ВDC; 38x38x20мм; 20,39м3/ч; 44дБА; 26AWG"
    #test_parameters = "{'Тип микросхемы':: 'driver', 'Вид микросхемы': 'транзисторная матрица','Выходной ток': '0,5А', 'Выходное напряжение': '50В', 'Кол-во каналов': '7','Монтаж': 'THT', 'Рабочая температура': '-40...85°C', 'Входное напряжение': '30В', 'Корпус': 'DIP16'}"
    
    test_parameters="Категория продукта круговой мил / технические характеристики соединителя, Количество позиций 5 position, Тип контакта socket (female), Схема установки 15-05, Продукт plugs, Размер оболочки 15, Форма оболочки straight, Тип аварийной лампы mil-dtl-38999 iii, Тип выводов crimp, Вид монтажа free hanging, Номинальный ток 13 a, Покрытие контакта gold, Материал контакта copper alloy, Серия d38999 iii, Угол монтажа straight, Покрытие оболочки olive drab cadmium, Вид сопряжения self locking, Квалификация erel, Материал корпуса aluminum, Размер фабричной упаковки 1"
    print(product_name(test_descr),'\n',product_type(test_parameters))
    
    
    