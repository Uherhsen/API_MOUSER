# API_MOUSER
Finding products on MOUSER by API and using parsing. Automatic creation of product descriptions with GUI.

В файле get_mouser_product_data.py следует заменить API_KEY на свой. Для получения ключа нужно зарегистрироваться в кабинете разработчика сайта MOUSER.
Страница с информацией:
https://ru.mouser.com/api-search/
Руководство: 
https://api.mouser.com/api/docs/ui/index

Программа API_MOUSER использует API сайта https://ru.mouser.com/ и простой парсер для получения параметров товара с сайта по артикулу. Программа автоматически выделяет и 
записывает категории товаров в базу json, после чего пользователь привязывает шаблон для автоматического описания (шаблоны хранятся в этой же базе). 
Вся информация полезная информация по продуктам сохраняется в файле productdata.xlsx. Для удобства работы создан графический интерфейс с редактором простых шаблонов.  
