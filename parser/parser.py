from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

from webdriver_manager.chrome import ChromeDriverManager

from bs4 import BeautifulSoup

import time
from datetime import datetime

import pandas
from pandas.io.excel import ExcelWriter

import httplib2
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery

import logging.config
from logging_config import dict_config

from decouple import config

logging.config.dictConfig(dict_config)
logger = logging.getLogger('parser')


class ParserWB:
    """
    Class parser wildberries
    """

    def __init__(self, url: str) -> None:
        self.headers = {
            'User-Agent': config('USER_AGENT'),
            'Accept-Language': 'ru'
        }
        self.options = Options()
        self.options.add_argument(f'{self.headers}')
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.options)
        self.url = url
        self.brand_name_list = []
        self.goods_id_list = []
        self.goods_names_list = []
        self.prices_list = []
        self.url_list = []
        self.check_date_list = []
        self.credentials_file = 'wb-project-367406-427fe6028f5b.json'
        self.spreadsheet_id = config('SPREADSHEET_ID')

    def load_page(self) -> str:
        """
        Get the html code of the page
        :return: HTML code
        :rtype: str
        """

        try:
            self.driver.get(url=self.url)
            time.sleep(5)
            return self.driver.page_source

        except Exception as exception:
            logger.debug('Start process load_page')
            logger.exception('Enter correct url', exc_info=exception)

        finally:
            self.driver.close()
            self.driver.quit()

    def parser_page(self, html: str) -> None:
        """
        Page parser
        :param html: HTML code
        :type: str
        """

        soup = BeautifulSoup(html, 'lxml')
        container = soup.find_all(class_='product-card j-card-item j-good-for-listing-event')
        for block in container:
            self.parse_block(block=block)

    def parse_block(self, block: BeautifulSoup) -> None:
        """
        Single product div parser
        :param block: single product div
        :type: BeautifulSoup
        """

        brand_names = block.find(class_='brand-name')
        if not brand_names:
            logger.error('No brand_names')
            return
        brand_name = brand_names.text

        goods_id = block.attrs['data-popup-nm-id']
        if not goods_id:
            logger.error('No goods_id')
            return

        goods_names = block.find(class_='goods-name')
        if not goods_names:
            logger.error('No goods_names')
            return
        goods_name = goods_names.text

        prices = block.find(class_='price__lower-price')
        if not prices:
            logger.error('No prices')
            return
        price = prices.text.strip().replace(u'\xa0₽', '')

        url_goods = block.find(class_='product-card__wrapper')
        if not url_goods:
            logger.error('No url_goods')
            return
        url = url_goods.find('a').get('href')

        check_date = datetime.now()

        self.brand_name_list.append(brand_name)
        self.goods_id_list.append(goods_id)
        self.goods_names_list.append(goods_name)
        self.prices_list.append(price)
        self.url_list.append(url)
        self.check_date_list.append(check_date.strftime("%d.%m.%Y_%H:%M"))

        logger.info(f'{brand_name}, {goods_id}, {goods_name}, {price}, {url}, {check_date.strftime("%d.%m.%Y_%H:%M")}')
        logger.info('=' * 100)

    def save_result_excel(self) -> None:
        """
        Save result in xlsx format
        """

        df = pandas.DataFrame({'Бренд': self.brand_name_list,
                               'Артикул': self.goods_id_list,
                               'Наименование товара': self.goods_names_list,
                               'Цена, руб.': self.prices_list,
                               'Ссылка': self.url_list,
                               'Дата проверки': self.check_date_list
                               })

        try:
            with ExcelWriter(path='goods.xlsx', mode='a', if_sheet_exists='replace') as writer:
                df.sample(len(self.brand_name_list)).to_excel(writer, sheet_name=self.brand_name_list[0], index=False)

        except PermissionError as error:
            logger.debug('Start process save_result_excel')
            logger.error(f'Close the goods.xlsx file and restart the application {error}')

        except FileNotFoundError:
            df.to_excel('goods.xlsx', sheet_name=self.brand_name_list[0], index=False)

    def save_result_google_table(self) -> None:
        """
        Save result in google table
        """

        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_file,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])

        http_auth = credentials.authorize(httplib2.Http())

        service = discovery.build('sheets', 'v4', http=http_auth)

        body = {'valueInputOption': 'USER_ENTERED',
                'data': [{'range': f'{self.brand_name_list[0]}!A1:F1',
                          'majorDimension': 'ROWS',
                          'values': [['Бренд', 'Артикул', 'Наименование товара', 'Цена, руб.', 'Ссылка',
                                      'Дата проверки']]
                          },
                         {'range': f'{self.brand_name_list[0]}!A2:F{len(self.brand_name_list) + 1}',
                          'majorDimension': 'COLUMNS',
                          'values': [self.brand_name_list, self.goods_id_list, self.goods_names_list, self.prices_list,
                                     self.url_list, self.check_date_list]
                          }
                         ]
                }

        try:
            service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()

        except Exception as error:
            logger.debug('Start process save_result_google_table')
            logger.error(f'Missing sheet with brand name {error}')

    def run(self) -> None:
        """
        Launching the parser
        """

        text = self.load_page()
        self.parser_page(html=text)

        logger.info(f'Получили {len(self.brand_name_list)} элементов')

        self.save_result_excel()
        self.save_result_google_table()
