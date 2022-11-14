from parser.parser import ParserWB


def main():
    url_list = ['https://www.wildberries.ru/brands/vegagreen',
                'https://www.wildberries.ru/brands/naturalno',
                'https://www.wildberries.ru/brands/prosto-zdorovo'
                ]

    for url in url_list:
        client = ParserWB(url=url)
        client.run()


if __name__ == '__main__':
    main()
