import datetime
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import openpyxl


def write_to_excel(data, file_name='output.xlsx', sheet_name='Продукты'):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name

    column_widths = [0] * len(data[0])

    for row_data in data:
        worksheet.append(row_data)
        for i, value in enumerate(row_data):
            if len(str(value)) > column_widths[i]:
                column_widths[i] = len(str(value))

    for i, width in enumerate(column_widths):
        worksheet.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = width + 2

    workbook.save(file_name)


def collect_data():
    cur_time = datetime.datetime.now().strftime('%d_%m_%Y_%H_%M')
    user_agent = UserAgent()

    headers = {
        'Accept': 'application/json, text/plain, */*',
        'User-Agent': user_agent.random
    }

    cookies = {
        'mg_geo_id': '12505'
    }

    response = requests.get(url='https://5ka.ru/special_offers', headers=headers, cookies=cookies)

    soup = BeautifulSoup(response.text, 'lxml')

    cards = soup.find_all('div', class_='product-card item')

    data = [['Продукт', 'Процент скидки', 'Старая цена', 'Новая цена', 'Время акции']]

    for card in cards:
        card_title = card.find('div', class_='item-name').text.strip()

        card_discount = card.find('div', class_='discount-hint hint').text.strip()

        card_price_old = card.find('span', class_='price-regular').text.strip().replace(' ', '')

        card_price_integer = card.find('div', class_='price-discount').find('span').text.strip()

        card_price_decimal = card.find('div', class_='price-discount').find('span',
                                                                            class_='price-discount_cents').text.strip()

        card_price = f'{card_price_integer}.{card_price_decimal}'.strip().replace(' ', '')

        card_sale_date = card.find('div', class_='item-date').text.strip().replace('\n', ' ')

        data.append([card_title, card_discount, card_price_old.replace('\n', ' '), card_price, card_sale_date])

    write_to_excel(data, f'{cur_time}.xlsx', 'Продукты')

    print(f'Файл {cur_time}.xlsx успешно записан!')


def main():
    collect_data()


if __name__ == '__main__':
    main()
