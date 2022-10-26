import datetime
import time
from calendar import monthrange

import pandas as pd
import openpyxl
from openpyxl.styles.numbers import BUILTIN_FORMATS

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager


PATH = './output.xlsx'


class Moex:
    """Основной класс программы"""

    def __init__(self, rate):
        self.rate = rate.replace('/', '_')

    def get_chrome_driver(self, site_url):
        chrome_options = Options()
        chrome_options.add_argument('start-maximized')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                  options=chrome_options)
        driver.get(site_url)
        return driver

    def agree(self, driver):
        agreement = '//div[@class="disclaimer__buttons"]/a[text()[contains(.,' \
                              '"Согласен")]][1]'
        if driver.find_element("xpath", agreement):
            to_agree = driver.find_element("xpath", agreement)
            to_agree.click()

    def rate_path(self, driver):
        menu = driver.find_element("xpath", '//div[@class="col-md-4 header__button"]')
        menu.click()
        driver.implicitly_wait(10)
        futures_market = driver.find_element("xpath",
                                             '//div[@class="header-burger header-burger__desktop"]'
                                             '//li/a[@href="/ru/derivatives/"]')
        futures_market.click()
        time.sleep(1)
        self.agree(driver)
        indicative_rates = driver.find_element("xpath",
                                               f'//div[a[@href="/ru/derivatives/currency-rate.aspx'
                                               f'?currency=USD_RUB"]]//span')
        indicative_rates.click()

        select_option = Select(driver.find_element("name", 'ctl00$PageContent$CurrencySelect'))
        select_option.select_by_value(f'{self.rate}')
        time.sleep(1)
        selected_rate = select_option.first_selected_option.accessible_name.split(' ')[0]
        if selected_rate != f"{self.rate.replace('_', '/')}":
            driver = self.get_chrome_driver(
                f'https://www.moex.com/ru/derivatives/currency-rate.aspx?currency={self.rate}')
            time.sleep(1)
            self.agree(driver)
        return driver

    def last_month_quotes(self, driver):
        driver = self.rate_path(driver)

        now = datetime.datetime.now()
        target_month = now.month - 1
        # Номер предыдущего месяца
        target_year = now.year if target_month != 0 else now.year - 1
        # Необходимый год
        n_days = monthrange(target_year, target_month)[1]
        # Количество дней в предыдущем месяце

        year1 = driver.find_element("id", 'd1year')
        selected_year1 = int(Select(year1).first_selected_option.accessible_name)
        if target_year < selected_year1:
            year1.send_keys(Keys.UP)
            year1.send_keys(Keys.ENTER)

        year2 = driver.find_element("id", 'd2year')
        selected_year2 = int(Select(year2).first_selected_option.accessible_name)
        if target_year < selected_year2:
            year2.send_keys(Keys.UP)
            year2.send_keys(Keys.ENTER)

        month1 = driver.find_element("id", 'd1month')
        month1.click()
        options = [opt.accessible_name for opt in Select(month1).options]
        selected_month1 = Select(month1).first_selected_option.accessible_name
        if options.index(selected_month1) == target_month:
            month1.send_keys(Keys.UP)
            month1.send_keys(Keys.ENTER)
        elif options.index(selected_month1) == 0:
            for _ in range(11):
                month1.send_keys(Keys.DOWN)
            month1.send_keys(Keys.ENTER)

        month2 = driver.find_element("id", 'd2month')
        month2.click()
        options = [opt.accessible_name for opt in Select(month2).options]
        selected_month2 = Select(month2).first_selected_option.accessible_name
        if options.index(selected_month2) == target_month:
            month2.send_keys(Keys.UP)
            month2.send_keys(Keys.ENTER)
        elif options.index(selected_month2) == 0:
            for _ in range(11):
                month2.send_keys(Keys.DOWN)
            month2.send_keys(Keys.ENTER)

        days1 = driver.find_element("id", 'd1day')
        days1.click()
        selected_day1 = int(Select(days1).first_selected_option.accessible_name)
        if selected_day1 > 1:
            for _ in range(selected_day1 - 1):
                days1.send_keys(Keys.UP)
            days1.send_keys(Keys.ENTER)

        days2 = driver.find_element("id", 'd2day')
        days2.click()
        selected_day2 = int(Select(days2).first_selected_option.accessible_name)
        if selected_day2 < n_days:
            for _ in range(n_days - selected_day2):
                days2.send_keys(Keys.DOWN)
            days2.send_keys(Keys.ENTER)
        elif selected_day2 > n_days:
            for _ in range(selected_day2 - n_days):
                days2.send_keys(Keys.UP)
            days2.send_keys(Keys.ENTER)

        show_data_button = driver.find_element("xpath", '//input[@value="Показать"]')
        show_data_button.click()
        return driver

    def get_dataframe(self, driver):
        rate = self.rate.replace('_', '/')
        driver = self.last_month_quotes(driver)
        data = driver.find_elements("xpath", '//table[@class="tablels"]//tr')
        tr_data = [tr.text.split(' ') for tr in data]
        tr_data[1].insert(0, f'{tr_data[0][0]} {rate}')
        tr_data[1][3] = f'Курс {rate}'
        tr_data[1][4] = f'Время {rate}'
        df_data = []
        for lst in tr_data[1:]:
            lst.pop(1)
            lst.pop(1)
            df_data.append(lst)
        df = pd.DataFrame(df_data[1:], columns=df_data[0])
        return df

    def close_browser(self, driver):
        driver.quit()

    def xlsx_write(self, dataframe, excel_path):
        # Создаем excel writer object
        writer = pd.ExcelWriter(excel_path)
        # Записываем dataframe в excel
        dataframe.to_excel(writer, index=False)
        # Устанавливаем автовыравнивание ширины колонок
        self.auto_adjust(dataframe, writer)
        # Сохраняем excel
        writer.save()

    def auto_adjust(self, dataframe, writer):
        for column in dataframe:
            column_width = max(dataframe[column].astype(str).map(len).max(), len(column))
            col_idx = dataframe.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

    def xlsx_append(self, dataframe, excel_path):
        df_excel = pd.read_excel(excel_path)
        result = pd.concat([df_excel, dataframe], axis=1)
        self.xlsx_write(result, excel_path)

    def calculation(self, excel_path):
        df_excel = pd.read_excel(excel_path)
        result_lst = []
        for column in df_excel.columns:
            listing = df_excel[column].tolist()
            result_lst.append(listing)

        result_column = [float(x.replace(',', '.')) * float(y.replace(',', '.'))
                         for x, y in zip(result_lst[1], result_lst[4])]
        df = pd.DataFrame(result_column, columns=['Результат'])
        return df

    @staticmethod
    def xlsx_num_format(cell_literal, excel_path):
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        for row in ws[cell_literal][1:]:
            row.number_format = BUILTIN_FORMATS[2]
        wb.save(excel_path)

    @staticmethod
    def num_rows(excel_path):
        df_excel = pd.read_excel(excel_path)
        n = df_excel.shape[0]
        es = ['а', 'и', '']
        n = n % 100
        if 11 <= n <= 19:
            s = es[2]
        else:
            i = n % 10
            if i == 1:
                s = es[0]
            elif i in [2, 3, 4]:
                s = es[1]
            else:
                s = es[2]
        return f'{n} строк{s}'


if __name__ == '__main__':
    endpoint = 'https://www.moex.com/'

    # Создаем объект класса основной программы для 'USD/RUB':
    rate1 = 'USD/RUB'
    moex1 = Moex(rate1)

    # 1.Открыть https://www.moex.com (точка входа):
    driver1 = moex1.get_chrome_driver(endpoint)
    # 2.Перейти по следующим элементам: Меню -> Срочный рынок -> Индикативные курсы,
    # 3.В выпадающем списке выбрать валюты: USD / RUB - Доллар США к российскому рублю,
    # 4.Сформировать данные за предыдущий месяц:
    df1 = moex1.get_dataframe(driver1)
    # 5. Скопировать данные в Excel; Столбцы в Excel:
    # (A) Дата USD/RUB – Дата
    # (B) Курс USD/RUB – Значение из Курс основного клиринга
    # (C) Время USD/RUB – Время из Курс основного клиринга
    moex1.xlsx_write(df1, PATH)
    moex1.close_browser(driver1)

    # 6. Повторить шаги для валют JPY/RUB - Японская йена к российскому рублю:
    rate2 = 'JPY/RUB'
    moex2 = Moex(rate2)
    driver2 = moex2.get_chrome_driver(endpoint)
    df2 = moex2.get_dataframe(driver2)
    # 7. Скопировать данные в Excel; Столбцы в Excel:
    # (D) Дата JPY/RUB – Дата
    # (E) Курс JPY/RUB – Значение из Курс основного клиринга
    # (F) Время JPY/RUB – Время из Курс основного клиринга
    moex2.xlsx_append(df2, PATH)

    # 8. Для каждой строки полученного файла поделить курс USD/RUB на JPY/RUB,
    # полученное значение записать в ячейку (G) Результат,
    # 9. Выровнять – автоширина:
    moex2.xlsx_append(moex2.calculation(PATH), PATH)
    moex2.close_browser(driver2)
    # 10. Формат чисел – финансовый,
    # 11. Проверить, что автосумма в Excel распознаёт ячейки как числовой формат:
    Moex.xlsx_num_format('G', PATH)
