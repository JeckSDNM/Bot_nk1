import sys
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from random import randint
from char_params import Person


all_t_q = Person.char_params()


class Bot:
    __emp_count = 0

    def __init__(self):
        Bot.__emp_count += 1
        self.__data_frame = self.__check_exel_doc()
        self.__is_exelDoc = self.__get_path()

        self.__student_code = ''
        self.__result = ''
        self.driver = None

        print('Бот создан')

    @staticmethod
    def __interval():
        sleep(randint(2, 5))

    @staticmethod
    def __start_driver():
        print('Запуск драйвера...')
        try:
            return webdriver.Chrome()
        except:
            try:
                return webdriver.Firefox()
            except:
                print("Не найден подходязий браузер. Версия драйвера могла устареть.")
                return sys.exit()

    '''Проверка/создание дока для сохранения результатов'''
    def __check_exel_doc(self):
        if not os.path.isdir("exel_doc"):
            os.mkdir("exel_doc")

        if not os.path.exists(f'exel_doc/Result{Bot.__emp_count}.xlsx'):
            d = {'Номер': [], 'Логин': [], 'Результат': []}
            pd.DataFrame(data=d).to_excel(f'exel_doc/Result{Bot.__emp_count}.xlsx')

        return pd.read_excel(f'exel_doc/Result{Bot.__emp_count}.xlsx', index_col=0)

    # ######################################## Работа со списком exel
    '''Задает exel документ'''
    def __get_path(self):
        path_exel_doc = input('Введите путь к таблице логинов(пропустиде для ручного ввода): ').strip()
        if os.path.exists(path_exel_doc):
            self.__exelDoc = path_exel_doc
            self.__get_code_list()
            return True

        elif len(path_exel_doc) < 4:
            return False
        else:
            print('Файла не существует.')
            self.__get_path()

    '''Выборка столбца логинов если есть файл'''
    def __get_code_list(self):
        exel = pd.read_excel(self.__exelDoc)
        name_column = 'Лгин'
        if name_column not in exel.columns:
            while name_column not in exel.columns:
                name_column = input('Не найден столбец логинов.\nВведите название столбца.\n=>')

        log_list1 = exel[name_column].tolist()
        log_list2 = self.__data_frame['Логин'].tolist()
        self.__code_list = list((set(log_list1) - set(log_list2))) + list((set(log_list2) - set(log_list1)))
        self.__index_code = -1
        # self.__max_index_code = len(self.__code_list)

    # ######################################################## сохранение
    '''Сохранение данных в файл отсчета'''
    def __saved_data(self):
        self.__data_frame = self.__data_frame.append({'Логин': self.__student_code, 'Результат': self.__result}, ignore_index=True)
        self.__data_frame.to_excel(f'exel_doc/Result{Bot.__emp_count}.xlsx')

    # ######################################################## Основа
    def __get_student_code(self):
        self.__index_code += 1
        return self.__code_list[self.__index_code]

    def __set_student_code(self):
        if self.__is_exelDoc:
            self.__student_code = self.__get_student_code()
        else:
            self.__student_code = input('Введите код студента: ')

    '''Вводит код студента'''
    def __inp_code(self):
        self.__set_student_code()
        elem = self.driver.find_element(By.NAME, 'oLogin')
        elem.send_keys(self.__student_code)
        self.driver.find_element(By.TAG_NAME, 'button').click()
        try:
            self.driver.find_element(By.XPATH, '/html/body/div/main/div/div/div/div/div/h3[2]')
            logic = True
        except:
            logic = False

        if logic:
            print(self.__student_code, ': Код не верный или уже используется')
            self.__result = 'Неверный'
            self.__saved_data()
            self.__inp_code()

    '''Проходит 2ю и 3ю страницы'''
    def __sec_the_pages(self):
        self.driver.find_element(By.NAME, 'sex').click()
        elem = self.driver.find_element(By.NAME, 'age')
        sleep(1)
        elem.send_keys(randint(17, 19))
        elem.send_keys(Keys.RETURN)
        self.__interval()

        self.driver.find_element(By.TAG_NAME, 'button').click()
        sleep(1)

    '''Неубиваемый кусок, всегда работает'''
    def __test(self):
        text_q = self.driver.find_element(By.CLASS_NAME, 'test_block').find_element(By.TAG_NAME, 'div').text
        but_arr = self.driver.find_element(By.CLASS_NAME, 'answers').find_elements(By.TAG_NAME, 'button')

        for k, v in all_t_q.items():
            if text_q == k:
                # print(text_q)
                self.__interval()
                but_arr[int(v)].click()
                sleep(1)
                but_arr[-1].click()

            elif text_q not in all_t_q:
                print('Нет в ', k)
                but_arr[-3].click()

    '''Устанавливает результат'''
    def __set_result(self):
        result = self.driver.find_element(By.XPATH, '/html/body/div[2]/main/div[1]/div/div/div/div[2]/div[2]/div[1]')
        if result.text.find('Вы успешно прошли социально-психологический тест.') != -1:
            a = 'Успех'
        else:
            a = 'Провал'
        self.__result = 'Пройдено' + a

    def __up_link(self):
        self.driver.get('https://spt.kuzrc.ru/')

    def _start(self):
        if self.driver is None:
            self.driver = self.__start_driver()
        self.__up_link()

        self.__inp_code()
        self.__sec_the_pages()

        for i in range(140):
            self.__test()
        sleep(0.5)

        self.__set_result()
        self.__saved_data()


bot1 = Bot()
bot1._start()
