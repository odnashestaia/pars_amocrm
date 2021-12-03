from selenium import webdriver
import openpyxl
import time


def login_in(login, password, get):
    """ вход на сайт """
    driver.get(get)

    """ Вход в систему """
    driver.find_element_by_name('password').send_keys(password)
    driver.implicitly_wait(10)
    driver.find_element_by_name('username').send_keys(login)
    driver.implicitly_wait(10)
    driver.find_element_by_class_name('auth_form__submit').click()

    """ Ввод кода """
    # time.sleep(120)


def pars_info():
    """ Парс """
    for i in driver.find_elements_by_class_name('list-row'):
        l = []
        l.append(i.find_elements_by_class_name('list-row__cell')[3].text)
        l.append(i.find_elements_by_class_name('list-row__cell')[2].text)
        l.append(i.find_elements_by_class_name('list-row__cell')[4].text)
        main_l.append(l)
    return main_l


def write_data(l):
    """ Запись в эксельку """
    print(len(l))
    write = openpyxl.Workbook()
    page = write.active
    for num, var in enumerate(l):
        page.append(list(var))
    write.save(filename="Сделки из AMOcrm.xlsx")


def main():
    get = str(input("Введите ссылку на amoCRM: "))
    login = str(input("Ведите логин: "))
    password = str(input("Ведите пароль: "))
    what = str(input("Парсить все страницы? (да\нет) "))
    login_in(login, password, get)
    if what == 'да':
        try:
            p = driver.find_elements_by_class_name('pagination-link__wrapper')[-1].text
        except IndexError:
            print('Страниц кажись меньше (перезапусти программу и нажми нет')
        for i in range(1, int(p) + 1):
            print(i)
            pars_info()  # запись в список
            if '?' in get:
                driver.get(get.split('?')[0] + 'page/' + str(i + 1) + '?' + get.split('?')[1])  # переход на новую страницу
    elif what == "нет":
        pars_info()
    write_data(main_l)


if __name__ == '__main__':
    main_l = []
    driver = webdriver.Chrome()
    main()
    driver.close()
