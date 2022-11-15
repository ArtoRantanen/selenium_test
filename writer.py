import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui
from selenium.webdriver.common.by import By
import time

login = 'password'
password = 'login'

filename = 'Tubes_data.xlsx'
data = pd.read_excel(filename)
i = 0


def get_data(data, i):
    list = []
    list.append(data['Name'][i])
    list.append(data['Price'][i])
    list.append(data['Марка'][i])
    list.append(data['Размер'][i])
    list.append(data['НТП'][i])
    list.append(data['Тип'][i])
    list.append(data['Способ производства'][i])
    list.append(data['Артикул'][i])
    if data['Описание'][i] is None:
        list.append('')
    else:
        list.append(data['Описание'][i])
    return list


def parcer(data):
    if data[6] == '-':
        text = """<div class="table-responsive">
        <table class="table table-bordered table-hover table-condensed">
        <tbody>
        <tr>
        <td>
        <p>Марка</p>
        </td>
        <td>
        <p>""" + str(data[2]) + """</p>
        </td>
        </tr>
        <tr>
        <td>
        <p>Размер</p>
        </td>
        <td>
        <p>""" + str(data[3]) + """"</p>
        </td>
        </tr>
        <tr>
        <td>
        <p>НТД</p>
        </td>
        <td>
        <p>""" + str(data[4]) + """</p>
        </td>
        </tr>
        <tr>
        <td>
        <p>Тип</p>
        </td>
        <td>
        <p>""" + str(data[5]) + """</p>
        </td>
        </tr>
        <tr>
        <td>
        <p>Артикул</p>
        </td>
        <td>
        <p>""" + str(data[7]) + """</p>
        </td>
        </tr>
        </tbody>
        </table>"""
    else:
        text = """<div class="table-responsive">
                <table class="table table-bordered table-hover table-condensed">
                <tbody>
                <tr>
                <td>
                <p>Марка</p>
                </td>
                <td>
                <p>""" + str(data[2]) + """</p>
                </td>
                </tr>
                <tr>
                <td>
                <p>Размер</p>
                </td>
                <td>
                <p>""" + str(data[3]) + """"</p>
                </td>
                </tr>
                <tr>
                <td>
                <p>НТД</p>
                </td>
                <td>
                <p>""" + str(data[4]) + """</p>
                </td>
                </tr>
                <tr>
                <td>
                <p>Тип</p>
                </td>
                <td>
                <p>""" + str(data[5]) + """</p>
                </td>
                </tr>
                <tr>
                <td>
                <p>Способ производства</p>
                </td>
                <td>
                <p>""" + str(data[6]) + """"</p>
                </td>
                </tr>
                <tr>
                <td>
                <p>Артикул</p>
                </td>
                <td>
                <p>""" + str(data[7]) + """</p>
                </td>
                </tr>
                </tbody>
                </table>"""
    return text


Delivery = """<h3>Условия доставки</h3>
<p>Для наших клиентов предлагаем выгодные условия доставки. Заказчикам в Санкт-Петербурге и в Ленинградской области предлагаем доставку в день заказа. Сроки доставки в регионы уточняйте у менеджеров при оформлении заказа.</p>
<p>У нас есть собственный автопарк со всей техникой, необходимой для оперативной погрузки, доставки и разгрузки металлопроката. Стоимость доставки уточняйте у менеджеров при оформлении заказа.</p>
<p><a class="btn btn-color" href="">Подробнее о доставке</a></p>"""

Metal = """<h3>Обработка металлов</h3>
<p>В нашей компании вы можете заказать дополнительные услуги по обработке металлов. Предлагаем широкий спектр услуг: от резки металла до лазерной сварки и химического анализа. В нашем распоряжении есть все оборудование, необходимое для того, чтобы выполнить обработку металлов качественно и быстро.</p>
<p><a class="btn btn-color" href="">Подробнее об услугах</a></p>"""

s = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s)
driver.maximize_window()

driver.get("")
driver.find_element(By.ID, "user_login").send_keys(login)
driver.find_element(By.ID, "user_pass").send_keys(password)
# id = 'user_login'
# id = 'user_pass'
time.sleep(5)
button = driver.find_element(By.ID, 'wp-submit')
button.click()
driver.get("")

while i < len(data['Name']):
    try:
        calls = get_data(data, i)
        Name, Price, Description = calls[0], calls[1], calls[8]
        Table = parcer(calls)
        if calls[2] == "Договорная":
            Before_price = ""
        else:
            Before_price = "от"
        time.sleep(1)
        target = driver.find_element(By.ID, "title")
        driver.find_element(By.ID, "title").send_keys(Name)
        driver.find_element(By.ID, "content-html").click()
        driver.find_element(By.ID, "content").send_keys(Description)
        driver.find_element(By.ID, "acf-field_6357bbb523918").send_keys(Before_price)
        driver.find_element(By.ID, "acf-field_5991273caef20").send_keys(Price)
        driver.find_element(By.LINK_TEXT, "Добавить изображение").click()
        driver.find_element(By.ID, "media-search-input").send_keys("2.1.1")
        pyautogui.press('TAB')
        with pyautogui.hold('CTRL'):
            pyautogui.press(['ENTER'])
        for j in range(11):
            pyautogui.press('TAB')
        pyautogui.press(['ENTER'])
        driver.find_element(By.ID, "acf-editor-54-html").click()
        driver.find_element(By.ID, "acf-editor-54").send_keys(Table)
        driver.find_element(By.ID, "acf-editor-55-html").click()
        driver.find_element(By.ID, "acf-editor-55").send_keys(Delivery)
        for j in range(3):
            pyautogui.press('TAB')
        pyautogui.press(['ENTER'])
        driver.find_element(By.ID, "acf-editor-56").send_keys(Metal)
        actions = ActionChains(driver)
        actions.move_to_element(target)
        actions.perform()
        driver.find_element(By.ID, "publish").click()
        i = i + 1
        driver.get("")
    except:
        driver.get("")
        time.sleep(1)
        pyautogui.press(['ENTER'])
        time.sleep(1)
        print(i)
        i = i + 1
driver.quit()
