from selenium import webdriver
from selenium.webdriver.common.by import By
import keyboard
import time
import openpyxl

## Открываем или создаем файл Excel
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Результаты поиска"

link = "https://www.citilink.ru/"
try:
    browser = webdriver.Chrome()
    browser.get(link)
    butt = browser.find_element(By.CSS_SELECTOR, ".app-catalog-144309a")
    butt.click()
    pole = browser.find_element(By.CSS_SELECTOR, ".css-1u9ewb3")
    pole.send_keys("RYZEN")
    keyboard.press("Enter")

    # Получаем результаты поиска
    time.sleep(5)  # Ждем некоторое время, чтобы страница загрузилась полностью
    items = browser.find_elements(By.CSS_SELECTOR, ".ProductCardHorizontal__content")

    # Записываем данные в файл Excel
    for index, item in enumerate(items, start=1):
        name = item.find_element(By.CSS_SELECTOR, ".ProductCardHorizontal__name").text
        price = item.find_element(By.CSS_SELECTOR, ".ProductCardHorizontal__price-current_current-price").text

        sheet[f"A{index}"] = name
        sheet[f"B{index}"] = price

finally:
    # Сохраняем файл Excel
    workbook.save("результаты_поиска.xlsx")
    time.sleep(5)  # Даем время для сохранения файла перед закрытием браузера
    browser.quit()
