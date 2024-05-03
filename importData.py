from typing import Generator, List

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from selenium import webdriver
from selenium.webdriver.common.by import By

import config
from Person import Person
from config import target_url


def import_check_data(username: str, password: str, file_path: str, mode: int):
    switcher = {
        0: import_new,
        1: import_cover
    }
    importer = switcher.get(mode)
    importer(username, password, file_path)


def import_new(username: str, password: str, file_path: str):
    driver: webdriver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(target_url)
    driver.find_element(by=By.ID, value="pwd").send_keys(password)
    # persons: List[Person] = read_excel(file_path)
    # for person in persons:
    #     config.current_row_index += 1


def import_cover(username: str, password: str, file_path: str):
    print(username, password)


def read_excel(file_path: str) -> List[Person]:
    workbook: Workbook = load_workbook(file_path)
    sheet: Worksheet = workbook.active
    rows: Generator[tuple[Cell, ...], None, None] = sheet.iter_rows(min_row=config.current_row_index)
    persons: List[Person] = []
    for row in rows:
        person = Person(row)
        persons.append(person)
    return persons
