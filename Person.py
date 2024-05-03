import random

from openpyxl.cell import Cell


class Person:
    def __init__(self, row: tuple[Cell, ...]):
        self.name: str = row[1].value
        self.id: str = row[2].value
        self.check_date: str = row[3].value
        self.elder: str = row[4].value
        self.hypertension: str = row[5].value
        self.diabetes: str = row[6].value
        self.temperature: str = self.__temperature(row[7].value)
        self.pulse_rate: str = self.__pulse_rate(row[8].value)
        self.breath_rate: str = self.__breath_rate(row[9].value)
        self.right_systolic_pressure: str = row[10].value

    def __breath_rate(self, excel_value: str) -> str:
        if excel_value is None:
            return self.__random_breath_rate()
        return excel_value

    @staticmethod
    def __random_breath_rate():
        return str(random.randint(16, 20))

    def __pulse_rate(self, excel_value: str) -> str:
        if excel_value is None:
            return self.__random_pulse_rate()
        return excel_value

    @staticmethod
    def __random_pulse_rate() -> str:
        return str(random.randint(60, 100))

    def is_elder(self) -> bool:
        return self.elder.strip() == "老年人"

    def __temperature(self, excel_value: str) -> str:
        if excel_value is None:
            return self.__random_temperature()
        return excel_value

    @staticmethod
    def __random_temperature() -> str:
        return str(round(random.uniform(36.5, 37), 1))
