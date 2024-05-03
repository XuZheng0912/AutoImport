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
        self.right_diastolic_pressure: str = row[11].value
        self.left_systolic_pressure: str = row[12].value
        self.left_diastolic_pressure: str = row[13].value
        self.height: str = row[14].value
        self.weight: str = row[15].value
        self.waistline: str = row[16].value
        self.teeth_left_up: str = self.__teeth(row[17].value)
        self.teeth_left_down: str = self.__teeth(row[18].value)
        self.teeth_right_up: str = self.__teeth(row[19].value)
        self.teeth_right_down: str = self.__teeth(row[20].value)
        self.left_eye: str = row[21].value
        self.right_eye: str = row[22].value
        self.electrocardiogram: str = row[23].value
        self.x_ray: str = row[24].value
        self.b_ultrasonic: str = row[25].value
        self.medicine_name_1: str = row[27].value
        self.medicine_use_method_1: str = row[28].value
        self.medicine_use_mount_1: str = row[29].value
        self.medicine_use_time_1: str = row[30].value
        self.medicine_compliance_1: str = row[31].value
        self.medicine_name_2: str = row[32].value
        self.medicine_use_method_2: str = row[33].value
        self.medicine_use_mount_2: str = row[34].value
        self.medicine_use_time_2: str = row[35].value
        self.medicine_compliance_2: str = row[36].value
        self.medicine_name_3: str = row[37].value
        self.medicine_use_method_3: str = row[38].value
        self.medicine_use_mount_3: str = row[39].value
        self.medicine_use_time_3: str = row[40].value
        self.medicine_compliance_3: str = row[41].value
        self.medicine_name_4: str = row[42].value
        self.medicine_use_method_4: str = row[43].value
        self.medicine_use_mount_4: str = row[44].value
        self.medicine_use_time_4: str = row[45].value
        self.medicine_compliance_4: str = row[46].value

    def __teeth(self, excel_value: str) -> str:
        if excel_value is None:
            return self.__random_teeth()
        return excel_value

    @staticmethod
    def __random_teeth():
        return str(random.randint(0, 4))

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
