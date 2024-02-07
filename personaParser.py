import openpyxl
import os
# pip install Levenshtein
from Levenshtein import distance


class PersonaParcer:
    def __init__(self, filePath: str, parseFlag: bool = True):
        self.filePath = filePath
        self.workbook = openpyxl.load_workbook(self.filePath)
        self.worksheet = self.workbook['Полномочия']
        self.personaDict = {}
        self.personaSet = set()

        if parseFlag:
            self.parser()

    def parser(self) -> None:
        """
        Данная функция предназначена для парсинга листа "Полномочия", то есть парсинга по персонам. Данная функция
        ничего не возвращает.
        """

        def getValue(row, index: int, defaultMessage: str = "нет") -> str:
            """
            Данная функция предназначена для получения значения из ячейки. Если значение None, то заменяем его на defaultMessage.

            :param
            row (docx.rows): строчка, из которой будет извлечен элемент.
            index (int): индекс извлекаемого элемента в строчке.

            :return
            (str): значение в ячейке в формате строки (defaultMessage в случае отсутствия значения)
            """

            return str(row[index].value) if row[index].value is not None else defaultMessage

        # Поскольку функция может быть вызвана многократна, реализовано обнуление словаря! Данные не сохраняются!
        self.personaDict = dict()

        for row in self.worksheet.iter_rows(min_row=2):
            # Проверка на конец списка
            if row[0].value is None:
                break

            # Стоит ли дописать до Data class? Как будто бы нет
            # Формируем словарь из данных в строке
            data = {
                "1. Должность: ": getValue(row, 1) + ";",
                "2. Полномочия получены: ": getValue(row, 2) + ";",
                "3. Дата приказа об установлении полномочий: ": getValue(row, 3) + ";",
                "4. Номер приказа об установлении полномочий: ": getValue(row, 4) + ";",
                "5. Статус приказа: ": getValue(row, 5) + ";",
                "6. Тип приказа: ": getValue(row, 6) + ";",
                "7. Возложение обязанностей на период временного отсутствия: ": getValue(row, 7) + ";",
                "8. Руководимые подразделения: ": getValue(row, 8) + ";",
                "9. Руководимые работники: ": getValue(row, 9) + ";",
                "10. Координируемые подразделения: ": getValue(row, 10) + ";",
                "11. Координируемые работники: ": getValue(row, 11) + ";",
                "12. Координируемые направление деятельности: ": getValue(row, 12) + ";",
                "13. Руководимые направления деятельности: ": getValue(row, 13) + ";",
                "14. Права работодателя: ": getValue(row, 14) + ";",
                "15. Право подписи, в рамках возложенных обязанностей и предоставленных полномочий: ": getValue(row,
                                                                                                                15) + ";"
            }

            resLine = "\n".join('{}{}'.format(key, val) for key, val in data.items())
            self.personaDict.setdefault(row[0].value, []).append(resLine)

        # Обновим также список персон
        self.updatePersonas(self.personaDict)

    def updatePersonas(self, persnDict: dict) -> None:
        """
        Данная функция предназначена для обновления множества персон.

        :param
        persnDict (dict): словарь с распарсенными данными персон.

        :return
        (None): данная функция ничего не возвращает, она только выполняет обновление.
        """
        self.personaSet = set(persnDict.keys())

    def find(self, targetLine: str) -> list:
        """
        Данная функция ищет по образцу персону в множестве. В случае если есть точное совпадение, возвращается один
        человек, в ином - три самых близких. Думается мне сделать так, что если совпадение неточное, но очень близкое
        к тому, то все равно возвращать одного человека, но пока это мысли..

        :param
        targetLine (str): входная строка, по которой производится поиск.

        :return:
        (lst): список ближайших по-расстоянию Левенштейна элементов из множества personaSet к элементу targetLine.
        """

        # Сортируем всех персон по расстоянию
        personasSorted = sorted(self.personaSet, key=lambda x: distance(x, targetLine))
        minDistance = distance(personasSorted[0], targetLine)

        if minDistance == 0:
            # Если минимум равен 0, возвращаем один элемент
            return [personasSorted[0]]
        else:
            # Если минимум не равен 0, возвращаем три элемента
            return personasSorted[:3]

    def show(self, fio: str) -> list:
        """ 
        Данная функция предназначена для вывода данных для отдельной персоны.
        
        :param
        fio (str): входная строка, по которой будет выводиться результат.

        :return
        (str): данные для персоны в строчном формате.

        """

        # # В случае если эта функция будет использоваться вне связки с find, то проверка необходима.
        # if fio not in self.personaSet:
        #     print("Персона не найдена!")
        #     return [""]

        result = [fio]

        for index, value in enumerate(self.personaDict[fio]):
            temp = ""
            # temp += f"{index + 1}) "
            temp += str(value) + "\n"

            if index != len(self.personaDict[fio]) - 1:
                temp += "\n"

            result.append(temp)

        return result

    def __close_workbook(self):
        self.workbook.close()

    def __str__(self):
        result = ""

        for key, item in self.personaDict.items():
            result += key + ": " + "\n"
            for index, value in enumerate(item):
                result += f"{index + 1}) "
                result += str(value) + "\n"

                if index != len(item) - 1:
                    result += "\n"
            result += "----\n"

        return result

    def __del__(self):
        self.__close_workbook()


if __name__ == '__main__':
    # Инициализация. В ней автоматически происходит парсинг.
    filePath = "Data.xlsx"
    authorityParser = PersonaParcer(filePath)

    # Поиск. Изначально передается в функцию find для поиска человека через наименьшее расстояние по Левенштейну.
    # После чего визуализируем данные из словаря
    # test1 = "Агамвир1зфыян Ивфгорь Рубевнови4фч"
    test2 = "ГуsЫбин Д1митsрий ВлЫадиdaмирови412ч"
    print(f"Входная строка: {test2}")

    f = authorityParser.find(test2)
    a = authorityParser.show(f[0])
    print(a)

    # Также класс является printable
    print(authorityParser)
