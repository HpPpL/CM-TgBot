import openpyxl
import os
from Levenshtein import distance


class StructureParser:
    def __init__(self, filePath: str, parseFlag: bool = True):
        self.filePath = filePath
        self.workbook = openpyxl.load_workbook(self.filePath)
        self.worksheet = self.workbook['Оргструктура']
        self.structureDict = {}
        self.structureSet = set()

        if parseFlag:
            self.parser()

    def parser(self) -> None:
        """
        Данная функция предназначена для парсинга листа "Полномочия" по структурам. Данная функция ничего не возвращает.
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

        for row in self.worksheet.iter_rows(min_row=2):
            if row[0].value is None:
                break

            # Уделить дополнительное время парсингу задач отделения
            data = {
                "1. Код ИС-ПРО: ": getValue(row, 1) + ";",
                "2. Вышестоящее подразделение: ": getValue(row, 2) + ";",
                "3. Вышестоящее подразделение верхнего уровня: ": getValue(row, 3) + ";",
                "4. Кампус: ": getValue(row, 4) + ";",
                "5. Самостоятельное: ": getValue(row, 5) + ";",
                "6. Статус: ": getValue(row, 6) + ";",
                "7. Вид подразделения: ": getValue(row, 7) + ";",
                "8. Тип подразделения: ": getValue(row, 8) + ";",
                "9. Основной вид деятельности: ": getValue(row, 9) + ";",
                "10. Дополнительные виды деятельности: ": getValue(row, 10) + ";",
                "11. Задачи подразделения: ": getValue(row, 11) + ";",
                "12. Дата создания: ": getValue(row, 12) + ";",
                "13. Номер приказа о создании: ": getValue(row, 13) + ";"
            }

            resLine = "\n".join('{}{}'.format(key, val) for key, val in data.items())
            self.structureDict.setdefault(row[0].value, []).append(resLine)

        self.updateStructure(self.structureDict)

    def updateStructure(self, strctDict: dict) -> None:
        """
        Данная функция предназначена для обновления множества структур.

        :param
        strctDict (dict): словарь с распарсенными данными структур.

        :return
        (None): данная функция ничего не возвращает, она только выполняет обновление.
        """
        self.structureSet = set(strctDict.keys())

    def find(self, targetLine: str) -> list:
        """
        Данная функция ищет по образцу струтктуру в множестве. В случае если есть точное совпадение, возвращается 
        одна структура, в ином - три самых близких.
    
        :param
        targetLine (str): входная строка, по которой производится поиск.
    
        :return:
        (lst): список ближайших по-расстоянию Левенштейна элементов из множества structureSet к элементу targetLine.
        """

        # Сортируем все структуры по расстоянию
        structureSorted = sorted(self.structureSet, key=lambda x: distance(x, targetLine))
        minDistance = distance(structureSorted[0], targetLine)

        if minDistance == 0:
            # Если минимум равен 0, возвращаем один элемент
            return [structureSorted[0]]
        else:
            # Если минимум не равен 0, возвращаем три элемента
            return structureSorted[:3]

    def show(self, structureName: str) -> list:
        """
        Данная функция предназначена для вывода данных для отдельной структуры.

        :param
        structureName (str): входная строка, по которой будет выводиться результат.

        :return
        (str): данные для структуры в строчном формате.

        """
        result = [structureName]

        for index, value in enumerate(self.structureDict[structureName]):
            temp = ""
            # result += f"{index + 1}) "
            temp += str(value) + "\n"

            if index != len(self.structureDict[structureName]) - 1:
                temp += "\n"

            result.append(temp)

        return result

    def __close_workbook(self):
        self.workbook.close()

    def __str__(self):
        pass

    def __del__(self):
        self.__close_workbook()

    def __call__(self, *args, **kwargs):
        # Стоит ли?
        pass


if __name__ == '__main__':
    filePath = "Data.xlsx"
    structureParser = StructureParser(filePath)

    test1 = 'Кaоdмпьютерwный цеeнтр'
    print(f"Входная строка: {test1}")
    f = structureParser.find(test1)
    r = structureParser.show(f[0])

    print(r)
