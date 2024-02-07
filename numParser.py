import openpyxl
import os
from Levenshtein import distance


class NumParser:
    def __init__(self, filePath: str, parseFlag: bool=True):
        self.filePath = filePath
        self.workbook = openpyxl.load_workbook(self.filePath)
        self.worksheet = self.workbook['Оргструктура']
        self.numDict = {}
        self.numSet = set()

        if parseFlag:
            self.parser()

    def parser(self) -> None:
        """
        Данная функция предназначена для парсинга листа "Полномочия" по шифрам. Данная функция ничего не возвращает.
        """
        
        def getValue(row, index, defaultMessage="нет") -> str:
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
            if row[1].value is None:
                break

            # Уделить дополнительное время парсингу задач отделения
            data = {
                "1. Наименование: ": getValue(row, 0) + ";",
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
            self.numDict.setdefault(row[1].value, []).append(resLine)

        self.updateNum(self.numDict)

    def updateNum(self, numDict: dict) -> None:
        """
        Данная функция предназначена для обновления множества шифров.

        :param
        numDict (dict): словарь с распарсенными данными шифров.

        :return
        (None): данная функция ничего не возвращает, она только выполняет обновление.
        """
        self.numSet = set(numDict.keys())

    def find(self, targetLine: str) -> list:
        """
        Данная функция ищет по образцу шифр в множестве. В случае если есть точное совпадение, возвращается 
        один шифр, в ином - три самых близких.

        :param
        targetLine (str): входная строка, по которой производится поиск.

        :return:
        (lst): список ближайших по-расстоянию Левенштейна элементов из множества numSet к элементу targetLine.
        """
        
        # Сортируем все структуры по расстоянию
        numSorted = sorted(self.numSet, key=lambda x: distance(x, targetLine))
        minDistance = distance(numSorted[0], targetLine)

        if minDistance == 0:
            # Если минимум равен 0, возвращаем один элемент
            return [numSorted[0]]
        else:
            # Если минимум не равен 0, возвращаем три элемента
            return numSorted[:3]

    def show(self, numName: str) -> list:
        """
        Данная функция предназначена для вывода данных для отдельной структуры по ее шифру.

        :param
        structureName (str): входная строка, по которой будет выводиться результат.

        :return
        (str): данные для структуры в строчном формате.

        """

        result = [numName]

        for index, value in enumerate(self.numDict[numName]):
            temp = ""
            # result += f"{index + 1}) "
            temp += str(value) + "\n"

            if index != len(self.numDict[numName]) - 1:
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
    numParser = NumParser(filePath)

    test1 = '1a0.f0s.05.12d01.d03'
    print(f"Входная строка: {test1}")
    f = numParser.find(test1)
    r = numParser.show(f[0])
    print(r)
