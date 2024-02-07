from personaParser import PersonaParcer
from structureParser import StructureParser
from numParser import NumParser
from Levenshtein import distance
import time


class Finder:
    def __init__(self, filepath):
        start = time.time()
        print("Начата подготовка базы данных..")

        self.filepath = filepath
        self.__personaParser = PersonaParcer(filepath)
        self.__structureParser = StructureParser(filepath)
        self.__numParser = NumParser(filepath)

        print(f"Подготовка окончена! Время обработки составило: {round(time.time() - start, 2)} с.")

    def classifier(self, line: str) -> list:
        """
        Данная функция предназначена для распределения запроса по типу (персона, структур, шифр, вхождение в другую
        строку и unclassified), и возвращению ближайших результатов согласно типу запроса.

        :param
        line (str): входная строка, которую мы будем классифицировать.

        :return
        (list):
        A) Возврщаем тип + результат:
        ["persona"] / ["structure"] / ["num"] + [Result]
        Б) "substring" + результат + тип:
        ["substring"] + [Result] + ["persona"] / ["structure"] / ["num"]
        В) в случае, если вопрос не классифицирован:
        ["unclassified"]
        """

        # Первая проверка на то, является ли данная строка шифром, если это так, тогда она состоит из точек и цифр (
        # в целом, она должна состоять на 100 процентов, но допускаем ошибки, поэтому пусть больше половины будут
        # точками или цифрами). Либо допускается вариант, когда line начинается со слов "Без кода".

        if "Без кода" in line or sum(c.isdigit() or c == '.' for c in line) / len(line) >= 0.5:
            numFind = self.__numParser.find(line)
            return ["num"] + numFind

        # В ином случае, мы работаем со структурой, или персоной по названию.
        else:
            # Формируем поиск по множества персон и структур, после чего находим ближайшие элементы из каждого.
            personaFind = self.__personaParser.find(line)
            structureFind = self.__structureParser.find(line)

            # Теперь проверим, кто оказался ближе всего к line. Сверяться будем по 0 элементу, так как у него
            # наименьшее расстояние в силу алгоритма find (там сортировка результатов по кратчайшему расстоянию).
            personaMinDistance, structureMinDistance = distance(personaFind[0], line), distance(structureFind[0], line)

            # Если расстояние слишком большое (больше половины строки) до ближайших элементов, то делаем
            # дополнительную проверку: Если количество слов в строке равно одному, то проверяем, является ли строка
            # Фамилией целиком, без неточностей! Проверка не пройдена => ["unclassified"].
            # Если в строке больше одного слова, то проверяем на вхождение в какую-либо строку в базе знаний.

            if min(personaMinDistance, structureMinDistance) >= len(line) // 2:
                # В строке одно слово => проверяем на вхождение в список фамилий.
                if len(line.split()) == 1:
                    family_dict = dict()
                    for persona in self.__personaParser.personaDict.keys():
                        # Проверяем, входит ли line в фамилию
                        if line.lower() in persona.split()[0].lower():
                            return ["substring", self.__personaParser.show(persona), 'persona']

                    # Если не нашли фамилию, то unclassified
                    return ["unclassified"]

                # Запустим проверку на вхождение в список персон
                for persona in self.__personaParser.personaSet:
                    if line.lower() in persona.lower():
                        return ["substring", self.__personaParser.show(persona), 'persona']

                # Запустим проверку на вхождение в список структур
                for structure in self.__structureParser.structureSet:
                    if line.lower() in structure.lower():
                        return ["substring", self.__structureParser.show(structure), 'structure']

                # Запустим проверку на вхождение в список шифров
                for num in self.__numParser.numSet:
                    if line.lower() in num.lower():
                        return ["substring", self.__numParser.show(num), 'num']

                # Если вхождений нет
                return ["unclassified"]

            return ["persona"] + personaFind if personaMinDistance < structureMinDistance else ["structure"] + structureFind

    def asker(self, data: list) -> list:
        """
        Данная функция предназначена для формирования ответа по обработанному запросу.
        :param
        data (list): список с классом, и лучшим / лучшими результатами.

        :return:
        Подготовленные данные к выводу.
        """

        # Проверка на то, что вопрос не найден
        if data[0] == "unclassified":
            return ["Данные не найдены!"]

        # Проверка на вхождение запроса в другие строки:
        if data[0] == "substring":
            data[0] = "Точного совпадения нет.. но найдено вхождение в строку!"
            return data

        if len(data) == 2:  # Значит нашли точное совпадение!
            result = ["Найдено точное совпадение!"]
        else:  # Значит точного совпадения нет => есть три ближайших результата! len(data) == 4
            result = ["Точного совпадения нет.. Ближайшие результаты:"]

        if data[0] == "persona":
            for persona in data[1:]:
                result.append(self.__personaParser.show(persona))

        elif data[0] == "structure":
            for struct in data[1:]:
                result.append(self.__structureParser.show(struct))

        elif data[0] == "num":
            for num in data[1:]:
                result.append(self.__numParser.show(num))

        return result

    def search(self, targetLine: str) -> list:
        """
        Данный поиск позволяет находить элемент, и выводить для него данные через общий интерфейс.
        :param targetLine:
        :return:
        """

        # Получаем класс запроса вместе с лучшим/лучшими результатами.
        questionClass = self.classifier(targetLine)
        answer = self.asker(questionClass)
        return answer


if __name__ == "__main__":
    finder = Finder("Data.xlsx")
    testPersonaNotFullMatch, testStructureNotFullMatch, testNumNotFullMatch = "Агамвир1зфыян Ивфгорь Рубевнови4фч", 'секsadтор "Пр1иказы"', "a01.11b1a.03.0в1"
    testPersonaFullMatch, testStructureFullMatch, testNumFullMatch = "Агамирзян Игорь Рубенович", 'сектор "Приказы"', "01.111.03.01"
    testPersonaIncome, testStructureIncome = "Агамирзян", "евроатлантический"

    # Блок тестов №1. Определение типа запроса
    print(f"Test 1. Входная строка: {testPersonaNotFullMatch = }. Результат поиска: {finder.classifier(testPersonaNotFullMatch)[0]}")
    print(f"Test 2. Входная строка: {testStructureNotFullMatch = }. Результат поиска: {finder.classifier(testStructureNotFullMatch)[0]}")
    print(f"Test 3. Входная строка: {testNumNotFullMatch = }. Результат поиска: {finder.classifier(testNumNotFullMatch)[0]}")

    # Блок тестов №2. Результаты поиска для неточного совпадения.
    print(finder.search(testPersonaNotFullMatch))
    print(finder.search(testStructureNotFullMatch))
    print(finder.search(testNumNotFullMatch))

    # Блок тестов №3. Результаты поиска для точного совпадения.
    print(finder.search(testPersonaFullMatch))
    print(finder.search(testStructureFullMatch))
    print(finder.search(testNumFullMatch))

    # Блок тестов №4. Результаты для вхождения в строку.
    print(finder.search(testPersonaIncome))
    # print(finder.search(testStructureIncome))