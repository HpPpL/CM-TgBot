from pathlib import Path

import openpyxl


OUTPUT_PATH = Path(__file__).resolve().parents[1] / "Data.xlsx"


def build_persona_sheet(workbook):
    worksheet = workbook.active
    worksheet.title = "Полномочия"
    worksheet.append(
        [
            "ФИО",
            "Должность",
            "Полномочия получены",
            "Дата приказа об установлении полномочий",
            "Номер приказа об установлении полномочий",
            "Статус приказа",
            "Тип приказа",
            "Возложение обязанностей на период временного отсутствия",
            "Руководимые подразделения",
            "Руководимые работники",
            "Координируемые подразделения",
            "Координируемые работники",
            "Координируемые направление деятельности",
            "Руководимые направления деятельности",
            "Права работодателя",
            "Право подписи",
        ]
    )
    worksheet.append(
        [
            "Иванов Иван Иванович",
            "Руководитель демонстрационного отдела",
            "В рамках тестового набора данных",
            "2024-01-10",
            "DEMO-001",
            "Действует",
            "Демонстрационный приказ",
            "нет",
            "Демонстрационный отдел",
            "нет",
            "нет",
            "нет",
            "Административная поддержка",
            "Демо-направление",
            "нет",
            "нет",
        ]
    )


def build_structure_sheet(workbook):
    worksheet = workbook.create_sheet("Оргструктура")
    worksheet.append([f"Колонка {index}" for index in range(1, 42)])

    row = ["нет"] * 41
    row[0] = "Демонстрационный отдел"
    row[1] = "01.01.01"
    row[2] = "Демонстрационный департамент"
    row[3] = "Демонстрационный блок"
    row[4] = "Москва"
    row[5] = "да"
    row[7] = "Отдел"
    row[9] = "Демонстрация работы поиска"
    row[10] = "Тестовые сценарии"
    row[11] = "Показывает формат данных без раскрытия реальных записей"
    row[12] = "2024-01-01"
    row[13] = "DEMO-STRUCT-001"
    row[40] = "Иванов Иван Иванович"
    worksheet.append(row)


def main():
    workbook = openpyxl.Workbook()
    build_persona_sheet(workbook)
    build_structure_sheet(workbook)
    workbook.save(OUTPUT_PATH)
    print(f"Sample workbook created: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
