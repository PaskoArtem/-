
# Задача 2. Формирование нового сводного файла


import os
import xlrd #чтобы читать xls
from openpyxl import Workbook


# Настройки пути к файлам и список столбцов
AUTUMN_FILE = "data/Список 15вар - осень.xls"   # файл Осень
SPRING_FILE = "data/Список 15вар - весна.xls"   # файл Весна
OUTPUT_FILE = "Сводный текущий год.xlsx"        # результат

# 11 исходных столбцов в файлах осень и весна
source_columns = [
    "Дисциплина", "Вид нагрузки", "Группы", "Нагрузка", "Поток",
    "Количества часов по УП", "Табельный №", "ФИО", "Должность",
    "Желаемая аудитория", "Описание (пожелания по дням недели, времени,"
    "требования по техническому обеспечению аудитории)",
]

output_header = ["семестр"] + source_columns #В сводном файле слева добавляем столбец «семестр»



# Чтение файла (осень или весна)
def read_semester_file(path, semester):
    """Читает лист и возвращает строки нагрузки.
    К каждой строке слева добавляется семестр ('осень' или 'весна').
    Строка 0 (шапка) и строка-итог в конце файла пропускаются.
    """
    book = xlrd.open_workbook(path) #присваевается файл? открываем "книгу"
    sheet = book.sheet_by_name("TDSheet") # (table document sheet) присваиваем лист из книги и дальше работаем с ним

    rows = [] #пустой список для готовых строк
    for r in range(1, sheet.nrows):# строка 0 — пропускается (цикл перебирает строки)
        discipline = str(sheet.cell_value(r, 0)).strip() # из столбца 0 (дисциплина) убираем лишние пробелы и превращаем с строку
        load_type = str(sheet.cell_value(r, 1)).strip() # из столбца 1 (вид нагрузки) убираем лишние пробелы и превращаем с строку

        # Строка-итог в конце: «Дисциплина» и «Вид нагрузки» пустые,
        # а в «Нагрузке» лежит общая сумма — пропускаем
        if discipline == "" and load_type == "":
            continue

        values = [] # пустой список для значения одной строки
        for c in range(sheet.ncols): # внутренний цикл перебирает столбцы в строках
            value = sheet.cell_value(r, c)
            if isinstance(value, str):
                value = value.strip()  # убираем лишние пробелы, если str
            values.append(value) # почищенные значения кладем в список values

        rows.append([semester] + values) # кладем готтовую строку в общий список rows и слева дописываем семестр
    return rows


# Построение сводного файла
def build_output_file():
    """Складывает «Осень» + «Весна» и сохраняет в OUTPUT_FILE"""
    autumn_rows = read_semester_file(AUTUMN_FILE, "осень")# читает файл и пишет осень
    spring_rows = read_semester_file(SPRING_FILE, "весна")# читает файл и пишет весна

    # Целиком копируем: сначала все строки осени, затем все строки весны
    # Структуру строк не меняем
    all_rows = autumn_rows + spring_rows

    workbook = Workbook() # создаем новую пустую excel лист
    sheet = workbook.active # лист по умолчанию кладем в sheet
    sheet.title = "Сводная" # даем имя

    sheet.append(output_header) # добавляем в конец листа шапку (1 строку)
    for row in all_rows: # цикл для готовых элементов списка
        sheet.append(row) # дописывает строки в лист

    workbook.save(OUTPUT_FILE) # сохранение таблицы в файл
    return autumn_rows, spring_rows # возваращаем кортежом строки осени и весны, чтобы посчитать их количество и сумму для проверки


# Проверка корректности: сумма часов должна сойтись с изначальным количеством
def total_hours(rows): #функция считает сумму часов
    """Сумма по столбцу «Нагрузка» (индекс 4 после добавления семестра)."""
    return round(sum(r[4] for r in rows if isinstance(r[4], (int, float))), 2) # складывает только целые или дробные ЧИСЛА, округляет до 2 знаков после запятой


# Точка входа
def main():
    autumn_rows, spring_rows = build_output_file()

    print("Сводный файл сформирован:", os.path.abspath(OUTPUT_FILE))
    print("Строк осеннего файла:", len(autumn_rows))
    print("Строк весеннего файла:", len(spring_rows))
    print("Всего строк:", len(autumn_rows) + len(spring_rows))
    print("Сумма часов (осень):", total_hours(autumn_rows))
    print("Сумма часов (весна):", total_hours(spring_rows))


if __name__ == "__main__":
    main()