import xlrd
import re

#СКРИПТ ДЛЯ РАБОТЫ С XLS расширением

# Открываем файлы Excel и получаем доступ к активным листам
wbs = input("Введите путь к файлу(/путь/имяфайла.xls): ")
wb1 = xlrd.open_workbook(wbs)
ws1 = wb1.sheet_by_index(0)

wbr = input("Введите путь к файлу(/путь/имяфайла.xls): ")
wb2 = xlrd.open_workbook(wbr)
ws2 = wb2.sheet_by_index(0)

# Диапазон ячеек, из которых нужно извлекать значения
cell_range1 = 'A1:A200'

# Содержимое ячеек из файла 1
words_to_find = {}

for row in range(ws2.nrows):
    for cell in ws2.row(row):
        if cell.value:
            #регулярное выражение для разбиения содержимого ячейки на слова
            words_in_cell = re.findall(r'\b\w+\b', str(cell.value))
            #Добавляем каждое найденное слово в диапазоне ячеек в словарь words_to_find со значением-списком ячеек
            #в которых оно было найдено
            for word in words_in_cell:
                word = word.lower()
                if word in words_to_find:
                    words_to_find[word].append(cell.value)
                else:
                    words_to_find[word] = [cell.value]

# Ищем слова в файле 2 и выводим результаты
found_words = []
not_found_words = {}

cell_range2 = 'A1:A200'

for row in range(ws1.nrows):
    for cell in ws1.row(row):
        if cell.value:
            #регулярное выражение для разбиения содержимого ячейки на слова
            words_in_cell = re.findall(r'\b\w+\b', str(cell.value))
            #Приводим слова к нижнему регистру и ищем совпадения
            #В каждой ячейке ищем все слова и сохраняем только первое найденное совпадение,
            #чтобы избежать дублирования найденных слов в списке found_words
            for word in words_in_cell:
                word = word.lower()
                if word in words_to_find and word not in found_words:
                    found_words.append(word)

            #Добавляем ячейку и ее содержимое в словарь not_found_words,
            #если не было найдено ни одного совпадения
            if all(word.lower() not in words_to_find for word in words_in_cell):
                not_found_words[cell.value] = words_in_cell


print("\nНайденные слова в файле 1 и соответствующие им ячейки в [] из файла 2:\n")
for word in found_words:
    print(word + ":", words_to_find[word])
    print("\n ")
print("\nНе найденные слова:\n")
for key, value in not_found_words.items():
    print(key)
    print("\n ")
