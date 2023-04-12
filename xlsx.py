import openpyxl
import re

#СКРИПТ ДЛЯ РАБОТЫ С XLSX РАСШИРЕНИЕМ

#Открываем файлы Excel и получаем доступ к активным листам
wbs = input("Введите путь к файлу 1(/путь/имяфайла.xlsx): ")
wb1 = openpyxl.load_workbook(wbs)
#выбираем либо определенный лист wb1.['лист'], либо активный как указан в ws1
ws1 = wb1.active

wbr = input("Введите путь к файлу 2(/путь/имяфайла.xlsx): ")
wb2 = openpyxl.load_workbook(wbr)
ws2 = wb2.active

#Диапазон ячеек, из которых нужно извлекать значения
cell_range1 = 'A1:A2000'

#Содержимое ячеек из файла 1
words_to_find = {}

for row in ws2[cell_range1]:
    for cell in row:
        if cell.value:
            #Регулярное выражение для разбиения содержимого ячейки на слова
            words_in_cell = re.findall(r'\b\w+\b', cell.value)
            #Добавляем каждое найденное слово в диапазоне ячеек в словарь words_to_find со значением-списком ячеек
            #в которых оно было найдено
            for word in words_in_cell:
                word = word.lower()
                if word in words_to_find:
                    words_to_find[word].append(cell.value)
                else:
                    words_to_find[word] = [cell.value]


#Ищем слова в файле 2 и выводим результаты
found_words = []
not_found_words = {}

cell_range2 = 'A1:A2000'

for row in ws1[cell_range2]:
    for cell in row:
        if cell.value:
            #Регулярное выражение для разбиения содержимого ячейки на слова
            words_in_cell = re.findall(r'\b\w+\b', cell.value)
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
#print(found_words + ":", words_to_find[word])
for word in found_words:
    print(word + ":", words_to_find[word])
    print("\n ")
print("\nНе найденные слова:\n")
for key, value in not_found_words.items():
    print(key)
    #print(key + ":", ', '.join(value))
    print("\n ")
