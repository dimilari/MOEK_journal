import os
import openpyxl
from last_month import days_in_last_month, last_month

# находим значение прошлого месяца в формате last_month = 'YYYYMM'
print(last_month)

# определяем путь к рабочему столу
desktop = os.environ['USERPROFILE'] + '\Desktop'

# если существует файл {last_month}.txt на рабочем столе
if os.path.exists(f'{desktop}/{last_month}.txt'):

    # парсим данные по расходу тепловой энергии за прошлый месяц из файла {last_month}.txt
    with open(f'{desktop}/{last_month}.txt', 'r') as file:

        # создаём список из строк текстового файла 'KM5.txt'
        lines = file.readlines()

        # дата начала расчетного периода
        date_start = lines[10][28:38]
        print(date_start)

        # дата окончания расчетного периода
        date_finish = lines[10][44:54]
        print(date_finish)

        # количество дней для составления отчёта
        days = int(date_finish[:2]) - int(date_start[:2]) + 1
        print(days)

        # парсим количество потребленной тепловой энергии за расчётный период
        total_Q = float(lines[19 + days][14:23].replace(' ', ''))
        print(total_Q)

        # парсим количество пройденного через систему теплоносителя за расчётный период
        total_M1 = float(lines[19 + days][24:32].replace(' ', ''))
        print(total_M1)

        # парсим показание расхода тепла на начало расчётного периода
        start_Q = lines[27 + days][34:44].replace(' ', '')
        print(start_Q)

        # парсим показание расхода тепла на конец расчётного периода
        stop_Q = lines[26 + days][34:44].replace(' ', '')
        print(stop_Q)

        # парсим количество пройденного через систему теплоносителя на начало расчётного периода
        start_М1 = lines[27 + days][45:54].replace(' ', '')
        print(start_М1)

        # парсим количество пройденного через систему теплоносителя на конец расчётного периода
        stop_М1 = lines[26 + days][45:54].replace(' ', '')
        print(stop_М1)

        # создаем список булевых значений ежедневного расхода теплоносителя
        list_Q = []
        count = float(start_Q)
        for i in range(18, 18 + days):
            count += float((lines[i][14:23]).replace(' ', ''))
            list_Q.append(float((lines[i][14:23]).replace(' ', '')) != 0)
        print(list_Q)

        # находим первое вхождение True в списке list_Q для его замены на start_Q
        first_True = (list_Q.index(True))

        # находим последнее вхождение True в списке list_Q для его замены на stop_Q
        last_True = -list_Q[::-1].index(True)-1

        # заменяем первое вхождение True в списке list_Q на start_Q
        list_Q[first_True] = float(start_Q)

        # заменяем последнее вхождение True в списке list_Q на stop_Q
        list_Q[last_True] = float(stop_Q)
        print(list_Q)

        # оставшееся количество True в списке list_Q
        days_true = list_Q.count(True)
        print(days_true)

        # сколько Гкал нужно прибавлять ежедневно в отчёте если был расход теплоэнергии в этот день
        delta_Q = round(total_Q / (days_true + 1), 3)
        print(delta_Q)

        # заменяем все предшествующие первому вхождению True элементы списка list_Q на start_Q
        count = 0
        for i in list_Q:
            if i == False:
                list_Q[count] = float(start_Q)
            else:
                break

        # окончательный список list_Q для заполнения файла отчёта
        count = 0
        for i in list_Q:
            # вместо True подставляем значение, которое больше предыдущего на delta_Q
            # округляем его до трёх знаков после запятой
            if i == True:
                list_Q[count] = round(list_Q[count - 1] + delta_Q, 3)
            # вместо False подставляем значение, которое равно предыдущему
            # округляем его до трёх знаков после запятой
            elif i == False:
                list_Q[count] = round(list_Q[count - 1], 3)
            count += 1


else:
    # путь к папке "Журнал"
    path_journal = r'D:\0=0=0=0=0\ГерПан#\. МОЭК\_Отчёт\Журнал'
    # имя последней папки в папке "Журнал"
    target_folder_name = os.listdir(path_journal)[-1]
    # имя последнего файла в последней папке внутри папки "Журнал"
    target_file_name = os.listdir(path_journal + '\\' + target_folder_name)[-1]
    # путь к файлу отчёта прошлого месяца
    target_file_path = path_journal + '\\' + target_folder_name + '\\' + target_file_name
    print(target_file_path)

    # открываем файл отчёт прошлого месяца
    book = openpyxl.load_workbook(filename=target_file_path, read_only=True)
    sheet = book.active
    print(sheet.max_row)
    print(type(sheet.max_row))
    print(sheet[sheet.max_row - 2][1].value)
    list_Q = [sheet[sheet.max_row - 2][1].value] * days_in_last_month
print(list_Q)
print(len(list_Q))

