import os
import openpyxl
from last_month import days_in_last_month, last_month


# определяем путь к рабочему столу
desktop = os.environ['USERPROFILE'] + '\Desktop'

# путь к папке "Журнал"
path_journal = r'D:\0=0=0=0=0\ГерПан#\. МОЭК\_Отчёт\Журнал'
# имя последней папки в папке "Журнал"
target_folder_name = os.listdir(path_journal)[-1]
# имя последнего файла в последней папке внутри папки "Журнал"
target_file_name = os.listdir(path_journal + '\\' + target_folder_name)[-1]
# путь к последнему существующему файлу отчёта
target_file_path = path_journal + '\\' + target_folder_name + '\\' + target_file_name

# если существует файл {last_month}.txt на рабочем столе
if os.path.exists(f'{desktop}/{last_month}.txt'):

    # парсим данные по расходу тепловой энергии за прошлый месяц из файла {last_month}.txt
    with open(f'{desktop}/{last_month}.txt', 'r') as file:

        # создаём список из строк текстового файла 'KM5.txt'
        lines = file.readlines()

        # дата начала расчетного периода
        date_start = lines[10][28:38]

        # дата окончания расчетного периода
        date_finish = lines[10][44:54]

        # количество дней для составления отчёта
        days = int(date_finish[:2]) - int(date_start[:2]) + 1

        # парсим количество потребленной тепловой энергии за расчётный период
        total_Q = float(lines[19 + days][14:23].replace(' ', ''))

        # парсим циркуляцию теплоносителя за расчётный период
        total_M1 = float(lines[19 + days][24:32].replace(' ', ''))

        # парсим показание расхода тепла на начало расчётного периода
        start_Q = lines[27 + days][34:44].replace(' ', '')

        # парсим показание расхода тепла на конец расчётного периода
        stop_Q = lines[26 + days][34:44].replace(' ', '')

        # парсим циркуляцию теплоносителя на начало расчётного периода
        start_M1 = lines[27 + days][45:54].replace(' ', '')

        # парсим циркуляцию теплоносителя на конец расчётного периода
        stop_M1 = lines[26 + days][45:54].replace(' ', '')

        # создаём список едеждевного расхода тепла
        daily_usage_Q = []
        for i in range(18, 18 + days):
            daily_usage_Q.append(float(lines[i][14:23].replace(' ', '')))

        # создаём список едеждевной циркуляции теплоносителя
        daily_usage_M1 = []
        for i in range(18, 18 + days):
            daily_usage_M1.append(float(lines[i][24:32].replace(' ', '')))

        # создаем список ежедневных показаний счётчика расхода тепла
        # на нулевое место этого списка ставим значение показания счётчика
        # расхода тепла на конец прошлого периода
        list_Q = [float(start_Q)]
        change_daily_counter = float(start_Q)
        for i in range(1, len(daily_usage_Q) - 1):
            if daily_usage_Q[i] == 0.0:
                change_daily_counter += daily_usage_Q[i-1]
            else:
                change_daily_counter += daily_usage_Q[i]
            list_Q.append(round(change_daily_counter, 3))
        list_Q.append(float(stop_Q))

        # создаем список ежедневных показаний счётчика циркуляции теплоносителя
        # на нулевое место этого списка ставим значение показания счётчика
        # циркуляции теплоносителя на конец прошлого периода
        list_M1 = [float(start_M1)]
        change_daily_counter = float(start_M1)
        for i in range(1, len(daily_usage_M1) - 1):
            if daily_usage_M1[i] == 0.0:
                change_daily_counter += daily_usage_M1[i-1]
            else:
                change_daily_counter += daily_usage_M1[i]
            list_M1.append(round(change_daily_counter, 3))
        list_M1.append(float(stop_M1))

        # парсим время работы приборов учёта
        list_T = []
        for i in range(18, 18 + days):
            list_T.append(float(lines[i][64:70].replace(' ', '')))

else:
    # открываем последний существующий файл отчёта
    book = openpyxl.load_workbook(filename=target_file_path, read_only=True)
    sheet = book.active
    print(sheet.max_row)
    print(type(sheet.max_row))
    print(sheet[sheet.max_row - 2][1].value)
    list_Q = [sheet[sheet.max_row - 2][1].value] * days_in_last_month
    list_M1 = [sheet[sheet.max_row - 2][2].value] * days_in_last_month
    list_T = [0.0] * days_in_last_month
print(list_Q)
print(list_M1)
print(list_T)


