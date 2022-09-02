"""
Для запуска программы необходимо наличие файла 'YYYYMM.txt' с данными по
расходу тепловой энергии за прошлый месяц на рабочем столе (если его не
будет на рабочем столе, то программа посчитает нулевой расход теплоэнергии
за отчётный период) и введены данные за прошлый месяц в файл учёта воды
'_Показания всех счётчиков.xlsm'
Если в заполняемых бланках журнала(YYYYMM.xlsm) и справки(YYYYMM.docx)
появляются записи в виде строки 'error', значит нужно искать ошибку
в корректности дат отчётного периода файла 'YYYYMM.txt' на рабочем
столе и (или) в файле '_Показания всех счётчиков.xlsm' не была произведена
запись расхода воды за оформляемый период
"""

import os
from win32com.client import Dispatch
from time import sleep
from variables import \
    path_journal, \
    last_month_date_MM, \
    last_month_date_YYYY, \
    last_month_date_YYYYMM, \
    this_project_path, \
    month_ru, \
    path_certificate
from parsing_water_excel import \
    last_in, \
    last_out, \
    now_in, \
    now_out, \
    water_rate
from parsing_txt import start_Q, stop_Q, total_Q
from docxtpl import DocxTemplate
from parsing_txt import list_Q, list_M1, list_T
from parsing_water_excel import list_in, list_out


# функция заполняет ячейки данными из списка some_list по вертикали от пятой
# ячейки в колонке some_column до (длины списка + 5)
# аргументы - буква колонки (some_column), список (some_list)
def write_in(some_column, some_list):
    for i in range(len(some_list)):
        ws.Cells(i+5, some_column).Value = some_list[i]

# функция для открытия окна windows explorer по заданному пути path
def open_window(path):
    path = os.path.realpath(path)
    os.startfile(path)

#---------------------------------------
'''в случае отсутствия в папке 'Журнал' папки с названием года
соответствующего дате файла отчёта, создаём такую папку'''
if not os.path.exists(fr'{path_journal}\{last_month_date_YYYY}'):
    os.mkdir(fr'{path_journal}\{last_month_date_YYYY}')
#---------------------------------------
'''открываем шаблон 'template.xlsm' из данного проекта, запускаем 
в нём макрос для формирования журнала учёта тепловой энергии, 
заполняем сформированный журнал учёта тепловой энергии и сохраняем его 
с актуальным на данный момент именем в предназначенную для него директорию'''
# запускаем приложение excel
xl = Dispatch('Excel.Application')
# открываем файл excel 'template.xlsm' в котором находится макрос
wb = xl.Workbooks.Open(rf'{this_project_path}\Template.xlsm', False, True)
# задаём активный лист в открытой книге excel
ws = wb.ActiveSheet
# запускаем в открытом файле макрос 'Macros_journal'
xl.Run('Macros_journal')
# задержка для предотвращения ошибки работы макроса 'Macros_journal'
# макрос добавлял в два раза больше строк в шаблон 'template.xlsm'
sleep(0.5)
# заполняем таблицу файла отчёта
write_in(2, list_Q)
write_in(3, list_M1)
write_in(4, list_T)
write_in(5, list_in)
write_in(6, list_out)
# сохраняем заполненный шаблон журнала отчёта в файл с актуальным на
# данный момент именем в предназначенную для него директорию
wb.SaveAs(Filename:=fr'{path_journal}\{last_month_date_YYYY}\{last_month_date_YYYYMM}.xlsm')
# закрываем изменённый файл
wb.Close()
# закрываем приложение excel
xl.Quit()
#---------------------------------------
'''случае отсутствия в папке "СправкаТепло" вложенной папки с названием 
года соответствующего дате файла отчёта, создаём такую папку'''
if not os.path.exists(fr'{path_certificate}\{last_month_date_YYYY}'):
    os.mkdir(fr'{path_certificate}\{last_month_date_YYYY}')
#---------------------------------------
'''открываем шаблон 'template.docx' из данного проекта,
заменяем в нём все шаблонные фразы и сохраняем его с актуальным 
на данный момент именем в предназначенную для него директорию'''
# открываем файл template.docx для внесения в него изменений
doc = DocxTemplate(fr'{this_project_path}\template.docx')
# определяем словарь для замены шаблонных фраз на нужную нам информацию
context = {
    'month'       : month_ru[last_month_date_MM],
    'year'        : last_month_date_YYYY,
    'heat_start'  : str(start_Q).replace('.', ','),
    'heat_stop'   : str(stop_Q).replace('.', ','),
    'total_heat'  : str(total_Q).replace('.', ',').ljust(5, '0'),
    'start_in'    : str(round(last_in)).replace('.', ',').rjust(5, '0'),
    'start_out'   : str(round(last_out)).replace('.', ',').rjust(5, '0'),
    'stop_in'     : str(round(now_in)).replace('.', ',').rjust(5, '0'),
    'stop_out'    : str(round(now_out)).replace('.', ',').rjust(5, '0'),
    'total_water' : str(water_rate).replace('.', ',')
}
# заменяем шаблонные фразы в template.docx на нужную нам информацию
doc.render(context)
# сохраняем заполненный шаблон справки по расходу тепловой энергии в
# файл с актуальным на данный момент именем в предназначенную для него директорию
doc.save(fr'{path_certificate}\{last_month_date_YYYY}\{last_month_date_YYYYMM}.docx')
#---------------------------------------
'''Открываем папки с созданными файлами отчёта для human контроля корректности 
и последующей их отправки по электронной почте'''
open_window(rf"D:\0=0=0=0=0\ГерПан#\. МОЭК\_Отчёт\СправкаТепло\{last_month_date_YYYY}")
open_window(rf"D:\0=0=0=0=0\ГерПан#\. МОЭК\_Отчёт\Журнал\{last_month_date_YYYY}")
open_window(rf"D:\0=0=0=0=0\ГерПан#\. МОЭК\_Отчёт\РаспечаткаКМ-5\{last_month_date_YYYY}")


