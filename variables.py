"""Данный скрипт опрделяет переменные, словари, функции для работы соседних скриптов"""

from datetime import datetime, date, timedelta
from os import path, environ
from sys import argv


# функция для подсчёта дней между датами
def diff_dates(start_day, stop_day):
    start = datetime.strptime(start_day, '%d.%m.%Y').date()
    stop = datetime.strptime(stop_day, '%d.%m.%Y').date()
    return (stop - start).days


# определяем текущую дату
today = date.today()
# заменяем в текущей дате день на 1
first_day_of_this_month = today.replace(day=1)
# отнимаем один день от изменённой выше даты
last_day_of_last_month = first_day_of_this_month - timedelta(days=1)
# определяем количество дней в прошлом месяце
days_in_last_month = int(last_day_of_last_month.strftime('%d'))
# определяем формат даты прошлого месяца как YYYYMM
last_month_date_YYYYMM = int(last_day_of_last_month.strftime('%Y%m'))
# определяем формат даты прошлого месяца как MM
last_month_date_MM = last_day_of_last_month.strftime('%m')
# определяем формат даты прошлого месяца как YYYY
last_month_date_YYYY = last_day_of_last_month.strftime('%Y')
# определяем первый день прошлого месяца
first_day_of_last_month = last_day_of_last_month.replace(day=1)
# определяем последний день позапрошлого месяца
last_day_of_month_before_last = first_day_of_last_month - timedelta(days=1)
# определяем формат даты позапрошлого месяца как YYYYMM
month_before_last_date_YYYYMM = int(last_day_of_month_before_last.strftime('%Y%m'))
# определяем формат даты позапрошлого месяца как YYYY
month_before_last_date_YYYY = int(last_day_of_month_before_last.strftime('%Y'))

# определяем путь на рабочий стол
desktop = environ['USERPROFILE'] + '\Desktop'
# определяем путь к данному проекту
this_project_path = path.dirname(path.abspath(argv[0]))
# путь к папке "Журнал"
path_journal = r'D:\0=0=0=0=0\ГерПан#\. МОЭК\_Отчёт\Журнал'
# путь к папке "СправкаТепло"
path_certificate = r'D:\0=0=0=0=0\ГерПан#\. МОЭК\_Отчёт\СправкаТепло'

# определяем словарь с названиями месяцев на русском языке по их номерам
month_ru = {
    '01' : 'январь',
    '02' : 'февраль',
    '03' : 'март',
    '04' : 'апрель',
    '05' : 'май',
    '06' : 'июнь',
    '07' : 'июль',
    '08' : 'август',
    '09' : 'сентябрь',
    '10' : 'октябрь',
    '11' : 'ноябрь',
    '12' : 'декабрь'
}

