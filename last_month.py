import datetime
# определяем текущую дату
today = datetime.date.today()
# заменяем в текущей дате день на 1
first_day_of_this_month = today.replace(day=1)
# отнимаем один день от изменённой выше даты
last_day_of_last_month = first_day_of_this_month - datetime.timedelta(days=1)
# определяем количество дней в прошлом месяце
days_in_last_month = int(last_day_of_last_month.strftime('%d'))
# определяем формат даты прошлого месяца как YYYYMM
last_month = int(last_day_of_last_month.strftime('%Y%m'))

