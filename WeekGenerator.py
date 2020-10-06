import datetime

def weekGenerator(date):
    this_day = date.weekday()
    monday = date - datetime.timedelta(days = this_day)
    dates = [(monday + datetime.timedelta(days=i)).strftime('%d/%m/%Y') for i in range(5)]
    return dates

today = datetime.date(2020,10,12)

dates = weekGenerator(today)
print(dates)
