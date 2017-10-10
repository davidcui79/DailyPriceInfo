import time
import os

def insert_log(fp, string, stdout=True):

    try:
        date_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        print(date_time + ' ' + string)
        fp.write(date_time + ' ' +  string + '\n')
    except Exception:
        print('Cannot insert log to file ' + fp)


def last_month(month= time.strftime('%m', time.localtime(time.time()))):
    int_month = int(month)
    if int_month == 1:
        return str(12)
    else:
        return str(int_month - 1).zfill(2)

def first_date_of_last_month():
    month = time.strftime('%m', time.localtime(time.time()))
    year = time.strftime('%Y', time.localtime(time.time()))
    if int(month) == 1:
        year = str(int(year) - 1)
        return str(year)+'-12-01'
    else:
        month = str(int(month) - 1).zfill(2)
        return year+'-'+month+'-01'

def first_date_of_previous_month(year, month):
    if int(month) == 1:
        year = str(int(year) - 1)
        return str(year)+'-12-01'
    else:
        month = str(int(month) - 1).zfill(2)
        return year+'-'+month+'-01'

def this_year():
    return time.strftime('%Y', time.localtime(time.time()))

def get_user_home():
    return os.path.expandvars('$HOME')

def get_path_to_data_dir():
    path = os.path.join(get_user_home(), 'DailyPriceInfo_data')
    if not os.path.exists(path):
        os.mkdir(path)
    return path