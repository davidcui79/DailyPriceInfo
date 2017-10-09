import tushare
import xlwt
import xlrd
from xlutils.copy import copy
import os
import time
from utils import insert_log
#from utils import first_date_of_last_month
import utils
import sys
import shutil
import PIL.Image
import requests
import traceback

#UP = 0
#DOWN = 1
#category = ['up', 'down']
#qualified = []

#for i in range(2):
#    qualified.append([])

def comp(x, y):
    id1 = x[0]
    id2 = y[0]
    if id1 < id2:
        return -1
    if id1 > id2:
        return 1
    if id1 == id2:
        return 0


date = time.strftime('%Y-%m-%d',time.localtime(time.time()))
REFERENCE_DIR = 'reference_' + date

LOG_FILE = ''
#
# important do not delete the directory, because the call of tushare.get_hist_data() is very
# slow and tend of fail somtimes, it's important to keep the reference.xls and resume from
# there
#
def init():
    if not os.path.exists(REFERENCE_DIR):
        os.mkdir(REFERENCE_DIR)

    log_fp = os.path.join(REFERENCE_DIR, 'log.txt')
    # create log file, overwrite mode
    global LOG_FILE
    LOG_FILE = open(log_fp, 'wb')

def generate_reference_xls():
    existing_rows = []

    #open the reference.xls file
    global date
    excel_fp = os.path.join(REFERENCE_DIR, 'reference_'+ date +'.xls')

    #create the xls file and the read_worksheet if it doesn't yet exist
    if not os.path.exists(excel_fp):
        write_workbook = xlwt.Workbook()
        write_workbook.add_sheet('all')
        write_workbook.save(excel_fp)

    #get the baseline existing_rows from the xls file
    read_workbook = xlrd.open_workbook(excel_fp)
    read_worksheet = read_workbook.sheet_by_name('all')
    rows = read_worksheet.nrows
    for i in range(0, rows):
        existing_rows.append(read_worksheet.row_values(i))
        #existing_rows.sort(comp)
    insert_log(LOG_FILE, 'Read ' + format(read_worksheet.nrows, "") + ' rows from file ' + excel_fp)

    write_workbook = copy(read_workbook)
    write_worksheet = write_workbook.get_sheet('all')

    #to skip the existing data we need to have the existing_rows of ID in existing_rows
    #we already know the format of existing_rows is [[id1, x, y, z, ...], [id2, x, y, z, ...] ...]
    ids_in_worksheet =[]
    for i in existing_rows:
        ids_in_worksheet.append(i[0])

    #get all stock info
    stock_info = tushare.get_stock_basics()
    insert_log (LOG_FILE, 'There are ' + format(len(stock_info), "") + ' items from tushare.get_stock_basics()')

    count = 0

    for id in stock_info.index:
        count += 1
        insert_log(LOG_FILE, 'processing ' + format(count, "") + ': ' + id)

        if id in ids_in_worksheet:
            insert_log(LOG_FILE, 'Already has ' + id + ' skip it')
            continue

        #test code
        #if count > 50:
        #     break

        month = utils.last_month()
        try:
            history_data = tushare.get_hist_data(id, start=utils.first_date_of_last_month(), retry_count=5, pause=1)
            #print history_data.columns
            #print history_data[0:4]
            close_price = history_data[u'close']
            print history_data
        except Exception:
            ##  nothing to handle
            insert_log(LOG_FILE, 'Exception when handling ' + id)
            info = sys.exc_info()
            print (info[0], ':', info[1])
            continue

        continous_up = False
        continous_down = False

        #only need to analyze if we have at least 4 sample
        if(len(close_price) >= 4):
            continous_up = True
            continous_down = True
            for i in range(0, 3):
                if(close_price[i] < close_price[i+1]):
                    continous_up = False
                    break

            for i in range(0, 3):
                if(close_price[i] > close_price[i+1]):
                    continous_down = False
                    break

        #row = read_worksheet.nrows
        #read_worksheet.write(row, 0, id)
        try:
            record = []

            date = close_price.keys()[0]

            three_days_ago = 'NA'
            if len(close_price.keys()) >= 4:
                three_days_ago = close_price.keys()[3]

            open_price = history_data[u'open'][0]
            high = history_data[u'high'][0]
            low = history_data[u'low'][0]
            price_change = history_data[u'price_change'][0]
            volume = history_data[u'volume'][0]
            p_change = history_data[u'p_change'][0]
            ma5 = history_data[u'ma5'][0]
            ma10 = history_data[u'ma10'][0]
            ma20 = history_data[u'ma20'][0]
            v_ma5 = history_data[u'v_ma5'][0]
            v_ma10 = history_data[u'v_ma10'][0]
            v_ma20 = history_data[u'v_ma20'][0]
            turnover = history_data[u'turnover'][0]

            trend = ''
            #
            #[id, 3_day_trend, date, open price, close price, high, low, volume, price_change, p_change, ma5, ma10, ma20, v_ma5, v_ma10, v_ma20, turnover]
            #
            if(continous_up):
                trend = 'up'
            elif (continous_down):
                trend = 'down'
            else:
                trend = 'NA'

            record.append(id)
            record.append(trend)
            record.append(date)
            record.append(three_days_ago)
            record.append(open_price)
            record.append(close_price[0])
            record.append(high)
            record.append(low)
            record.append(volume)
            record.append(price_change)
            record.append(p_change)
            record.append(ma5)
            record.append(ma10)
            record.append(ma20)
            record.append(v_ma5)
            record.append(v_ma10)
            record.append(v_ma20)
            record.append(turnover)

            for i in range(len(record)):
                write_worksheet.write(rows, i, record[i])

            rows += 1

            write_workbook.save(excel_fp)
            insert_log(LOG_FILE, 'written to file ' + excel_fp)

        except Exception, e:
            insert_log(LOG_FILE, 'Exception when handling id ' + id)
            info = sys.exc_info()
            print traceback.print_exc()
            continue

        #existing_rows.append([id, trend, date, open_price, close_price[0], high, low, price_change, ma5, ma10, ma20, v_ma5, v_ma10, v_ma20, turnover])
        insert_log(LOG_FILE, id + ' 3 day trend is ' + trend)


    insert_log(LOG_FILE, 'Finished populating reference.xls')

#
# this function creates directory JPEGImages and downloads all the k line images to it
# if the JPEGImages directory already exists, it will be deleted so every time it's a
# full download of all images
#
def download_daily_image():
    JPEG_DIR = os.path.join(REFERENCE_DIR, 'JPEGImages')

    if os.path.exists(JPEG_DIR):
        shutil.rmtree(JPEG_DIR)
    os.mkdir(JPEG_DIR)

    stock_info = tushare.get_stock_basics()

    # download gif and save as jpeg for each id from tushare
    for id in stock_info.index:
        # if id != '002270': continue

        # insert prefix
        if ((id.find('300') == 0 or id.find('00') == 0)):
            id = 'sz' + id
        elif (id.find('60') == 0):
            id = 'sh' + id
        else:
            LOG_ID_NOT_HANDLED = 'Don\'t know how to handle ID:' + id + ', skip and continue'
            insert_log(LOG_FILE, LOG_ID_NOT_HANDLED)
            continue

        insert_log(LOG_FILE, 'Start downloading gif for ' + id)

        # download the gif file
        url = 'http://image.sinajs.cn/newchart/daily/n/' + id + '.gif'
        gif = requests.get(url)

        try:
            # save gif file to folder IMAGE_DATA_DIR
            gif_file_name = os.path.join(JPEG_DIR, id + '.gif')
            gif_fp = open(gif_file_name, 'wb')
            gif_fp.write(gif.content)
            gif_fp.close()
            insert_log(LOG_FILE, 'Complete downloading ' + id + '.gif')

            # save the gif to jpeg
            # need the convert otherwise will report raise IOError("cannot write mode %s as JPEG" % im.mode)
            im = PIL.Image.open(gif_file_name)
            im = im.convert('RGB')
            jpeg_file = os.path.join(JPEG_DIR, id + '.jpeg')
            im.save(jpeg_file)
            insert_log(LOG_FILE, 'Saved file ' + os.path.realpath(jpeg_file))
            # delete the gif file
            os.remove(gif_file_name)

        except IOError, err:
            insert_log(LOG_FILE, 'IOError when handling ' + id)
            info = sys.exc_info()
            print (info[0], ':', info[1])
            continue
        except Exception, e:
            insert_log(LOG_FILE, 'Exception when handling ' + id)
            info = sys.exc_info()
            print (info[0], ':', info[1])
            continue

if __name__ == "__main__":
    init()
    generate_reference_xls()
    download_daily_image()
    LOG_FILE.close()
