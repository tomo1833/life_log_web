from flask import Flask, make_response, render_template, request

import calendar
import datetime
import shutil
import json
import locale
import openpyxl
import os

app = Flask(__name__)

@app.route('/')
def indexpage():

    env_client_id = os.environ['CLIENT_ID']
    env_api_key = os.environ['API_KEY']

    return render_template('index.html', client_id=env_client_id, api_key=env_api_key)

@app.route('/caldownload', methods=['POST'])
def caldownload():
    locale.setlocale(locale.LC_TIME, 'ja_JP.utf8')

    result = request.get_data()
    dec = json.loads(result)
    
    # print(dec)
    # 日付取得
    today_time = datetime.datetime.now()
    today_date_str = today_time.strftime("%Y%m%d")
    today_time_str = today_time.strftime("%Y%m%d%H%M%S")
    yesterday_date_str = (today_time - datetime.timedelta(days=1)).strftime("%Y%m%d")

    # 前日：ディレクトリ削除/当日：ディレクトリ作成
    # TODO 削除は前日以前
    yestaday_directory_path = './' + yesterday_date_str
    today_directory_path = './' + today_date_str
    if (os.path.isdir(yestaday_directory_path)):
        shutil.rmtree(yestaday_directory_path)
    elif not (os.path.isdir(today_directory_path)):
        os.mkdir(today_directory_path)

    # 日付ファイル作成
    new_file_name = today_time_str + '_output.xlsx'
    new_file_full_path = today_directory_path + '/' + new_file_name

    # Excel file copy.
    shutil.copy('./output.xlsx', new_file_full_path)

    # WorkBookロード
    wb = openpyxl.load_workbook(new_file_full_path)
    sheet = wb['Sheet']

    # WorkBookセル書き込み
    # 初期値
    year = 2020
    month = 1
    last_data = dec[len(dec) - 1 ]
    if 'requestdate' in last_data:
        year = int(last_data['requestdate'][:4])
        month = int(last_data['requestdate'][4:])

    for dateIndex in range(calendar.monthrange(year, month)[0], calendar.monthrange(year, month)[1]):
        sheet['A' + str(dateIndex + 5) ] = dateIndex + 1
        sheet['C' + str(dateIndex + 5) ] = datetime.date(year, month, dateIndex + 1).strftime('%a')
        for event in dec:
            if 'start' in event:
                startTime = datetime.datetime.strptime(event['start']['dateTime'], '%Y-%m-%dT%H:%M:%S%z')
                endTime = datetime.datetime.strptime(event['end']['dateTime'], '%Y-%m-%dT%H:%M:%S%z')
                if dateIndex + 1 == startTime.day:
                    sheet['D' + str(dateIndex + 5) ] = '○'
                    sheet['E' + str(dateIndex + 5) ] = startTime.strftime('%H:%M')
                    sheet['F' + str(dateIndex + 5) ] = endTime.strftime('%H:%M')

    # WorkBook書き込み
    wb.save(new_file_full_path)

    # スプレットシートダウンロード
    XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response = make_response()
    response.data = open(new_file_full_path, "rb").read()
    downloadFileName = 'new_file_name'  
    response.headers['Content-Disposition'] = 'attachment; filename=' + downloadFileName
    response.mimetype = XLSX_MIMETYPE

    return response
    

# run the app.
if __name__ == "__main__":
    app.debug=True
    app.run(host='0.0.0.0')
