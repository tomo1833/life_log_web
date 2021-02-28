import calendar
import datetime
from flask import Flask, make_response, render_template, request
import json
import locale
import openpyxl

app = Flask(__name__)

@app.route('/')
def indexpage():
    return render_template('index.html')

@app.route('/caldownload', methods=['POST'])
def caldownload():
    #locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
    print(locale.getlocale(locale.LC_TIME))
    locale.setlocale(locale.LC_TIME, 'ja_JP.utf8')
    print(locale.getlocale(locale.LC_TIME))

    result = request.get_data()
    dec = json.loads(result)

    for event in dec:
         print(event['start'])
         a = datetime.datetime.strptime(event['start']['dateTime'], '%Y-%m-%dT%H:%M:%S%z')
         print(a)
    wb = openpyxl.load_workbook('output.xlsx')
    sheet = wb['Sheet']

    for dateIndex in range(calendar.monthrange(2021, 2)[0], calendar.monthrange(2021, 2)[1]):
        sheet['A' + str(dateIndex + 5) ] = dateIndex + 1
        sheet['C' + str(dateIndex + 5) ] = datetime.date(2021, 2, dateIndex + 1).strftime('%a')
        for event in dec:
            startTime = datetime.datetime.strptime(event['start']['dateTime'], '%Y-%m-%dT%H:%M:%S%z')
            endTime = datetime.datetime.strptime(event['end']['dateTime'], '%Y-%m-%dT%H:%M:%S%z')
            if dateIndex + 1 == startTime.day:
                sheet['D' + str(dateIndex + 5) ] = '○'
                sheet['E' + str(dateIndex + 5) ] = startTime.strftime('%H:%M')
                sheet['F' + str(dateIndex + 5) ] = endTime.strftime('%H:%M')

    wb.save('output.xlsx')

    XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response = make_response()
    response.data = open("output.xlsx", "rb").read()
    downloadFileName = 'output.xlsx'  
    response.headers['Content-Disposition'] = 'attachment; filename=' + downloadFileName
    response.mimetype = XLSX_MIMETYPE
    return response
    

# run the app.
if __name__ == "__main__":
    app.debug=True
    app.run(host='0.0.0.0')
