import flask
import werkzeug
from werkzeug import *
import xlrd
import os
from flask import (
    redirect,
    render_template,
    request,
    url_for,
)
from flask import Flask
import csv
app = Flask(__name__)
ALLOWED_EXTENSIONS = set(['csv', 'xlsx', 'xls'])

@app.route('/', methods = ['GET'])
def reciever():
    return render_template('default.html')


@app.route('/csvparse', methods = ['POST'])
def csv():
    file = request.files['file']
    if file and allowed_file(file.filename):
        extention = str(file.filename.rsplit('.', 1)[1].lower())
        if extention == "xlsx" or extention == "xls":
            dirToSave = str(os.getcwd())
            filename = secure_filename(file.filename)
            dirToSave += "/spreadsheets/" + filename
            file.save(dirToSave)
            workbook = xlrd.open_workbook(dirToSave)
            worksheet = workbook.sheet_by_index(0)
            #must make sure that the rows are filled.
            for rx in range(worksheet.nrows):
                allFound = True
                try:
                    int(worksheet.cell_value(rowx = rx, colx = 0))
                except (IndexError, ValueError):
                    allFound = False
                    break
                try:
                    int(worksheet.cell_value(rowx = rx, colx = 1))
                except (IndexError, ValueError):
                    allFound = False
                    break
                try:
                    int(worksheet.cell_value(rowx = rx, colx = 2))
                except (IndexError, ValueError):
                    allFound = False
                    break
                if allFound:
                    print(int(worksheet.cell_value(rowx = rx, colx = 0)))
                    print(int(worksheet.cell_value(rowx = rx, colx = 1)))
                    print(int(worksheet.cell_value(rowx = rx, colx = 2)))

            #deleted file
            os.remove(dirToSave)
    return render_template('default.html')



def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


if __name__ == "__main__":
    app.run()
