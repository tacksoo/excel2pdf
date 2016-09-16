# Started based on this sample:
# http://stackoverflow.com/questions/27628053/uploading-and-downloading-files-with-flask
from flask import Flask, make_response, request
import openpyxl # using version 1.6.1

app = Flask(__name__)


def get_column_by_name(sheet,name):
  ''' Return Excel column (A,B,C...) based on column heading '''
  for c in range(ord('A'),ord('Z')+1):
    # hard coded to row two, find a better way to do this
    if sheet.cell( chr(c) + '2' ).value == name:
      return chr(c)

def get_names(sheet,column):
  names = ''
  # lastname
  temp = ''
  count = 3
  while sheet.cell(column + str(count)).value != None:
    names += sheet.cell(column +str(count)).value
    names += "\n"
    count += 1
  return names

@app.route('/')
def form():
    return """
        <html>
            <body>
                <h1>Upload directory excel file</h1>

                <form action="/transform" method="post" enctype="multipart/form-data">
                    <input type="file" name="data_file" />
                    <input type="submit" />
                </form>

                <h1>Upload courses excel file</h1>

                <form action="/results" method="post" enctype="multipart/form-data">
                    <input type="file" name="data_file" />
                    <input type="submit" />
                </form>
            </body>
        </html>
    """

@app.route('/transform', methods=["POST"])
def transform_view():
    file = request.files['data_file']
    if not file:
        return "No file"

    # openpyxl documentation found here: https://media.readthedocs.org/pdf/openpyxl/latest/openpyxl.pdf
    wb = openpyxl.load_workbook(file)
    sheet = wb.get_sheet_by_name('Last Name')
    lastname_column = get_column_by_name(sheet,'Last Name')
    names = get_names(sheet,lastname_column)

    response = make_response(names)
    response.headers["Content-Disposition"] = "attachment; filename=part_time_faculty.csv"
    return response