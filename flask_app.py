# Started based on this sample:
# http://stackoverflow.com/questions/27628053/uploading-and-downloading-files-with-flask
from flask import Flask, make_response, request
import openpyxl

app = Flask(__name__)


def findColumnByName(sheet,name):
  for c in range(ord('A'),ord('Z')+1):
    if sheet.cell( chr(c) + '1' ).value == name:
      return chr(c)

@app.route('/')
def form():
    return """
        <html>
            <body>
                <h1>Do something with Excel Spreadsheet</h1>

                <form action="/transform" method="post" enctype="multipart/form-data">
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
    sheet = wb.get_sheet_by_name('MAIN')
    column = findColumnByName(sheet,'INSTRUCTOR')

    response = make_response(str(column))
    response.headers["Content-Disposition"] = "attachment; filename=result.csv"
    return response