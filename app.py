from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook
import csv
import os

app = Flask(__name__)

# Route to upload the Excel file
@app.route('/')
def upload_file():
    return render_template('index.html')

# Route to process the Excel file and convert it to CSV
@app.route('/convert', methods=['POST'])
def convert_excel_to_csv():
    # Check if the post request has the file part
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']

    # Check if the file is selected
    if file.filename == '':
        return "No selected file"
    
    # Load the uploaded Excel file
    wb = load_workbook(file)
    sheet = wb.active

    # Prepare CSV data
    csv_data = []
    for value in sheet.iter_rows(values_only=True):
        csv_data.append(list(value))

    # Save CSV file in the current directory
    csv_filename = 'Payroll.csv'
    with open(csv_filename, 'w', newline="") as csv_obj:
        writer = csv.writer(csv_obj)
        writer.writerows(csv_data)

    # Send the CSV file as a response
    return send_file(csv_filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
