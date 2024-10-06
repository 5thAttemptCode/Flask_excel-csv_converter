from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook
import csv
import io

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

    # Prepare CSV data in-memory using io.StringIO
    output = io.StringIO()
    writer = csv.writer(output)
    
    for row in sheet.iter_rows(values_only=True):
        writer.writerow(list(row))
    
    output.seek(0)  # Go back to the start of the StringIO object
    
    # Send the CSV file as an attachment
    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='Payroll.csv'
    )

if __name__ == '__main__':
    app.run(debug=True)
