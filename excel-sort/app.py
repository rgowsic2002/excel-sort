import os
import io
import zipfile
import tempfile
import xlsxwriter
import pandas as pd
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files['file']
        if excel_file:
            temp_dir = tempfile.TemporaryDirectory()
            file_path = os.path.join(temp_dir.name, excel_file.filename)
            excel_file.save(file_path)

            # Read the Excel file
            df = pd.read_excel(file_path, sheet_name=0)

            # Filter data by "SeqStatusLabel" column
            df_filtered = df[df["SeqStatusLabel"].isin(["In-Process", "Not Started"])]

            # Group data by "WCdesc" column
            df_grouped = df_filtered.groupby("WCdesc")

            # Create separate Excel files for each unique "WCdesc" value
            excel_files = {}
            for name, group in df_grouped:
                # Sort data by "RawDesc" column
                group_sorted = group.sort_values(by=["RawDesc"])
                output = io.BytesIO()
                group_sorted.to_excel(output, index=False)
                output.seek(0)
                excel_files[name] = output

            # Zip the Excel files
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, mode='w') as zip_file:
                for name, data in excel_files.items():
                    zip_file.writestr(f"{name}.xlsx", data.read())
            zip_buffer.seek(0)

            # Provide a link for users to download the zip file
            return send_file(zip_buffer, as_attachment=True, download_name="output.zip")

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)