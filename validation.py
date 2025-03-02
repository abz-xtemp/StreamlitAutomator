import xlwings as xw
import pandas as pd
import tempfile
import time
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import logging
import io
import traceback

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def process(validation_file):
    """
    Performs validation checks on the given Excel file.
    Identifies missing values, formula inconsistencies, and generates a validation report.
    Highlights error cells in the original file.
    Returns bytes of the report file, or an error message within bytes.
    """
    output = io.BytesIO()  # Initialize output as BytesIO

    try:
        start_time = time.time()

        # Save uploaded file to a temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            temp_file.write(validation_file.read())
            temp_file_path = temp_file.name

        # Open the workbook
        wb = xw.Book(temp_file_path)
        sheet = wb.sheets[0]

        # Read data into DataFrame
        data = sheet.used_range.value
        df = pd.DataFrame(data[1:], columns=data[0])

        # Identify missing values
        missing_values = df.isnull().sum().sum()
        invalid_rows = df[df.isnull().any(axis=1)]

        # Highlight error cells
        red_fill = xw.utils.rgb_to_int((255, 0, 0))
        for index, row in invalid_rows.iterrows():
            for col_name, value in row.items():
                if pd.isnull(value):
                    col_index = df.columns.get_loc(col_name) + 1
                    row_index = index + 2
                    sheet.range(row_index, col_index).color = red_fill

        # Create validation report
        report_wb = Workbook()
        report_ws = report_wb.active
        report_ws.title = "Validation Report"

        report_ws.append(["Validation Summary"])
        report_ws.append(["File Name", "Uploaded File"])
        report_ws.append(["Total Rows", df.shape[0]])
        report_ws.append(["Missing Values", missing_values])
        report_ws.append(["Invalid Rows", len(invalid_rows)])
        report_ws.append([])

        if not invalid_rows.empty:
            report_ws.append(["Invalid Rows Data"])
            for row in invalid_rows.itertuples(index=False, name=None):
                report_ws.append(row)

        # Save the report to BytesIO
        report_wb.save(output)
        output.seek(0)
        report_bytes = output.read()

        wb.save(temp_file_path)
        wb.close()

        end_time = time.time()
        execution_time = round(end_time - start_time, 2)

        file_size = len(report_bytes)
        logging.info(f"Report file size: {file_size} bytes")

        return report_bytes  # Return report bytes

    except Exception as e:
        error_message = f"Error in validation automation: {str(e)}\n{traceback.format_exc()}"
        logging.exception(error_message)
        # Create an excel file with the error message.
        error_df = pd.DataFrame([error_message], columns=['Error Message'])
        error_output = io.BytesIO()
        with pd.ExcelWriter(error_output, engine='openpyxl') as writer:
            error_df.to_excel(writer, index=False, sheet_name='Error')
        error_output.seek(0)
        return error_output.read()