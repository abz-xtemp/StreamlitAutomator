import win32com.client as win32
import shutil
import tempfile
import logging
import io
import traceback

def process(template_file):
    """
    Automates the roll-over process for Excel files.
    Returns bytes of a status message, or an error message within bytes.
    """
    output = io.BytesIO()  # Initialize output as BytesIO

    try:
        # Save uploaded file to a temporary path
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_xlsx:
            temp_xlsx.write(template_file.read())
            template_path = temp_xlsx.name

        xl = win32.Dispatch("Excel.Application")
        xl.Application.Calculation = -4135  # Set Calculation mode to Manual
        xl.DisplayAlerts = False
        xl.Visible = False

        wb1 = xl.Workbooks.Open(template_path, False, False, None)
        wb1.Password = "*"

        # Define Sheets
        sheet1 = wb1.Sheets("Roll_Over")
        row_count = sheet1.UsedRange.Rows.Count

        for i in range(2, row_count + 1):
            try:
                source = sheet1.Cells(i, 2).Value
                file_name = sheet1.Cells(i, 3).Value
                destination = sheet1.Cells(i, 4).Value
                dest_file_name = sheet1.Cells(i, 5).Value

                if not source:
                    continue

                # Copy file from source to destination
                shutil.copyfile(f"{source}\\{file_name}", f"{destination}\\{dest_file_name}")
                logging.info(f"File rolled over to: {destination}\\{dest_file_name}")

            except Exception as e:
                logging.error(f"Error processing row {i}: {str(e)}")
                continue

        wb1.Close(SaveChanges=False)
        xl.Quit()

        message = "Roll Over Process Completed Successfully."
        output.write(message.encode('utf-8'))
        return output.getvalue()  # Return bytes of success message

    except Exception as e:
        error_message = f"Error in Roll Over Automation: {str(e)}\n{traceback.format_exc()}"
        logging.exception(error_message)
        output.write(error_message.encode('utf-8'))
        return output.getvalue() # Return bytes of error message