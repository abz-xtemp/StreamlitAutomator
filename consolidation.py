import logging
import win32com.client as win32
import tempfile

def process(uploaded_file, password=""):
    """
    Consolidates data from multiple Excel files by copying specified sheets into a destination file.

    Parameters:
        uploaded_file: Streamlit file uploader object.
        password (str, optional): Password for opening protected Excel files.

    Returns:
        str: Success message or error details.
    """
    try:
        if not uploaded_file:
            return "No file uploaded."

        # Save the uploaded file temporarily
        temp_dir = tempfile.mkdtemp()
        file_path = f"{temp_dir}/{uploaded_file.name}"
        
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Initialize Excel application
        xl = win32.Dispatch("Excel.Application")
        xl.DisplayAlerts = False
        xl.Visible = False

        # Open the main consolidation template file
        wb_main = xl.Workbooks.Open(file_path, False, False, None, Password=password)
        sheet_main = wb_main.Sheets("Consolidation")
        row_count = sheet_main.UsedRange.Rows.Count

        for i in range(2, row_count + 1):
            try:
                # Read consolidation details
                source_file = sheet_main.Cells(i, 3).Value
                template_file = sheet_main.Cells(i, 5).Value
                tab_names = [sheet_main.Cells(i, j).Value for j in range(6, 11)]
                destination_file = sheet_main.Cells(i, 12).Value

                # Open source workbook
                wb_source = xl.Workbooks.Open(source_file, False, False, None, Password=password)

                # Copy sheets from source workbook
                for tab_name in filter(None, tab_names):  
                    try:
                        ws_source = wb_source.Worksheets(tab_name)
                        ws_source.Copy(After=wb_main.Sheets(wb_main.Sheets.Count))
                    except Exception as e:
                        logging.error(f"Error copying sheet '{tab_name}': {str(e)}")

                # Change links in the workbook to use the new template
                wb_main.ChangeLink(source_file, template_file, Type=1)

                # Save the consolidated workbook
                wb_main.SaveAs(destination_file)
                logging.info(f"File saved at {destination_file}")

            except Exception as e:
                logging.exception(f"Error processing row {i}: {str(e)}")
                continue

        wb_main.Close()
        xl.Quit()
        return "Consolidation process completed successfully."

    except Exception as e:
        logging.exception("Error in consolidation automation.")
        return f"Error: {str(e)}"
