import win32com.client as win32
import tempfile
import logging

def process(source_file, reporting_month, entity, template_file, key_sheet, other_sheets):
    """
    Automates the Staging Process for Excel files.

    Parameters:
        source_file (BytesIO): Uploaded Excel source file.
        reporting_month (str): Reporting month value.
        entity (str): Entity name.
        template_file (BytesIO): Uploaded template file.
        key_sheet (str): Name of the key sheet.
        other_sheets (str): Name of additional sheets.

    Returns:
        str: Status message indicating completion or errors.
    """
    try:
        # Save uploaded files to temporary paths
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_source:
            temp_source.write(source_file.read())
            source_path = temp_source.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_template:
            temp_template.write(template_file.read())
            template_path = temp_template.name

        # Initialize Excel Application
        xl = win32.Dispatch('Excel.Application')
        xl.Workbooks.Add()
        xl.Calculation = win32.constants.xlCalculationManual
        xl.DisplayAlerts = False
        xl.Visible = False  # Run in background

        # Open the user-provided template
        wb1 = xl.Workbooks.Open(source_path, False, False, None)
        wb1.Password = 'password'  # Change if required

        # Open Staging Sheet
        sheet1 = wb1.Sheets("Stagingfile")

        # Open control sheets and update values
        wb2 = xl.Workbooks.Open(template_path, False, False, None)
        sheet2 = wb2.Sheets("Control Sheet")
        sheet3 = wb2.Sheets(key_sheet)
        sheet4 = wb2.Sheets(other_sheets)

        sheet2.Cells(2, 2).Value = reporting_month
        sheet2.Cells(16, 2).Value = entity
        sheet2.EnableCalculation = True
        sheet2.Calculate()

        sheet3.EnableCalculation = True
        sheet3.Calculate()

        sheet4.EnableCalculation = True
        sheet4.Calculate()

        # Loop through staging file rows
        row_count = sheet1.UsedRange.Rows.Count
        for i in range(2, row_count + 1):
            try:
                source = sheet1.Cells(i, 2).Value
                file_name = sheet1.Cells(i, 3).Value
                reporting_date = sheet1.Cells(i, 4).Value
                entity_cut = sheet1.Cells(i, 5).Value
                template_file_path = sheet1.Cells(i, 6).Value
                template_file_name = sheet1.Cells(i, 7).Value
                key_sheet = sheet1.Cells(i, 8).Value
                tab1 = sheet1.Cells(i, 9).Value
                tab2 = sheet1.Cells(i, 10).Value
                tab3 = sheet1.Cells(i, 11).Value
                tab4 = sheet1.Cells(i, 12).Value
                tab5 = sheet1.Cells(i, 13).Value
                destination = sheet1.Cells(i, 14).Value
                destination_file = sheet1.Cells(i, 15).Value

                if source is None:
                    continue

                logging.info(f"Processing Row {i}: {file_name} -> {destination_file}")

                # Open source and template files
                wb_source = xl.Workbooks.Open(source, False, False, None, Password='password')
                wb_template = xl.Workbooks.Open(template_file_path, False, False, None, Password='password')

            except Exception as e:
                logging.error(f"Error processing row {i}: {str(e)}")

        # Close workbooks safely
        wb1.Close(SaveChanges=True)
        wb2.Close(SaveChanges=True)
        xl.Quit()

        return "Staging Automation Completed Successfully."

    except Exception as e:
        logging.exception(f"Staging Automation Error: {str(e)}")
        return f"Error in Staging Automation: {str(e)}"
