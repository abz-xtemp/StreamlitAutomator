import xlwings as xw
import pandas as pd
from time import sleep
import logging

def process(template_file):
    """
    Automates Power Query operations in Excel.
    Streamlit-compatible version using uploaded file-like objects.

    Args:
        template_file (BytesIO): The uploaded Excel template file.

    Returns:
        str: Success message or error details.
    """
    try:
        logging.info("Opening template file...")

        # Open template file in-memory
        with open("temp_template.xlsx", "wb") as f:
            f.write(template_file.getvalue())

        temp_book = xw.Book("temp_template.xlsx")
        temp_sheet = temp_book.sheets['PW Query']

        # Read the sheet's used range
        data_feed = temp_sheet.used_range.value
        temp_book.close()

        # Convert to DataFrame
        input_df = pd.DataFrame(data_feed)
        input_df = input_df.set_axis(input_df.iloc[1], axis=1).iloc[2:]

        # Extract required file paths
        main_file_loc = input_df['Latest Source File'][0]
        stripped_data_loc = input_df['Stripped Data'][0]
        setup_pq_loc = input_df['Setup File'][0]

        target_sheets = ['Total Summary', 'MVOI']

        # Open main and stripped data files
        main_file = xw.Book(main_file_loc)
        stripped_data = xw.Book(stripped_data_loc)

        for sheet in target_sheets:
            curr_sheet = main_file.sheets[sheet]
            dest_sheet = stripped_data.sheets[sheet]
            dest_sheet.range("A1").value = curr_sheet.used_range.value

        # Save and close files
        main_file.close()
        stripped_data.save(stripped_data_loc)
        stripped_data.close()

        # Open Setup File and Refresh Queries
        setup_pq = xw.Book(setup_pq_loc)
        setup_pq.api.RefreshAll()
        sleep(30)  # Allow time for refresh
        setup_pq.save(setup_pq_loc)
        setup_pq.close()

        return "Power Query automation completed successfully!"

    except Exception as e:
        logging.error(f"Error in Power Query automation: {str(e)}")
        return f"Error: {str(e)}"
