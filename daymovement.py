import pandas as pd
import re
import io

def process(file1, file2, sheet_name, cell_range):
    """
    Calculates the daily movement and returns it as Excel bytes.
    """
    try:
        df1 = pd.read_excel(file1, sheet_name=sheet_name, engine="openpyxl")
        df2 = pd.read_excel(file2, sheet_name=sheet_name, engine="openpyxl")

        start_col, start_row, end_col, end_row = parse_cell_range(cell_range)
        df1_values = df1.iloc[start_row:end_row, start_col:end_col].astype(float)
        df2_values = df2.iloc[start_row:end_row, start_col:end_col].astype(float)
        day_movement = df2_values - df1_values

        df = pd.DataFrame(day_movement)

        # Convert DataFrame to Excel bytes
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Day Movement')
        output.seek(0)
        return output.read()

    except Exception as e:
        print(f"Error: {e}")
        return None

def parse_cell_range(cell_range):
    """
    Converts an Excel cell range (e.g., "A1:K26") into row and column indices.

    :param cell_range: A string representing the cell range (e.g., "A1:K26").
    :return: A tuple (start_col_index, start_row_index, end_col_index, end_row_index)
    """
    # Match the cell range using regular expression
    match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", cell_range)
    if not match:
        raise ValueError("Invalid cell range format. Use format like 'A1:D10'.")

    # Extract start and end column and row information
    start_col, start_row, end_col, end_row = match.groups()

    # Convert column letters to indices (A -> 0, B -> 1, ..., Z -> 25, AA -> 26, etc.)
    def col_to_index(col):
        index = 0
        for char in col:
            index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1

    # Convert the row and column values from Excel to pandas indices
    start_col_index = col_to_index(start_col)
    start_row_index = int(start_row) - 1  # Excel rows are 1-indexed
    end_col_index = col_to_index(end_col) + 1  # End column should be exclusive
    end_row_index = int(end_row)  # End row is inclusive

    return start_col_index, start_row_index, end_col_index, end_row_index