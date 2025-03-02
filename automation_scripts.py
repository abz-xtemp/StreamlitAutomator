# automation_scripts.py
# automation_scripts.py
# automation_scripts.py
import os
import tempfile
from daymovement import process as day_movement_process
from exceltoppt import process as excel_to_ppt_process
from ppttopdf import process as ppt_to_pdf_process
from updateppt_ppt2ppt import process as update_ppt_process
from mergeppt import process as merge_ppt_process
from powerquery import process as power_query_process
from validation import process as validation_process
from consolidation import process as consolidation_process
from rollover import process as roll_over_process
from staging import process as staging_process
from trendcheck import process as trend_check_process

def save_uploaded_file(uploaded_file):
    if uploaded_file is not None:
        temp_dir = tempfile.mkdtemp()
        file_path = os.path.join(temp_dir, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None

def day_movement(file1, file2, sheet_name, cell_range):
    result = day_movement_process(file1, file2, sheet_name, cell_range)
    if isinstance(result, bytes):
        temp_dir = tempfile.mkdtemp()
        output_file_path = os.path.join(temp_dir, "day_movement_output.xlsx")
        with open(output_file_path, "wb") as f:
            f.write(result)
        return output_file_path
    else:
        return result

def excel_to_ppt(ppt_file, excel_file, sheet_name, cell_range, slide_number, height, width, left, top, password):
    with open(excel_file, "rb") as excel_obj:
        if ppt_file:
            with open(ppt_file, "rb") as ppt_obj:
                return excel_to_ppt_process(ppt_obj, excel_obj, sheet_name, cell_range, slide_number, height, width, left, top, password)
        else:
            return excel_to_ppt_process(None, excel_obj, sheet_name, cell_range, slide_number, height, width, left, top, password)

def ppt_to_pdf(ppt_file_path):
    with open(ppt_file_path, 'rb') as file_object:
        return ppt_to_pdf_process(file_object)

def update_ppt(ppt_file_path, slides, new_order):
    with open(ppt_file_path, "rb") as ppt_obj:
        return update_ppt_process(ppt_obj, slides, new_order)

def merge_ppt(ppt_a_path, ppt_b_path, slide_index, merge_index):
    with open(ppt_a_path, "rb") as ppt_a_obj, open(ppt_b_path, "rb") as ppt_b_obj:
        return merge_ppt_process(ppt_a_obj, ppt_b_obj, slide_index, merge_index)

def power_query(source_file, stripped_data, setup_file):
    result = power_query_process(source_file, stripped_data, setup_file)
    if isinstance(result, bytes):
        temp_dir = tempfile.mkdtemp()
        output_file_path = os.path.join(temp_dir, "power_query_output.xlsx")
        with open(output_file_path, "wb") as f:
            f.write(result)
        return output_file_path
    else:
        return result

def validation(input_file):
    result = validation_process(input_file)
    if isinstance(result, bytes):
        temp_dir = tempfile.mkdtemp()
        output_file_path = os.path.join(temp_dir, "validation_output.xlsx")
        with open(output_file_path, "wb") as f:
            f.write(result)
        return output_file_path
    else:
        return result

def consolidation(excel_file, template_file, sheet_names):
    result = consolidation_process(excel_file, template_file, sheet_names)
    if isinstance(result, bytes):
        temp_dir = tempfile.mkdtemp()
        output_file_path = os.path.join(temp_dir, "consolidation_output.xlsx")
        with open(output_file_path, "wb") as f:
            f.write(result)
        return output_file_path
    else:
        return result

def roll_over(input_file):
    result = roll_over_process(input_file)
    if isinstance(result, bytes):
        temp_dir = tempfile.mkdtemp()
        output_file_path = os.path.join(temp_dir, "roll_over_output.xlsx")
        with open(output_file_path, "wb") as f:
            f.write(result)
        return output_file_path
    else:
        return result

def staging(input_file):
    result = staging_process(input_file)
    if isinstance(result, bytes):
        temp_dir = tempfile.mkdtemp()
        output_file_path = os.path.join(temp_dir, "staging_output.xlsx")
        with open(output_file_path, "wb") as f:
            f.write(result)
        return output_file_path
    else:
        return result

def trend_check(input_file):
    result = trend_check_process(input_file)
    if isinstance(result, bytes):
        temp_dir = tempfile.mkdtemp()
        output_file_path = os.path.join(temp_dir, "trend_check_output.xlsx")
        with open(output_file_path, "wb") as f:
            f.write(result)
        return output_file_path
    else:
        return result



# import os
# import tempfile
# from daymovement import process as day_movement_process
# from exceltoppt import process as excel_to_ppt_process
# from ppttopdf import process as ppt_to_pdf_process
# from updateppt_ppt2ppt import process as update_ppt_process
# from mergeppt import process as merge_ppt_process
# from powerquery import process as power_query_process
# from validation import process as validation_process
# from consolidation import process as consolidation_process
# from rollover import process as roll_over_process
# from staging import process as staging_process
# from trendcheck import process as trend_check_process

# def save_uploaded_file(uploaded_file):
#     if uploaded_file is not None:
#         temp_dir = tempfile.mkdtemp()
#         file_path = os.path.join(temp_dir, uploaded_file.name)
#         with open(file_path, "wb") as f:
#             f.write(uploaded_file.getbuffer())
#         return file_path
#     return None

# def day_movement(file1, file2, sheet_name, cell_range):
#     return day_movement_process(file1, file2, sheet_name, cell_range)

# def excel_to_ppt(ppt_file, excel_file, sheet_name, cell_range, slide_number, height, width, left, top, password):
#     # Since exceltoppt.py expects file objects, open the files before passing
#     with open(excel_file, "rb") as excel_obj:
#         if ppt_file:
#             with open(ppt_file, "rb") as ppt_obj:
#                 return excel_to_ppt_process(ppt_obj, excel_obj, sheet_name, cell_range, slide_number, height, width, left, top, password)
#         else:
#             return excel_to_ppt_process(None, excel_obj, sheet_name, cell_range, slide_number, height, width, left, top, password)

# def ppt_to_pdf(ppt_file_path):
#     # Open the file and pass the file object to ppt_to_pdf_process
#     with open(ppt_file_path, 'rb') as file_object:
#         return ppt_to_pdf_process(file_object)

# def update_ppt(ppt_file_path, slides, new_order):
#     # Open the file and pass the file object to update_ppt_process
#     with open(ppt_file_path, "rb") as ppt_obj:
#         return update_ppt_process(ppt_obj, slides, new_order)

# def merge_ppt(ppt_a_path, ppt_b_path, slide_index, merge_index):
#     # Open both files and pass the file objects to merge_ppt_process
#     with open(ppt_a_path, "rb") as ppt_a_obj, open(ppt_b_path, "rb") as ppt_b_obj:
#         return merge_ppt_process(ppt_a_obj, ppt_b_obj, slide_index, merge_index)

# def power_query(source_file, stripped_data, setup_file):
#     return power_query_process(source_file, stripped_data, setup_file)

# def validation(input_file):
#     return validation_process(input_file)

# def consolidation(excel_file, template_file, sheet_names):
#     return consolidation_process(excel_file, template_file, sheet_names)

# def roll_over(input_file):
#     return roll_over_process(input_file)

# def staging(input_file):
#     return staging_process(input_file)

# def trend_check(input_file):
#     return trend_check_process(input_file)