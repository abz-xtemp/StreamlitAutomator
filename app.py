# streamlit_app.py
import streamlit as st
import os
import tempfile
import pandas as pd
from pathlib import Path
from automation_scripts import *

# Configure the page
st.set_page_config(page_title="Automation Hub", layout="wide")


def save_uploaded_file(uploaded_file):
    """Save uploaded file to a temporary directory and return the file path"""
    if uploaded_file is not None:
        temp_dir = tempfile.mkdtemp()
        file_path = os.path.join(temp_dir, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None


def provide_download_button(
    file_output, label="Download Processed File", file_name="processed_file.xlsx"
):
    """Create a download button for the processed file, now handles bytes or filepath"""
    if file_output:
        if isinstance(file_output, bytes):
            st.download_button(
                label=label,
                data=file_output,
                file_name=file_name,
                mime="application/octet-stream",
            )
        elif os.path.exists(file_output):
            with open(file_output, "rb") as f:
                st.download_button(
                    label=label,
                    data=f,
                    file_name=os.path.basename(file_output),
                    mime="application/octet-stream",
                )
        else:  # If file_output is string but not filepath
            st.error(file_output)
    else:
        st.error("No output file generated.")


def display_error_for_missing_inputs(required_inputs):
    """Display error if any required inputs are missing"""
    missing = [name for name, value in required_inputs.items() if not value]
    if missing:
        st.error(f"Missing required input(s): {', '.join(missing)}")
        return True
    return False


def main():
    st.title("🔄 Automation Hub")
    st.subheader("Run automation tasks on your Excel and PowerPoint files")

    # Create a sidebar for task selection
    with st.sidebar:
        st.header("Task Selection")
        automation_task = st.selectbox(
            "Select an automation task",
            [
                "Day Movement",
                "Excel to PPT",
                "PPT to PDF",
                "Update PPT",
                "Merge PPT",
                "Power Query",
                "Validation",
                "Consolidation",
                "Roll Over",
                "Staging",
                "Trend Check",
            ],
        )

    # Main area for task-specific inputs
    with st.container():
        st.write(f"## {automation_task}")

        # Initialize output_file to None
        output_file = None

        # Different input fields based on selected task
        if automation_task == "Day Movement":
            file1 = st.file_uploader(
                "Upload Excel File 1", type=["xls", "xlsx"], key="dm_file1"
            )
            file2 = st.file_uploader(
                "Upload Excel File 2", type=["xls", "xlsx"], key="dm_file2"
            )
            sheet_name = st.text_input("Enter Sheet Name", key="dm_sheet")
            cell_range = st.text_input(
                "Enter Cell Range(ex., A1: B10)", key="dm_range"
            )

            required = {
                "Excel File 1": file1,
                "Excel File 2": file2,
                "Sheet Name": sheet_name,
                "Cell Range": cell_range,
            }

            if st.button("Run Day Movement") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Processing Day Movement..."):
                    file1_path = save_uploaded_file(file1)
                    file2_path = save_uploaded_file(file2)
                    output_file = day_movement(
                        file1_path, file2_path, sheet_name, cell_range
                    )
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Day Movement Completed Successfully!")

        elif automation_task == "Excel to PPT":
            ppt_file = st.file_uploader(
                "Upload PPT File (Optional)", type=["pptx"], key="ep_ppt"
            )
            excel_file = st.file_uploader(
                "Upload Excel File", type=["xls", "xlsx"], key="ep_excel"
            )
            sheet_name = st.text_input("Enter Sheet Name", key="ep_sheet")
            cell_range = st.text_input("Enter Cell Range", key="ep_range")

            col1, col2 = st.columns(2)
            with col1:
                slide_number = st.number_input(
                    "Slide Number", min_value=1, value=1, key="ep_slide"
                )
                height = st.number_input(
                    "Slide Height", min_value=1, value=400, key="ep_height"
                )
                width = st.number_input(
                    "Slide Width", min_value=1, value=600, key="ep_width"
                )

            with col2:
                left = st.number_input("Image Left Position", value=50, key="ep_left")
                top = st.number_input("Image Top Position", value=50, key="ep_top")
                password = st.text_input(
                    "Excel Password (Optional)", type="password", key="ep_pass"
                )

            required = {
                "Excel File": excel_file,
                "Sheet Name": sheet_name,
                "Cell Range": cell_range,
            }

            if st.button("Run Excel to PPT") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Converting Excel to PPT..."):
                    ppt_path = save_uploaded_file(ppt_file) if ppt_file else None
                    excel_path = save_uploaded_file(excel_file)
                    output_file = excel_to_ppt(
                        ppt_path,
                        excel_path,
                        sheet_name,
                        cell_range,
                        slide_number,
                        height,
                        width,
                        left,
                        top,
                        password,
                    )
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Excel to PPT Conversion Completed!")

        elif automation_task == "PPT to PDF":
            ppt_file = st.file_uploader("Upload PPT File", type=["pptx"], key="p2p_ppt")

            required = {"PPT File": ppt_file}

            if st.button("Convert to PDF") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Converting PPT to PDF..."):
                    ppt_path = save_uploaded_file(ppt_file)
                    output_file = ppt_to_pdf(ppt_path)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("PPT to PDF Conversion Completed!")

        elif automation_task == "Update PPT":
            ppt_file = st.file_uploader("Upload PPT File", type=["pptx"], key="up_ppt")
            slides = st.text_input(
                "Enter Slides to Update (Comma Separated)", key="up_slides"
            )
            new_order = st.text_input(
                "Enter New Slide Order (Comma Separated)", key="up_order"
            )

            required = {
                "PPT File": ppt_file,
                "Slides to Update": slides,
                "New Slide Order": new_order,
            }

            if st.button("Update PPT") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Updating PPT..."):
                    ppt_path = save_uploaded_file(ppt_file)
                    output_file = update_ppt(ppt_path, slides, new_order)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("PPT Updated Successfully!")

        elif automation_task == "Merge PPT":
            ppt_a = st.file_uploader("Upload First PPT", type=["pptx"], key="mp_ppt_a")
            ppt_b = st.file_uploader("Upload Second PPT", type=["pptx"], key="mp_ppt_b")

            col1, col2 = st.columns(2)
            with col1:
                slide_index = st.number_input(
                    "Slide Number from PPT A", min_value=1, value=1, key="mp_slide"
                )
            with col2:
                merge_index = st.number_input(
                    "Merge at Index in PPT B", min_value=1, value=1, key="mp_merge"
                )

            required = {"First PPT": ppt_a, "Second PPT": ppt_b}

            if st.button("Merge PPTs") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Merging PPTs..."):
                    ppt_a_path = save_uploaded_file(ppt_a)
                    ppt_b_path = save_uploaded_file(ppt_b)
                    output_file = merge_ppt(
                        ppt_a_path, ppt_b_path, slide_index, merge_index
                    )
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("PPTs Merged Successfully!")

        elif automation_task == "Power Query":
            source_file = st.file_uploader(
                "Upload Source File", type=["xls", "xlsx"], key="pq_source"
            )
            stripped_data = st.file_uploader(
                "Upload Stripped Data", type=["xls", "xlsx"], key="pq_stripped"
            )
            setup_file = st.file_uploader(
                "Upload Setup File", type=["xls", "xlsx"], key="pq_setup"
            )

            required = {
                "Source File": source_file,
                "Stripped Data": stripped_data,
                "Setup File": setup_file,
            }

            if st.button("Run Power Query") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Running Power Query..."):
                    source_path = save_uploaded_file(source_file)
                    stripped_path = save_uploaded_file(stripped_data)
                    setup_path = save_uploaded_file(setup_file)
                    output_file = power_query(source_path, stripped_path, setup_path)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Power Query Completed!")

        elif automation_task == "Validation":
            validation_file = st.file_uploader(
                "Upload Validation File", type=["xls", "xlsx"], key="val_file"
            )

            required = {"Validation File": validation_file}

            if st.button("Run Validation") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Running Validation..."):
                    validation_path = save_uploaded_file(validation_file)
                    output_file = validation(validation_path)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Validation Completed!")

        elif automation_task == "Consolidation":
            excel_file = st.file_uploader(
                "Upload Excel File", type=["xls", "xlsx"], key="con_excel"
            )
            template_file = st.file_uploader(
                "Upload Template File", type=["xls", "xlsx"], key="con_template"
            )
            sheet_names = st.text_input(
                "Enter Sheet Names (Comma Separated)", key="con_sheets"
            )

            required = {
                "Excel File": excel_file,
                "Template File": template_file,
                "Sheet Names": sheet_names,
            }

            if st.button("Run Consolidation") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Running Consolidation..."):
                    excel_path = save_uploaded_file(excel_file)
                    template_path = save_uploaded_file(template_file)
                    output_file = consolidation(excel_path, template_path, sheet_names)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Consolidation Completed!")

        elif automation_task == "Roll Over":
            input_file = st.file_uploader(
                "Upload Input File", type=["xls", "xlsx"], key="ro_file"
            )

            required = {"Input File": input_file}

            if st.button("Run Roll Over") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Running Roll Over..."):
                    input_path = save_uploaded_file(input_file)
                    output_file = roll_over(input_path)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Roll Over Completed!")

        elif automation_task == "Staging":
            input_file = st.file_uploader(
                "Upload Input File", type=["xls", "xlsx"], key="stg_file"
            )

            required = {"Input File": input_file}

            if st.button("Run Staging") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Running Staging..."):
                    input_path = save_uploaded_file(input_file)
                    output_file = staging(input_path)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Staging Completed!")

        elif automation_task == "Trend Check":
            input_file = st.file_uploader(
                "Upload Input File", type=["xls", "xlsx"], key="tc_file"
            )

            required = {"Input File": input_file}

            if st.button("Run Trend Check") and not display_error_for_missing_inputs(
                required
            ):
                with st.spinner("Running Trend Check..."):
                    input_path = save_uploaded_file(input_file)
                    output_file = trend_check(input_path)
                    if (
                        not isinstance(output_file, str)
                        or os.path.exists(output_file)
                        or isinstance(output_file, bytes)
                    ):
                        st.success("Trend Check Completed!")

        # Provide download button if there's an output file
        if output_file:
            if output_file:
                st.subheader("Output")
                print(f"Output file type: {type(output_file)}") #Add this
                print(f"Output file value: {output_file}") #Add this
                if automation_task in ["Day Movement", "Power Query", "Validation", "Consolidation", "Roll Over", "Staging", "Trend Check"]:
                    provide_download_button(output_file, file_name="output.xlsx")
                elif automation_task == "Excel to PPT":
                    provide_download_button(output_file, file_name="output.pptx")
                elif automation_task == "PPT to PDF":
                    provide_download_button(output_file, file_name="output.pdf")
                elif automation_task in ["Update PPT", "Merge PPT"]:
                    provide_download_button(output_file, file_name="output.pptx")


if __name__ == "__main__":
    main()
