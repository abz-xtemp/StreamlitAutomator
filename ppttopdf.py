import win32com.client
import pythoncom  # This import is crucial
import tempfile
import os

def process(ppt_file):
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Save uploaded PPT as a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_ppt:
            temp_ppt.write(ppt_file.read())
            temp_ppt_path = temp_ppt.name
            print(f"Temporary PPT saved at: {temp_ppt_path}")  # Debugging line

        # Create PowerPoint Application
        pptApp = win32com.client.Dispatch("PowerPoint.Application")
        pptApp.Visible = 1  # Run PowerPoint in foreground (for debugging)

        # Open the PPT
        presentation = pptApp.Presentations.Open(temp_ppt_path, WithWindow=False)

        # Generate temporary PDF path
        temp_pdf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
        print(f"Temporary PDF will be saved at: {temp_pdf_path}")  # Debugging line

        # Save as PDF
        presentation.SaveAs(temp_pdf_path, 32)  # 32 -> PDF format
        presentation.Close()
        pptApp.Quit()
        
        # Always clean up COM resources
        pythoncom.CoUninitialize()

        # Ensure PDF exists
        if os.path.exists(temp_pdf_path):
            return temp_pdf_path
        else:
            return "Error: PDF conversion failed."

    except Exception as e:
        # Make sure to uninitialize COM even if there's an error
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return f"Error: {str(e)}"





# import win32com.client
# import tempfile
# import os

# def process(ppt_file):
#     try:
#         # Save uploaded PPT as a temporary file
#         with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_ppt:
#             temp_ppt.write(ppt_file.read())
#             temp_ppt_path = temp_ppt.name
#             print(f"Temporary PPT saved at: {temp_ppt_path}")  # Debugging line

#         # Create PowerPoint Application
#         pptApp = win32com.client.Dispatch("PowerPoint.Application")
#         pptApp.Visible = 1  # Run PowerPoint in foreground (for debugging)

#         # Open the PPT
#         presentation = pptApp.Presentations.Open(temp_ppt_path, WithWindow=False)

#         # Generate temporary PDF path
#         temp_pdf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
#         print(f"Temporary PDF will be saved at: {temp_pdf_path}")  # Debugging line

#         # Save as PDF
#         presentation.SaveAs(temp_pdf_path, 32)  # 32 -> PDF format
#         presentation.Close()
#         pptApp.Quit()

#         # Ensure PDF exists
#         if os.path.exists(temp_pdf_path):
#             return temp_pdf_path
#         else:
#             return "Error: PDF conversion failed."

#     except Exception as e:
#         return f"Error: {str(e)}"








