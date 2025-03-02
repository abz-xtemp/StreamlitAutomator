import win32com.client
import pythoncom
import tempfile
import time
import os

def process(ppt_file_A, ppt_file_B, slide_to_merge, merge_position):
    """
    Copies a slide from presentation A and inserts it at a specific position in presentation B.
    
    Parameters:
        ppt_file_A (BytesIO): First PowerPoint file containing the slide to copy.
        ppt_file_B (BytesIO): Second PowerPoint file where the slide will be inserted.
        slide_to_merge (int): Slide number in presentation A to copy.
        merge_position (int): Position in presentation B to insert the copied slide.
    
    Returns:
        str: Path to the merged PowerPoint file or error message.
    """
    pptApp = None
    presentation_A = None
    presentation_B = None
    
    try:
        # Initialize COM - this is the missing step in the original code
        pythoncom.CoInitialize()
        
        # Save uploaded files as temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_A, \
             tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_B:

            temp_A.write(ppt_file_A.read())
            temp_B.write(ppt_file_B.read())

            temp_A_path = temp_A.name
            temp_B_path = temp_B.name
            
        print(f"Temp files created: {temp_A_path}, {temp_B_path}")

        # Create PowerPoint Application
        pptApp = win32com.client.Dispatch("PowerPoint.Application")
        pptApp.Visible = True  # Run in foreground for stability
        
        # Allow PowerPoint to fully initialize
        time.sleep(1)

        # Open presentations - explicitly set ReadOnly to False
        presentation_A = pptApp.Presentations.Open(temp_A_path, ReadOnly=False, WithWindow=True)
        time.sleep(0.5)  # Give time to fully load
        
        presentation_B = pptApp.Presentations.Open(temp_B_path, ReadOnly=False, WithWindow=True)
        time.sleep(0.5)  # Give time to fully load

        slides_A = presentation_A.Slides
        slides_B = presentation_B.Slides
        
        # Make sure one of the presentations is active
        presentation_A.Windows(1).Activate()
        time.sleep(0.2)

        # Validate slide and merge position
        if slide_to_merge < 1 or slide_to_merge > slides_A.Count:
            raise ValueError(f"Invalid slide number {slide_to_merge} in PPT A (which has {slides_A.Count} slides)!")
        
        if merge_position < 1 or merge_position > slides_B.Count + 1:
            raise ValueError(f"Invalid merge position {merge_position} in PPT B (which has {slides_B.Count} slides)!")

        # Copy the slide from A and paste it in B
        print(f"Copying slide {slide_to_merge} from presentation A")
        slides_A(slide_to_merge).Copy()
        time.sleep(0.5)  # Give time for copy operation to complete
        
        # Activate presentation B to ensure paste works correctly
        presentation_B.Windows(1).Activate()
        time.sleep(0.3)
        
        print(f"Pasting slide to position {merge_position} in presentation B")
        # Use the Paste method with Index parameter
        slides_B.Paste(Index=merge_position)
        time.sleep(0.5)  # Give time for paste operation to complete

        # Create output filename
        merged_ppt_file = os.path.join(os.path.dirname(temp_B_path), "merged_presentation.pptx")
        
        # Save the merged file
        print(f"Saving merged presentation to {merged_ppt_file}")
        presentation_B.SaveAs(merged_ppt_file)
        time.sleep(0.5)  # Give time for save operation to complete

        # Close presentations
        presentation_A.Close()
        presentation_B.Close()
        
        # Quit PowerPoint
        pptApp.Quit()
        
        # Clean up temporary files
        try:
            os.remove(temp_A_path)
            os.remove(temp_B_path)
        except:
            pass

        # Always clean up COM resources
        pythoncom.CoUninitialize()
        
        return merged_ppt_file

    except Exception as e:
        error_msg = f"Error: {str(e)}"
        print(error_msg)
        
        # Clean up resources
        try:
            if presentation_A is not None:
                presentation_A.Close()
        except:
            pass
            
        try:
            if presentation_B is not None:
                presentation_B.Close()
        except:
            pass
            
        try:
            if pptApp is not None:
                pptApp.Quit()
        except:
            pass
            
        # Make sure to uninitialize COM if it was initialized
        try:
            pythoncom.CoUninitialize()
        except:
            pass
            
        return error_msg