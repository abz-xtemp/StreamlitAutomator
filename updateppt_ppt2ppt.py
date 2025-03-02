import win32com.client
import pythoncom
import tempfile
import os
import time

def process(ppt_file, slides_to_update, new_order):
    """
    Updates the order of specific slides in a PowerPoint file.
    
    Parameters:
        ppt_file (BytesIO): Uploaded PowerPoint file.
        slides_to_update (str): Comma-separated string of slide numbers to update.
        new_order (str): Comma-separated string representing the new order of slides.
    
    Returns:
        str: Path to the updated PowerPoint file or error message.
    """
    # Initialize PowerPoint application and presentation variables
    pptApp = None
    presentation = None
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Parse inputs
        try:
            slides_to_update = [int(x.strip()) for x in slides_to_update.split(',')]
            new_order_indices = [int(x.strip()) for x in new_order.split(',')]
            
            # Validate that both lists have the same length
            if len(slides_to_update) != len(new_order_indices):
                raise ValueError("The number of slides to update and new positions must match.")
        except ValueError as e:
            return f"Error parsing input: {str(e)}"
        
        # Save uploaded PPT as a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_ppt:
            temp_ppt.write(ppt_file.read())
            temp_ppt_path = temp_ppt.name
            print(f"Temporary PPT saved at: {temp_ppt_path}")  # Debugging line

        # Create PowerPoint Application
        pptApp = win32com.client.Dispatch("PowerPoint.Application")
        pptApp.Visible = True  # Run PowerPoint in foreground for better stability
        
        # Allow PowerPoint to fully initialize
        time.sleep(1)

        # Open the PPT - explicitly setting ReadOnly to False to ensure we can modify it
        presentation = pptApp.Presentations.Open(temp_ppt_path, ReadOnly=False, WithWindow=True)
        time.sleep(0.5)  # Give PowerPoint time to fully load the presentation
        
        slides = presentation.Slides
        num_slides = slides.Count
        
        print(f"Presentation loaded with {num_slides} slides")  # Debugging line
        
        # Validate slide indices
        if max(slides_to_update) > num_slides or min(slides_to_update) < 1:
            return f"Error: Slide indices out of range. Presentation has {num_slides} slides."
        
        if max(new_order_indices) > num_slides or min(new_order_indices) < 1:
            return f"Error: New order indices out of range. Presentation has {num_slides} slides."
        
        # METHOD 1: Using MoveTo approach (simpler and more reliable)
        # Create a copy of the presentation to avoid reference issues
        temp_save_path = temp_ppt_path.replace(".pptx", "_temp.pptx")
        presentation.SaveAs(temp_save_path)
        presentation.Close()
        time.sleep(0.5)
        
        # Reopen the presentation to ensure fresh object references
        presentation = pptApp.Presentations.Open(temp_save_path, ReadOnly=False, WithWindow=True)
        slides = presentation.Slides
        
        # Check if we have a valid window reference
        if presentation.Windows.Count > 0:
            presentation.Windows(1).Activate()
        
        # Process slide moves one at a time
        # Sort by source position (reverse) if moving backward to avoid index conflicts
        move_list = list(zip(slides_to_update, new_order_indices))
        
        # Different sorting strategies depending on direction of movement
        moving_forward = [pair for pair in move_list if pair[0] < pair[1]]
        moving_backward = [pair for pair in move_list if pair[0] > pair[1]]
        staying_same = [pair for pair in move_list if pair[0] == pair[1]]
        
        # Process slides that are moving backward first (sorted by source in descending order)
        for old_idx, new_idx in sorted(moving_backward, key=lambda x: x[0], reverse=True):
            print(f"Moving slide {old_idx} to position {new_idx}")
            try:
                # Check if slide exists before attempting to move it
                slide = slides(old_idx)
                slide.MoveTo(new_idx)
                time.sleep(0.2)  # Add small delay between operations
            except Exception as e:
                print(f"Error moving slide {old_idx} to {new_idx}: {str(e)}")
                raise
        
        # Then process slides that are moving forward (sorted by source in ascending order)
        for old_idx, new_idx in sorted(moving_forward, key=lambda x: x[0]):
            print(f"Moving slide {old_idx} to position {new_idx}")
            try:
                # Due to previous movements, we need to recalculate the current position
                # of the slide that was originally at old_idx
                slide_to_move = None
                for i in range(1, slides.Count + 1):
                    if i not in [pair[1] for pair in moving_backward]:
                        if i == old_idx:
                            slide_to_move = i
                            break
                
                if slide_to_move is not None:
                    slides(slide_to_move).MoveTo(new_idx)
                    time.sleep(0.2)  # Add small delay between operations
            except Exception as e:
                print(f"Error moving slide {old_idx} to {new_idx}: {str(e)}")
                raise
        
        # Save the updated file
        updated_ppt_path = temp_ppt_path.replace(".pptx", "_updated.pptx")
        presentation.SaveAs(updated_ppt_path)
        time.sleep(0.5)  # Give time to save
        
        # Close and clean up
        presentation.Close()
        pptApp.Quit()
        
        # Clean up temporary files
        try:
            os.remove(temp_save_path)
        except:
            pass
            
        # Always clean up COM resources
        pythoncom.CoUninitialize()
        
        # Return the path to the updated PowerPoint file
        return updated_ppt_path

    except Exception as e:
        # Log the detailed error
        error_msg = f"Error: {str(e)}"
        print(error_msg)
        
        # Clean up resources
        try:
            if presentation is not None:
                presentation.Close()
        except:
            pass
            
        try:
            if pptApp is not None:
                pptApp.Quit()
        except:
            pass
            
        # Make sure to uninitialize COM even if there's an error
        try:
            pythoncom.CoUninitialize()
        except:
            pass
            
        return error_msg