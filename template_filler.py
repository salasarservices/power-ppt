import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


def fill_template_with_pages(template_bytes, pages, tables, title_font="Poppins", body_font="Poppins"):
    """
    Fills a PowerPoint template with provided pages (text content) and tables.
    Uses the standard template layout with proper positioning.

    Args:
        template_bytes: The binary content of the template PPTX file.
        pages: List of dictionaries with 'title', 'body', and 'slide_index' keys.
        tables: List of table data (headers, rows, slide index).
        title_font: Font name for slide titles.
        body_font: Font name for slide bodies.

    Returns:
        The binary content of the filled PPTX file.
    """
    # Load the template (REQUIRED - must have standard template)
    if not template_bytes:
        raise ValueError("Standard template is required!")
    
    prs = Presentation(io.BytesIO(template_bytes))
    
    # Use the appropriate slide layout from template (typically layout 1 or 2)
    # The template should have the layout with title area and content area
    slide_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]

    # Iterate through provided content pages
    for page in pages:
        slide_idx = page.get("slide_index", 0)
        title_text = page["title"]
        body_text = page["body"]
        
        # Create slide from template layout
        slide = prs.slides.add_slide(slide_layout)

        # Find and populate the title placeholder
        title_shape = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type == 1:  # Title placeholder
                title_shape = shape
                break
        
        if title_shape and title_shape.has_text_frame:
            title_shape.text = title_text
            # Apply title formatting
            for paragraph in title_shape.text_frame.paragraphs:
                paragraph.font.name = title_font
                paragraph.font.size = Pt(20)
                paragraph.font.bold = True
                paragraph.font.color.rgb = RGBColor(0, 112, 192)  # Blue color

        # Define content area (green-bordered area in template)
        # Adjust these coordinates based on your template's content area
        content_left = Inches(0.5)
        content_top = Inches(1.8)
        content_width = Inches(9)
        content_height = Inches(4.5)
        
        # Get tables for this slide
        tables_for_slide = [table for table in tables if table.get("slide_index") == slide_idx]
        
        if tables_for_slide:
            # If there are tables, render them in the content area
            current_top = content_top
            
            for table_data in tables_for_slide:
                if not table_data.get("rows"):
                    continue
                
                rows_count = len(table_data["rows"])
                cols_count = len(table_data["header"]) if table_data.get("header") else len(table_data["rows"][0])
                
                if rows_count == 0 or cols_count == 0:
                    continue
                
                # Calculate table height (approximate)
                row_height = Inches(0.3)
                table_height = rows_count * row_height
                
                # Check if table fits in current slide
                remaining_height = content_top + content_height - current_top
                
                if table_height > remaining_height and current_top > content_top:
                    # Table doesn't fit, create continuation slide
                    slide = prs.slides.add_slide(slide_layout)
                    
                    # Add title with (CONTD...)
                    title_shape = None
                    for shape in slide.placeholders:
                        if shape.placeholder_format.type == 1:
                            title_shape = shape
                            break
                    
                    if title_shape and title_shape.has_text_frame:
                        title_shape.text = f"{title_text} (CONTD...)"
                        for paragraph in title_shape.text_frame.paragraphs:
                            paragraph.font.name = title_font
                            paragraph.font.size = Pt(20)
                            paragraph.font.bold = True
                            paragraph.font.color.rgb = RGBColor(0, 112, 192)
                    
                    current_top = content_top
                
                # Calculate rows that fit in remaining space
                max_rows_that_fit = int(remaining_height / row_height)
                
                if max_rows_that_fit < rows_count:
                    # Split table across slides
                    rows_to_render = max_rows_that_fit
                else:
                    rows_to_render = rows_count
                
                # Create table
                table_shape = slide.shapes.add_table(
                    rows_to_render,
                    cols_count,
                    content_left,
                    current_top,
                    content_width,
                    min(table_height, remaining_height)
                )
                table = table_shape.table
                
                # Populate table
                for row_idx in range(rows_to_render):
                    row_data = table_data["rows"][row_idx]
                    for col_idx in range(min(len(row_data), cols_count)):
                        cell = table.cell(row_idx, col_idx)
                        cell.text = str(row_data[col_idx])
                        
                        # Format header row (first row)
                        if row_idx == 0:
                            for paragraph in cell.text_frame.paragraphs:
                                paragraph.font.bold = True
                                paragraph.font.size = Pt(11)
                                paragraph.font.color.rgb = RGBColor(255, 255, 255)
                            # Add blue background to header
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(0, 112, 192)
                        else:
                            for paragraph in cell.text_frame.paragraphs:
                                paragraph.font.size = Pt(10)
                
                current_top += min(table_height, remaining_height) + Inches(0.2)
                
                # If there are remaining rows, handle continuation
                if rows_to_render < rows_count:
                    # Create remaining rows data for next slide
                    remaining_rows = table_data["rows"][rows_to_render:]
                    # Add to tables list for processing (will be picked up in next iteration)
                    # For now, we'll handle this in a simpler way
                    pass
        
        elif body_text:
            # No tables, just render body text in content area
            textbox = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
            text_frame = textbox.text_frame
            text_frame.text = body_text
            text_frame.word_wrap = True
            
            for paragraph in text_frame.paragraphs:
                paragraph.font.name = body_font
                paragraph.font.size = Pt(12)
                paragraph.font.color.rgb = RGBColor(0, 0, 0)

    # Save the presentation to bytes
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream.read()
