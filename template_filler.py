import io
from pptx import Presentation
from pptx.util import Inches, Pt


def fill_template_with_pages(template_bytes, pages, tables, title_font="Arial", body_font="Arial"):
    """
    Fills a PowerPoint template with provided pages (text content) and tables.
    Adjusts content placement based on the template's structure.

    Args:
        template_bytes: The binary content of the template PPTX file.
        pages: List of dictionaries with 'title', 'body', and 'slide_index' keys.
        tables: List of table data (headers, rows, slide index).
        title_font: Font name for slide titles.
        body_font: Font name for slide bodies.

    Returns:
        The binary content of the filled PPTX file.
    """
    # Load the template
    if template_bytes:
        prs = Presentation(io.BytesIO(template_bytes))
    else:
        prs = Presentation()  # Create a blank presentation if no template is provided

    # Iterate through provided content pages
    for page in pages:
        slide_idx = page.get("slide_index", 0)
        
        # Use the first slide layout (typically title + content)
        slide_layout = prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)

        # Add title
        title_shape = None
        for shape in slide.shapes:
            if shape.is_placeholder:
                phf = shape.placeholder_format
                if phf.type == 1:  # Title placeholder
                    title_shape = shape
                    break
        
        if title_shape and title_shape.has_text_frame:
            title_shape.text = page["title"]
            title_shape.text_frame.paragraphs[0].font.name = title_font
            title_shape.text_frame.paragraphs[0].font.size = Pt(24)
        else:
            # If no title placeholder, create a text box for the title
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
            title_frame = title_box.text_frame
            title_frame.text = page["title"]
            title_frame.paragraphs[0].font.name = title_font
            title_frame.paragraphs[0].font.size = Pt(24)
            title_frame.paragraphs[0].font.bold = True

        # Add body content
        body_top = Inches(1.5)
        body_left = Inches(0.5)
        body_width = Inches(9)
        body_height = Inches(4)
        
        if page["body"]:
            textbox = slide.shapes.add_textbox(body_left, body_top, body_width, body_height)
            text_frame = textbox.text_frame
            text_frame.text = page["body"]
            text_frame.word_wrap = True
            for paragraph in text_frame.paragraphs:
                paragraph.font.name = body_font
                paragraph.font.size = Pt(14)

        # Handle rendering tables specific to the slide
        tables_for_slide = [table for table in tables if table.get("slide_index") == slide_idx]
        
        table_top = Inches(6) if page["body"] else body_top
        
        for table_data in tables_for_slide:
            if not table_data.get("rows"):
                continue
                
            rows_count = len(table_data["rows"])
            cols_count = len(table_data["header"]) if table_data.get("header") else 0
            
            # Skip if invalid table dimensions
            if rows_count == 0 or cols_count == 0:
                continue

            # Create the table with proper dimensions
            table_shape = slide.shapes.add_table(
                rows_count, 
                cols_count, 
                Inches(0.5), 
                table_top, 
                Inches(9), 
                Inches(2)
            )
            table = table_shape.table

            # Populate all rows (including header as first row)
            for row_idx, row_data in enumerate(table_data["rows"]):
                # Ensure we don't exceed the number of columns
                for col_idx in range(min(len(row_data), cols_count)):
                    cell = table.cell(row_idx, col_idx)
                    cell.text = str(row_data[col_idx])
                    
                    # Format header row (first row) with bold text
                    if row_idx == 0:
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.font.bold = True
                            paragraph.font.size = Pt(12)
                    else:
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.font.size = Pt(11)
            
            # Move next table down
            table_top += Inches(2.5)

    # Save the presentation to bytes
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream.read()
