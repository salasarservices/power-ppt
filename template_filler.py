import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


def fill_template_with_pages(template_bytes, pages, tables, title_font="Poppins", body_font="Poppins"):
    """
    Fills a PowerPoint template with provided pages and tables.
    """
    if not template_bytes:
        raise ValueError("Standard template is required!")
    
    prs = Presentation(io.BytesIO(template_bytes))
    slide_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]

    # Content area dimensions
    content_left = Inches(0.5)
    content_top = Inches(1.8)
    content_width = Inches(9)
    content_height = Inches(4.5)

    for page in pages:
        slide_idx = page.get("slide_index", 0)
        title_text = page["title"]
        body_text = page["body"]
        
        slide = prs.slides.add_slide(slide_layout)

        # Set title
        title_shape = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type == 1:
                title_shape = shape
                break
        
        if title_shape and title_shape.has_text_frame:
            title_shape.text = title_text
            for paragraph in title_shape.text_frame.paragraphs:
                paragraph.font.name = title_font
                paragraph.font.size = Pt(20)
                paragraph.font.bold = True
                paragraph.font.color.rgb = RGBColor(0, 112, 192)

        # Get tables for this slide
        tables_for_slide = [table for table in tables if table.get("slide_index") == slide_idx]
        
        if tables_for_slide:
            current_top = content_top
            
            for table_data in tables_for_slide:
                if not table_data.get("rows"):
                    continue
                
                rows_count = len(table_data["rows"])
                cols_count = len(table_data["rows"][0]) if table_data["rows"] else 0
                
                if rows_count == 0 or cols_count == 0:
                    continue
                
                table_shape = slide.shapes.add_table(
                    rows_count,
                    cols_count,
                    content_left,
                    current_top,
                    content_width,
                    Inches(0.3) * rows_count
                )
                table = table_shape.table
                
                for row_idx in range(rows_count):
                    row_data = table_data["rows"][row_idx]
                    for col_idx in range(min(len(row_data), cols_count)):
                        cell = table.cell(row_idx, col_idx)
                        cell.text = str(row_data[col_idx])
                        
                        if row_idx == 0:
                            for paragraph in cell.text_frame.paragraphs:
                                paragraph.font.bold = True
                                paragraph.font.size = Pt(11)
                                paragraph.font.color.rgb = RGBColor(255, 255, 255)
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(0, 112, 192)
                        else:
                            for paragraph in cell.text_frame.paragraphs:
                                paragraph.font.size = Pt(10)
                
                current_top += Inches(0.3) * rows_count + Inches(0.2)
        
        elif body_text:
            textbox = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
            text_frame = textbox.text_frame
            text_frame.text = body_text
            text_frame.word_wrap = True
            
            for paragraph in text_frame.paragraphs:
                paragraph.font.name = body_font
                paragraph.font.size = Pt(12)
                paragraph.font.color.rgb = RGBColor(0, 0, 0)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream.read()
