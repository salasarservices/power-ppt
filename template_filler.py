import io
from pptx import Presentation
from pptx.util import Inches, Pt


def fill_template_with_pages(template_bytes, pages, tables, title_font="Arial", body_font="Arial"):
    """
    Fills a PowerPoint template with provided pages (text content) and tables.
    Adjusts content placement based on the template's structure.

    Args:
        template_bytes: The binary content of the template PPTX file.
        pages: List of dictionaries with 'title' and 'body' keys.
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
    for i, page in enumerate(pages):
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Use blank slide layout

        # Add title
        title_box = slide.shapes.title
        if title_box:
            title_box.text = page["title"]
            title_box.text_frame.paragraphs[0].font.name = title_font
            title_box.text_frame.paragraphs[0].font.size = Pt(24)

        # Add body content
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4))
        text_frame = textbox.text_frame
        text_frame.text = page["body"]
        for paragraph in text_frame.paragraphs:
            paragraph.font.name = body_font
            paragraph.font.size = Pt(16)

        # Handle rendering tables specific to the slide
        tables_for_slide = [table for table in tables if table["slide_index"] == i]
        for table_data in tables_for_slide:
            rows = len(table_data["rows"])
            cols = len(table_data["header"])

            # Create the table
            table = slide.shapes.add_table(rows, cols, Inches(1), Inches(6), Inches(8), Inches(2)).table

            # Format table header
            for col_idx, header_text in enumerate(table_data["header"]):
                cell = table.cell(0, col_idx)
                cell.text = header_text
                cell.text_frame.paragraphs[0].font.bold = True

            # Populate row data
            for row_idx, row_data in enumerate(table_data["rows"], start=1):
                for col_idx, cell_text in enumerate(row_data):
                    table.cell(row_idx, col_idx).text = cell_text

    # Save the presentation to bytes
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream.read()
