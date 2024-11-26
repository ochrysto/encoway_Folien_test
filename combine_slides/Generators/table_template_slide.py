from pptx.util import Inches


def create_table_slide(presentation, data):
    table_slide_layout = presentation.slide_layouts[16]
    table_slide = presentation.slides.add_slide(table_slide_layout)
    title = table_slide.placeholders[0]
    title.text = data["headline"]
    subtitle = table_slide.placeholders[36]
    subtitle.text = data["sub_header"]

    rows = len(data["table_content"]) + 1 # +1 for table head
    cols = len(data["table_content"][0])
    left = Inches(0.37)
    top = Inches(1.82)
    width = Inches(12.6)
    height = Inches(4.5)
    graphic_frame = table_slide.shapes.add_table(rows, cols, left, top, width, height)
    table = graphic_frame.table
    # table head
    key_names = list(data["table_content"][0].keys())
    for i, key in enumerate(key_names):
        table.cell(0, i).text = key
    # table body
    for i, entry in enumerate(data["table_content"], start=1):
        for j, key in enumerate(key_names):
            table.cell(i, j).text = entry[key]



