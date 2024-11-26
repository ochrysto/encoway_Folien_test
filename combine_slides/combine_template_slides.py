from pptx import Presentation
from combine_slides.Generators import title_template_slide, table_template_slide, slide_9_template_slide, \
    end_template_slide, chart_template_slide
import json

data_path = "../Data/data.json"
# open Data JSON
with open(data_path, 'r') as data_objects:
    data = json.load(data_objects)

# load template
presentation = Presentation(data["template"])
# create slides
for slide_data in data["slides"]:
    slide_format = slide_data["slide_format"]
    if slide_format == "title":
        title_template_slide.create_title_slide(presentation, slide_data)
    elif slide_format == "introduction":
        slide_9_template_slide.create_slide_9(presentation, slide_data)
    elif slide_format == "table":
        table_template_slide.create_table_slide(presentation, slide_data)
    elif slide_format == "chart":
        chart_template_slide.create_chart_slide(presentation, slide_data)
    elif slide_format == "end":
        end_template_slide.create_end_slde(presentation, slide_data)

# save presentation
presentation.save(f"../Outcome/{data["file_name"]}.pptx")
