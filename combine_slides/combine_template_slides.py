from pptx import Presentation
from combine_slides.Generators import title_slide, table_slide, info_slide, \
    end_slide, chart_slide
from Util import slide_cleaner
import json

data_path = "../Data/baseball_rules.json"
# open Data JSON
with open(data_path, 'r') as data_objects:
    data = json.load(data_objects)

# load template
presentation = Presentation(data["template"])
# create slides
for slide_data in data["slides"]:
    slide_format = slide_data["slide_format"]
    if slide_format == "title":
        title_slide.create_title_slide(presentation, slide_data)
    elif slide_format == "info":
        info_slide.create_slide_9(presentation, slide_data)
    elif slide_format == "table":
        table_slide.create_multi_slide_table(presentation, slide_data)
    elif slide_format == "chart":
        chart_slide.create_chart_slide(presentation, slide_data)
    elif slide_format == "end":
        end_slide.create_end_slde(presentation, slide_data)

slide_cleaner.remove_empty_placeholder(presentation)

# save presentation
presentation.save(f"../Outcome/{data["file_name"]}.pptx")
