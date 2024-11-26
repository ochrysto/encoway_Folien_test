
def create_title_slide(presentation, data):
   title_slide_layout = presentation.slide_layouts[0]
   title_slide = presentation.slides.add_slide(title_slide_layout)
   title = title_slide.shapes.title
   title.text = data["headline"]
   subtitle = title_slide.placeholders[10]
   subtitle.text = f"{data["name"]} | {data["date"]}"