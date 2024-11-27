
def create_end_slde(presentation, data):
    end_slide_layout = presentation.slide_layouts[19]
    end_slide = presentation.slides.add_slide(end_slide_layout)
    title = end_slide.placeholders[0]
    title.text = data["headline"]