
def remove_empty_placeholder(presentation):
    for slide in presentation.slides:
        for placeholder in slide.placeholders:
            if not placeholder.has_text_frame or not placeholder.text_frame.text.strip():
                sp = placeholder.element
                sp.getparent().remove(sp)


def remove_unused_slides(presentation):
    slides_to_remove =[]
    for slide in presentation.slides:
        if not slide.shapes.title and not any(shape.has_text_frame and shape.text_frame.text.strip() for shape in slide.shapes):
            slides_to_remove.append(slide)

    for slide in slides_to_remove:
        slide_id = slide.slide_id
        presentation.part.drop_rel(slide_id)
        presentation.slides._sldIdLst.remove(slide._element)


def clean_presentation(presentation):
    remove_empty_placeholder(presentation)
    remove_unused_slides(presentation)
