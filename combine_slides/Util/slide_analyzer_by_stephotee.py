from pptx import Presentation

from combine_slides.Analysers.slide_reader_v2 import shape

source1 = "../../Source/Baseball Rules!.pptx"
source2 = "../../Source/HiTech GmbH Vision.pptx"

def analyze_single_slide(filepath, layout_index):
    # Load your presentation template
    prs = Presentation(filepath)
    # Choose the slide layout index you're interested in
    # Slide layouts are indexed starting from 0

    # Get the layout
    slide_layout = prs.slide_layouts[layout_index]

    # Print placeholder details for the chosen layout
    print(f"Details for layout {layout_index}: {slide_layout.name}")
    for placeholder in slide_layout.placeholders:
        print(f"Placeholder index: {placeholder.placeholder_format.idx}, Type: {placeholder.placeholder_format.type}, "
              f"Name: '{placeholder.name}'")

def analyze_presentation(filepath):
    presentation = Presentation(filepath)
    for slide_index, slide in enumerate(presentation.slides):
        # layout_index = slide.slide_layout.slide_layout_id - 1  # Adjusting for 0-based index
        print(f"\nAnalyzing slide {slide_index + 1} :")
        analyze_single_slide(filepath, slide_index)



analyze_presentation(source1)