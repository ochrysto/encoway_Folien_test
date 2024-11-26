from pptx import Presentation

# Load your presentation template
prs = Presentation("slide-9-test.pptx")

# Choose the slide layout index you're interested in
# Slide layouts are indexed starting from 0
layout_index = 19  # Change this to the index of the layout you're interested in

# Get the layout
slide_layout = prs.slide_layouts[layout_index]

# Print placeholder details for the chosen layout
print(f"Details for layout {layout_index}: {slide_layout.name}")
for placeholder in slide_layout.placeholders:
    print(f"Placeholder index: {placeholder.placeholder_format.idx}, Type: {placeholder.placeholder_format.type}, "
          f"Name: '{placeholder.name}' dimensions in Inches: left: {placeholder.left / 914400} "
          f"top: {placeholder.top / 914400} width: {placeholder.width / 914400} height: {placeholder.height / 914400}")