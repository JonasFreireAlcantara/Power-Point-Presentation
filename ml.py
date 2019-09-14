from pptx.util import Inches

'''
    Sequence of slide layouts from python-pptx guide
    https://python-pptx.readthedocs.io/en/latest/user/slides.html

    Title (presentation title slide)
    Title and Content
    Section Header (sometimes called Segue)
    Two Content (side by side bullet textboxes)
    Comparison (same but additional title for each side by side content box)
    Title Only
    Blank
    Content with Caption
    Picture with Caption
'''


def add_background(slide, image):
    pass


def add_title(slide, text):
    title = slide.shapes.title
    title.text = text


def add_stanza(slide, text):
    left = Inches(1)
    top = Inches(2)
    width = Inches(3)
    height = Inches(4)
    stanza = slide.shapes.add_textbox(left, top, width, height)
    stanza.text = text


def add_logo_bottom_right_corner(presentation, image):
    pass


def add_slide_with_title_and_content(presentation, title_text, content_text):
    pass