from pptx.util import Inches, Cm, Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.text import MSO_AUTO_SIZE

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
    title.text = text.upper()
    paragraph = title.text_frame.paragraphs[0]
    paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
    paragraph.font.size = Pt(40)
    paragraph.font.bold = True
    title.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE



def add_stanza(slide, text):
    left = Cm(1.20)
    top = Cm(5.80)
    width = Cm(23.00)
    height = Cm(12.00)
    stanza = slide.shapes.add_textbox(left, top, width, height)
    stanza.text = text
    for paragraph in stanza.text_frame.paragraphs:
        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
        paragraph.font.size = Pt(30)

    stanza.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def add_logo_bottom_right_corner(presentation, image):
    pass


def add_slide_with_title_and_content(presentation, title_text, content_text):
    pass