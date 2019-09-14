from pptx import Presentation

import ml


prs = Presentation()
title_only_layout = prs.slide_layouts[5]

slide = prs.slides.add_slide(title_only_layout)
ml.add_title(slide, 'Slide 1')
ml.add_stanza(slide, 'açslkfajjlkçkljçlkjçjklçlkçsdklf\naçlkfjaçlsdkfj\naçlkjfçaskld\naçlskfdj')


slide = prs.slides.add_slide(title_only_layout)
ml.add_title(slide, 'Slide 2')

prs.save('test-presentation.pptx')