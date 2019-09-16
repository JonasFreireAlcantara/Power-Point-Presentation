from pptx import Presentation

import ml


dictionary = ml.read_file('./digno_e_o_Senhor.txt')

title = dictionary[ml.TITLE]
stanzas = dictionary[ml.STANZA]

presentation = Presentation()
for stanza in stanzas:
    ml.add_slide_with_title_and_content(presentation, './faces.gif', './nasa.jpg', title, stanza)

presentation.save('test-presentation.pptx')
