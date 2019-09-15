from pptx import Presentation

import ml


prs = Presentation()
title_only_layout = prs.slide_layouts[5]

# Adicionar um novo slide com o layout TITLE_ONLY_LAYOUT
slide = prs.slides.add_slide(title_only_layout)
# Setar o titulo para o slide
ml.add_title(slide, 'Digno é o Senhor')
# Adicionar uma estrofe para o slide
ml.add_stanza(slide, 'Graças eu te dou, Pai\nPelo preço que pagou\nSacrifício de amor\nQue me comprou\nUngido do Senhor')

slide = prs.slides.add_slide(title_only_layout)
ml.add_title(slide, 'Digno é o Senhor')
ml.add_stanza(slide, 'Pelos cravos em Suas mãos\nGraças eu te dou, ó meu Senhor\nLavou minha mente e coração\nMe deu perdão')

# Salvar a apresentação em arquivo com extensão pptx
prs.save('test-presentation.pptx')