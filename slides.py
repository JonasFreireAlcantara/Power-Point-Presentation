import sys

import ml


def usage():
    print(f'\nUtilização:\n  {sys.argv[0]} <<letras_arquivo>> <<background_imagem>> <<logo_imagem>> <<nome_slide_final>>\n')


if len(sys.argv) != 5:
    usage()
    exit(0)

filename = sys.argv[1]
background_image = sys.argv[2]
logo_image = sys.argv[3]
final_slide_name = sys.argv[4]

dictionary = ml.read_file(filename)
title = dictionary[ml.TITLE]
stanzas = dictionary[ml.STANZA]

presentation = ml.create_presentation_slide(title, stanzas, background_image, logo_image)

presentation.save(final_slide_name)
