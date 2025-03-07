from constants import *
from classes import *

t = Tatics()
log = Logger()
ft = Filetype()
o = Office()
g = Games()

i = f'{folder_tatics}/{name_folder}/'

l_moves = t.get_tatics_moves(path_pgn)

l_fens = t.get_tatics_fens(path_pgn)

path_tatics_docx = f'{folder_tatics}/{name_folder}/{name_folder}_1.docx'
o.create_word(path_tatics_docx)

word = Office().config_word()
doc = Office().open_word(word, path_tatics_docx)

for j, i in enumerate(l_moves):

    file_svg = f'{path_tatics_folder}puzzle_{j+1:04d}.svg'

    if not os.path.isfile(file_svg):

        p_pdf = f'{file_svg[:-3]}pdf'

        pn = f'{(j+1):04d}'

        side_to_move = t.display_board(l_fens[j])

        t.create_tatics_pdf(folder_tatics, l_fens[j], name_folder, pn, side_to_move)

        log.insert('i', f'{p_pdf} was created')

        ft.crop_pdf(f'{p_crop_2}{name_folder}', pn)

        ft.pdf_to_svg(p_pdf, file_svg)

        os.remove(p_pdf)

    else:

        log.insert('i', f'{i}pdf was ALREADY created')

    list_svg = ft.get_list_files(f'{folder_tatics}/{name_folder}', '.svg')
    for index, svg_item in enumerate(list_svg):

        g.edit_word_games(
        l_moves[j], name_folder, f'#{j+1:04d}.svg', 
        file_svg, path_tatics_docx, doc
        )

        doc.Save()
        doc.Close()
        word.Quit()