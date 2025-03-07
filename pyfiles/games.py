# from settings import *
from constants import *
from classes import *

g = Games()
e = Engines()
ft = Filetype()
t = Tatics()
o = Office()

l_folders = [file for file in os.listdir(folder_pgns)]
l_docx = ft.get_list_files(folder_pgns, '.docx')
l_pdf = ft.get_list_files(folder_pgns, '.pdf')

for index, name_game_folder in enumerate(l_folders):

    path_folder_game = f'{folder_pgns}/{name_game_folder}/'
    path_game = f'{path_folder_game}{name_game_folder}.pgn'
    names = g.get_names(name_game_folder)
    
    main_line = e.create_main_line_df(
        path_game, p_fenbase, names
        )
    
    variants_bm = e.create_variants(
        "fen_pre_move", main_line, "variants_bm", p_fenbase, "top_three_moves_pre_move"
        )
    
    variants_af = e.create_variants(
    "fen_after_move", variants_bm, "variants_af", p_fenbase, "top_three_moves_after_move"
    )

    #creating lists  of dataframe
    l_fens = variants_af['fen_pre_move'].values.tolist()
    l_moves_number = variants_af['moves_number'].values.tolist()
    l_variants_bm = variants_af['variants_bm'].values.tolist()
    l_variants_af = variants_af['variants_af'].values.tolist()
    combined_variants = [f"{a} {b}" for a, b in zip(l_variants_bm, l_variants_af)]

    for df_row in range(len(variants_af)):

        mov_number = l_moves_number[df_row][:-1]

        #create match puzzles
        file_svg = f'{folder_pgns}{name_game_folder}/puzzle_{mov_number}.svg'

        if not os.path.isfile(file_svg):

            p_pdf = f'{file_svg[:-3]}pdf'

            side_to_move = t.display_board(l_fens[df_row])

            t.create_tatics_pdf(folder_pgns, l_fens[df_row], name_game_folder, mov_number, 
                                side_to_move)

            ft.crop_pdf(f'{p_crop}{name_game_folder}', mov_number)

            ft.pdf_to_svg(p_pdf, file_svg)

            os.remove(p_pdf)

    #OPEN WORD FILE
    list_svg = ft.get_list_files(path_folder_game, '.svg')
    list_svg = sorted(list_svg, key=lambda x: int(''.join(filter(str.isdigit, x.split('_')[1]))))

    for index, svg_item in enumerate(list_svg):

        list_docx = Filetype().get_list_files(folder_pgns, '.docx')

        if not list_docx:
            path_games_docx = f'{folder_pgns}games_0001.docx'
            o.create_word(path_games_docx)

        else:
            if len(list_docx) > 1:
                name_docx = list_docx[-1]
                path_games_docx = f'{folder_pgns}{name_docx}'
            else:
                name_docx = list_docx[0]
                path_games_docx = f'{folder_pgns}{name_docx}'

        docx_size = os.path.getsize(path_games_docx)

        if docx_size > word_size_limit:
            path_games_docx = g.create_new_filename(name_docx)

        path_svg_image = f'{path_folder_game}{svg_item}'
        name_subtitle = svg_item[svg_item.index('puzzle'):]

        word = Office().config_word()
        doc = Office().open_word(word, path_games_docx)

        g.edit_word_games(
        combined_variants[index], name_game_folder, name_subtitle, 
        path_svg_image, path_games_docx, doc
        )

        doc.Save()
        doc.Close()
        word.Quit()

    print("pause")