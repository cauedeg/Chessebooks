from constants import *

class Tatics:
    def get_tatics_fens(self, path):

        with open(path, 'r') as arquivo:
            linhas = arquivo.readlines()

        linhas = [linha.strip() for linha in linhas]
        df = pd.DataFrame(linhas, columns=['data'])

        substrings = [
            r'\[Eve',
            r'\[Dat',
            r'\[Sit',
            r'\[Round',
            r'\[Whit',
            r'\[Resu',
            r'\[Blac',
            r'\[Ply',
            r'\[Set',
        ]

        for item in substrings:
            df = df[~df['data'].str.contains(item, na=False)]

        df = df[(df != '').any(axis=1)]

        l_fen = df['data'][df['data'].str.startswith('[FEN')].tolist()
        l_fen = [item[5:-1] for item in l_fen]

        l_fen_2 = []
        for fen in l_fen:
            fen_no_quotes = fen.strip('"')
            l_fen_2.append(fen_no_quotes)

        return l_fen_2

    def get_tatics_moves(self, path):

        """
        Função que realiza uma operação com múltiplos parâmetros.

        Parâmetros esperados em **kwargs:
        - path (str): path of pgn file of tatics.

        """

        with open(path, 'r') as arquivo:
            linhas = arquivo.readlines()

        linhas = [linha.strip() for linha in linhas]
        df = pd.DataFrame(linhas, columns=['data'])

        substrings = [
            r'\[Eve',
            r'\[Dat',
            r'\[Sit',
            r'\[Round',
            r'\[Whit',
            r'\[Resu',
            r'\[Blac',
            r'\[Ply',
            r'\[Set',
            r'\[FEN',
        ]

        for item in substrings:
            df = df[~df['data'].str.contains(item, na=False)]

        df = df[(df != '').any(axis=1)]

        combined_lines = []
        current_combined = ""

        for line in df["data"]:
            if "{[%evp" in line:
        # Salvar a linha anterior se não estiver vazia
                if current_combined:
                    combined_lines.append(current_combined)
        # Iniciar uma nova combinação
                current_combined = line
            else:
        # Adicionar a linha atual à combinação
                current_combined += " " + line

        # Adicionar a última combinação
        if current_combined:
            combined_lines.append(current_combined)

        # Criar um novo DataFrame com as linhas combinadas
        df = pd.DataFrame(combined_lines, columns=["data"])

        l_moves = df['data'].tolist()

        return l_moves

    def create_tatics_pdf(self, path, item, name_folder, p_number,
                          side_to_move):

        """
        Função que realiza uma operação com múltiplos parâmetros.

        Parâmetros esperados em **kwargs:
        - path (str): path of pgn file of tatics.
        - item (str): list item of fens list
        -
        """

        with open(
            f'{path}/{name_folder}/Puzzle {p_number}.tex',
            'a+',
            encoding='utf-8',
        ) as f:

            f.write(r'\documentclass[a4paper,14pt]{extarticle}' + '\n')
            f.write(r'\usepackage[utf8]{inputenc}' + '\n')
            f.write(r'\usepackage{amsmath}' + '\n')
            f.write(r'\usepackage{xskak}' + '\n')
            f.write(r'\usepackage{graphicx}' + '\n')
            f.write(r'\usepackage{chessboard}' + '\n')
            f.write(r'\usepackage{fancyhdr}' + '\n')
            f.write(
                r'\usepackage[left=0.1in, right=0.1in, top=0.2in, bottom=0.2in]{geometry}'
                + '\n'
            )
            f.write(r'\usepackage{hyperref}' + '\n')
            f.write(r'\begin{document}' + '\n')

            f.write(r'\thispagestyle{empty}' + '\n')

            if side_to_move == 'b':
                f.write(r'\chessboard[inverse,' + '\n')

            else:
                f.write(r'\chessboard[' + '\n')
                
            f.write(r'    setfen=' + item + ',' + '\n')
            f.write(r'    boardfontsize=55,' + '\n')
            f.write(r'    showmover=true,' + '\n')
            f.write(r'    showboard=false]' + '\n')

            f.write(r'\end{document}')

            f.seek(0)
            txt_content = f.read()
            txt_content = txt_content.replace('#', r'\#').replace('$', r'\$')

        f.close()

        subprocess.run(
            [
                'pdflatex',
                '-output-directory',
                f'{path}/{name_folder}',
                f'{path}/{name_folder}/Puzzle {p_number}.tex',
            ]
        )

        # Remover arquivos temporários
        arquivos_a_excluir = [
            f'{path}/{name_folder}/Puzzle {p_number}.tex',
            f'{path}/{name_folder}/Puzzle {p_number}.aux',
            f'{path}/{name_folder}/Puzzle {p_number}.log',
            f'{path}/{name_folder}/Puzzle {p_number}.out',
        ]

        for arquivo in arquivos_a_excluir:
            if os.path.exists(arquivo):
                os.remove(arquivo)

        Logger().insert('i', f'Puzzle created {arquivo}')

    def create_word_tatics(
        self, path_tatics_folder, path_docx, responses, pini, pfin, name_docx
    ):

        """
        Create word file of tatics

        Parâmetros esperados em **kwargs:
        - path_tatics_folder (str): path of pgn file of tatics.
        - path_docx (str): path of word file.
        - responses (list): tatics moves list.
        - pini (int): first puzzle to be inserted
        - pfin (int): last puzzle to be inserted
        - name_docx (str): word file name
        """

        if not os.path.exists(path_docx):

            doc = Document()
            doc.save(path_docx)

            # Inicializa o Word
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False  # Executa o Word em modo invisível

            # Abre o documento existente
            print(f'Abrindo o documento existente: {path_docx}')
            doc = word.Documents.Open(path_docx)

        else:
            doc = Document(path_docx)
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False  # Executa o Word em modo invisível

            # Abre o documento existente
            print(f'Abrindo o documento existente: {path_docx}')
            doc = word.Documents.Open(path_docx)

        # Obtém todos os arquivos SVG na pasta
        svg_files = [
            f for f in os.listdir(path_tatics_folder) if f.endswith('.svg')
        ]
        puzzles = svg_files[pini - 1 : pfin]

        for index, item in enumerate(puzzles):
            # Subtítulo para cada imagem
            if index > 0:

                break_paragraph = doc.Content.Paragraphs.Add()
                break_paragraph.Range.InsertBreak(win32.constants.wdPageBreak)
                subtitle_text = f'Puzzle {item[7:11]}'  # Ex: Subtítulo 1, Subtítulo 2, etc.
                subtitle_paragraph = doc.Content.Paragraphs.Add()
                subtitle_paragraph.Range.Text = subtitle_text
                subtitle_paragraph.Range.Style = 'Título 2'
                subtitle_paragraph.Range.InsertParagraphAfter()
                subtitle_paragraph.Range.InsertParagraphAfter()  # Insere uma nova linha após o subtítulo

            else:

                title_paragraph = doc.Content.Paragraphs.Add()
                title_paragraph.Range.Text = name_docx
                title_paragraph.Range.Style = 'Título 1'
                title_paragraph.Range.InsertParagraphAfter()
                title_paragraph.Range.InsertParagraphAfter()  # Insere uma nova linha após o título principal

                subtitle_text = f'Puzzle {item[7:11]}'  # Ex: Subtítulo 1, Subtítulo 2, etc.
                subtitle_paragraph = doc.Content.Paragraphs.Add()
                subtitle_paragraph.Range.Text = subtitle_text
                subtitle_paragraph.Range.Style = 'Título 2'
                subtitle_paragraph.Range.InsertParagraphAfter()
                subtitle_paragraph.Range.InsertParagraphAfter()  # Insere uma nova linha após o subtítulo

            # Caminho do arquivo SVG
            svg_path = os.path.join(path_tatics_folder, item)
            print(f'Inserindo imagem SVG: {svg_path}')

            # Insere a imagem SVG imediatamente após o subtítulo
            shape_range = (
                subtitle_paragraph.Range
            )  # Insere a imagem após o subtítulo
            inline_shape = shape_range.InlineShapes.AddPicture(
                FileName=svg_path
            )
            inline_shape.Width = 466.66  # Ajuste o tamanho conforme necessário
            inline_shape.Height = 400
            shape_range.InsertParagraphAfter()

            break_paragraph = doc.Content.Paragraphs.Add()
            break_paragraph.Range.InsertBreak(win32.constants.wdPageBreak)

            response_text = (
                responses[index]
                if index < len(responses)
                else 'Resposta padrão'
            )
            response_paragraph = doc.Content.Paragraphs.Add()
            response_paragraph.Range.Text = response_text

            response_paragraph.Range.InsertParagraphAfter()

        # Salva o documento
        doc.Save()
        doc.Close()
        word.Quit()

    def display_board(self, fen):
        # Cria o tabuleiro a partir da FEN
        board = chess.Board(fen)
        
        # Determina o lado a partir do campo FEN (quem joga)
        if board.turn == chess.BLACK:
            side_to_move = 'b'
        else:
            side_to_move = 'w'

        return side_to_move

    def get_word(self, indice, folder_name):

        intervalo = 500
        inicio = (indice // intervalo) * intervalo + 1
        fim = inicio + intervalo - 1
        return f"{folder_name}_{inicio}_{fim}.docx"

class Games:
    def get_str_game_moves(self, game_path):

        with open(game_path, 'r') as arquivo:
            linhas = arquivo.readlines()

        linhas = [linha.strip() for linha in linhas]
        df = pd.DataFrame(linhas, columns=['Conteúdo'])
        substrings = [
            r'\[Event',
            r'\[Site',
            r'\[Date',
            r'\[Site',
            r'\[Round',
            r'\[White',
            r'\[Black',
            r'\[Result',
            r'\[ECO',
            r'\[Time',
            r'\[EndTime',
            r'\[Termination',
        ]

        for item in substrings:
            df = df[~df['Conteúdo'].str.contains(item, na=False)]

        df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
        df = df.dropna(how='all')
        game = df.to_string(index=False, header=False)
        game = game.replace('\n', '').replace('  ', '')

        return game

    def get_names(self, string):
        base = string.split('_', 1)[1].replace('.pgn', '')
        names = base.split('_vs_')
        return names

    def get_str_first_move(self, game_path):

        game = self.get_str_game_moves(game_path)
        names = self.get_str_game_players(game_path)
        game_list = ' '.join(game.split())
        game_list = game_list.split()

        if names[0] == 'cauedeg':
            if game_list[2] == 'e5':
                mainline = 'e5_brancas'
            elif game_list[2] == 'd5':
                mainline = 'd5_brancas'
            elif game_list[2] == 'e6':
                mainline = 'e6_brancas'
            else:
                mainline = None
                log.insert('w', 'Linha principal ainda não inserida na base')
        else:
            if game_list[1] == 'e4':
                mainline = 'e4_pretas'
            elif game_list[1] == 'd4':
                mainline = 'd4_pretas'
            else:
                mainline = None
                log.insert('w', 'Linha principal ainda não inserida na base')

        return mainline

    def edit_word_games(
        self, text, title, subtitle, images, path_file_docx, doc,
    ):

        """
        Create word file of tatics

        Parâmetros esperados em **kwargs:
        - path_tatics_folder (str): path of pgn file of tatics.
        - path_docx (str): path of word file.
        - responses (list): tatics moves list.
        - pini (int): first puzzle to be inserted
        - pfin (int): last puzzle to be inserted
        - name_docx (str): word file name
        """

        title_exists = Office().check_title(doc, title)
        subtitle_exists = Office().check_subtitle(doc, title, subtitle)
        check_img_exists = Office().check_word_img(doc, title, subtitle, 466.66, 400)
        text_exists = Office().check_text(doc, title, subtitle, text)

        if not title_exists:

            end_range = doc.Content
            end_range.Collapse(0)

            Office().insert_word_title(doc, title)
            Office().insert_blank_paragraph(doc, 1)
            Office().insert_word_subtitle(doc, title, subtitle, 2)
            Office().insert_blank_paragraph(doc, 1)
            Office().insert_word_image(doc, images, 466.66, 400)
            Office().insert_blank_paragraph(doc, 1)
            Office().insert_text(doc, text)

        elif title_exists and not subtitle_exists:

            end_range = doc.Content
            end_range.Collapse(0)
        
            Office().insert_word_subtitle(doc, title, subtitle, 2)
            Office().insert_blank_paragraph(doc, 1)
            Office().insert_word_image(doc, images, 466.66, 400)
            Office().insert_blank_paragraph(doc, 1)
            Office().insert_text(doc, text)
        
        elif title_exists and subtitle_exists and not check_img_exists:

            end_range = doc.Content
            end_range.Collapse(0)

            Office().insert_word_image(doc, images, 466.66, 400)
            Office().insert_blank_paragraph(doc, 1)
            Office().insert_text(doc, text)

        elif title_exists and subtitle_exists and check_img_exists and not text_exists:

            end_range = doc.Content
            end_range.Collapse(0)

            Office().insert_text(doc, text)

    def open_word_file(self, file_docx):

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Executa o Word em modo invisível

        print(f'Abrindo o documento existente: {file_docx}')
        doc = word.Documents.Open(file_docx)

        return doc

    def create_new_filename(self, filename: str):

        match = re.search(r'(\d+)', filename)
        number = int(match.group(0)) + 1
        incremented_number = f"{number:04d}"

        return re.sub(r'\d+', incremented_number, filename, count=1)


class Engines:
    def evaluate_position(self, fen, var_n):

        stockfish.set_fen_position(fen)
        stockfish.set_depth(25)  # Configurar a profundidade da análise
        evaluation = stockfish.get_top_moves(var_n)
        stockfish.get_best_move()
        return evaluation

    def create_main_line_df(self, gamepath, fenbase, names):

        game = chess.pgn.read_game(open(gamepath))
        board = game.board()
        pgn_list = ' '.join(gamepath.split())
        pgn_list = pgn_list.split()

        l_score = []
        l_fen = []
        l_move = []
        l_top_3_moves = []
        l_alg_move = []

        l_uci_moves = list(game.mainline_moves())

        for move in game.mainline_moves():

            positions = pd.read_csv(fenbase)
            positions = positions.drop_duplicates(subset=['fen', 'move_rank'])
            exists = positions[
                (positions['fen'] == board.fen())
            ]

            if not exists.empty:

                if exists.values.tolist()[0][2]:
                    move1 = exists.values.tolist()[0][2]
                    move1 = ast.literal_eval(move1)

                if len(exists.values.tolist()) != 1:
                    move2 = exists.values.tolist()[1][2]
                    move2 = ast.literal_eval(move2)

                if len(exists.values.tolist()) > 2:
                    move3 = exists.values.tolist()[2][2]
                    move3 = ast.literal_eval(move3)

                eval_main_line = [move1, move2, move3]
                Logger().insert('d', f'Position found: {board.fen()}')

            else:
                eval_main_line = self.evaluate_position(board.fen(), 3)
                with open(
                    fenbase, mode='a', newline='', encoding='utf-8'
                ) as file:
                    writer = csv.writer(file)

                    for index, item in enumerate(eval_main_line):

                        writer.writerow(
                            [board.fen(), index+1, item]
                        )

                Logger().insert('d', f'Position evaluated: {board.fen()}')

            move_number = board.fullmove_number
            side_to_move = 'w' if board.turn else 'b'
            alg_move = board.san(move)

            if eval_main_line[0]['Centipawn'] is not None:
                eval_score = eval_main_line[0]['Centipawn'] / 100
                l_score.append(eval_score)
            else:
                eval_score = eval_main_line[0]['Mate'] * 1000
                l_score.append(eval_main_line[0]['Mate'] * 1000)

            l_move.append(f'{move_number}{side_to_move}.')
            l_alg_move.append(f'{alg_move}')
            l_fen.append(board.fen())
            l_top_3_moves.append(eval_main_line)
            board.push(move)


            if l_uci_moves.index(move)+1 == len(l_uci_moves):
                eval_last_line = self.evaluate_position(board.fen(), 3)

        #evaluating the last move
        if eval_last_line:
            if eval_last_line[0]['Centipawn'] is not None:
                eval_score = eval_last_line[0]['Centipawn'] / 100
                l_score.append(eval_score)
            else:
                eval_score = eval_last_line[0]['Mate'] * 1000
                l_score.append(eval_last_line[0]['Mate'] * 1000)
        else:
            l_score.append(0.0)

        df = pd.DataFrame(
            {
                'moves_number': l_move,
                'played_move': l_alg_move,
                'score_pre_move': l_score[:-1],
                'score_after_move': l_score[1:],
                'fen_pre_move': l_fen,
                'top_three_moves_pre_move': l_top_3_moves,
            }
        )

        df['dif'] = abs(round(df['score_pre_move'].diff(), 2))
        mark_index = df.index[df['dif'] > 0.6].tolist()
        list_index = [i - 1 for i in mark_index if i > 0]
        df_answer = df.loc[mark_index].reset_index(drop=True)
        df_answer = df_answer.drop(columns=['score_pre_move'])
        df_answer = df_answer.rename(
            columns={
                'fen_pre_move': 'fen_after_move',
                'top_three_moves_pre_move': 'top_three_moves_after_move',
            }
        )

        df_puzzle = df.loc[list_index].reset_index(drop=True)
        df_puzzle = df_puzzle.drop(columns=['dif'])

        dffin = pd.concat(
            [
                df_puzzle,
                df_answer['dif'],
                df_answer['fen_after_move'],
                df_answer['top_three_moves_after_move'],
            ],
            axis=1,
        )

        if names[0] == 'cauedeg':
            side_puzzle = 'w.'
        else:
            side_puzzle = 'b.'

        dffin = dffin[
            dffin['moves_number'].str.contains(side_puzzle, case=False, na=False)
        ]
        dffin = dffin.reset_index()
        dffin = dffin[
            [
                'fen_pre_move',
                'score_pre_move',
                'moves_number',
                'played_move',
                'score_after_move',
                'fen_after_move',
                'dif',
                'top_three_moves_pre_move',
                'top_three_moves_after_move',
            ]
        ]

        return dffin

    def create_variants(self, fen_column, df, var_column, fenbase, best_moves):

        df = df.copy() #para não alterar o dataframe original

        l_final_vars = []
        positions = pd.read_csv(fenbase)

        if best_moves == 'top_three_moves_pre_move':
            var_n = 3
        else:
            var_n = 1

        for r in range(len(df)):

            fen = df.loc[r][fen_column]
            b1 = chess.Board(fen)
            b2 = chess.Board(fen)
            b3 = chess.Board(fen)
            l_uci_1 = []
            l_uci_2 = []
            l_uci_3 = []
            l_boards = [b1, b2, b3]
            l_alg_moves = [l_uci_1, l_uci_2, l_uci_3]

            for i in range(var_n):

                str_move = df.loc[r][best_moves][i]['Move']
                uci_move = chess.Move.from_uci(str_move)
                move_number_var = l_boards[i].fullmove_number
                side_to_move_var = 'w' if l_boards[i].turn else 'b'
                l_alg_moves[i].append(
                    f'{move_number_var}{side_to_move_var}. {l_boards[i].san(uci_move)}'
                )
                Logger().insert(
                    'i',
                    f'[line {i+1}], puzzle {r+1} from {len(df)}, move {move_number_var}',
                )
                l_boards[i].push(uci_move)

                for _ in range(10):

                    exists = positions[
                        (positions['fen'] == l_boards[i].fen())
                        & (positions['move_rank'] == i+1)
                    ]

                    if not exists.empty:
                        ev_var_lines = exists.values.tolist()[0][2]
                        ev_var_lines = ast.literal_eval(ev_var_lines)
                        Logger().insert('d', f'Position Found: {l_boards[i].fen()}')

                    else:
                        ev_var_lines = self.evaluate_position(l_boards[i].fen(), 1)
                        if ev_var_lines:
                            with open(
                                fenbase, mode='a', newline='', encoding='utf-8'
                            ) as file:
                                writer = csv.writer(file)
                                writer.writerow(
                                    [
                                        l_boards[i].fen(),
                                        i+1,
                                        ev_var_lines[0],
                                    ]
                                )

                            Logger().insert(
                                'd', f'Position evaluated: {l_boards[i].fen()}'
                            )
                        else:
                            Logger().insert(
                                'd', f'Line is over: {l_boards[i].fen()}'
                            )

                    if ev_var_lines:

                        if type(ev_var_lines) == list:
                            var_move = ev_var_lines[0]['Move']
                        else:
                            var_move = ev_var_lines['Move']

                        uci_move_var = chess.Move.from_uci(var_move)
                        move_number_var = l_boards[i].fullmove_number
                        side_to_move_var = 'w' if l_boards[i].turn else 'b'
                        l_alg_moves[i].append(
                            f'{move_number_var}{side_to_move_var}. {l_boards[i].san(uci_move_var)}'
                        )
                        l_boards[i].push(uci_move_var)

                if ev_var_lines:
                    if type(ev_var_lines) == list:
                        ev_var_lines = ev_var_lines[0]

                    if ev_var_lines['Centipawn'] is not None:
                        l_alg_moves[i].append(
                            f"Evaluation: {ev_var_lines['Centipawn']/100}"
                        )
                    else:
                        l_alg_moves[i].append(
                            f"Evaluation: {ev_var_lines['Mate']}"
                        )

            if l_alg_moves[1]:
                partes = [
                    ' '.join(map(str, sublista)) for sublista in l_alg_moves
                ]
                l_alg_moves = ' || '.join(partes)
                l_alg_moves = f'{l_alg_moves}'
                l_final_vars.append(l_alg_moves)
            else:
                l_alg_moves = l_alg_moves[0]
                l_alg_moves = ' '.join(l_alg_moves)
                l_final_vars.append(f' || {l_alg_moves}')

            Logger().insert('i', f'Puzzle {r+1}/{len(df)} Finished')

        df[var_column] = l_final_vars

        return df


class Office_win32:
    def open_word(self, path_docx):

        # Abre o documento no Word
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(path_docx)
        return doc
    
    def close_word(self, doc):

        word = win32.Dispatch('Word.Application')
        doc.Close()
        word.Quit()

    def create_word(path_docx):

        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        doc = word.Documents.Add()

        doc.SaveAs(path_docx)
        doc.Close()
        word.Quit()

    def insert_paragraph_text(self, doc, text, text_style):

        new_paragraph = doc.Content.Paragraphs.Add()
        new_paragraph.Range.Text = text
        new_paragraph.Range.Style = text_style

    def insert_page_break(doc):

        """
        Insere um subtítulo após o último subtítulo existente dentro de um título específico.

        Args:
            doc: O documento Word.
            title: O título no qual o subtítulo será inserido.
            subtitle_text: O texto do subtítulo a ser inserido.
        """

        range_end = doc.Content
        range_end.Collapse(0)
        range_end.InsertBreak(4)

    def insert_blank_paragraph(doc, paragraphs_number):
        """
        Insere um número específico de parágrafos em branco no final do documento,
        garantindo que não haja parágrafos extras ou duplicados.

        Args:
            doc: O documento Word.
            paragraphs_number: O número de parágrafos em branco a serem inseridos.
        """
        range_end = doc.Content
        range_end.Collapse(0)  # Move o cursor para o final do conteúdo existente

        for _ in range(paragraphs_number):
            range_end.InsertParagraphAfter()

    def insert_word_title(doc, title):

        """
        Insere um título no documento Word usando win32com.

        Args:
            doc: O documento Word.
            title: O texto do título a ser inserido.
        """

        paragraph = doc.Content.Paragraphs.Add()
        paragraph.Range.Text = title
        paragraph.Style = doc.Styles("Título 1")
        paragraph.Range.Collapse(0)

    def insert_word_subtitle(self, doc, title, subtitle_text):
        """
        Insere um subtítulo após o último subtítulo existente dentro de um título específico.

        Args:
            doc: O documento Word.
            title: O título no qual o subtítulo será inserido.
            subtitle_text: O texto do subtítulo a ser inserido.
        """

        title_found = False
        last_subtitle_paragraph = None

        # Itera pelos parágrafos no documento
        for paragraph in doc.Paragraphs:
            text = paragraph.Range.Text.strip()
            style = paragraph.Style.NameLocal

            # Verifica se o parágrafo é o título especificado (Heading 1)
            if style == "Título 1" and text == title:
                title_found = True
                last_subtitle_paragraph = paragraph
                continue

            # Verifica se o parágrafo é um subtítulo (Heading 2)
            if title_found and style == "Título 2":
                last_subtitle_paragraph = paragraph
                continue

            # Se um novo título for encontrado, interrompe a busca
            if title_found and style == "Título 1" and text != title:
                break

        # Define o ponto de inserção
        if last_subtitle_paragraph:
            insert_after_range = last_subtitle_paragraph.Range
        elif title_found:
            insert_after_range = paragraph.Range
        else:
            raise ValueError(f"O título '{title}' não foi encontrado no documento.")

        # Move o cursor para depois do último subtítulo ou título
        insert_after_range.InsertAfter("\n" + subtitle_text)
        insert_after_range.Collapse(0)

        # Define o estilo do novo parágrafo como "Heading 2"

        insert_after_range.Style = doc.Styles("Título 2")
        print(f"Subtítulo '{subtitle_text}' inserido com sucesso.")
        insert_after_range.Collapse(0)

    def insert_word_image(
            doc, path_svg, width, height
            ):

        subtitle_paragraph = doc.Content.Paragraphs.Add()

        shape_range = (
            subtitle_paragraph.Range
        )  # Insere a imagem após o subtítulo
        inline_shape = shape_range.InlineShapes.AddPicture(
            FileName=path_svg
        )
        inline_shape.Width = width  # Ajuste o tamanho conforme necessário
        inline_shape.Height = height

    def check_title(self, doc, title):

        title_exists = False

        for i in range(1, doc.Paragraphs.Count + 1):  # Iterando pelos índices de 1 a Paragraphs.Count
            paragraph = doc.Paragraphs(i)
            text = paragraph.Range.Text.strip()  # Obtendo o texto do parágrafo
            style = paragraph.Range.Style.NameLocal  # Obtendo o estilo do parágrafo

            if style == 'Título 1' and text == title:  # Substitua 'Título 1' pelo nome correto
                title_exists = True
                break

        return title_exists
    
    def check_subtitle(doc, title, subtitle):
        
        """
        Procura um subtítulo dentro de um título específico em um documento Word.

        Args:
            doc: O documento Word.
            title: O título no qual buscar o subtítulo.
            subtitle: O subtítulo a ser procurado.

        Returns:
            bool: True se o subtítulo for encontrado dentro do título, False caso contrário.
        """
        title_found = False
        subtitle_exists = False

        for i in range(1, doc.Paragraphs.Count + 1):  # Iterando pelos índices de 1 a Paragraphs.Count
            paragraph = doc.Paragraphs(i)
            text = paragraph.Range.Text.strip()  # Obtendo o texto do parágrafo
            style = paragraph.Range.Style.NameLocal  # Obtendo o estilo do parágrafo

            # Verifica se o parágrafo atual é o título especificado
            if style == 'Título 1' and text == title:  # Substitua 'Título 1' pelo nome correto do estilo
                title_found = True
                continue

            # Verifica se o parágrafo é o subtítulo e está dentro do título encontrado
            if title_found and style == 'Título 2' and text == subtitle:  # Substitua 'Título 2' pelo nome correto do estilo
                subtitle_exists = True
                print("Subtitulo encontrado")
                break  # Interrompe a busca ao encontrar o subtítulo

            # Se um novo título for encontrado, interrompe a busca no título atual
            if title_found and style == 'Título 1' and text != title:
                break

        return subtitle_exists
    
    def check_word_img(self, doc, title, subtitle, ex_width, ex_height):
        # Abre o Word e o documento
        
        title_found = False
        subtitle_found = False

        # Itera pelos parágrafos do documento
        for paragraph in doc.Paragraphs:
            text = paragraph.Range.Text.strip()
            style = paragraph.Range.Style.NameLocal

            # Localiza o título desejado
            if style == 'Título 1' and text == title:
                title_found = True
                subtitle_found = False  # Redefine a busca do subtítulo
                continue

            # Se estiver no título correto, procura o subtítulo
            if title_found and style == 'Título 2' and text == subtitle:
                subtitle_found = True
                continue

            # Quando o subtítulo correto for encontrado, busca a imagem dentro dele
            if subtitle_found:
                # Itera pelas imagens associadas ao subtítulo
                for shape in paragraph.Range.InlineShapes:
                    width = shape.Width
                    height = shape.Height

                    # Verifica se existe uma imagem com as dimensões esperadas
                    if abs(width - ex_width) < 1 and abs(height - ex_height) < 1:
                        return True

        # Fecha o documento se não encontrar a imagem
        return False

    def check_text_in_subtitle(self, doc, title_text, subtitle_text, search_text):
        # Abre o Word e o documento

        title_found = False
        subtitle_found = False
        text_found = False

        # Itera pelos parágrafos do documento
        for paragraph in doc.Paragraphs:
            text = paragraph.Range.Text.strip()
            style = paragraph.Range.Style.NameLocal

            # Localiza o título desejado
            if style == 'Título 1' and text == title_text:
                title_found = True
                subtitle_found = False  # Redefine a busca do subtítulo
                continue

            # Se estiver no título correto, procura o subtítulo
            if title_found and style == 'Título 2' and text == subtitle_text:
                subtitle_found = True
                continue

            # Quando estiver no subtítulo correto, busca pelo texto desejado
            if subtitle_found and search_text in text:
                text_found = True
                break

        # Retorna o resultado
        if text_found:
            return True
        else:
            return False

    def check_text_in_title(self, doc, title_text, search_text):
        # Inicializa as variáveis de controle
        title_found = False
        text_found = False

        # Itera pelos parágrafos do documento
        for paragraph in doc.Paragraphs:
            text = paragraph.Range.Text.strip()
            style = paragraph.Range.Style.NameLocal

            # Localiza o título desejado
            if style == 'Título 1' and text == title_text:
                title_found = True
                continue  # Continua a busca pelo conteúdo dentro do título

            # Se estiver no título correto, verifica a presença do texto
            if title_found and search_text in text:
                text_found = True
                break

        # Retorna o resultado
        return text_found

    def insert_text(self, doc, text):
                    
        response_paragraph = doc.Content.Paragraphs.Add()
        response_paragraph.Range.Text = text
        response_paragraph.Range.InsertParagraphAfter()
        response_paragraph.Range.InsertBreak(win32com.client.constants.wdPageBreak)
        
    def is_word_file_open(self, file_name):
        # Converte o nome do arquivo para minúsculo para evitar problemas de comparação
        file_name = file_name.lower()
        
        # Itera pelos processos em execução
        for process in psutil.process_iter(['name', 'cmdline']):
            try:
                # Verifica se o processo é do Microsoft Word
                if process.info['name'] and 'winword' in process.info['name'].lower():
                    # Verifica os argumentos do processo para ver se o arquivo está sendo usado
                    if process.info['cmdline']:
                        for arg in process.info['cmdline']:
                            if file_name in arg.lower():
                                return True  # O arquivo está aberto
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass  # Ignora erros ao acessar processos
        
        return False  # O arquivo não está aberto

class Office_python_docx:

    def create_word(path_docx):

        doc = Document()
        doc.save(path_docx)

    def insert_paragraph_text(self, doc, text, text_style):

        paragraph = doc.add_paragraph(text)
        paragraph.style = text_style

class CustomFormatter(logging.Formatter):

    grey = '\x1b[38;20m'
    yellow = '\x1b[33;20m'
    red = '\x1b[31;20m'
    bold_red = '\x1b[31;1m'
    reset = '\x1b[0m'
    format = (
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s (%(filename)s)'
    )

    FORMATS = {
        logging.DEBUG: Style.BRIGHT
        + format
        + Style.RESET_ALL,  # Cor verde para DEBUG
        logging.INFO: Style.BRIGHT
        + Fore.BLUE
        + format
        + Style.RESET_ALL,  # INFO em branco e negrito
        logging.WARNING: Fore.YELLOW
        + format
        + Style.RESET_ALL,  # WARNING em amarelo
        logging.ERROR: Fore.RED
        + format
        + Style.RESET_ALL,  # ERROR em vermelho
        logging.CRITICAL: Style.BRIGHT
        + Fore.RED
        + format
        + Style.RESET_ALL,  # CRITICAL em vermelho e negrito
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno)
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)

class Logger:

    def __init__(self):
        self.logger = logging.getLogger('My_app')
        self.logger.setLevel(logging.DEBUG)

        # Verifica se o logger já tem um handler
        if not self.logger.hasHandlers():
            # create console handler with a higher log level
            ch = logging.StreamHandler()
            ch.setLevel(logging.DEBUG)
            ch.setFormatter(CustomFormatter())
            self.logger.addHandler(ch)

    def insert(self, type, message):
        if type == 'd':
            self.logger.debug(message)
        elif type == 'i':
            self.logger.info(message)
        elif type == 'w':
            self.logger.warning(message)
        elif type == 'e':
            self.logger.error(message)
        elif type == 'c':
            self.logger.critical(message)

class Filetype:

    def pdf_to_svg(self, input_pdf, output_svg):
        subprocess.run([
            "inkscape", input_pdf, "--export-type=svg", "--export-filename", output_svg
        ], check=True)

    def crop_pdf(self, input_pdf_path, width_left_margin,
                lenght_inferior_margin, width_board, 
                lenght_board
                ):

        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        # Coordenadas específicas baseadas na análise do PDF original
        width_left_margin = 75   # Ajuste para começar após a margem esquerda
        lenght_inferior_margin = 315   # Ajuste para cortar a parte inferior branca
        width_board = 520  # Largura suficiente para o tabuleiro e indicador
        lenght_board = 470   # Altura ajustada para incluir o indicador de quem joga

        # Loop por todas as páginas do PDF
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]

            # Define o novo mediabox ajustado
            page.mediabox.lower_left = (
                width_left_margin,
                lenght_inferior_margin
            )
            page.mediabox.upper_right = (
                width_left_margin + width_board,
                lenght_inferior_margin + lenght_board
            )

            # Adiciona a página cortada ao novo PDF
            writer.add_page(page)

        # Salva o PDF cortado
        with open(input_pdf_path, "wb") as output_file:
            writer.write(output_file)

        dimensoes_finais = (
            writer.pages[0].mediabox.width,
            writer.pages[0].mediabox.height
        )
        return dimensoes_finais[0], dimensoes_finais[1]

    def get_dimensions(self, caminho_pdf: str, pagina: int = 0):
        # Lê o PDF
        reader = PdfReader(caminho_pdf)
        
        # Verifica se a página solicitada existe
        if pagina >= len(reader.pages):
            raise ValueError(f"O PDF possui apenas {len(reader.pages)} páginas. Página {pagina} é inválida.")
        
        # Obtém o mediabox da página especificada
        page = reader.pages[pagina]
        width = page.mediabox.width
        lenght = page.mediabox.height
        
        return width, lenght

    def optimize_svg(path_folder):

        command = [svgo_path, '-f', path_folder]

        try:
            # Executa o SVGO
            result = subprocess.run(command, check=True, capture_output=True, text=True)
            print('SVGO output:', result.stdout)

            # Renomeia os arquivos otimizados
            for filename in os.listdir(path_folder):
                if filename.endswith('.svg') and not filename.startswith('o_'):
                    original_path = os.path.join(path_folder, filename)
                    optimized_path = os.path.join(path_folder, f'o_{filename}')
                    os.rename(original_path, optimized_path)
                    print(f'Renamed: {filename} -> o_{filename}')

        except subprocess.CalledProcessError as e:
            print('Error running SVGO:')
            print(e.stderr)
    
    def get_list_files(self, fo_path, term=None):

        if term is None:
            return [file for file in os.listdir(fo_path) if not file.startswith('~$')]
        else:
            return [file for file in os.listdir(fo_path) if file.endswith(term) and not file.startswith('~$')]
