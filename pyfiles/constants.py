import ast
import chess
import chess.pgn
# from stockfish import Stockfish
from stockfish import *
import pandas as pd

import io
import pandas as pd
import numpy as np
import re
import subprocess
from docx.shared import Inches
import pypandoc

import os
import json
import csv
import hashlib
from pypdf import PdfReader, PdfWriter
from docx import Document
from docx.shared import Inches, Pt

import sys
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

import win32com.client as win32
from colorama import Fore, Style, init
import logging
import pandas as pd
import subprocess
import os
import shutil

import psutil
import xml.etree.ElementTree as ET

import win32com.client

sys.path.append(
    os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
)

# name_folder = '600_1200_elem_atq_dup_peao'

fo_chessebooks = Path(__file__).parent
fo_projects = fo_chessebooks.parent
fo_python = fo_projects.parent
fo_catalog = fo_python.parent
fo_git = fo_catalog.parent

folder_work = fo_git.parent
folder_chess = f'{folder_work}/chess'
folder_tatics = f'{folder_chess}/tatics/'
folder_pgns = f'{folder_chess}/games/pgns/'

engine = f'{fo_projects}/inputs/sf16.exe'
p_fenbase = f'{fo_projects}/inputs/fenbase.csv'

############# word.py inputs #############################
pini = 1
pfin = 500
word_size_limit = 200000000
# name_docx = f'{name_folder}_{pini}_{pfin}'
# file_tatics_docx = f'{name_docx}.docx'
# path_pgn = f'{folder_tatics}/{name_folder}/{name_folder}.pgn'
# path_tatics_folder = f'{folder_tatics}/{name_folder}/'
# path_docx_tatics = f'{path_tatics_folder}{file_tatics_docx}'
##########################################################

svgo_path = 'C:/Users/cauedeg/AppData/Roaming/npm/svgo.cmd'
p_crop = 'C:/Users/cauedeg/OneDrive/work/chess/games/pgns/'
p_crop_2 = 'C:/Users/cauedeg/OneDrive/work/chess/tatics/'

stockfish = Stockfish(
    engine, parameters={'Threads': 8, 'Minimum Thinking Time': 30}
)

