"""
Build script for creating AnkiCardGenerator binary.
Run: python build_exe.py
Output: dist/AnkiCardGenerator_v{version}_{platform}.exe (or no extension on Linux/Mac)
"""
import PyInstaller.__main__
import tkinterdnd2
import pykakasi
import os
import sys
import platform

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gen_anki import __version__

tkdnd_path = os.path.dirname(tkinterdnd2.__file__)
pykakasi_path = os.path.dirname(pykakasi.__file__)

# Detect platform for filename
system = platform.system().lower()
if system == 'windows':
    plat = 'win'
elif system == 'darwin':
    plat = 'mac'
else:
    plat = 'linux'

# Set data separator (: for Linux/Mac, ; for Windows)
sep = ';' if system == 'windows' else ':'

# Clean up previous .spec files
import glob
for spec_file in glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'AnkiCardGenerator_*.spec')):
    os.remove(spec_file)

app_name = f'AnkiCardGenerator_v{__version__}_{plat}'

PyInstaller.__main__.run([
    'gen_anki_gui.py',
    '--onefile',
    '--windowed',
    f'--name={app_name}',
    f'--add-data={tkdnd_path}{sep}tkinterdnd2',
    f'--add-data={pykakasi_path}{sep}pykakasi',
    '--hidden-import=tkinterdnd2',
    '--hidden-import=genanki',
    '--hidden-import=gtts',
    '--hidden-import=deep_translator',
    '--hidden-import=openpyxl',
    '--hidden-import=requests',
    '--hidden-import=beautifulsoup4',
    '--hidden-import=pykakasi',
    '--hidden-import=imageio',
    '--hidden-import=imageio_ffmpeg',
    '--collect-all=tkinterdnd2',
    '--collect-all=pykakasi',
    '--collect-all=imageio',
    '--collect-all=imageio_ffmpeg',
    '--noconfirm',
])
