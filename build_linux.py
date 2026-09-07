"""
Build script for creating AnkiCardGenerator Linux (Ubuntu) binary.
Run: python build_linux.py
Output: dist/AnkiCardGenerator_v{version}_linux   (standalone executable, no extension)

Requirements:
    pip install pyinstaller tkinterdnd2 pykakasi

System packages (Ubuntu) needed for the Tkinter GUI + drag-and-drop:
    sudo apt-get install python3-tk tk-dev

Notes:
    - This builds a single-file executable for the current Linux architecture
      (e.g. x86_64). Build on the same architecture you intend to distribute to.
    - Linux uses ':' as the PyInstaller --add-data separator.
    - Run the result with:  ./dist/AnkiCardGenerator_v{version}_linux
      (you may need: chmod +x dist/AnkiCardGenerator_v{version}_linux)
"""
import PyInstaller.__main__
import tkinterdnd2
import pykakasi
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gen_anki import __version__

tkdnd_path = os.path.dirname(tkinterdnd2.__file__)
pykakasi_path = os.path.dirname(pykakasi.__file__)

app_name = f'AnkiCardGenerator_v{__version__}_linux'

# Linux uses ':' as the data separator
sep = ':'

pyinstaller_args = [
    'gen_anki_gui.py',
    '--onefile',
    '--windowed',                    # GUI app, no console window
    f'--name={app_name}',
    f'--add-data={tkdnd_path}{sep}tkinterdnd2',
    f'--add-data={pykakasi_path}{sep}pykakasi',
    '--hidden-import=tkinterdnd2',
    '--hidden-import=genanki',
    '--hidden-import=gtts',
    '--hidden-import=deep_translator',
    '--hidden-import=openpyxl',
    '--hidden-import=requests',
    '--hidden-import=bs4',
    '--hidden-import=pykakasi',
    '--hidden-import=imageio',
    '--hidden-import=imageio_ffmpeg',
    # Optional TTS / video features (bundled so they work out of the box)
    '--hidden-import=edge_tts',
    '--hidden-import=pyttsx3',
    '--hidden-import=moviepy',
    '--hidden-import=PIL',
    '--collect-all=tkinterdnd2',
    '--collect-all=pykakasi',
    '--collect-all=imageio',
    '--collect-all=imageio_ffmpeg',
    '--collect-all=edge_tts',
    '--collect-all=pyttsx3',
    '--collect-all=moviepy',
    '--collect-all=PIL',
    '--noconfirm',
    '--clean',                       # Clean PyInstaller cache before building
]

# Optional: add an icon if one exists (PNG works for Linux)
icon_path = os.path.join(os.path.dirname(__file__), 'icon.png')
if os.path.exists(icon_path):
    pyinstaller_args.append(f'--icon={icon_path}')

print(f"Building {app_name} for Linux...")
print(f"  tkinterdnd2: {tkdnd_path}")
print(f"  pykakasi:    {pykakasi_path}")
print()

PyInstaller.__main__.run(pyinstaller_args)

print()
print("=" * 60)
print("Build complete!")
print(f"  Binary: dist/{app_name}")
print(f"  Run it: chmod +x dist/{app_name} && ./dist/{app_name}")
print("=" * 60)
