"""
Build script for creating AnkiCardGenerator macOS .app bundle.
Run: python build_mac.py
Output: dist/AnkiCardGenerator_v{version}_mac.app  (double-click to run)
        dist/AnkiCardGenerator_v{version}_mac       (standalone CLI binary)

Requirements:
    pip install pyinstaller tkinterdnd2 pykakasi

Notes:
    - This creates a macOS .app bundle (--windowed) suitable for double-clicking.
    - The original build_exe.py is configured for Windows; this file is for macOS.
    - On Apple Silicon (M1/M2/M3), the binary is built for the current architecture.
    - To create a universal binary, install dependencies for both archs and use
      --target-architecture universal2 (requires both arm64 and x86_64 Python).
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

app_name = f'AnkiCardGenerator_v{__version__}_mac'

# macOS uses ':' as the data separator
sep = ':'

pyinstaller_args = [
    'gen_anki_gui.py',
    '--onefile',
    '--windowed',                    # Creates .app bundle on macOS
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
    '--collect-all=tkinterdnd2',
    '--collect-all=pykakasi',
    '--noconfirm',
    '--clean',                       # Clean PyInstaller cache before building
]

# Optional: add an icon if one exists
icon_path = os.path.join(os.path.dirname(__file__), 'icon.icns')
if os.path.exists(icon_path):
    pyinstaller_args.append(f'--icon={icon_path}')

print(f"Building {app_name} for macOS...")
print(f"  tkinterdnd2: {tkdnd_path}")
print(f"  pykakasi:    {pykakasi_path}")
print()

PyInstaller.__main__.run(pyinstaller_args)

print()
print("=" * 60)
print(f"Build complete!")
print(f"  App bundle: dist/{app_name}.app")
print(f"  Binary:     dist/{app_name}")
print("=" * 60)
