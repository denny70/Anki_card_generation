"""
Anki Card Generator GUI
------------------------
Drag and drop cards_input.xlsx onto the GUI to generate Anki cards with audio.
Settings are saved automatically and restored on next launch.
"""

import json
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

# Import gen_anki functions
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gen_anki import read_and_generate_csv, create_anki_package, set_log_callback, __version__

# Settings file path — save in user home so it persists with exe
SETTINGS_DIR = os.path.join(os.path.expanduser('~'), '.anki_card_generator')
os.makedirs(SETTINGS_DIR, exist_ok=True)
SETTINGS_FILE = os.path.join(SETTINGS_DIR, 'gui_settings.json')


def get_downloads_folder() -> str:
    """Get the Downloads folder path for Windows, macOS, and Linux."""
    import platform
    system = platform.system()
    if system == 'Windows':
        # Use Windows known folder path
        import ctypes
        from ctypes import wintypes
        GUID = ctypes.c_char * 16
        try:
            SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
            # FOLDERID_Downloads = {374DE290-123F-4565-9164-39C4925E467B}
            downloads_guid = ctypes.create_string_buffer(
                b'\x90\xe2\x4d\x37\x3f\x12\x65\x45\x91\x64\x39\xc4\x92\x5e\x46\x7b'
            )
            path_ptr = ctypes.c_wchar_p()
            SHGetKnownFolderPath(downloads_guid, 0, None, ctypes.byref(path_ptr))
            result = path_ptr.value
            ctypes.windll.ole32.CoTaskMemFree(path_ptr)
            if result:
                return result
        except Exception:
            pass
        # Fallback
        return os.path.join(os.path.expanduser('~'), 'Downloads')
    elif system == 'Darwin':
        return os.path.join(os.path.expanduser('~'), 'Downloads')
    else:
        # Linux: check XDG user dirs, fallback to ~/Downloads
        try:
            import subprocess
            result = subprocess.run(
                ['xdg-user-dir', 'DOWNLOAD'],
                capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except Exception:
            pass
        return os.path.join(os.path.expanduser('~'), 'Downloads')


def load_settings() -> dict:
    """Load saved settings from JSON file."""
    defaults = {
        'deck_name': 'Anki_deck',
        'language': 'auto - Auto-detect',
        'reverse': True,
        'last_input_file': '',
    }
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                defaults.update(saved)
    except Exception:
        pass
    return defaults


def save_settings(settings: dict):
    """Save settings to JSON file."""
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class AnkiGeneratorGUI:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title(f"Anki Card Generator v{__version__}")
        self.root.geometry("700x600")
        self.root.resizable(True, True)

        # Load saved settings
        self.settings = load_settings()

        self.input_file = tk.StringVar(value=self.settings.get('last_input_file', ''))
        self.deck_name = tk.StringVar(value=self.settings.get('deck_name', 'Anki_deck'))
        self.output_dir = tk.StringVar(value=self.settings.get('output_dir', get_downloads_folder()))
        self.cache_dir = tk.StringVar(value=self.settings.get('cache_dir', ''))
        self.reverse_var = tk.BooleanVar(value=self.settings.get('reverse', True))
        self.recall_var = tk.BooleanVar(value=self.settings.get('recall', True))
        self.regenerate_var = tk.BooleanVar(value=False)
        self.data_only_var = tk.BooleanVar(value=False)

        self._build_ui()

        # Restore last input file display
        last_file = self.settings.get('last_input_file', '')
        if last_file and os.path.exists(last_file):
            self.drop_label.config(text=f"File loaded:\n{os.path.basename(last_file)}", bg="#e8f5e9")
        else:
            # Clear invalid saved path
            self.input_file.set('')

        # Save settings on close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        # --- Tabbed Interface ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Tab 1: Anki Card Generator
        self.tab_anki = tk.Frame(self.notebook)
        self.notebook.add(self.tab_anki, text="  Anki Generator  ")
        self._build_anki_tab()

        # Tab 2: Text-to-Speech
        self.tab_tts = tk.Frame(self.notebook)
        self.notebook.add(self.tab_tts, text="  Text to Speech  ")
        self._build_tts_tab()

    def _build_anki_tab(self):
        # --- Drop Zone ---
        drop_frame = tk.LabelFrame(self.tab_anki, text="Drag Excel/CSV File Here", padx=10, pady=10)
        drop_frame.pack(fill=tk.X, padx=10, pady=10)

        self.drop_label = tk.Label(
            drop_frame,
            text="Drag and drop Excel/CSV file here\nor click Browse",
            font=("Arial", 14),
            bg="#f0f8ff",
            fg="#333",
            relief="groove",
            width=50,
            height=4,
        )
        self.drop_label.pack(fill=tk.BOTH, expand=True)

        # Enable drag and drop
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self._on_drop)

        # Browse button
        browse_btn = ttk.Button(drop_frame, text="Browse...", command=self._browse_file)
        browse_btn.pack(pady=5)

        # File path display
        file_frame = tk.Frame(self.tab_anki)
        file_frame.pack(fill=tk.X, padx=10)
        tk.Label(file_frame, text="Input File:").pack(side=tk.LEFT)
        tk.Entry(file_frame, textvariable=self.input_file, state='readonly', width=60).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # --- Settings ---
        settings_frame = tk.LabelFrame(self.tab_anki, text="Settings", padx=10, pady=5)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)

        # Deck name
        row1 = tk.Frame(settings_frame)
        row1.pack(fill=tk.X, pady=2)
        tk.Label(row1, text="Deck Name:", width=15, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row1, textvariable=self.deck_name, width=30).pack(side=tk.LEFT)

        # Language
        row2 = tk.Frame(settings_frame)
        row2.pack(fill=tk.X, pady=2)
        tk.Label(row2, text="Language:", width=15, anchor='w').pack(side=tk.LEFT)
        lang_options = [
            "auto - Auto-detect",
            "vi - Vietnamese",
            "en - English",
            "ja - Japanese",
            "ko - Korean",
            "zh-CN - Chinese (Simplified)",
            "zh-TW - Chinese (Traditional)",
            "th - Thai",
            "id - Indonesian",
            "ms - Malay",
            "fr - French",
            "de - German",
            "es - Spanish",
            "pt - Portuguese",
            "it - Italian",
            "ru - Russian",
            "ar - Arabic",
            "hi - Hindi",
        ]
        self.lang_combo = ttk.Combobox(row2, values=lang_options, width=30, state='readonly')
        self.lang_combo.set(self.settings.get('language', 'auto - Auto-detect'))
        self.lang_combo.pack(side=tk.LEFT)

        # Output folder
        row_out = tk.Frame(settings_frame)
        row_out.pack(fill=tk.X, pady=2)
        tk.Label(row_out, text="Output Folder:", width=15, anchor='w').pack(side=tk.LEFT)
        tk.Entry(row_out, textvariable=self.output_dir, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(row_out, text="Browse...", command=self._browse_output).pack(side=tk.LEFT, padx=2)

        # Audio cache folder (default = same as output folder)
        row_cache = tk.Frame(settings_frame)
        row_cache.pack(fill=tk.X, pady=2)
        tk.Label(row_cache, text="Cache Folder:", width=15, anchor='w').pack(side=tk.LEFT)
        self.cache_entry = tk.Entry(row_cache, textvariable=self.cache_dir, width=40)
        self.cache_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(row_cache, text="Browse...", command=self._browse_cache).pack(side=tk.LEFT, padx=2)
        # Hint on a new line below
        row_cache_hint = tk.Frame(settings_frame)
        row_cache_hint.pack(fill=tk.X)
        tk.Label(row_cache_hint, text="                (blank = same as output folder)", fg="#888", font=("Arial", 8)).pack(anchor='w')

        # Checkboxes
        row3 = tk.Frame(settings_frame)
        row3.pack(fill=tk.X, pady=2)
        tk.Checkbutton(row3, text="Generate reverse cards (opposite word as front)", variable=self.reverse_var).pack(anchor='w')

        row5 = tk.Frame(settings_frame)
        row5.pack(fill=tk.X, pady=2)
        tk.Checkbutton(row5, text="Generate recall cards (meaning as front, word as back)", variable=self.recall_var).pack(anchor='w')

        row6 = tk.Frame(settings_frame)
        row6.pack(fill=tk.X, pady=2)
        tk.Checkbutton(row6, text="Regenerate all data (clear existing data, re-fetch from internet)", variable=self.regenerate_var).pack(anchor='w')

        # --- Buttons ---
        btn_frame = tk.Frame(self.tab_anki)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        self.gen_btn = ttk.Button(btn_frame, text="Generate Anki Cards", command=self._generate)
        self.gen_btn.pack(side=tk.LEFT, padx=5)

        self.csv_btn = ttk.Button(btn_frame, text="Generate Data Only", command=self._generate_data_only)
        self.csv_btn.pack(side=tk.LEFT, padx=5)

        self.install_btn = ttk.Button(btn_frame, text="Install Packages", command=self._install_packages)
        self.install_btn.pack(side=tk.LEFT, padx=5)

        # --- Log ---
        log_frame = tk.LabelFrame(self.tab_anki, text="Log", padx=5, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, font=("Consolas", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _build_tts_tab(self):
        """Build the Text-to-Speech tab."""
        # --- Engine + Language Selection ---
        top_frame = tk.Frame(self.tab_tts)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        # TTS Engine
        tk.Label(top_frame, text="Engine:", width=8, anchor='w').pack(side=tk.LEFT)
        tts_engine_options = [
            "gTTS (Google)",
            "Edge-TTS (Microsoft)",
            "pyttsx3 (Offline)",
        ]
        self.tts_engine_combo = ttk.Combobox(top_frame, values=tts_engine_options, width=20, state='readonly')
        self.tts_engine_combo.set(self.settings.get('tts_engine', 'gTTS (Google)'))
        self.tts_engine_combo.pack(side=tk.LEFT, padx=5)

        # Language
        lang_frame = tk.Frame(self.tab_tts)
        lang_frame.pack(fill=tk.X, padx=10, pady=2)

        tk.Label(lang_frame, text="Language:", width=10, anchor='w').pack(side=tk.LEFT)
        tts_lang_options = [
            "auto - Auto-detect",
            "vi - Vietnamese",
            "en - English",
            "ja - Japanese",
            "ko - Korean",
            "zh-CN - Chinese (Simplified)",
            "zh-TW - Chinese (Traditional)",
            "th - Thai",
            "id - Indonesian",
            "fr - French",
            "de - German",
            "es - Spanish",
            "pt - Portuguese",
            "it - Italian",
            "ru - Russian",
        ]
        self.tts_lang_combo = ttk.Combobox(lang_frame, values=tts_lang_options, width=25, state='readonly')
        self.tts_lang_combo.set(self.settings.get('tts_language', 'auto - Auto-detect'))
        self.tts_lang_combo.pack(side=tk.LEFT, padx=5)

        # Output folder
        tk.Label(lang_frame, text="Output:").pack(side=tk.LEFT, padx=(10, 0))
        self.tts_output_dir = tk.StringVar(value=self.settings.get('tts_output_dir', get_downloads_folder()))
        tk.Entry(lang_frame, textvariable=self.tts_output_dir, width=25).pack(side=tk.LEFT, padx=2)
        ttk.Button(lang_frame, text="...", width=3, command=self._browse_tts_output).pack(side=tk.LEFT)

        # --- Text Editor ---
        editor_frame = tk.LabelFrame(self.tab_tts, text="Paste text here (sentences, words, paragraphs)", padx=5, pady=5)
        editor_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tts_text = tk.Text(editor_frame, wrap=tk.WORD, font=("Arial", 14), height=12)
        self.tts_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tts_text, command=self.tts_text.yview)
        self.tts_text.configure(yscrollcommand=scrollbar.set)

        # --- Buttons ---
        btn_frame = tk.Frame(self.tab_tts)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(btn_frame, text="Generate Audio", command=self._generate_tts).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Generate Video", command=self._generate_video).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Clear", command=lambda: self.tts_text.delete('1.0', tk.END)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Install Packages", command=self._install_tts_packages).pack(side=tk.LEFT, padx=5)

        # Note about video engine
        note_frame = tk.Frame(self.tab_tts)
        note_frame.pack(fill=tk.X, padx=10)
        tk.Label(note_frame, text="Note: Video always uses Edge-TTS (required for word-by-word timing sync)",
                 fg="#888", font=("Arial", 9)).pack(anchor='w')

        # --- TTS Log ---
        tts_log_frame = tk.LabelFrame(self.tab_tts, text="Log", padx=5, pady=5)
        tts_log_frame.pack(fill=tk.X, padx=10, pady=5)

        self.tts_log = tk.Text(tts_log_frame, height=4, wrap=tk.WORD, font=("Consolas", 10))
        self.tts_log.pack(fill=tk.BOTH, expand=True)

    def _browse_tts_output(self):
        """Open folder dialog for TTS output."""
        folder = filedialog.askdirectory(
            title="Select TTS Output Folder",
            initialdir=self.tts_output_dir.get()
        )
        if folder:
            self.tts_output_dir.set(folder)

    def _install_tts_packages(self):
        """Install TTS-specific packages."""
        self.tts_log.insert(tk.END, "Installing TTS packages...\n")
        self.tts_log.see(tk.END)
        thread = threading.Thread(target=self._do_install_tts, daemon=True)
        thread.start()

    def _do_install_tts(self):
        """Background TTS package installation."""
        import subprocess
        import importlib

        packages = {
            'gTTS': 'gtts',
            'edge-tts': 'edge_tts',
            'pyttsx3': 'pyttsx3',
            'moviepy': 'moviepy',
            'Pillow': 'PIL',
        }
        try:
            for pip_name, import_name in packages.items():
                try:
                    importlib.import_module(import_name)
                    self.tts_log.insert(tk.END, f"  {pip_name} - already installed, skipping.\n")
                    self.tts_log.see(tk.END)
                    continue
                except ImportError:
                    pass

                self.tts_log.insert(tk.END, f"  Installing {pip_name}...\n")
                self.tts_log.see(tk.END)
                result = subprocess.run(
                    [sys.executable, '-m', 'pip', 'install', pip_name],
                    capture_output=True, text=True, timeout=120
                )
                if result.returncode == 0:
                    self.tts_log.insert(tk.END, f"    {pip_name} OK\n")
                else:
                    self.tts_log.insert(tk.END, f"    {pip_name} FAILED: {result.stderr.strip()}\n")
                self.tts_log.see(tk.END)

            self.tts_log.insert(tk.END, "All TTS packages installed!\n")
            self.tts_log.see(tk.END)
        except Exception as e:
            self.tts_log.insert(tk.END, f"ERROR: {str(e)}\n")
            self.tts_log.see(tk.END)

    def _generate_tts(self):
        """Generate audio from text in the editor."""
        text = self.tts_text.get('1.0', tk.END).strip()
        if not text:
            messagebox.showwarning("No Text", "Please paste or type text first.")
            return

        # Get language
        lang = self.tts_lang_combo.get().split(' - ')[0].strip()

        # Auto-detect if needed
        if lang == 'auto':
            from gen_anki import detect_language
            lang = detect_language(text)
            self.tts_log.insert(tk.END, f"[Auto-detect] Language: {lang}\n")

        # Generate in background
        self.tts_log.insert(tk.END, f"Generating audio ({lang})...\n")
        self.tts_log.see(tk.END)

        thread = threading.Thread(target=self._do_tts, args=(text, lang), daemon=True)
        thread.start()

    def _do_tts(self, text, lang):
        """Background TTS generation."""
        try:
            import time
            import re

            output_dir = self.tts_output_dir.get()
            os.makedirs(output_dir, exist_ok=True)

            # Generate filename from first few words + timestamp
            clean_text = re.sub(r'[^\w\s]', '', text[:30]).strip()
            clean_text = re.sub(r'\s+', '_', clean_text)
            timestamp = time.strftime('%Y%m%d_%H%M%S')
            filename = f"tts_{clean_text}_{timestamp}.mp3"
            output_path = os.path.join(output_dir, filename)

            engine = self.tts_engine_combo.get()

            if engine == 'Edge-TTS (Microsoft)':
                import asyncio
                import edge_tts
                # Map language to Edge-TTS voice
                voice_map = {
                    'vi': 'vi-VN-HoaiMyNeural', 'en': 'en-US-JennyNeural',
                    'ja': 'ja-JP-NanamiNeural', 'ko': 'ko-KR-SunHiNeural',
                    'zh-CN': 'zh-CN-XiaoxiaoNeural', 'zh-TW': 'zh-TW-HsiaoChenNeural',
                    'th': 'th-TH-PremwadeeNeural', 'id': 'id-ID-GadisNeural',
                    'fr': 'fr-FR-DeniseNeural', 'de': 'de-DE-KatjaNeural',
                    'es': 'es-ES-ElviraNeural', 'pt': 'pt-BR-FranciscaNeural',
                    'it': 'it-IT-ElsaNeural', 'ru': 'ru-RU-SvetlanaNeural',
                }
                voice = voice_map.get(lang, 'en-US-JennyNeural')
                communicate = edge_tts.Communicate(text, voice)
                asyncio.run(communicate.save(output_path))

            elif engine == 'pyttsx3 (Offline)':
                import pyttsx3
                tts_engine = pyttsx3.init()
                tts_engine.save_to_file(text, output_path)
                tts_engine.runAndWait()

            else:
                # Default: gTTS (Google)
                from gtts import gTTS
                tts = gTTS(text=text, lang=lang)
                tts.save(output_path)

            self.tts_log.insert(tk.END, f"Saved: {output_path}\n")
            self.tts_log.see(tk.END)
            self.root.after(0, lambda: messagebox.showinfo("Done", f"Audio saved:\n{filename}"))

        except Exception as e:
            self.tts_log.insert(tk.END, f"ERROR: {str(e)}\n")
            self.tts_log.see(tk.END)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def _generate_video(self):
        """Generate karaoke-style video with highlighted text synced to audio."""
        text = self.tts_text.get('1.0', tk.END).strip()
        if not text:
            messagebox.showwarning("No Text", "Please paste or type text first.")
            return

        # Get language
        lang = self.tts_lang_combo.get().split(' - ')[0].strip()
        if lang == 'auto':
            from gen_anki import detect_language
            lang = detect_language(text)
            self.tts_log.insert(tk.END, f"[Auto-detect] Language: {lang}\n")

        self.tts_log.insert(tk.END, f"Generating video ({lang})... This may take a moment.\n")
        self.tts_log.see(tk.END)

        thread = threading.Thread(target=self._do_video, args=(text, lang), daemon=True)
        thread.start()

    def _do_video(self, text, lang):
        """Background video generation with word-by-word highlighting."""
        try:
            import asyncio
            import edge_tts
            import time
            import re
            import numpy as np
            from PIL import Image, ImageDraw, ImageFont
            from moviepy import AudioFileClip, VideoClip

            output_dir = self.tts_output_dir.get()
            os.makedirs(output_dir, exist_ok=True)

            # Generate filename
            clean_text = re.sub(r'[^\w\s]', '', text[:30]).strip()
            clean_text = re.sub(r'\s+', '_', clean_text)
            timestamp = time.strftime('%Y%m%d_%H%M%S')

            audio_path = os.path.join(output_dir, f"_temp_video_audio_{timestamp}.mp3")
            video_path = os.path.join(output_dir, f"video_{clean_text}_{timestamp}.mp4")

            # Edge-TTS voice map
            voice_map = {
                'vi': 'vi-VN-HoaiMyNeural', 'en': 'en-US-JennyNeural',
                'ja': 'ja-JP-NanamiNeural', 'ko': 'ko-KR-SunHiNeural',
                'zh-CN': 'zh-CN-XiaoxiaoNeural', 'zh-TW': 'zh-TW-HsiaoChenNeural',
                'th': 'th-TH-PremwadeeNeural', 'id': 'id-ID-GadisNeural',
                'fr': 'fr-FR-DeniseNeural', 'de': 'de-DE-KatjaNeural',
                'es': 'es-ES-ElviraNeural', 'pt': 'pt-BR-FranciscaNeural',
                'it': 'it-IT-ElsaNeural', 'ru': 'ru-RU-SvetlanaNeural',
            }
            voice = voice_map.get(lang, 'en-US-JennyNeural')

            # Step 1: Get audio + word timestamps from Edge-TTS
            self.tts_log.insert(tk.END, "  Generating audio with timestamps...\n")
            self.tts_log.see(tk.END)

            word_timings = []

            async def get_tts_with_timestamps():
                communicate = edge_tts.Communicate(text, voice, boundary='WordBoundary')
                with open(audio_path, 'wb') as audio_file:
                    async for chunk in communicate.stream():
                        if chunk['type'] == 'audio':
                            audio_file.write(chunk['data'])
                        elif chunk['type'] == 'WordBoundary':
                            word_timings.append({
                                'text': chunk['text'],
                                'offset': chunk['offset'] / 10000000.0,
                                'duration': chunk['duration'] / 10000000.0,
                            })

            asyncio.run(get_tts_with_timestamps())

            if not word_timings:
                self.tts_log.insert(tk.END, "  ERROR: No word timestamps received.\n")
                self.tts_log.see(tk.END)
                return

            self.tts_log.insert(tk.END, f"  Got {len(word_timings)} word timestamps.\n")
            self.tts_log.see(tk.END)

            # Step 2: Create video frames with word highlighting
            self.tts_log.insert(tk.END, "  Creating video frames...\n")
            self.tts_log.see(tk.END)

            # Video settings
            width, height = 1280, 720
            bg_color = (30, 30, 30)  # Dark background
            text_color = (255, 255, 255)  # White text
            highlight_color = (255, 200, 0)  # Yellow highlight
            font_size = 48

            # Try to load a font that supports CJK/Vietnamese
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except Exception:
                try:
                    font = ImageFont.truetype("C:/Windows/Fonts/msgothic.ttc", font_size)
                except Exception:
                    font = ImageFont.load_default()

            # Get audio duration
            audio_clip = AudioFileClip(audio_path)
            total_duration = audio_clip.duration

            # Build word list with positions
            words = [w['text'] for w in word_timings]

            def make_frame(t):
                """Generate a frame at time t with highlighted current word."""
                img = Image.new('RGB', (width, height), bg_color)
                draw = ImageDraw.Draw(img)

                # Find which word is currently being spoken
                current_word_idx = -1
                for i, wt in enumerate(word_timings):
                    if t >= wt['offset']:
                        current_word_idx = i
                    else:
                        break

                # Build layout: calculate positions for all words first
                lines = text.split('\n')
                line_height = font_size + 24
                layout = []  # [(x, y, word_text, word_idx)]
                y = 0
                word_idx = 0

                for line in lines:
                    if not line.strip():
                        y += line_height // 2
                        continue

                    x = 60
                    line_words = line.split()

                    for lw in line_words:
                        bbox = draw.textbbox((0, 0), lw + ' ', font=font)
                        word_width = bbox[2] - bbox[0]

                        if x + word_width > width - 60:
                            x = 60
                            y += line_height

                        layout.append((x, y, lw, word_idx))
                        x += word_width
                        word_idx += 1

                    y += line_height

                # Calculate scroll offset to keep current word visible
                if current_word_idx >= 0 and current_word_idx < len(layout):
                    current_y = layout[current_word_idx][1]
                    # Keep current word in the middle area of screen
                    visible_height = height - 120
                    if current_y > visible_height // 2:
                        scroll_offset = current_y - visible_height // 2
                    else:
                        scroll_offset = 0
                else:
                    scroll_offset = 0

                # Draw words with scroll offset
                for x, y_pos, word_text, w_idx in layout:
                    draw_y = y_pos - scroll_offset + 60

                    # Skip if outside visible area
                    if draw_y < -font_size or draw_y > height:
                        continue

                    # Color
                    if w_idx == current_word_idx:
                        color = highlight_color
                    elif w_idx < current_word_idx:
                        color = (180, 180, 180)
                    else:
                        color = text_color

                    draw.text((x, draw_y), word_text + ' ', font=font, fill=color)

                return np.array(img)

            # Step 3: Create video
            self.tts_log.insert(tk.END, "  Rendering video...\n")
            self.tts_log.see(tk.END)

            video_clip = VideoClip(make_frame, duration=total_duration)
            video_clip = video_clip.with_audio(audio_clip)
            video_clip.write_videofile(video_path, fps=24, codec='libx264',
                                       audio_codec='aac', logger=None,
                                       threads=4, preset='ultrafast')

            # Cleanup temp audio
            if os.path.exists(audio_path):
                os.remove(audio_path)

            self.tts_log.insert(tk.END, f"  Video saved: {video_path}\n")
            self.tts_log.see(tk.END)
            self.root.after(0, lambda: messagebox.showinfo("Done", f"Video saved:\n{os.path.basename(video_path)}"))

        except Exception as e:
            self.tts_log.insert(tk.END, f"ERROR: {str(e)}\n")
            self.tts_log.see(tk.END)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def _on_drop(self, event):
        """Handle drag and drop of file."""
        file_path = event.data
        # Remove curly braces if present (Windows drag-drop format)
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        file_path = file_path.strip()

        if file_path.lower().endswith('.csv') or file_path.lower().endswith('.xlsx'):
            self.input_file.set(file_path)
            self.output_dir.set(os.path.dirname(file_path))
            self.drop_label.config(text=f"File loaded:\n{os.path.basename(file_path)}", bg="#e8f5e9")
            self._log(f"File loaded: {file_path}")
        else:
            messagebox.showwarning("Invalid File", "Please drop a .xlsx or .csv file")

    def _browse_file(self):
        """Open file dialog to select CSV."""
        file_path = filedialog.askopenfilename(
            title="Select cards_input.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=os.path.dirname(os.path.abspath(__file__))
        )
        if file_path:
            self.input_file.set(file_path)
            self.output_dir.set(os.path.dirname(file_path))
            self.drop_label.config(text=f"File loaded:\n{os.path.basename(file_path)}", bg="#e8f5e9")
            self._log(f"File loaded: {file_path}")

    def _browse_output(self):
        """Open folder dialog to select output directory."""
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_dir.get()
        )
        if folder:
            self.output_dir.set(folder)
            self.root.update_idletasks()
            self._log(f"Output folder: {folder}")

    def _browse_cache(self):
        """Open folder dialog to select audio cache directory."""
        initial = self.cache_dir.get() or self.output_dir.get()
        folder = filedialog.askdirectory(
            title="Select Audio Cache Folder",
            initialdir=initial
        )
        if folder:
            self.cache_dir.set(folder)
            self.cache_entry.icursor(tk.END)
            self.cache_entry.xview_moveto(1.0)
            self.root.update_idletasks()
            self._log(f"Cache folder: {folder}")

    def _log(self, message):
        """Append message to log."""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def _generate(self):
        """Generate both xlsx and Anki package."""
        self.data_only_var.set(False)
        self._run_generation()

    def _generate_data_only(self):
        """Generate xlsx only."""
        self.data_only_var.set(True)
        self._run_generation()

    def _install_packages(self):
        """Install required Python packages."""
        self.install_btn.config(state='disabled')
        self._log("---")
        self._log("Installing required packages...")
        thread = threading.Thread(target=self._do_install, daemon=True)
        thread.start()

    def _do_install(self):
        """Background package installation."""
        import subprocess
        import importlib

        # Map pip package name -> import name
        packages = {
            'genanki': 'genanki',
            'gTTS': 'gtts',
            'deep-translator': 'deep_translator',
            'openpyxl': 'openpyxl',
            'tkinterdnd2': 'tkinterdnd2',
            'pykakasi': 'pykakasi',
            'edge-tts': 'edge_tts',
            'pyttsx3': 'pyttsx3',
        }
        try:
            for pip_name, import_name in packages.items():
                # Check if already installed
                try:
                    importlib.import_module(import_name)
                    self._log(f"  {pip_name} - already installed, skipping.")
                    continue
                except ImportError:
                    pass

                self._log(f"  Installing {pip_name}...")
                result = subprocess.run(
                    [sys.executable, '-m', 'pip', 'install', pip_name],
                    capture_output=True, text=True, timeout=120
                )
                if result.returncode == 0:
                    self._log(f"    {pip_name} OK")
                else:
                    self._log(f"    {pip_name} FAILED: {result.stderr.strip()}")

            self._log("All packages installed!")
            self.root.after(0, lambda: messagebox.showinfo("Success", "All packages installed successfully!"))
        except Exception as e:
            self._log(f"ERROR: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: self.install_btn.config(state='normal'))

    def _run_generation(self):
        """Run the generation in a background thread."""
        input_path = self.input_file.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid CSV input file")
            return

        self.gen_btn.config(state='disabled')
        self.csv_btn.config(state='disabled')
        self._log("---")
        self._log("Starting generation...")

        thread = threading.Thread(target=self._do_generate, daemon=True)
        thread.start()

    def _do_generate(self):
        """Background generation task."""
        # Set log callback so gen_anki messages appear in GUI
        set_log_callback(self._log)
        try:
            input_path = self.input_file.get()
            output_dir = self.output_dir.get()
            os.makedirs(output_dir, exist_ok=True)
            input_basename = os.path.splitext(os.path.basename(input_path))[0]
            generated_csv_path = os.path.join(output_dir, f"{input_basename}_gen.xlsx")

            # Step 1: Generate CSV
            self._log(f"Reading: {input_path}")
            lang = self.lang_combo.get().split(' - ')[0].strip()
            cards = read_and_generate_csv(input_path, generated_csv_path, self.reverse_var.get(), lang, self.regenerate_var.get())
            self._log(f"Total cards: {len(cards)}")

            # Step 2: Generate Anki package (unless data-only)
            if not self.data_only_var.get():
                self._log("Generating audio and Anki package...")
                # Extract language code from combobox (e.g. "vi - Vietnamese" -> "vi")
                lang = self.lang_combo.get().split(' - ')[0].strip()
                # Cache folder: use setting, or default to output folder
                cache_dir = self.cache_dir.get().strip() or output_dir
                result = create_anki_package(
                    cards, self.deck_name.get(), lang, output_dir, self.recall_var.get(), cache_dir
                )
                self._log(f"Anki package created: {result}")
            else:
                self._log("Data only mode - skipping Anki package")

            self._log("Done!")
            self._save_settings()
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Generated {len(cards)} cards!"))

        except Exception as e:
            self._log(f"ERROR: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

        finally:
            self.root.after(0, self._enable_buttons)

    def _enable_buttons(self):
        self.gen_btn.config(state='normal')
        self.csv_btn.config(state='normal')

    def _save_settings(self):
        """Save current GUI settings to file."""
        settings = {
            'deck_name': self.deck_name.get(),
            'language': self.lang_combo.get(),
            'reverse': self.reverse_var.get(),
            'recall': self.recall_var.get(),
            'last_input_file': self.input_file.get(),
            'output_dir': self.output_dir.get(),
            'cache_dir': self.cache_dir.get(),
            'tts_engine': self.tts_engine_combo.get(),
            'tts_language': self.tts_lang_combo.get(),
            'tts_output_dir': self.tts_output_dir.get(),
        }
        save_settings(settings)

    def _on_close(self):
        """Save settings and close."""
        self._save_settings()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = AnkiGeneratorGUI()
    app.run()
