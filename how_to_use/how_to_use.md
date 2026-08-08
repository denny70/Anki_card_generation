# Anki Card Generator - How to Use (v1.7)

## Overview

This tool generates Anki flashcard packages (.apkg) with embedded TTS audio from an Excel file. No internet is needed when reviewing cards in Anki because the audio files are baked directly into the package.

Internet is required during generation to auto-translate and fill in card data.

## Files

| File | Purpose |
|------|---------|
| `gen_anki.py` | Command-line card generator (Engine) |
| `gen_anki_gui.py` | GUI version (drag and drop) |
| `AnkiCardGenerator_v*_win.exe` | Standalone GUI for Windows (no Python needed) |
| `build_exe.py` | Build script for creating the exe |
| `cards_input.xlsx` | Your input file template |
| `how_to_use.md` | This documentation |
| `CHANGELOG.md` | Version history |

## Quick Start

### Option 1: GUI (Drag and Drop)

```
python gen_anki_gui.py
```
Or double-click `AnkiCardGenerator_v*_win.exe`

1. Drag your Excel/CSV file onto the GUI window (or click Browse)
2. Set deck name and language (dropdown)
3. Click "Generate Anki Cards" or "Generate Data Only"
4. Import the `.apkg` file into Anki

### Option 2: Command Line

```
python gen_anki.py --input cards_input.xlsx --lang vi --deck "VN_class"
```

## Input File Format (.xlsx or .csv)

Excel file with a header row. Columns are matched by **header name** (not position) — you can reorder columns freely.

**Minimum requirement: only `front_card` is needed.** All other fields are optional — the script auto-fills them from the internet.

| Column | Description | Required | Auto-filled? |
|--------|-------------|----------|--------------|
| `front_card` | Word to learn (shown on front of card) | Yes | No |
| `back_card_meaning` | English meaning | No | Yes (Google Translate) |
| `back_card_sentence` | Example sentence in source language | No | Yes (Tatoeba) |
| `back_card_sentence_meaning` | English translation of the sentence | No | Yes |
| `back_card_opposite` | Opposite/antonym word | No | Yes (adjectives only) |
| `back_card_opposite_meaning` | English meaning of opposite word | No | Yes |

### Example Input

| front_card | back_card_meaning | back_card_sentence | back_card_sentence_meaning | back_card_opposite | back_card_opposite_meaning |
|---|---|---|---|---|---|
| dày | | | | | |
| nhanh | | | | | |
| bàn | | | | | |

### Rules

- If a field already has data → script keeps it (no internet call for that field)
- If all fields are filled → script skips that row entirely (fast)
- Duplicate front words → automatically removed (first occurrence kept)
- `back_card_opposite` only auto-filled for adjectives with clear antonyms
- Example sentences prefer 8-12 words, fallback to 6-8, then 4-6
- Example sentences must contain the front word
- Columns matched by header name — reorder freely

## Output Files

| File | Description |
|------|-------------|
| `<input_name>_gen.xlsx` | Filled data for review (e.g., `VN_class_0711_gen.xlsx`) |
| `<deck_name>.apkg` | Anki package with embedded audio |

## Incremental Updates (Re-importing)

When you add new words and re-generate:
- Same front word → same note ID → Anki skips it (preserves review history)
- New words → new cards added to the deck
- **No duplicates** on re-import

**Important:** Use the same deck name each time. If you changed deck names before and have duplicates, delete the old deck in Anki and re-import fresh.

## Data Sources

The script uses multiple free online sources (no API keys needed):

| Source | What it provides | Priority |
|--------|-----------------|----------|
| **Tatoeba** | Real example sentences + translations (400+ languages) | 1st |
| **Free Dictionary API** | Antonyms, fallback sentences | 2nd |
| **Google Translate** | Word meanings, sentence translations | 3rd |

Only the sources needed for empty fields are called — no wasted internet requests.

## Card Layout

### Front Side
- Word in source language (large, 40px)
- Audio play button (centered)

### Back Side
- English meaning (grey, 24px)
- Example sentence (blue, bold, 28px) + audio
- Sentence translation (grey, 18px)
- Opposite word (pink/red, 30px) + audio
- Opposite meaning (grey, 18px, below audio)

## Supported Languages

| Code | Language |
|------|----------|
| vi | Vietnamese |
| en | English |
| ja | Japanese |
| ko | Korean |
| zh-CN | Chinese (Simplified) |
| zh-TW | Chinese (Traditional) |
| th | Thai |
| id | Indonesian |
| ms | Malay |
| fr | French |
| de | German |
| es | Spanish |
| pt | Portuguese |
| it | Italian |
| ru | Russian |
| ar | Arabic |
| hi | Hindi |

## Command Line Options

| Option | Default | Description |
|--------|---------|-------------|
| `--input` | (required) | Path to input .xlsx or .csv file |
| `--lang` | `vi` | TTS language code |
| `--deck` | `Anki_deck` | Anki deck name |
| `--output` | Downloads folder | Output directory |
| `--reverse` | yes (default) | Generate reverse cards (opposite word becomes front) |
| `--no-reverse` | | Do not generate reverse cards |
| `--data-only` | | Only generate _gen.xlsx, no .apkg file |

## Examples

### Generate cards (default, with reverse)
```
python gen_anki.py --input VN_class_0711.xlsx --lang vi --deck "VN_class"
```

### Generate data only (review before making Anki package)
```
python gen_anki.py --input cards_input.xlsx --lang vi --data-only
```

### Generate without reverse cards
```
python gen_anki.py --input cards_input.xlsx --lang ko --deck "Korean" --no-reverse
```

## Workflow

1. Edit input `.xlsx` — add front words (and any known data)
2. Run with `--data-only` to preview the generated data
3. Review `*_gen.xlsx` in Excel
4. Run without `--data-only` to generate the Anki package
5. Double-click the `.apkg` file to import into Anki
6. When adding new words: just add to input file, regenerate, and re-import (only new cards added)

## GUI Features

- Drag and drop `.xlsx` or `.csv` files
- Language dropdown (17 languages)
- Output folder selector (default: Downloads)
- Reverse cards toggle
- Real-time log showing translation progress
- Settings memory (persists across restarts, even for exe)
- Install Packages button (checks if installed, skips if present)
- Version display in title bar

## Building the Exe

```
python build_exe.py
```

Output: `dist/AnkiCardGenerator_v{version}_{platform}.exe`

Platform is auto-detected (win/mac/linux). Build on each platform separately.

## Dependencies

Click "Install Packages" in the GUI, or manually:
```
pip install genanki gTTS deep-translator openpyxl tkinterdnd2
```

| Package | Purpose |
|---------|---------|
| `genanki` | Creates Anki packages |
| `gTTS` | Google Text-to-Speech (generates audio) |
| `deep-translator` | Google Translate for meanings |
| `openpyxl` | Read/write Excel .xlsx files |
| `tkinterdnd2` | Drag and drop support (GUI only) |

## Notes

- Internet required during generation (translation + audio)
- No internet needed when reviewing in Anki (audio embedded)
- All text normalized to Unicode NFC
- Use `.xlsx` format to avoid character encoding issues
- Output defaults to Downloads folder (cross-platform)
- No hardcoded dictionaries — all data from online sources
- Settings stored at `~/.anki_card_generator/gui_settings.json`
