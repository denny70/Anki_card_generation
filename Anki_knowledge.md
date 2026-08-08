# Anki Card Generator — Knowledge Base

## Overview

Generates Anki flashcard packages (.apkg) with embedded TTS audio from an Excel/CSV input file.
Two main components: `gen_anki.py` (engine) and `gen_anki_gui.py` (GUI with 2 tabs: Anki Generator + TTS/Video).

## Version

Current: v2.2.6

## Excel Input Format

| Column | Required | Description |
|--------|----------|-------------|
| `front_card` | Yes | The word to learn (Vietnamese, Japanese, etc.) |
| `front_card_reading` | No | Hiragana reading for Japanese (auto-generated if empty) |
| `back_card_meaning` | No | English meaning (auto-filled via Google Translate) |
| `back_card_sentence` | No | Example sentence in source language (auto-filled via Tatoeba) |
| `back_card_sentence_meaning` | No | English translation of sentence (auto-filled) |
| `back_card_opposite` | No | Opposite/antonym word (auto-filled via Free Dictionary API) |
| `back_card_opposite_meaning` | No | English meaning of opposite word (auto-filled) |
| `transfer_language` | No | Target language for meaning translation (default: English). E.g. "vn" → translate to Vietnamese instead of English. **[PENDING IMPLEMENTATION]** |

## Key Features

### Audio Cache
- Cache folder: `<cache_dir>/anki_audio_cache/{lang}_{md5_of_word}.mp3`
- Same word = reuse cached audio (skip gTTS download)
- Default cache location: same as output folder
- GUI has "Cache Folder" field with Browse button

### N/A Logic
- If any field can't be found after internet lookup → fill "N/A"
- Next run: sees "N/A" as filled → skips internet lookup entirely (fast)
- "N/A" is case-insensitive: "N/A", "n/a", " N/A " all treated the same
- On Anki card: "N/A" fields are treated as empty (won't display)

### Skip Logic (Per Row)
1. If ALL fields are filled (including "N/A") → skip internet lookup entirely
2. Auto-detect language only runs when fields are missing
3. Only search internet for fields that are actually empty

### Sentence Source (Internet Only)
- Tatoeba (real example sentences from native speakers) — priority 1
- Free Dictionary API (English example translated to source language) — priority 2
- No code-generated template sentences (removed)
- If nothing found → "N/A"

### Opposite Word Lookup
- Uses the **English meaning** (back_card_meaning) to find antonyms
- Looks up in Free Dictionary API
- If meaning is multi-word (e.g. "clouds, rain"), uses first word only
- Translates English antonym back to source language

### Japanese Support
- Auto-generates hiragana reading via pykakasi
- Furigana (ruby tags) on front word and sentence
- Works in "auto" language mode (detects Japanese from kanji/hiragana/katakana)

### Regenerate Mode
- GUI checkbox: "Regenerate all data"
- Clears all existing data (except front_card) and re-fetches from internet
- Useful when you want to refresh all translations

### Output Excel (_gen.xlsx)
- Saved to Output Folder
- Copies column widths from input
- Freeze panes: B2 (row 1 + column A frozen)
- Follows input header order

### Anki Card Templates

**Recognition Card:**
- Front: Word (with furigana) + Audio
- Back: Meaning + Sentence + Sentence Audio + Sentence Meaning + Opposite + Opposite Audio + Opposite Meaning

**Recall Card:**
- Front: Meaning + Front Audio
- Back: Word (with furigana) + Sentence + Sentence Audio + Sentence Meaning + Opposite + Opposite Audio + Opposite Meaning

**Reverse Card (from opposite word):**
- Same structure as Recognition, but with opposite word as front

### Deterministic IDs
- Same deck name → same deck ID (Anki updates existing deck)
- Same front word → same note ID (preserves review history on re-import)
- Only new cards are added when re-importing

## GUI Settings

- Saved in `~/.anki_card_generator/gui_settings.json`
- Persists: deck name, language, output folder, cache folder, reverse/recall checkboxes, TTS settings
- Output folder defaults to source Excel file's directory when file is loaded

## Build

```
python build_exe.py
```

- Output: `dist/AnkiCardGenerator_v{version}_{platform}.exe`
- Cleans previous .spec files before building
- Bundles: tkinterdnd2, pykakasi, imageio, imageio_ffmpeg

## Data Sources

| Source | Used For | Notes |
|--------|----------|-------|
| Google Translate | Meaning, sentence translation | Works for all languages |
| Tatoeba | Example sentences | Real sentences from native speakers, 400+ languages |
| Free Dictionary API | Antonyms, English examples | English words only, no API key needed |

## Pending Features

- [ ] `transfer_language` column: override target language for meaning translation per word (e.g. "vn" → translate to Vietnamese instead of English)
