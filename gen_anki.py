"""
Anki Card Generator with Embedded Audio
----------------------------------------
Generates .apkg files with TTS audio embedded directly in the card.
Audio files are included inside the Anki package so no internet is needed during review.

Workflow:
    1. You put front_card words (and any known data) in cards_input.xlsx
       (minimum: only front_card is required, all other fields are optional)
    2. Script reads cards_input.xlsx, fills in empty fields, writes cards_generated.xlsx
    3. Script generates .apkg from cards_generated.xlsx

    cards_input.xlsx     = your original input (never modified by script)
    cards_generated.xlsx = auto-filled version used for Anki generation

Card layout:
    Front: Vietnamese word + audio
    Back:  English meaning + example sentence + sentence meaning + opposite word

Usage:
    python gen_anki.py --input cards_input.xlsx --lang vi --deck "Vietnamese"
    python gen_anki.py --input cards_input.xlsx --lang vi --deck "Vietnamese" --no-reverse
    python gen_anki.py --input cards_input.xlsx --lang vi --deck "Vietnamese" --data-only

CSV format (5 columns, first row is header):
    front_card,back_card_meaning,back_card_sentence,back_card_sentence_meaning,back_card_opposite
    dày,thick,Cuốn sách này rất dày,This book is very thick,mỏng
    xin chào,,,,
"""

__version__ = '2.2.6'

import argparse
import csv
import os
import random
import shutil
import sys
import unicodedata
from pathlib import Path

import genanki
from gtts import gTTS
from deep_translator import GoogleTranslator
from openpyxl import Workbook, load_workbook

# Japanese furigana support (optional, only loaded for Japanese)
_kakasi_instance = None

def _get_kakasi():
    """Lazy-load pykakasi for Japanese reading generation."""
    global _kakasi_instance
    if _kakasi_instance is None:
        try:
            from pykakasi import kakasi
            _kakasi_instance = kakasi()
        except ImportError:
            _kakasi_instance = False  # Mark as unavailable
    return _kakasi_instance


def generate_furigana_html(word: str) -> str:
    """Generate HTML with ruby tags for furigana (reading above kanji)."""
    kks = _get_kakasi()
    if not kks:
        return word

    result = kks.convert(word)
    html_parts = []
    for item in result:
        orig = item['orig']
        hira = item['hira']
        # Only add furigana if the original contains kanji (differs from hiragana)
        if orig != hira and any('\u4e00' <= c <= '\u9fff' for c in orig):
            html_parts.append(f'<ruby>{orig}<rt>{hira}</rt></ruby>')
        else:
            html_parts.append(orig)
    return ''.join(html_parts)


def get_reading(word: str) -> str:
    """Get hiragana reading for a Japanese word using pykakasi."""
    kks = _get_kakasi()
    if not kks:
        return ''
    result = kks.convert(word)
    return ''.join(item['hira'] for item in result)


def normalize_text(text: str) -> str:
    """Normalize Unicode text to NFC form (composed characters)."""
    return unicodedata.normalize('NFC', text)


def detect_language(text: str) -> str:
    """
    Auto-detect language from text characters.
    Returns language code (ja, ko, zh-CN, vi, th, etc.) or 'en' as fallback.
    """
    for c in text:
        # Japanese: has hiragana or katakana
        if '\u3040' <= c <= '\u309f' or '\u30a0' <= c <= '\u30ff':
            return 'ja'
        # Korean: hangul
        if '\uac00' <= c <= '\ud7af' or '\u1100' <= c <= '\u11ff' or '\u3130' <= c <= '\u318f':
            return 'ko'
        # Thai
        if '\u0e00' <= c <= '\u0e7f':
            return 'th'
        # Arabic
        if '\u0600' <= c <= '\u06ff':
            return 'ar'
        # Hindi/Devanagari
        if '\u0900' <= c <= '\u097f':
            return 'hi'

    # Check for Vietnamese diacritics (Latin with special chars)
    vietnamese_chars = set('ăâđêôơưàảãáạằẳẵắặầẩẫấậèẻẽéẹềểễếệìỉĩíịòỏõóọồổỗốộờởỡớợùủũúụừửữứựỳỷỹýỵ')
    if any(c.lower() in vietnamese_chars for c in text):
        return 'vi'

    # Check for CJK characters only (likely Chinese if no hiragana/katakana)
    if any('\u4e00' <= c <= '\u9fff' for c in text):
        return 'zh-CN'

    return 'en'


def _is_na(val: str) -> bool:
    """Check if a value is N/A (case-insensitive, whitespace-tolerant)."""
    return val.strip().lower() == 'n/a' if val else False


def _audio_cache_key(word: str, lang: str) -> str:
    """Generate a cache filename: {lang}_{md5_of_word}.mp3"""
    import hashlib
    word_hash = hashlib.md5(word.encode('utf-8')).hexdigest()
    return f"{lang}_{word_hash}.mp3"


def generate_audio(word: str, lang: str, output_path: str, cache_dir: str = '') -> str:
    """
    Generate TTS audio file for a word. Returns the output path.
    If cache_dir is provided:
      - Check cache first → copy cached file to output_path (skip gTTS)
      - If not cached → generate via gTTS, save to output_path AND cache
    Cache is stored in: cache_dir/anki_audio_cache/{lang}_{word_hash}.mp3
    """
    import shutil

    if cache_dir:
        subfolder = 'anki_audio_cache'
        cache_dir = os.path.join(cache_dir, subfolder)
        os.makedirs(cache_dir, exist_ok=True)
        cache_file = os.path.join(cache_dir, _audio_cache_key(word, lang))

        if os.path.exists(cache_file):
            # Cache hit — copy to output
            shutil.copy2(cache_file, output_path)
            return output_path

        # Cache miss — generate and save to both output and cache
        tts = gTTS(text=word, lang=lang)
        tts.save(output_path)
        shutil.copy2(output_path, cache_file)
        return output_path

    # No cache — original behavior
    tts = gTTS(text=word, lang=lang)
    tts.save(output_path)
    return output_path


# Global log callback — set by GUI to capture progress messages
_log_callback = None


def set_log_callback(callback):
    """Set a callback function for log messages. Used by GUI."""
    global _log_callback
    _log_callback = callback


def log_message(msg: str):
    """Print a log message and send to callback if set."""
    print(msg)
    if _log_callback:
        _log_callback(msg)


# =============================================================================
# DATA SOURCES — multiple free resources for word lookup
# =============================================================================

def _source_google_translate(word: str, source_lang: str) -> dict:
    """Source 1: Google Translate — works for any language pair."""
    try:
        meaning = GoogleTranslator(source=source_lang, target='en').translate(word)
        return {'meaning': meaning or ''}
    except Exception:
        return {}


def _source_tatoeba(word: str, source_lang: str) -> dict:
    """
    Source 2: Tatoeba — community-maintained real example sentences.
    Returns a short, natural sentence containing the word + English translation.
    Supports 400+ languages. No API key needed.
    """
    import requests
    MAX_WORDS = 20

    # Tatoeba uses ISO 639-3 codes (3 letters)
    lang_map = {
        'vi': 'vie', 'en': 'eng', 'ja': 'jpn', 'ko': 'kor',
        'zh-CN': 'cmn', 'zh-TW': 'cmn', 'th': 'tha', 'id': 'ind',
        'ms': 'zsm', 'fr': 'fra', 'de': 'deu', 'es': 'spa',
        'pt': 'por', 'it': 'ita', 'ru': 'rus', 'ar': 'ara', 'hi': 'hin',
    }
    tatoeba_lang = lang_map.get(source_lang, source_lang)

    try:
        # Collect all valid sentences across pages, then pick best length
        all_sentences = []

        for page in range(1, 8):  # Try up to 7 pages
            url = f"https://tatoeba.org/en/api_v0/search?from={tatoeba_lang}&to=eng&query=%22{word}%22&sort=relevance&page={page}"
            resp = requests.get(url, timeout=10)
            if resp.status_code != 200:
                break

            data = resp.json()
            results = data.get('results', [])
            if not results:
                break

            for sentence in results:
                text = sentence.get('text', '')
                if word.lower() not in text.lower():
                    continue
                word_count = len(text.split())
                if word_count < 4 or word_count > MAX_WORDS:
                    continue

                # Get English translation
                translations = sentence.get('translations', [])
                if translations and translations[0]:
                    for trans in translations[0]:
                        if trans.get('lang') == 'eng':
                            en_text = trans.get('text', '')
                            if en_text and len(en_text.split()) <= MAX_WORDS:
                                all_sentences.append({
                                    'sentence': text,
                                    'sentence_en': en_text,
                                    'word_count': word_count,
                                })
                                break

            # Stop searching more pages only if we found an ideal (8-12 word) sentence
            has_ideal = any(8 <= s['word_count'] <= 12 for s in all_sentences)
            if has_ideal:
                break

        # Pick best sentence by preferred length: 8-12 > 6-8 > 4-6
        if all_sentences:
            # Try 8-12 words first
            for s in all_sentences:
                if 8 <= s['word_count'] <= 12:
                    return {'sentence': s['sentence'], 'sentence_en': s['sentence_en']}
            # Then 6-8
            for s in all_sentences:
                if 6 <= s['word_count'] <= 8:
                    return {'sentence': s['sentence'], 'sentence_en': s['sentence_en']}
            # Then 4-6
            for s in all_sentences:
                if 4 <= s['word_count'] <= 6:
                    return {'sentence': s['sentence'], 'sentence_en': s['sentence_en']}
            # Fallback: return longest available
            all_sentences.sort(key=lambda x: x['word_count'], reverse=True)
            return {'sentence': all_sentences[0]['sentence'], 'sentence_en': all_sentences[0]['sentence_en']}

    except Exception:
        pass
    return {}


def _source_free_dictionary_api(english_word: str) -> dict:
    """
    Source 2: Free Dictionary API (dictionaryapi.dev)
    Returns definition, example sentence, synonyms, antonyms.
    Only works with English words — use after translating to English.
    No API key needed. Prefers short examples (max 20 words).
    """
    import requests
    MAX_WORDS = 20
    try:
        url = f"https://api.dictionaryapi.dev/api/v2/entries/en/{english_word.lower().strip()}"
        resp = requests.get(url, timeout=10)
        if resp.status_code != 200:
            return {}

        data = resp.json()
        if not data or not isinstance(data, list):
            return {}

        result = {}
        entry = data[0]

        # Get first short definition and example
        all_examples = []
        for meaning_group in entry.get('meanings', []):
            for defn in meaning_group.get('definitions', []):
                example = defn.get('example', '')
                if example:
                    # Split by semicolons and take individual sentences
                    for part in example.split(';'):
                        part = part.strip()
                        if part and len(part.split()) <= MAX_WORDS:
                            all_examples.append(part)
                if not result.get('definition'):
                    result['definition'] = defn.get('definition', '')

            # Get antonyms
            antonyms = meaning_group.get('antonyms', [])
            if antonyms and not result.get('antonym_en'):
                result['antonym_en'] = antonyms[0]

            # Also check definition-level antonyms
            for defn in meaning_group.get('definitions', []):
                if defn.get('antonyms') and not result.get('antonym_en'):
                    result['antonym_en'] = defn['antonyms'][0]

        # Pick the shortest suitable example
        if all_examples:
            all_examples.sort(key=lambda x: len(x.split()))
            result['example_en'] = all_examples[0]

        return result
    except Exception:
        return {}


def _source_wiktionary(english_word: str) -> dict:
    """
    Source 3: Wiktionary via simple REST API.
    Returns definition and example if available.
    """
    import requests
    try:
        url = f"https://en.wiktionary.org/api/rest_v1/page/definition/{english_word.lower().strip()}"
        resp = requests.get(url, timeout=10, headers={'Accept': 'application/json'})
        if resp.status_code != 200:
            return {}

        data = resp.json()
        result = {}

        for lang_section in data.get('en', []):
            for defn in lang_section.get('definitions', []):
                # Get definition text (strip HTML)
                import re
                definition = re.sub(r'<[^>]+>', '', defn.get('definition', ''))
                if definition and not result.get('definition'):
                    result['definition'] = definition

                # Get example sentences
                examples = defn.get('examples', [])
                if examples and not result.get('example_en'):
                    result['example_en'] = re.sub(r'<[^>]+>', '', examples[0])

        return result
    except Exception:
        return {}


def lookup_word_data(word: str, source_lang: str) -> dict:
    """
    Look up word data from multiple free sources. Tries each source and
    combines the best results. Handles errors/timeouts gracefully.

    Returns dict with keys: meaning, sentence, sentence_en, antonym_en
    """
    result = {'meaning': '', 'sentence': '', 'sentence_en': '', 'antonym_en': ''}

    # Step 1: Get English meaning via Google Translate (works for all languages)
    log_message(f"  [Google Translate] Looking up: '{word}' ...")
    gt_data = _source_google_translate(word, source_lang)
    if gt_data.get('meaning'):
        result['meaning'] = gt_data['meaning']

    # Step 2: Get real example sentence from Tatoeba (best quality, natural sentences)
    log_message(f"  [Tatoeba] Looking up sentence for: '{word}' ...")
    tatoeba_data = _source_tatoeba(word, source_lang)
    if tatoeba_data.get('sentence'):
        result['sentence'] = tatoeba_data['sentence']
        result['sentence_en'] = tatoeba_data.get('sentence_en', '')
        log_message(f"    Found: '{result['sentence']}'")

    # Step 3: Look up English word in Free Dictionary API (for antonyms)
    if result['meaning']:
        english_word = result['meaning'].lower().strip()
        # Only look up single words (not phrases)
        if ' ' not in english_word and len(english_word) < 20:
            log_message(f"  [Free Dictionary API] Looking up: '{english_word}' ...")
            fd_data = _source_free_dictionary_api(english_word)
            if fd_data.get('antonym_en'):
                result['antonym_en'] = fd_data['antonym_en']

            # If no Tatoeba sentence, try Free Dict example as fallback
            if not result['sentence'] and fd_data.get('example_en'):
                result['sentence_en'] = fd_data['example_en']

    return result


def generate_sentence(word: str, meaning: str, source_lang: str, example_en: str = '') -> tuple[str, str]:
    """
    Generate an example sentence. Prefers real examples from dictionary APIs.
    Falls back to creating a sentence pattern if no real example found.
    Limits sentences to max 20 words for easy memorization.
    Returns (source_language_sentence, english_sentence).
    """
    MAX_WORDS = 20

    try:
        # If we have a real English example from dictionary, check length first
        if example_en:
            # Clean up: take only first sentence if multiple joined by semicolons
            clean_example = example_en.split(';')[0].split('.')[0].strip()
            if clean_example and len(clean_example.split()) <= MAX_WORDS:
                source_sentence = GoogleTranslator(source='en', target=source_lang).translate(clean_example)
                # Validate: sentence must contain the front word
                if (source_sentence and len(source_sentence.split()) <= MAX_WORDS
                        and word.lower() in source_sentence.lower()):
                    return source_sentence, clean_example

        # Fallback: use Google Translate to get a sentence
        # Translate "I use [meaning] every day" as a generic pattern
        if not meaning:
            return '', ''

        meaning_lower = meaning.lower().strip()

        # Create a simple English sentence with the meaning, translate to source language
        en_sentence = f"I use {meaning_lower} every day"
        try:
            source_sentence = GoogleTranslator(source='en', target=source_lang).translate(en_sentence)
            if source_sentence and word.lower() in source_sentence.lower():
                return source_sentence, en_sentence
        except Exception:
            pass

        # If that didn't contain the word, try another pattern
        en_sentence = f"This {meaning_lower} is very good"
        try:
            source_sentence = GoogleTranslator(source='en', target=source_lang).translate(en_sentence)
            if source_sentence and word.lower() in source_sentence.lower():
                return source_sentence, en_sentence
        except Exception:
            pass

    except Exception:
        pass
    return '', ''


def get_opposite_word(word: str, meaning: str, antonym_en: str, source_lang: str) -> str:
    """
    Get the opposite/antonym of a word.
    Uses antonym from Free Dictionary API and translates to source language.
    Returns empty string if no clear opposite exists.
    """
    try:
        if not meaning:
            return ''

        # If the Free Dictionary API gave us an antonym, use it
        if antonym_en:
            # Translate the English antonym to source language
            opposite_source = GoogleTranslator(source='en', target=source_lang).translate(antonym_en)
            if (opposite_source and
                    opposite_source.lower() != word.lower() and
                    len(opposite_source.split()) <= 2):
                return opposite_source

    except Exception:
        pass
    return ''


def fill_empty_fields(front: str, meaning: str, sentence: str, sentence_meaning: str, opposite: str, opposite_meaning: str = '', reading: str = '', source_lang: str = 'vi') -> dict:
    """
    Fill in empty fields using multiple online sources.
    Priority: Tatoeba (sentences) > Free Dictionary API (antonyms) > Google Translate (fallback).
    Keep existing data as-is. Internet connection required for auto-fill.
    Handles errors/timeouts gracefully — skips failed lookups.
    """

    # Skip internet lookup entirely if all fields are already filled
    # "N/A" counts as filled (means we searched before and found nothing)
    if meaning and sentence and sentence_meaning and opposite and opposite_meaning and reading:
        log_message(f"  All fields filled for '{front}', skipping internet lookup.")
        return {
            'front': front,
            'reading': reading,
            'meaning': meaning,
            'sentence': sentence,
            'sentence_meaning': sentence_meaning,
            'opposite': opposite,
            'opposite_meaning': opposite_meaning,
        }

    # Only search internet for what's actually missing
    word_data = {'meaning': '', 'sentence': '', 'sentence_en': '', 'antonym_en': ''}

    # Need meaning? → Google Translate only
    if not meaning:
        log_message(f"  [Google Translate] Looking up: '{front}' ...")
        gt_data = _source_google_translate(front, source_lang)
        if gt_data.get('meaning'):
            word_data['meaning'] = gt_data['meaning']
            meaning = word_data['meaning']

    # Need sentence? → Tatoeba first
    if not sentence:
        log_message(f"  [Tatoeba] Looking up sentence for: '{front}' ...")
        tatoeba_data = _source_tatoeba(front, source_lang)
        if tatoeba_data.get('sentence'):
            word_data['sentence'] = tatoeba_data['sentence']
            word_data['sentence_en'] = tatoeba_data.get('sentence_en', '')
            log_message(f"    Found: '{word_data['sentence']}'")

    # Need opposite? → Free Dictionary API (use English meaning to find antonym)
    if not opposite:
        # Use meaning to look up antonym — try first word if multi-word
        english_word = meaning.lower().strip() if meaning else ''
        # If meaning has comma or space, try first word only
        if ',' in english_word:
            english_word = english_word.split(',')[0].strip()
        elif ' ' in english_word:
            english_word = english_word.split()[0].strip()
        if english_word and len(english_word) < 20:
            log_message(f"  [Free Dictionary API] Looking up antonym for: '{english_word}' ...")
            fd_data = _source_free_dictionary_api(english_word)
            if fd_data.get('antonym_en'):
                word_data['antonym_en'] = fd_data['antonym_en']
            # Also grab example sentence if we still need one
            if not sentence and not word_data.get('sentence') and fd_data.get('example_en'):
                word_data['sentence_en'] = fd_data['example_en']

    # Fill meaning if empty
    if not meaning:
        meaning = word_data.get('meaning', '')
    if not meaning:
        meaning = 'N/A'

    # Fill sentence if empty — prefer Tatoeba (real natural sentences)
    if not sentence:
        # First try: use Tatoeba sentence directly (already in source language)
        if word_data.get('sentence'):
            sentence = word_data['sentence']
            if not sentence_meaning:
                sentence_meaning = word_data.get('sentence_en', '')
        elif word_data.get('sentence_en'):
            # Second try: translate Free Dictionary English example to source language
            log_message(f"  Translating dictionary example for: '{front}' ...")
            en_example = word_data['sentence_en']
            try:
                sentence = GoogleTranslator(source='en', target=source_lang).translate(en_example) or ''
                if sentence:
                    sentence_meaning = en_example
            except Exception:
                sentence = ''

        # If still nothing from internet — mark N/A
        if not sentence:
            sentence = 'N/A'

    # Fill sentence_meaning if empty (and sentence exists)
    if not sentence_meaning and sentence and not _is_na(sentence):
        log_message(f"  Translating sentence for: '{front}' ...")
        try:
            sentence_meaning = GoogleTranslator(source=source_lang, target='en').translate(sentence) or ''
        except Exception:
            sentence_meaning = ''
    if not sentence_meaning:
        sentence_meaning = 'N/A'

    # Fill opposite if empty — try to find antonym
    if not opposite:
        log_message(f"  Finding opposite for: '{front}' ...")
        effective_meaning = meaning if not _is_na(meaning) else ''
        opposite = get_opposite_word(front, effective_meaning, word_data.get('antonym_en', ''), source_lang)
        if opposite:
            log_message(f"    Found opposite: '{opposite}'")
        else:
            log_message(f"    No clear opposite, marking N/A.")
            opposite = 'N/A'
            opposite_meaning = 'N/A'

    # Fill opposite_meaning if opposite exists but meaning is empty
    if opposite and not _is_na(opposite) and not opposite_meaning:
        try:
            opposite_meaning = GoogleTranslator(source=source_lang, target='en').translate(opposite) or ''
        except Exception:
            opposite_meaning = ''
    if not opposite_meaning:
        opposite_meaning = 'N/A'

    # Fill reading for Japanese (auto-generate from kanji)
    # If source_lang is 'auto', detect from front word
    effective_lang = source_lang
    if source_lang == 'auto':
        effective_lang = detect_language(front)
    if not reading and effective_lang == 'ja':
        log_message(f"  Generating reading for: '{front}' ...")
        reading = get_reading(front)
        if not reading:
            reading = 'N/A'
    elif not reading and effective_lang != 'ja':
        # Non-Japanese: reading not applicable, mark N/A to skip future lookups
        reading = 'N/A'

    return {
        'front': front,
        'reading': reading,
        'meaning': meaning,
        'sentence': sentence,
        'sentence_meaning': sentence_meaning,
        'opposite': opposite,
        'opposite_meaning': opposite_meaning,
    }


def get_reverse_meaning(opposite_word: str, source_lang: str = 'vi') -> str:
    """Get the English meaning of the opposite word via Google Translate."""
    try:
        result = GoogleTranslator(source=source_lang, target='en').translate(opposite_word)
        return result if result else ''
    except Exception:
        return ''


def read_and_generate_csv(input_path: str, output_path: str, reverse: bool, source_lang: str = 'vi', regenerate: bool = False) -> list[dict]:
    """
    Read cards_input (.xlsx or .csv), fill empty fields, write cards_generated.xlsx.
    If reverse=True, also includes reverse (opposite) card rows.
    If regenerate=True, clears all existing data (except front_card) and re-fetches from internet.
    Uses online translation (Google Translate) for auto-filling empty fields.
    Returns the list of all cards (including reverse).
    """
    log_message("[Info] Internet connection required for auto-translating empty fields.")
    if regenerate:
        log_message("[Info] REGENERATE mode: clearing all existing data, re-fetching everything from internet.")
    if source_lang == 'auto':
        log_message("[Info] Language: auto-detect (from word characters)")
    else:
        log_message(f"[Info] Language: '{source_lang}'")
    cards = []
    seen_fronts = set()
    duplicates_removed = 0
    input_headers = None  # Store input header order

    # Read input file (supports .xlsx and .csv)
    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.xlsx':
        wb = load_workbook(input_path, read_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not all_rows:
            return []

        # Map header names to column indices
        headers = [str(h).strip().lower() if h else '' for h in all_rows[0]]
        input_headers = [str(h).strip() if h else '' for h in all_rows[0]]  # preserve original case
        col_map = {name: idx for idx, name in enumerate(headers)}

        def get_col(row, col_name):
            idx = col_map.get(col_name)
            if idx is not None and idx < len(row) and row[idx]:
                return normalize_text(str(row[idx]).strip())
            return ''

        for row in all_rows[1:]:  # skip header
            front = get_col(row, 'front_card')
            if not front:
                continue
            # Check for duplicate front word
            if front in seen_fronts:
                log_message(f"  [Duplicate] Skipping duplicate: '{front}'")
                duplicates_removed += 1
                continue
            seen_fronts.add(front)
            meaning = get_col(row, 'back_card_meaning')
            sentence = get_col(row, 'back_card_sentence')
            sentence_meaning = get_col(row, 'back_card_sentence_meaning')
            opposite = get_col(row, 'back_card_opposite')
            opposite_meaning = get_col(row, 'back_card_opposite_meaning')
            reading = get_col(row, 'front_card_reading')

            # If regenerate mode, clear all fields to force re-fetch from internet
            if regenerate:
                meaning = ''
                sentence = ''
                sentence_meaning = ''
                opposite = ''
                opposite_meaning = ''
                reading = ''

            # Auto-detect language per word if set to 'auto' (only needed if fields are missing)
            word_lang = source_lang
            if source_lang == 'auto':
                if not (meaning and sentence and sentence_meaning and opposite and opposite_meaning and reading):
                    word_lang = detect_language(front)
                    log_message(f"  [Auto-detect] '{front}' → {word_lang}")

            filled = fill_empty_fields(front, meaning, sentence, sentence_meaning, opposite, opposite_meaning, reading, word_lang)
            cards.append(filled)
    else:
        # CSV fallback
        with open(input_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            # Map header names to column indices
            header_row = next(reader, None)
            if not header_row:
                return []
            headers = [h.strip().lower() for h in header_row]
            input_headers = [h.strip() for h in header_row]  # preserve original case
            col_map = {name: idx for idx, name in enumerate(headers)}

            def get_csv_col(row, col_name):
                idx = col_map.get(col_name)
                if idx is not None and idx < len(row) and row[idx].strip():
                    return normalize_text(row[idx].strip())
                return ''

            for row in reader:
                front = get_csv_col(row, 'front_card')
                if not front:
                    continue
                # Check for duplicate front word
                if front in seen_fronts:
                    log_message(f"  [Duplicate] Skipping duplicate: '{front}'")
                    duplicates_removed += 1
                    continue
                seen_fronts.add(front)
                meaning = get_csv_col(row, 'back_card_meaning')
                sentence = get_csv_col(row, 'back_card_sentence')
                sentence_meaning = get_csv_col(row, 'back_card_sentence_meaning')
                opposite = get_csv_col(row, 'back_card_opposite')
                opposite_meaning = get_csv_col(row, 'back_card_opposite_meaning')
                reading = get_csv_col(row, 'front_card_reading')

                # If regenerate mode, clear all fields to force re-fetch from internet
                if regenerate:
                    meaning = ''
                    sentence = ''
                    sentence_meaning = ''
                    opposite = ''
                    opposite_meaning = ''
                    reading = ''

                # Auto-detect language per word if set to 'auto' (only needed if fields are missing)
                word_lang = source_lang
                if source_lang == 'auto':
                    if not (meaning and sentence and sentence_meaning and opposite and opposite_meaning and reading):
                        word_lang = detect_language(front)
                        log_message(f"  [Auto-detect] '{front}' → {word_lang}")

                filled = fill_empty_fields(front, meaning, sentence, sentence_meaning, opposite, opposite_meaning, reading, word_lang)
                cards.append(filled)

    if duplicates_removed > 0:
        log_message(f"[Info] Removed {duplicates_removed} duplicate(s). Processing {len(cards)} unique words.")

    # Generate reverse cards (deduplicating)
    all_cards = list(cards)
    if reverse:
        reverse_seen = set(card['front'] for card in cards)
        for card in cards:
            if card['opposite'] and not _is_na(card['opposite']) and card['opposite'] not in reverse_seen:
                opp = card['opposite']
                reverse_seen.add(opp)
                opp_meaning = get_reverse_meaning(opp, source_lang)
                opp_sentence = ''
                if not _is_na(card['sentence']) and card['front'] in card['sentence']:
                    opp_sentence = card['sentence'].replace(card['front'], opp)
                opp_sentence_meaning = ''
                # Translate reverse sentence if we generated one
                if opp_sentence:
                    try:
                        opp_sentence_meaning = GoogleTranslator(source=source_lang, target='en').translate(opp_sentence) or ''
                    except Exception:
                        opp_sentence_meaning = ''
                all_cards.append({
                    'front': opp,
                    'reading': 'N/A' if source_lang != 'ja' else '',
                    'meaning': opp_meaning or 'N/A',
                    'sentence': opp_sentence or 'N/A',
                    'sentence_meaning': opp_sentence_meaning or 'N/A',
                    'opposite': card['front'],
                    'opposite_meaning': card.get('meaning', '') or 'N/A',
                })

    # Write output xlsx following input header order and column widths
    output_xlsx = output_path.replace('.csv', '.xlsx') if output_path.endswith('.csv') else output_path
    if not output_xlsx.endswith('.xlsx'):
        output_xlsx = os.path.splitext(output_path)[0] + '.xlsx'

    # Read column widths from input file
    input_col_widths = {}
    if ext == '.xlsx':
        try:
            wb_in = load_workbook(input_path, read_only=False)
            ws_in = wb_in.active
            for col_letter, dim in ws_in.column_dimensions.items():
                if dim.width:
                    input_col_widths[col_letter] = dim.width
            wb_in.close()
        except Exception:
            pass

    # Map internal keys to header names
    key_to_header = {
        'front_card': 'front',
        'front_card_reading': 'reading',
        'back_card_meaning': 'meaning',
        'back_card_sentence': 'sentence',
        'back_card_sentence_meaning': 'sentence_meaning',
        'back_card_opposite': 'opposite',
        'back_card_opposite_meaning': 'opposite_meaning',
    }

    # Use input header order, or default if not available
    default_headers = ['front_card', 'front_card_reading', 'back_card_meaning', 'back_card_sentence',
                       'back_card_sentence_meaning', 'back_card_opposite', 'back_card_opposite_meaning']
    out_headers = input_headers if input_headers else default_headers

    # Ensure all required headers exist in output (add missing ones at end)
    out_headers_lower = [h.lower() for h in out_headers]
    for dh in default_headers:
        if dh.lower() not in out_headers_lower:
            out_headers.append(dh)

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = 'Cards'
    ws_out.append(out_headers)
    for card in all_cards:
        row_data = []
        for h in out_headers:
            h_lower = h.lower()
            key = key_to_header.get(h_lower, '')
            row_data.append(card.get(key, ''))
        ws_out.append(row_data)

    # Apply column widths from input file
    from openpyxl.utils import get_column_letter
    if input_col_widths:
        for col_letter, width in input_col_widths.items():
            ws_out.column_dimensions[col_letter].width = width
    else:
        # Auto-fit: set reasonable widths based on header length
        for i, h in enumerate(out_headers, 1):
            col_letter = get_column_letter(i)
            ws_out.column_dimensions[col_letter].width = max(len(h) + 4, 15)

    # Freeze header row and first column (B2 = row 1 + column A frozen)
    ws_out.freeze_panes = 'B2'

    wb_out.save(output_xlsx)

    log_message(f"Generated: {output_xlsx}")
    return all_cards


def create_anki_package(cards: list[dict], deck_name: str, lang: str, output_dir: str, recall: bool = True, cache_dir: str = ''):
    """
    Create an Anki .apkg file with embedded audio.
    Uses deterministic IDs so that:
    - Same deck name → same deck ID (Anki updates existing deck)
    - Same front word → same note ID (Anki preserves review history)
    - Only new cards are added when re-importing

    cache_dir: if provided, audio files are cached here by {lang}_{word_hash}.mp3.
               Same word reuses cached audio (skips gTTS call).
               Default (empty string) means no caching.
    """
    import hashlib

    def _stable_id(text: str) -> int:
        """Generate a stable integer ID from text (deterministic)."""
        h = hashlib.md5(text.encode('utf-8')).hexdigest()
        return int(h[:8], 16) + (1 << 30)  # Ensure it's in valid range

    deck_id = _stable_id(f"deck_{deck_name}")
    model_id = _stable_id(f"model_{deck_name}")

    # Build card templates
    templates = [
        {
            'name': 'Recognition',
            'qfmt': '<div style="font-size: 40px; text-align: center;">{{FrontFurigana}}</div>'
                    '<div style="text-align: center;">{{FrontAudio}}</div>',
            'afmt': '{{FrontSide}}<hr id="answer">'
                    '<div style="font-size: 24px; text-align: center; color: #555;">{{Meaning}}</div>'
                    '<br>'
                    '{{#Sentence}}<div style="font-size: 28px; text-align: center; color: #2196F3; font-weight: bold;">{{SentenceFurigana}}</div>'
                    '<div style="text-align: center;">{{SentenceAudio}}</div>{{/Sentence}}'
                    '{{#SentenceMeaning}}<div style="font-size: 18px; text-align: center; color: #888;">({{SentenceMeaning}})</div>{{/SentenceMeaning}}'
                    '<br>'
                    '{{#Opposite}}<div style="font-size: 30px; text-align: center; color: #E91E63;">↔ {{Opposite}}</div>'
                    '<div style="text-align: center;">{{OppositeAudio}}</div>'
                    '{{#OppositeMeaning}}<div style="font-size: 18px; text-align: center; color: #888;">({{OppositeMeaning}})</div>{{/OppositeMeaning}}'
                    '{{/Opposite}}',
        },
    ]

    if recall:
        templates.append({
            'name': 'Recall',
            'qfmt': '<div style="font-size: 24px; text-align: center; color: #555;">{{Meaning}}</div>'
                    '<div style="text-align: center;">{{FrontAudio}}</div>',
            'afmt': '{{FrontSide}}<hr id="answer">'
                    '<div style="font-size: 40px; text-align: center;">{{FrontFurigana}}</div>'
                    '<br>'
                    '{{#Sentence}}<div style="font-size: 28px; text-align: center; color: #2196F3; font-weight: bold;">{{SentenceFurigana}}</div>'
                    '<div style="text-align: center;">{{SentenceAudio}}</div>{{/Sentence}}'
                    '{{#SentenceMeaning}}<div style="font-size: 18px; text-align: center; color: #888;">({{SentenceMeaning}})</div>{{/SentenceMeaning}}'
                    '<br>'
                    '{{#Opposite}}<div style="font-size: 30px; text-align: center; color: #E91E63;">↔ {{Opposite}}</div>'
                    '<div style="text-align: center;">{{OppositeAudio}}</div>'
                    '{{#OppositeMeaning}}<div style="font-size: 18px; text-align: center; color: #888;">({{OppositeMeaning}})</div>{{/OppositeMeaning}}'
                    '{{/Opposite}}',
        })

    model = genanki.Model(
        model_id,
        f'{deck_name} Model',
        fields=[
            {'name': 'Front'},
            {'name': 'FrontFurigana'},
            {'name': 'Reading'},
            {'name': 'FrontAudio'},
            {'name': 'Meaning'},
            {'name': 'Sentence'},
            {'name': 'SentenceFurigana'},
            {'name': 'SentenceMeaning'},
            {'name': 'SentenceAudio'},
            {'name': 'Opposite'},
            {'name': 'OppositeMeaning'},
            {'name': 'OppositeAudio'},
        ],
        templates=templates,
    )

    deck = genanki.Deck(deck_id, deck_name)

    # Custom note class with deterministic GUID (preserves review history on re-import)
    class StableNote(genanki.Note):
        @property
        def guid(self):
            # Use only front word for stable GUID — same word = same note regardless of deck name
            return genanki.guid_for(self.fields[0])

    # Cards already include reverse cards from read_and_generate_csv
    all_cards = cards

    media_files = []
    temp_audio_dir = os.path.join(output_dir, '_temp_audio')
    os.makedirs(temp_audio_dir, exist_ok=True)

    for i, card in enumerate(all_cards):
        front_word = card['front']
        meaning = card['meaning'] if not _is_na(card['meaning']) else ''
        sentence = card['sentence'] if not _is_na(card['sentence']) else ''
        sentence_meaning = card['sentence_meaning'] if not _is_na(card['sentence_meaning']) else ''
        opposite = card['opposite'] if not _is_na(card['opposite']) else ''
        opposite_meaning = card.get('opposite_meaning', '')
        if _is_na(opposite_meaning):
            opposite_meaning = ''

        # Detect language for audio if set to 'auto'
        audio_lang = lang
        if lang == 'auto':
            audio_lang = detect_language(front_word)

        # Generate audio for front word
        front_audio_filename = f"front_{i}.mp3"
        front_audio_path = os.path.join(temp_audio_dir, front_audio_filename)
        generate_audio(front_word, audio_lang, front_audio_path, cache_dir)
        media_files.append(front_audio_path)

        # Generate audio for sentence (if provided)
        sentence_audio_field = ''
        if sentence:
            sentence_audio_filename = f"sentence_{i}.mp3"
            sentence_audio_path = os.path.join(temp_audio_dir, sentence_audio_filename)
            generate_audio(sentence, audio_lang, sentence_audio_path, cache_dir)
            media_files.append(sentence_audio_path)
            sentence_audio_field = f'[sound:{sentence_audio_filename}]'

        # Generate audio for opposite word (if provided)
        opposite_audio_field = ''
        if opposite:
            opposite_audio_filename = f"opposite_{i}.mp3"
            opposite_audio_path = os.path.join(temp_audio_dir, opposite_audio_filename)
            generate_audio(opposite, audio_lang, opposite_audio_path, cache_dir)
            media_files.append(opposite_audio_path)
            opposite_audio_field = f'[sound:{opposite_audio_filename}]'

        # Generate furigana HTML for Japanese, plain text for others
        reading = card.get('reading', '')
        if _is_na(reading):
            reading = ''
        if reading or any('\u4e00' <= c <= '\u9fff' for c in front_word):
            furigana_html = generate_furigana_html(front_word)
            sentence_furigana = generate_furigana_html(sentence) if sentence else ''
        else:
            furigana_html = front_word
            sentence_furigana = sentence

        note = StableNote(
            model=model,
            fields=[
                front_word,
                furigana_html,
                reading,
                f'[sound:{front_audio_filename}]',
                meaning,
                sentence,
                sentence_furigana,
                sentence_meaning,
                sentence_audio_field,
                opposite,
                opposite_meaning,
                opposite_audio_field,
            ],
        )
        deck.add_note(note)

    # Create package
    package = genanki.Package(deck)
    package.media_files = media_files

    output_filename = f"{deck_name.replace(' ', '_')}.apkg"
    output_path = os.path.join(output_dir, output_filename)
    package.write_to_file(output_path)

    # Clean up temp audio files
    for f in media_files:
        if os.path.exists(f):
            os.remove(f)
    if os.path.exists(temp_audio_dir):
        shutil.rmtree(temp_audio_dir, ignore_errors=True)

    log_message(f"Anki package created: {output_path}")
    log_message(f"Cards generated: {len(all_cards)}")
    log_message(f"[Info] Import into Anki: existing cards keep review history, only new cards are added.")
    return output_path


def main():
    parser = argparse.ArgumentParser(description='Generate Anki cards with embedded TTS audio')
    parser.add_argument('--input', type=str, required=True, help='Input CSV file (cards_input.csv)')
    parser.add_argument('--lang', type=str, default='auto', help='Language code for TTS (default: auto-detect)')
    parser.add_argument('--deck', type=str, default='Anki_deck', help='Deck name (default: Anki_deck)')
    parser.add_argument('--output', type=str, default=None, help='Output directory (default: Downloads folder)')
    parser.add_argument('--reverse', action='store_true', default=True, help='Generate reverse cards (default: yes)')
    parser.add_argument('--no-reverse', action='store_false', dest='reverse', help='Do not generate reverse cards')
    parser.add_argument('--recall', action='store_true', default=True, help='Generate recall cards (default: yes)')
    parser.add_argument('--no-recall', action='store_false', dest='recall', help='Do not generate recall cards')
    parser.add_argument('--data-only', action='store_true', default=False, help='Only generate cards_generated.xlsx, no Anki package')
    parser.add_argument('--cache-dir', type=str, default='', help='Audio cache folder. Same word reuses cached MP3 (default: same as output dir)')

    args = parser.parse_args()

    if args.output:
        output_dir = os.path.abspath(args.output)
    else:
        # Default to Downloads folder (cross-platform)
        output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    os.makedirs(output_dir, exist_ok=True)

    # Step 1: Read input and generate xlsx (input filename + _gen)
    input_path = os.path.abspath(args.input)
    input_basename = os.path.splitext(os.path.basename(input_path))[0]
    generated_filename = f"{input_basename}_gen.xlsx"
    generated_path = os.path.join(output_dir, generated_filename)
    cards = read_and_generate_csv(input_path, generated_path, args.reverse, args.lang)

    if not cards:
        print("Error: No cards to generate")
        sys.exit(1)

    # Step 2: Generate Anki package (unless --data-only)
    if not args.data_only:
        cache_dir = args.cache_dir if args.cache_dir else output_dir
        create_anki_package(cards, args.deck, args.lang, output_dir, args.recall, cache_dir)


if __name__ == '__main__':
    main()