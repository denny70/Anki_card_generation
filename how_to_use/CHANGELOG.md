# Changelog

## v2.2.7
- Added: Linux/Ubuntu build script (`build_linux.py`) producing a standalone binary
- Added: bundle optional TTS/video packages (edge-tts, pyttsx3, moviepy, Pillow) into all builds so every feature works out of the box
- Fixed: "Install Packages" button relaunched the app in a new window when run as a bundled binary; it is now hidden in the standalone app (still available when running from source)

## v2.2.6
- Fixed: Cache folder path not displaying in GUI Entry after browsing (Windows)
- Fixed: Output folder path display refresh on Windows
- UI: Moved cache folder hint to separate line for better layout

## v2.2.5
- Added "Regenerate all data" checkbox in GUI — clears all existing data (except front_card) and re-fetches everything from internet
- Removed hardcoded fallback sentence patterns ("I use X every day", "This X is very good") — all sentences now come from internet sources only (Tatoeba, Free Dictionary API)
- If no real sentence is found, field is marked "N/A" instead of generating a fake one
- Added KNOWLEDGE.md (KW) for internal project documentation

## v1.7
- Updated documentation and changelog
- Drop zone label updated to "Excel/CSV file"

## v1.6.1
- Fix: GUID now based on front word only (not deck name)
- Prevents duplicate cards when re-importing with different deck names
- Delete old deck in Anki and re-import to fix existing duplicates

## v1.6
- Incremental Anki updates: stable note IDs preserve review history on re-import
- Same word always generates the same note ID — Anki skips existing cards
- Only new cards are added when re-importing
- Platform name added to binary filename (win/mac/linux)

## v1.5
- Input columns matched by header name (not position) — reorder freely in Excel
- Output Excel preserves input column widths and column order
- Any missing columns auto-added at end of output

## v1.4
- Single version number (defined in gen_anki.py only)
- Settings saved to user home directory (~/.anki_card_generator/)
- Works with both Python script and standalone exe
- Removed old gui_settings.json from project folder

## v1.3
- Duplicate front word detection — duplicates skipped in output
- Log shows which duplicates were removed
- Summary: "[Info] Removed N duplicate(s). Processing X unique words."

## v1.2
- Added `back_card_opposite_meaning` column (English meaning of opposite word)
- Output file named after input file + "_gen" suffix
- Default deck name changed to "Anki_deck"
- Opposite meaning shown on Anki card below opposite audio

## v1.1
- Removed redundant "Data only" checkbox (button is enough)
- Clean UI

## v1.0
- Initial release
- GUI with drag-and-drop Excel/CSV input
- Multi-source data: Tatoeba (sentences) + Free Dictionary API (antonyms) + Google Translate (meanings)
- Sentences prefer 8-12 words, fallback to shorter
- Sentences must contain the front word
- Smart internet lookup: only calls sources for empty fields
- If all fields filled, skips internet entirely
- No hardcoded dictionaries — all data from internet
- Generates .apkg with embedded TTS audio (no internet needed during review)
- Reverse cards for words with opposites (deduplicated)
- Unicode NFC normalization for proper character display
- Output defaults to Downloads folder (Windows/macOS/Linux)
- Language dropdown in GUI (17 languages supported)
- Settings memory (deck name, language, output folder, last file)
- Install Packages button in GUI
- Build script for standalone .exe (PyInstaller)
