# Quick Start Guide

## 3 Steps to Create Anki Cards

### Step 1: Prepare Input File

Create an Excel file with your words in the `front_card` column. Leave other columns empty — the tool fills them automatically.

Example (`demo_vietnamese.xlsx`):

| front_card |
|-----------|
| xin chào |
| cảm ơn |
| đẹp |
| nhanh |
| ăn |

### Step 2: Run the Tool

**GUI:** Double-click `AnkiCardGenerator_v1.7_win.exe`, drag your file onto it, click "Generate Anki Cards"

**CLI:**
```
python gen_anki.py --input demo_vietnamese.xlsx --lang vi --deck "VN_demo"
```

### Step 3: Import into Anki

Double-click the generated `.apkg` file → Anki imports it automatically.

---

## Demo Files Included

| File | Language | Description |
|------|----------|-------------|
| `demo_vietnamese.xlsx` | Vietnamese | 5 words, minimal (only front words) |
| `demo_korean.xlsx` | Korean | 5 words, minimal |
| `demo_partial_data.xlsx` | Vietnamese | 3 words, some data pre-filled |

## Tips

- Only `front_card` is required — everything else is auto-generated
- Keep the same deck name when adding new words (preserves review history)
- Use "Generate Data Only" first to preview before creating the Anki package
- If you already know the meaning or opposite, fill it in — the tool won't overwrite it
