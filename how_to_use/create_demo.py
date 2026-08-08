"""Create demo input files for Anki Card Generator."""
from openpyxl import Workbook

HEADERS = ['front_card', 'front_card_reading', 'back_card_meaning', 'back_card_sentence',
           'back_card_sentence_meaning', 'back_card_opposite', 'back_card_opposite_meaning']

COL_WIDTHS = [15, 20, 20, 30, 30, 15, 20]


def set_col_widths(ws):
    from openpyxl.utils import get_column_letter
    for i, width in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = width


# Demo 1: Vietnamese (minimal - only front words)
wb = Workbook()
ws = wb.active
ws.title = 'Cards'
ws.append(HEADERS)
words_vi = ['xin chào', 'cảm ơn', 'đẹp', 'nhanh', 'ăn']
for w in words_vi:
    ws.append([w, '', '', '', '', '', ''])
set_col_widths(ws)
wb.save('demo_vietnamese.xlsx')
print('Created: demo_vietnamese.xlsx (5 words, minimal input)')

# Demo 2: Korean (minimal)
wb2 = Workbook()
ws2 = wb2.active
ws2.title = 'Cards'
ws2.append(HEADERS)
words_ko = ['안녕하세요', '감사합니다', '빠르다', '느리다', '먹다']
for w in words_ko:
    ws2.append([w, '', '', '', '', '', ''])
set_col_widths(ws2)
wb2.save('demo_korean.xlsx')
print('Created: demo_korean.xlsx (5 words, minimal input)')

# Demo 3: Vietnamese with some data pre-filled
wb3 = Workbook()
ws3 = wb3.active
ws3.title = 'Cards'
ws3.append(HEADERS)
ws3.append(['nóng', '', 'hot', '', '', 'lạnh', 'cold'])
ws3.append(['to',   '', '',    '', '', 'nhỏ',  ''])
ws3.append(['sách', '', '',    '', '', '',      ''])
set_col_widths(ws3)
wb3.save('demo_partial_data.xlsx')
print('Created: demo_partial_data.xlsx (3 words, some data pre-filled)')

print('\nDone! Use these files with AnkiCardGenerator.')
print('Example: python gen_anki.py --input demo_vietnamese.xlsx --lang vi --deck "VN_demo"')
