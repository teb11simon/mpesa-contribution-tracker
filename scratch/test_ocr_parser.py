import sys
sys.path.append('src')
from ocr_processor import OCRProcessor

sample_text = '''
100 x 2
100+2
100 2
100*2
100%2
200 ÷ 3
'''
processor = OCRProcessor()
breakdown = processor._parse_denominations(sample_text)
print('Breakdown:', breakdown)
''
