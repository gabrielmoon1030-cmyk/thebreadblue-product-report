"""
원부재료 이미지 전수 OCR 스캔 (EasyOCR)
- 2025/2026 폴더의 모든 이미지를 OCR 처리
- 결과를 JSON으로 저장
"""
import sys, io, os, json, glob, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import numpy as np
from PIL import Image
import easyocr

# 설정
IMAGE_DIRS = [
    r'C:/Users/Moon/OneDrive/03. 양지희_품질/6. 원부재료/2025',
    r'C:/Users/Moon/OneDrive/03. 양지희_품질/6. 원부재료/2026',
]
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), 'ocr_results.json')

def scan_all_images():
    reader = easyocr.Reader(['ko', 'en'], gpu=False, verbose=False)

    all_results = {}
    image_files = []

    for d in IMAGE_DIRS:
        for ext in ['*.jpg', '*.jpeg', '*.png']:
            image_files.extend(glob.glob(os.path.join(d, ext)))

    print(f'총 {len(image_files)}개 이미지 발견')
    print()

    for i, fpath in enumerate(sorted(image_files)):
        fname = os.path.basename(fpath)
        folder = os.path.basename(os.path.dirname(fpath))
        key = f'{folder}/{fname}'

        print(f'[{i+1}/{len(image_files)}] {key} ... ', end='', flush=True)

        try:
            img = np.array(Image.open(fpath))
            result = reader.readtext(img)

            texts = []
            for (bbox, text, conf) in result:
                texts.append({
                    'text': text,
                    'conf': round(conf, 3)
                })

            # 전체 텍스트 합치기 (신뢰도 0.3 이상만)
            full_text = ' '.join([t['text'] for t in texts if t['conf'] >= 0.3])

            all_results[key] = {
                'file': fpath,
                'texts': texts,
                'full_text': full_text
            }

            print(f'OK ({len(texts)}개 텍스트)')

        except Exception as e:
            print(f'ERROR: {e}')
            all_results[key] = {'file': fpath, 'error': str(e)}

    # 저장
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)

    print(f'\n완료! {OUTPUT_FILE} 에 저장됨')
    print(f'성공: {sum(1 for v in all_results.values() if "error" not in v)}건')
    print(f'실패: {sum(1 for v in all_results.values() if "error" in v)}건')

if __name__ == '__main__':
    start = time.time()
    scan_all_images()
    elapsed = time.time() - start
    print(f'총 소요시간: {elapsed:.0f}초')
