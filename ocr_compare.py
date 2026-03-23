"""
OCR 결과와 기존 ingredient_db.json 비교 분석
- 각 이미지의 OCR 텍스트에서 핵심 정보 추출
- 기존 DB와 대조하여 차이점 리포트
"""
import sys, io, os, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(SCRIPT_DIR, 'ocr_results.json'), 'r', encoding='utf-8') as f:
    ocr_data = json.load(f)

with open(os.path.join(SCRIPT_DIR, 'ingredient_db.json'), 'r', encoding='utf-8') as f:
    db = json.load(f)

# 이미지 파일명 → DB 키 매핑 (파일명에서 추출)
def filename_to_db_keys(fname):
    """파일명에서 가능한 DB 키 후보를 추출"""
    name = os.path.splitext(fname)[0]
    # 괄호 안 내용 포함/제외 버전 모두 생성
    candidates = [name]
    # 괄호 제거 버전
    clean = re.sub(r'[(\(].*?[)\)]', '', name).strip()
    if clean and clean != name:
        candidates.append(clean)
    return candidates

# DB의 모든 키를 정규화해서 빠른 검색
db_keys_normalized = {}
for k in db:
    db_keys_normalized[k.replace(' ', '').lower()] = k

def find_db_key(fname):
    """파일명으로 DB 키 찾기"""
    candidates = filename_to_db_keys(fname)
    for c in candidates:
        norm = c.replace(' ', '').lower()
        if norm in db_keys_normalized:
            return db_keys_normalized[norm]
        # 부분 매칭
        for dk_norm, dk_orig in db_keys_normalized.items():
            if norm in dk_norm or dk_norm in norm:
                return dk_orig
    return None

# 핵심 비교: OCR 텍스트에서 비율(%) 추출하여 DB와 대조
def extract_percentages(text):
    """텍스트에서 숫자% 패턴 추출"""
    return re.findall(r'(\d+\.?\d*)\s*%', text)

def extract_allergens_from_ocr(text):
    """OCR 텍스트에서 알레르기 정보 추출"""
    allergen_keywords = ['밀', '대두', '우유', '계란', '땅콩', '호두', '돼지고기',
                         '새우', '고등어', '복숭아', '토마토', '아황산', '닭고기',
                         '쇠고기', '오징어', '조개', '잣', '알류']
    found = []
    # "알레르기" 또는 "알러지" 근처 텍스트에서 찾기
    for kw in allergen_keywords:
        if kw in text:
            found.append(kw)
    return found

print("=" * 80)
print("OCR 결과 vs DB 대조 리포트")
print("=" * 80)
print()

issues = []
matched = 0
unmatched_files = []

for key, data in sorted(ocr_data.items()):
    if 'error' in data:
        continue

    fname = os.path.basename(key).replace('.jpg', '').replace('.png', '')
    full_text = data.get('full_text', '')

    # 변경공문, 영양성분 등 보조 이미지 스킵
    if any(skip in fname for skip in ['변경공문', '영양성분', '설명1', '설명2', '설명3', '전후']):
        continue

    db_key = find_db_key(fname)

    if not db_key:
        unmatched_files.append(fname)
        continue

    matched += 1
    db_entry = db[db_key]

    # 1. 비율 검증 - DB 성분의 비율과 OCR 텍스트의 비율 비교
    db_percentages = extract_percentages(db_entry.get('성분', ''))
    ocr_percentages = extract_percentages(full_text)

    for dp in db_percentages:
        if dp not in ocr_percentages and float(dp) > 1.0:
            # 큰 비율이 OCR에서 안 보이면 의심
            # OCR에서 유사한 값 찾기
            dp_float = float(dp)
            close_match = any(abs(float(op) - dp_float) < 1.0 for op in ocr_percentages if op.replace('.','').isdigit())
            if not close_match and dp_float > 5:
                issues.append({
                    'file': fname,
                    'db_key': db_key,
                    'type': '비율 불일치 의심',
                    'detail': f'DB에 {dp}% 있으나 OCR에서 미발견. OCR 비율: {ocr_percentages[:10]}'
                })

    # 2. 알레르기 교차 검증
    ocr_allergens = extract_allergens_from_ocr(full_text)
    db_allergens = db_entry.get('알레르기', '')

    # OCR에서 알레르기 키워드가 있는데 DB에 없는 경우
    for oa in ocr_allergens:
        if oa == '밀' and '소맥' in full_text and '밀' not in db_allergens:
            issues.append({
                'file': fname,
                'db_key': db_key,
                'type': '알레르기 누락 의심',
                'detail': f'OCR에서 "{oa}" 감지, DB 알레르기: "{db_allergens}"'
            })

    # 3. 원산지 검증 - OCR에서 국가명이 보이는데 DB와 다른 경우
    countries = ['미국', '중국', '호주', '프랑스', '독일', '이탈리아', '캐나다',
                 '뉴질랜드', '태국', '베트남', '인도네시아', '벨기에', '스페인',
                 '덴마크', '싱가포르', '콜롬비아', '튀르키예', '국산', '국내산']
    ocr_countries = [c for c in countries if c in full_text]
    db_origin = db_entry.get('원산지', '')

    if ocr_countries and db_origin:
        # DB 원산지에 OCR 국가가 하나도 안 들어있으면 의심
        origin_match = any(c in db_origin for c in ocr_countries)
        if not origin_match and len(ocr_countries) <= 3:
            issues.append({
                'file': fname,
                'db_key': db_key,
                'type': '원산지 불일치 의심',
                'detail': f'OCR 국가: {ocr_countries}, DB 원산지: "{db_origin}"'
            })

    # 4. 식품유형 검증
    food_types = ['밀가루', '설탕', '버터', '우유', '견과류', '당류', '잼류',
                  '초콜릿', '조림류', '유지류', '전분가공품', '식품첨가물']
    ocr_types = [ft for ft in food_types if ft in full_text]

# 결과 출력
print(f"매칭된 항목: {matched}건")
print(f"매칭 안 된 파일: {len(unmatched_files)}건")
print()

if unmatched_files:
    print("--- 매칭 안 된 파일 (DB에 없는 이미지) ---")
    for f in unmatched_files:
        print(f"  - {f}")
    print()

print(f"--- 발견된 이슈: {len(issues)}건 ---")
print()

for issue in issues:
    print(f"[{issue['type']}] {issue['db_key']}")
    print(f"  파일: {issue['file']}")
    print(f"  상세: {issue['detail']}")
    print()

# 5. 특별 검사: 모든 OCR full_text를 출력해서 수동 확인 가능하게
print("=" * 80)
print("주요 원재료 OCR 텍스트 (수동 확인용)")
print("=" * 80)

# 가장 중요한 항목들만 출력
priority_items = ['살균 조림류 무설탕 팥앙금', '데코젤뉴트럴', '데코화이트',
                  '슬라이스치즈', '식물성크림', '버터소스', '카야소스',
                  '담금주', '몽크슈 알룰로스', '피스타치오레진']

for key, data in sorted(ocr_data.items()):
    if 'error' in data:
        continue
    fname = os.path.basename(key).replace('.jpg', '').replace('.png', '')
    if any(p in fname for p in priority_items):
        print(f"\n--- {fname} ---")
        texts = data.get('texts', [])
        for t in texts:
            if t['conf'] >= 0.3:
                print(f"  [{t['conf']:.2f}] {t['text']}")
