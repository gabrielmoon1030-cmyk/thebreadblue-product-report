"""
원부재료 이미지 자동 동기화
- OneDrive 폴더 감시 → 새 이미지 OCR → ingredient_db.json 업데이트 → Git push
- Windows 작업 스케줄러로 매일 1회 실행
"""
import sys, io, os, json, glob, re, time, subprocess, logging
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ─── 경로 설정 ───
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(SCRIPT_DIR, 'ingredient_db.json')
SYNC_LOG_DIR = os.path.join(SCRIPT_DIR, 'sync_logs')
SYNC_STATE_PATH = os.path.join(SCRIPT_DIR, '.sync_state.json')

IMAGE_DIRS = [
    r'C:/Users/Moon/OneDrive/03. 양지희_품질/6. 원부재료/2025',
    r'C:/Users/Moon/OneDrive/03. 양지희_품질/6. 원부재료/2026',
]
IMAGE_EXTS = ['*.jpg', '*.jpeg', '*.png']

# 보조 이미지 스킵 키워드
SKIP_KEYWORDS = ['변경공문', '영양성분', '설명1', '설명2', '설명3', '전후', '성적서',
                 '표시사항변경', '변경_']

# ─── 로깅 설정 ───
os.makedirs(SYNC_LOG_DIR, exist_ok=True)
log_file = os.path.join(SYNC_LOG_DIR, f'sync_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger(__name__)


# ─── 동기화 상태 관리 ───
def load_sync_state():
    if os.path.exists(SYNC_STATE_PATH):
        with open(SYNC_STATE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"synced_files": {}}

def save_sync_state(state):
    with open(SYNC_STATE_PATH, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ─── DB 로드/저장 ───
def load_db():
    with open(DB_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_db(db):
    with open(DB_PATH, 'w', encoding='utf-8') as f:
        json.dump(db, f, ensure_ascii=False, indent=2)


# ─── 이미지 스캔 ───
def find_all_images():
    files = []
    for d in IMAGE_DIRS:
        if not os.path.exists(d):
            log.warning(f'폴더 없음: {d}')
            continue
        for ext in IMAGE_EXTS:
            files.extend(glob.glob(os.path.join(d, ext)))
    return sorted(files)

def should_skip(fname):
    return any(kw in fname for kw in SKIP_KEYWORDS)

def filename_to_ingredient_name(fpath):
    """파일명에서 원재료명 추출"""
    name = os.path.splitext(os.path.basename(fpath))[0]
    # 앞의 숫자+마침표 제거 (예: "1. 프랑스밀가루" → "프랑스밀가루")
    name = re.sub(r'^\d+\.\s*', '', name)
    # 괄호 안 부가정보 제거
    clean = re.sub(r'\s*[(\(].*?[)\)]\s*', '', name).strip()
    # 언더스코어 뒤 날짜/부가정보 제거 (예: "XXX_06월" → "XXX")
    clean = re.sub(r'_\d+월$', '', clean).strip()
    return clean if clean else name


# ─── 새 이미지 감지 ───
def find_db_match(ingredient_name, db_keys_lower_map):
    """DB에서 매칭되는 키 찾기 (정확 매칭 + 부분 매칭)"""
    norm = ingredient_name.replace(' ', '').lower()
    # 정확 매칭
    if norm in db_keys_lower_map:
        return db_keys_lower_map[norm]
    # 부분 매칭: DB키가 이름에 포함되거나 이름이 DB키에 포함
    for dk_norm, dk_orig in db_keys_lower_map.items():
        if len(dk_norm) >= 3 and (dk_norm in norm or norm in dk_norm):
            return dk_orig
    return None

def find_new_images(db, sync_state):
    all_images = find_all_images()
    db_keys_lower_map = {k.replace(' ', '').lower(): k for k in db}
    synced = set(sync_state.get("synced_files", {}).keys())

    new_images = []
    seen_names = set()  # 중복 파일명 방지 (2025/2026 동일 파일)

    for fpath in all_images:
        fname = os.path.basename(fpath)
        if should_skip(fname):
            continue
        if fpath in synced:
            continue
        ingredient_name = filename_to_ingredient_name(fpath)
        # 중복 방지
        if ingredient_name.lower() in seen_names:
            continue
        seen_names.add(ingredient_name.lower())
        # DB 매칭 (정확 + 부분)
        if find_db_match(ingredient_name, db_keys_lower_map):
            continue
        new_images.append(fpath)

    return new_images


# ─── OCR 처리 ───
def ocr_image(reader, fpath):
    """EasyOCR로 이미지 텍스트 추출"""
    import numpy as np
    from PIL import Image

    img = np.array(Image.open(fpath))
    result = reader.readtext(img)

    texts = []
    for (bbox, text, conf) in result:
        if conf >= 0.3:
            texts.append({'text': text, 'conf': round(conf, 3)})

    full_text = ' '.join(t['text'] for t in texts)
    return full_text, texts


# ─── OCR 텍스트 → DB 항목 파싱 ───
ALLERGEN_KEYWORDS = [
    '밀', '대두', '우유', '계란', '땅콩', '호두', '돼지고기', '새우',
    '고등어', '복숭아', '토마토', '아황산류', '닭고기', '쇠고기',
    '오징어', '조개류', '잣', '아몬드', '메밀', '게', '카카오',
]

COUNTRY_KEYWORDS = [
    '미국', '중국', '호주', '프랑스', '독일', '이탈리아', '캐나다',
    '뉴질랜드', '태국', '베트남', '인도네시아', '벨기에', '스페인',
    '덴마크', '싱가포르', '콜롬비아', '튀르키예', '터키', '일본',
    '국산', '국내산', '외국산', '네덜란드', '필리핀', '브라질',
    '멕시코', '인도', '영국', '페루', '이란',
]

FOOD_TYPE_MAP = {
    '밀가루': '밀가루', '소맥분': '밀가루', '쌀가루': '쌀가루',
    '설탕': '설탕', '백설탕': '설탕', '흑설탕': '설탕',
    '버터': '버터', '마가린': '유지류', '쇼트닝': '유지류',
    '우유': '우유', '연유': '연유', '크림': '유가공품',
    '초콜릿': '초콜릿', '카카오': '초콜릿',
    '잼': '잼류', '조림류': '조림류', '앙금': '조림류',
    '견과': '견과류', '아몬드': '견과류', '호두': '견과류',
    '치즈': '치즈', '요거트': '발효유',
    '이스트': '효모', '드라이이스트': '효모',
    '소금': '소금', '천일염': '소금',
}

def parse_ocr_to_entry(ingredient_name, full_text):
    """OCR 텍스트에서 DB 항목 필드를 추출 (best-effort)"""
    entry = {
        "식품유형": "",
        "원재료명": ingredient_name,
        "표시명": ingredient_name,
        "성분": "",
        "원산지": "",
        "알레르기": "",
        "_자동등록": True,
        "_등록일": datetime.now().strftime("%Y-%m-%d"),
        "_OCR원문": full_text[:500],
    }

    # 1. 알레르기 추출
    found_allergens = []
    for kw in ALLERGEN_KEYWORDS:
        if kw in full_text:
            found_allergens.append(kw)
    if found_allergens:
        entry["알레르기"] = ",".join(found_allergens)

    # 2. 원산지 추출
    found_countries = []
    for c in COUNTRY_KEYWORDS:
        if c in full_text:
            found_countries.append(c)
    if found_countries:
        entry["원산지"] = ",".join(found_countries[:3]) + "산"

    # 3. 식품유형 추출
    for keyword, food_type in FOOD_TYPE_MAP.items():
        if keyword in full_text or keyword in ingredient_name:
            entry["식품유형"] = food_type
            break

    # 4. 성분 추출 (원재료명 섹션에서)
    # "원재료명" 또는 "원료" 이후 텍스트에서 성분 목록 추출 시도
    origin_match = re.search(r'원재료[명]?\s*[:：]?\s*(.{10,200})', full_text)
    if origin_match:
        raw_ingredients = origin_match.group(1)
        # 다음 섹션 키워드 전까지 잘라내기
        for stop in ['영양성분', '보관방법', '유통기한', '소비기한', '제조원']:
            idx = raw_ingredients.find(stop)
            if idx > 0:
                raw_ingredients = raw_ingredients[:idx]
        entry["성분"] = raw_ingredients.strip()[:300]

    # 5. 표시명 — 식품유형이 있으면 사용
    if entry["식품유형"]:
        entry["표시명"] = entry["식품유형"]

    return entry


# ─── Git 자동 커밋 & 푸시 ───
def git_push(added_count):
    try:
        os.chdir(SCRIPT_DIR)
        date_str = datetime.now().strftime("%Y-%m-%d")

        subprocess.run(['git', 'add', 'ingredient_db.json'], check=True, capture_output=True)
        subprocess.run(['git', 'add', '.sync_state.json'], check=True, capture_output=True)

        msg = f"auto: 원재료 DB {added_count}건 자동 추가 ({date_str})"
        result = subprocess.run(
            ['git', 'commit', '-m', msg],
            capture_output=True, text=True, encoding='utf-8'
        )
        if result.returncode != 0:
            if 'nothing to commit' in result.stdout:
                log.info('변경사항 없음, 커밋 스킵')
                return True
            log.error(f'git commit 실패: {result.stderr}')
            return False

        result = subprocess.run(
            ['git', 'push'],
            capture_output=True, text=True, encoding='utf-8'
        )
        if result.returncode != 0:
            log.error(f'git push 실패: {result.stderr}')
            return False

        log.info(f'Git push 완료: {msg}')
        return True

    except Exception as e:
        log.error(f'Git 오류: {e}')
        return False


# ─── 메인 ───
def main():
    log.info('=' * 60)
    log.info('원부재료 자동 동기화 시작')
    log.info('=' * 60)

    db = load_db()
    sync_state = load_sync_state()
    log.info(f'현재 DB: {len(db)}건')

    # 새 이미지 감지
    new_images = find_new_images(db, sync_state)
    log.info(f'새 이미지: {len(new_images)}건')

    if not new_images:
        log.info('추가할 이미지 없음. 종료.')
        return

    # EasyOCR 로드 (새 이미지가 있을 때만)
    log.info('EasyOCR 로딩 중...')
    import easyocr
    reader = easyocr.Reader(['ko', 'en'], gpu=False, verbose=False)
    log.info('EasyOCR 준비 완료')

    added = 0
    for i, fpath in enumerate(new_images):
        fname = os.path.basename(fpath)
        ingredient_name = filename_to_ingredient_name(fpath)
        log.info(f'[{i+1}/{len(new_images)}] {fname} → "{ingredient_name}"')

        try:
            full_text, texts = ocr_image(reader, fpath)
            log.info(f'  OCR 완료: {len(texts)}개 텍스트 블록')

            entry = parse_ocr_to_entry(ingredient_name, full_text)
            db[ingredient_name] = entry
            added += 1

            # 동기화 상태 기록
            sync_state.setdefault("synced_files", {})[fpath] = {
                "synced_at": datetime.now().isoformat(),
                "ingredient_name": ingredient_name,
            }

            log.info(f'  → DB 추가 완료: 식품유형={entry["식품유형"] or "미분류"}, '
                     f'알레르기={entry["알레르기"] or "없음"}, '
                     f'원산지={entry["원산지"] or "미확인"}')

        except Exception as e:
            log.error(f'  처리 실패: {e}')

    # 저장
    if added > 0:
        save_db(db)
        save_sync_state(sync_state)
        log.info(f'DB 저장 완료: {len(db)}건 (신규 {added}건)')

        # Git push
        git_push(added)

        # 리뷰 필요 항목 알림
        review_items = [k for k, v in db.items() if isinstance(v, dict) and v.get('_자동등록')]
        if review_items:
            log.info('')
            log.info(f'⚠ 수동 검토 필요: {len(review_items)}건')
            for item in review_items:
                log.info(f'  - {item}')
            log.info('ingredient_db.json에서 "_자동등록" 항목을 검토 후 플래그를 제거하세요.')

    log.info('')
    log.info(f'동기화 완료! 로그: {log_file}')


if __name__ == '__main__':
    start = time.time()
    main()
    elapsed = time.time() - start
    log.info(f'총 소요시간: {elapsed:.0f}초')
