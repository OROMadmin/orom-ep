"""
build_naver_ep.py  v2 — 오롬 네이버 가격비교 EP 생성기 (Cafe24 미러 + 오버라이드)
================================================================================

구조
    Cafe24 Admin API  ──►  [미러 레이어]  Cafe24가 네이버로 보내던 EP와 동일한 값 생성
                            (2025-08 Cafe24 DBURL 사본 483건 역분석으로 규칙 확정)
    Cafe24 EP 사본    ──►  [보강 레이어]  API로 못 얻는 필드만 채움
                            (review_count / order_made / product_flag / event_words)
    overrides         ──►  [오버라이드]   우리가 바꾸고 싶은 필드만 덮어씀
                            (title / image_link / search_tag / attribute ...)
                     ──►  [검증] ──► naver_ep.tsv

미러 규칙 (사본 역분석 결과)
    id               = product_no                      (483/483 일치, link 내 번호와도 일치)
    price_pc/mobile  = price                           (483/483)
    benefit_price    = price                           (483/483)
    image_link       = 목록이미지(medium), http 스킴     (483/483)
    category         = 표시중 분류 중 가장 얕은 것, 동률 시 분류번호 작은 것 (표본 88건 중
                       구조 일치 100%, 이름 변경분 제외)
    condition        = '신상품' (중고류는 '중고')
    shipping         = 개별배송비 상품 → 0 / 그 외 판매가 ≥ 40,000 → 0, 미만 → 3,000
                       (470건 불일치 0) ※ 현재 정책과 같은지 반드시 확인
    brand / maker    = 브랜드·제조사 코드 → 이름 (오롬 / 자체브랜드 / 자체제작)
    minimum_purchase_quantity = minimum_quantity
    model_number     = model_name

원칙
    * Cafe24 원본은 절대 수정하지 않는다 (GET만)
    * 가격·배송비·링크·id는 오버라이드 불가 (가격 정합성 제재 방지)
    * 검증 실패 / 상품 수 급변 시 기존 파일 유지하고 중단

사용법
    python build_naver_ep.py                          # 생성
    python build_naver_ep.py --dry-run                # 검증만
    python build_naver_ep.py --limit 10               # 테스트
    python build_naver_ep.py --mirror-only            # 오버라이드 없이 순수 미러 (비교용)
    python build_naver_ep.py --compare 사본.xlsx       # Cafe24 EP와 필드별 일치율 비교
    python build_naver_ep.py --init-override          # 오버라이드 양식 생성
"""

import argparse
import csv
import glob
import io
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime

# ── 설정 ──────────────────────────────────────────────────────────────────────
MALL_URL = "https://orom.co.kr"
FREE_SHIPPING_THRESHOLD = 40000   # ⚠ 2025-08 사본 역산값. 현재 무료배송 기준과 같은지 확인할 것
BASE_SHIPPING_FEE = 3000
IMAGE_SCHEME = "http"             # Cafe24 EP와 동일하게 (바꾸면 네이버가 이미지 재수집)

# ── 데이터 소스 ──
#  (A) 로컬/Cowork: Cafe24 API 직접 + 로컬 overrides.xlsx + cafe24_snapshot*.xlsx
#  (B) 운영(시트 허브): 구글시트 탭을 '웹에 게시 → CSV'한 URL 3개를 환경변수로 지정
#      → Cafe24 키 없이 동작. 토큰 체인은 시트(Apps Script) 하나만 가진다.
SHEET_PRODUCTS_CSV = os.environ.get("OROM_SHEET_PRODUCTS_CSV", "")    # cafe24_products 탭
OVERRIDE_CSV_URL = os.environ.get("OROM_SHEET_OVERRIDES_CSV", "")     # overrides 탭
SHEET_SNAPSHOT_CSV = os.environ.get("OROM_SHEET_SNAPSHOT_CSV", "")    # cafe24_ep_snapshot 탭
OVERRIDE_XLSX = "overrides.xlsx"
CAFE24_SNAPSHOT_GLOB = "cafe24_snapshot*.xlsx"

# 우리 정책상 네이버로 내보내지 않을 분류 (Cafe24 미러와 의도적 차이)
EXCLUDE_CATS = ("컨텐츠", "콜라보레이션(미노출)", "기획전", "이벤트",
                "별도 결제", "미사용 카테고리", "사은품")

# ── Cafe24 헬퍼 탐색 (시트 모드에서는 불필요) ─────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
ca = None
if not SHEET_PRODUCTS_CSV:
    _CANDIDATES = [r"C:\Users\taeyo\OneDrive\6. 프로그래밍\@클로드\판매 데이터 분석\cafe24"]
    _CANDIDATES += glob.glob("/sessions/*/mnt/판매 데이터 분석/cafe24")
    for _d in _CANDIDATES:
        if os.path.isdir(_d):
            sys.path.insert(0, _d)
            _CWD = os.getcwd()
            os.chdir(_d)
            import cafe24_api as ca  # noqa: E402
            os.chdir(_CWD)
            break
    else:
        sys.exit("cafe24 헬퍼 폴더를 찾을 수 없습니다 (또는 OROM_SHEET_PRODUCTS_CSV 지정)")


def fetch_text(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "orom-ep"}), timeout=60) as r:
        return r.read().decode("utf-8")

OUT_TSV = os.path.join(BASE, "naver_ep.tsv")
LOG = os.path.join(BASE, "build.log")
CATCACHE = os.path.join(BASE, "_category_map.json")

# ── EP 컬럼 (Cafe24 사용 컬럼 + 보강/오버라이드용) ────────────────────────────
COLUMNS = [
    "id", "title", "price_pc", "price_mobile", "normal_price", "link", "mobile_link",
    "image_link", "add_image_link", "video_url",
    "category_name1", "category_name2", "category_name3", "category_name4",
    "naver_category", "condition", "order_made", "product_flag",
    "model_number", "brand", "maker", "origin", "event_words", "benefit_price",
    "search_tag", "minimum_purchase_quantity", "review_count", "shipping",
    "attribute", "age_group", "gender",
]
LOCKED = {"id", "price_pc", "price_mobile", "normal_price", "benefit_price",
          "link", "mobile_link", "shipping", "review_count"}
FROM_SNAPSHOT = ("review_count", "order_made", "product_flag", "event_words")

# ── 사본 이후 신상품용 추정 규칙 (2025-08 사본 역분석) ──────────────────────
#  사본에서 order_made/product_flag/event_words가 채워진 13개는 전부 232번 분류
#  '기업/단체 다이어리' 소속이며, 그 분류의 사본 상품은 13/13 모두 같은 값이었다.
#  → 사본에 없는 232번 소속 상품(예: 4057~4068 기업 명함지갑)에 같은 값을 적용.
#  review_count는 추정하지 않는다 (실제보다 많으면 몰 전체 노출중단 → 공란이 안전).
INFER_RULES = [
    {"category_no": 232,
     "set": {"order_made": "Y", "product_flag": "도매",
             "event_words": "[무료 샘플 제공] 직접 찾아뵙고 상담해 드립니다"}},
]
CHECK_IMAGES = os.environ.get("OROM_CHECK_IMAGES", "1") == "1"   # _dburl 이미지 404 → 원본 이미지로 대체

BANNED_TITLE = re.compile(
    r"(무료배송|당일배송|빠른배송|최저가|1\+1|2\+1|사은품|증정|할인|특가|세일|SALE"
    r"|이벤트|쿠폰|적립|베스트|BEST|한정|기획전|테스트)", re.I)
BANNED_CHARS = re.compile(r"[★☆♥♡●○◆◇■□▶◀※!?~@#$%^*+=|<>{}【】]")
CONDITION_MAP = {"N": "신상품"}          # 그 외(B/R/U/E/F/S) → '중고'


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ═════════════════════════════════════════════════════════════════════════════
# 1. Cafe24 원천 (읽기 전용)
# ═════════════════════════════════════════════════════════════════════════════
def get_all(path, key, params=None):
    out, off = [], 0
    while True:
        p = dict(params or {}); p.update(limit=100, offset=off)
        r = ca.get(path, p)
        items = r.get(key, []) if isinstance(r, dict) else []
        if not items:
            break
        out.extend(items)
        off += 100
        if len(items) < 100:
            break
    return out


def _cat_products(cn, retries=5):
    """분류 소속 상품번호 전체.
    ※ categories/{cn}/products 는 offset·since_product_no를 무시하고 최신 100개만 준다
      (2026-09-28 확인: 133번 분류 362개 중 100개만 반환) → products?category= 로 offset 페이징."""
    out, off = [], 0
    while True:
        items = None
        for a in range(retries):
            r = ca.get("products", {"category": cn, "fields": "product_no", "limit": 100, "offset": off})
            if isinstance(r, dict) and "products" in r:
                items = r["products"]; break
            time.sleep(1.5 * (a + 1))
        if items is None:
            raise RuntimeError(f"카테고리 {cn} 조회 실패")
        out.extend(items)
        if len(items) < 100:
            break
        off += 100
    return out


def product_category_map(cat_map):
    if os.path.exists(CATCACHE) and time.time() - os.path.getmtime(CATCACHE) < 6 * 3600:
        log("  카테고리 매핑 캐시 사용")
        return {int(k): v for k, v in json.load(open(CATCACHE, encoding="utf-8")).items()}
    mp, failed = {}, []
    for i, cn in enumerate(cat_map, 1):
        try:
            for it in _cat_products(cn):
                mp.setdefault(it["product_no"], []).append(cn)
        except RuntimeError:
            failed.append(cn)
        if i % 40 == 0:
            time.sleep(1.0)
    for cn in list(failed):
        time.sleep(3)
        try:
            for it in _cat_products(cn, retries=8):
                mp.setdefault(it["product_no"], []).append(cn)
            failed.remove(cn)
        except RuntimeError:
            pass
    if failed:
        raise SystemExit(f"[중단] 카테고리 {len(failed)}개 조회 실패 → 불완전한 EP 생성 방지: {failed[:10]}")
    json.dump({str(k): v for k, v in mp.items()}, open(CATCACHE, "w", encoding="utf-8"))
    return mp


def cat_path(cat_map, cn):
    c = cat_map.get(cn) or {}
    f = c.get("full_category_name") or {}
    return [f.get(str(i)) for i in range(1, 5) if f.get(str(i))]


def pick_category(cat_map, cns):
    """Cafe24 규칙: 상품 소속 분류 중 분류번호가 가장 작은 것 (표시여부·깊이 무관).
    2025-08 사본의 link 내 cate_no 기준 258/258 일치."""
    if not cns:
        return None, []
    cn = min(cns)
    return cn, cat_path(cat_map, cn)


def code_names(endpoint, key, code_field, name_field, fallback):
    try:
        rows = ca.get(endpoint, {"limit": 100}).get(key, [])
        m = {r[code_field]: r.get(name_field, "") for r in rows if r.get(code_field)}
        return {**fallback, **m}
    except Exception:
        return fallback


BRAND_FALLBACK = {"B000000B": "오롬", "B0000000": "자체브랜드"}
MAKER_FALLBACK = {"M000000C": "오롬", "M0000000": "자체제작"}


def load_products():
    """→ (products[list of dict], catinfo{product_no: {'cn','path','allnames'}})
    두 소스가 같은 형태를 돌려주므로 이후 로직은 동일하다."""
    catinfo = {}
    if SHEET_PRODUCTS_CSV:
        rows = list(csv.DictReader(io.StringIO(fetch_text(SHEET_PRODUCTS_CSV))))
        if not rows or "product_no" not in rows[0]:
            raise SystemExit("[중단] cafe24_products 시트 CSV 형식 오류")
        products = []
        for r in rows:
            if not str(r.get("product_no", "")).strip().isdigit():
                continue
            p = dict(r)
            p["product_no"] = int(r["product_no"])
            p["product_tag"] = [t for t in (r.get("product_tag") or "").split("|") if t]
            p["_brand"] = r.get("brand_name") or BRAND_FALLBACK.get(r.get("brand_code"), "")
            p["_maker"] = r.get("manufacturer_name") or MAKER_FALLBACK.get(r.get("manufacturer_code"), "")
            products.append(p)
            cn = r.get("ep_cate_no") or ""
            catinfo[p["product_no"]] = {
                "cn": int(float(cn)) if str(cn).strip() else None,
                "path": [x.strip() for x in (r.get("ep_category_path") or "").split(" > ") if x.strip()],
                "allnames": r.get("all_category_paths") or "",
                "cns": [int(float(x)) for x in str(r.get("all_category_nos") or "").split(",") if x.strip()],
            }
        return products, catinfo

    products = get_all("products", "products")
    cat_map = {c["category_no"]: c for c in get_all("categories", "categories")}
    prod_cats = product_category_map(cat_map)
    brands = code_names("brands", "brands", "brand_code", "brand_name", BRAND_FALLBACK)
    makers = code_names("manufacturers", "manufacturers", "manufacturer_code", "manufacturer_name", MAKER_FALLBACK)
    for p in products:
        p["_brand"] = brands.get(p.get("brand_code"), "")
        p["_maker"] = makers.get(p.get("manufacturer_code"), "")
        cns = prod_cats.get(p["product_no"], [])
        cn, path = pick_category(cat_map, cns)
        catinfo[p["product_no"]] = {
            "cn": cn, "path": path,
            "allnames": " | ".join(" > ".join(filter(None, cat_path(cat_map, c))) for c in sorted(cns)),
            "cns": sorted(cns),
        }
    return products, catinfo


# ═════════════════════════════════════════════════════════════════════════════
# 2. Cafe24 EP 사본 (API로 못 얻는 필드)
# ═════════════════════════════════════════════════════════════════════════════
def read_ep_table(path):
    """Cafe24 EP 사본 (xlsx 또는 tsv) → {id: {col: val}}"""
    if path.lower().endswith((".tsv", ".txt")):
        raw = open(path, "rb").read()
        for enc in ("utf-8", "euc-kr", "cp949"):
            try:
                txt = raw.decode(enc); break
            except UnicodeDecodeError:
                continue
        rows = list(csv.DictReader(io.StringIO(txt), delimiter="\t"))
    else:
        from openpyxl import load_workbook
        ws = load_workbook(path, data_only=True, read_only=True).active
        it = ws.iter_rows(values_only=True)
        hdr = [str(h).strip() if h is not None else "" for h in next(it)]
        rows = [dict(zip(hdr, ["" if v is None else str(v) for v in r])) for r in it]
    return {r["id"]: r for r in rows if r.get("id")}


def load_snapshot():
    if SHEET_SNAPSHOT_CSV:
        rows = list(csv.DictReader(io.StringIO(fetch_text(SHEET_SNAPSHOT_CSV))))
        snap = {r["id"]: r for r in rows if r.get("id")}
        log(f"  Cafe24 EP 사본: 시트 cafe24_ep_snapshot ({len(snap)}건)")
        return snap, "sheet"
    files = sorted(glob.glob(os.path.join(BASE, CAFE24_SNAPSHOT_GLOB)))
    if not files:
        log("  Cafe24 EP 사본 없음 → review_count 등 보강필드 비움 (⚠ 전환 전 필수)")
        return {}, None
    f = files[-1]                                   # 파일명 날짜순 (cafe24_snapshot_YYYY-MM-DD.xlsx)
    snap = read_ep_table(f)
    m = re.search(r"(\d{4}-\d{2}(?:-\d{2})?)", os.path.basename(f))
    log(f"  Cafe24 EP 사본: {os.path.basename(f)} ({len(snap)}건, 기준일 {m.group(1) if m else '파일명에 날짜 없음'})")
    if m:
        d = datetime.strptime(m.group(1) + ("-01" if len(m.group(1)) == 7 else ""), "%Y-%m-%d")
        if (datetime.now() - d).days > 14:
            log(f"  ⚠ 사본이 {(datetime.now() - d).days}일 지남 — 리뷰 수가 실제보다 적게 나감(많게는 안 나감), 신규상품은 리뷰수 공란")
    return snap, f


# ═════════════════════════════════════════════════════════════════════════════
# 3. 오버라이드
# ═════════════════════════════════════════════════════════════════════════════
def load_overrides():
    rows = []
    if OVERRIDE_CSV_URL:
        try:
            with urllib.request.urlopen(OVERRIDE_CSV_URL, timeout=30) as r:
                rows = list(csv.DictReader(io.StringIO(r.read().decode("utf-8"))))
            log(f"  오버라이드: 구글시트 CSV ({len(rows)}행)")
        except Exception as e:
            raise SystemExit(f"[중단] 오버라이드 시트를 읽지 못함: {e}")
    else:
        p = os.path.join(BASE, OVERRIDE_XLSX)
        if not os.path.exists(p):
            log("  오버라이드 없음")
            return {}
        from openpyxl import load_workbook
        ws = load_workbook(p, data_only=True).active
        it = ws.iter_rows(values_only=True)
        hdr = [str(h).strip() if h else "" for h in next(it)]
        rows = [dict(zip(hdr, r)) for r in it]
        log(f"  오버라이드: {OVERRIDE_XLSX} ({len(rows)}행)")
    ov, blocked = {}, set()
    for r in rows:
        pid = str(r.get("id") or "").strip()
        if not pid.isdigit():
            continue
        rec = {}
        for k, v in r.items():
            if not k or k == "id" or k.startswith("_") or v in (None, ""):
                continue
            if k in LOCKED:
                blocked.add(k); continue
            if k in COLUMNS:
                rec[k] = str(v).strip()
        if rec:
            ov[pid] = rec
    if blocked:
        log(f"  보호 필드 오버라이드 무시됨: {sorted(blocked)}")
    return ov


# ═════════════════════════════════════════════════════════════════════════════
# 4. 레코드 생성 (미러)
# ═════════════════════════════════════════════════════════════════════════════
def shipping_fee(p):
    if p.get("shipping_fee_by_product") == "T":
        return "0" if p.get("shipping_fee_type") == "T" else None   # None = 판단불가 → 제외
    return "0" if int(float(p.get("price") or 0)) >= FREE_SHIPPING_THRESHOLD else str(BASE_SHIPPING_FEE)


def mirror_record(p, ci):
    """p: 상품(dict), ci: {'cn': 대표분류번호, 'path': [분류명...], 'allnames': 전체소속분류명}"""
    pid = str(p["product_no"])
    price = str(int(float(p.get("price") or 0)))
    cn, cats = ci.get("cn"), (ci.get("path") or ["기타"])
    # 이미지: Cafe24가 네이버용으로 따로 만드는 '_dburl' 사본, http 스킴 (사본 483/483 동일 형식)
    img = p.get("list_image") or p.get("detail_image") or ""
    img = re.sub(r"(\.[A-Za-z0-9]+)$", r"_dburl\1", img)
    if IMAGE_SCHEME == "http":
        img = re.sub(r"^https://", "http://", img)
    tags = list(dict.fromkeys(t.strip() for t in (p.get("product_tag") or []) if t.strip()))
    # 링크: Cafe24와 동일한 추적 파라미터 필수
    #   cafe_mkt=naver_ks → detail.html의 utm 부여 스크립트 + Cafe24 네이버쇼핑 매출 집계(판매지수 EP)
    base = f"{MALL_URL}/product/detail.html?product_no={pid}&cate_no={cn or ''}&display_group=1&cafe_mkt=naver_ks"
    link = base + "&mkt_in=Y&ghost_mall_id=naver&ref=naver_open"
    mlink = base
    return {
        "id": pid,
        "title": (p.get("product_name") or "").strip(),
        "price_pc": price, "price_mobile": price, "benefit_price": price, "normal_price": "",
        "link": link, "mobile_link": mlink,
        "image_link": img, "add_image_link": "", "video_url": "",
        "category_name1": cats[0] if len(cats) > 0 else "",
        "category_name2": cats[1] if len(cats) > 1 else "",
        "category_name3": cats[2] if len(cats) > 2 else "",
        "category_name4": cats[3] if len(cats) > 3 else "",
        "naver_category": "",
        "condition": CONDITION_MAP.get(p.get("product_condition") or "N", "중고"),
        "order_made": "", "product_flag": "",
        "model_number": (p.get("model_name") or "").strip(),
        "brand": p.get("_brand") or "오롬",
        "maker": p.get("_maker") or "오롬",
        "origin": "", "event_words": "",
        # 네이버는 앞 10개·100자만 사용 → 실효값 그대로 전송
        "search_tag": _tags(tags),
        "minimum_purchase_quantity": str(p.get("minimum_quantity") or 1),
        "review_count": "",
        "shipping": shipping_fee(p),
        "attribute": "", "age_group": "", "gender": "",
    }


def _tags(tags):
    tags = tags[:10]
    s = "|".join(tags)
    while len(s) > 100 and tags:
        tags.pop(); s = "|".join(tags)
    return s


def exclude_reason(p, ci):
    if p.get("display") != "T": return "미진열"
    if p.get("selling") != "T": return "판매안함"
    if float(p.get("price") or 0) <= 0: return "가격0원"
    if p.get("sold_out") == "T": return "품절"
    # 소속 분류의 '최상위 이름'이 전부 비상품 분류일 때만 제외 (부분일치 금지:
    #  '다이어리 컨텐츠별 보기'가 '컨텐츠'에 걸려 2027 다이어리 40여 개가 빠지던 버그 수정, 2026-10-01)
    tops = {seg.split(" > ")[0].strip() for seg in (ci.get("allnames") or "").split(" | ") if seg.strip()}
    if tops and tops <= set(EXCLUDE_CATS): return "비상품분류(정책)"
    if re.search(r"테스트|개인결제", p.get("product_name") or ""): return "테스트/개인결제"
    if not (p.get("list_image") or p.get("detail_image")): return "이미지없음"
    if shipping_fee(p) is None: return "배송비판단불가"
    return ""


# ═════════════════════════════════════════════════════════════════════════════
# 5. 검증
# ═════════════════════════════════════════════════════════════════════════════
REQUIRED = ["id", "title", "price_pc", "link", "image_link", "category_name1", "shipping"]


def validate(rows):
    errs, warns, seen = [], [], set()
    for r in rows:
        pid = r["id"]
        for f in REQUIRED:
            if not str(r.get(f, "")).strip():
                errs.append(f"#{pid} 필수 누락 {f}")
        if pid in seen: errs.append(f"#{pid} id 중복")
        seen.add(pid)
        if not re.fullmatch(r"[A-Za-z0-9_\- ]{1,50}", pid): errs.append(f"#{pid} id 규격")
        t = r["title"]
        if len(t) > 100: errs.append(f"#{pid} 상품명 100자 초과")
        if BANNED_TITLE.search(t): warns.append(f"#{pid} 상품명 금지어 '{BANNED_TITLE.search(t).group()}'")
        if BANNED_CHARS.search(t): warns.append(f"#{pid} 상품명 특수문자 '{BANNED_CHARS.search(t).group()}'")
        if not r["image_link"].startswith(("http://", "https://")): errs.append(f"#{pid} 이미지 URL")
        if len(r["image_link"]) > 255: errs.append(f"#{pid} 이미지 URL 255자 초과")
        if r["add_image_link"] and len(r["add_image_link"].split("|")) > 10:
            errs.append(f"#{pid} 추가이미지 10개 초과")
        if not re.fullmatch(r"-1|\d{1,7}", r["shipping"] or ""): errs.append(f"#{pid} 배송비 형식")
        for f in ("order_made",):
            if r[f] not in ("", "Y"): errs.append(f"#{pid} {f}는 Y 또는 공란")
        if r["review_count"] and not r["review_count"].isdigit(): errs.append(f"#{pid} 리뷰수 형식")
        if "cafe_mkt=naver_ks" not in r["link"] or "cafe_mkt=naver_ks" not in r["mobile_link"]:
            errs.append(f"#{pid} 링크 추적 파라미터 누락 (판매지수·GA 유입 끊김)")
    # 네이버: 다수 상품에 동일 대표이미지 → 해당 이미지 상품 전체 노출중단
    from collections import Counter
    dup = [u for u, n in Counter(r["image_link"] for r in rows).items() if n > 1]
    for u in dup[:10]:
        ids = [r["id"] for r in rows if r["image_link"] == u]
        errs.append(f"대표이미지 중복 {len(ids)}개 상품 {ids[:5]} → {u[-40:]}")
    return errs, warns


# ═════════════════════════════════════════════════════════════════════════════
# 6. Cafe24 EP와 비교
# ═════════════════════════════════════════════════════════════════════════════
def compare(rows, snap_path):
    snap = read_ep_table(snap_path)
    ours = {r["id"]: r for r in rows}
    common = sorted(set(snap) & set(ours), key=int)
    norm = lambda s: (s or "").replace("내지", "속지").replace(" ", "")
    print(f"\n=== Cafe24 EP 비교: {os.path.basename(snap_path)} ===")
    print(f"  Cafe24 {len(snap)} / 우리 {len(ours)} / 공통 {len(common)} / "
          f"Cafe24에만 {len(set(snap)-set(ours))} / 우리에게만 {len(set(ours)-set(snap))}")
    print(f"\n  {'필드':28s} 일치율   (불일치 예시)")
    for col in COLUMNS:
        if col not in snap[next(iter(snap))]:
            continue
        diff = [(i, snap[i].get(col, ""), ours[i].get(col, "")) for i in common
                if (norm(snap[i].get(col, "")) != norm(ours[i].get(col, "")) if col.startswith("category")
                    else (snap[i].get(col, "") or "").strip() != (ours[i].get(col, "") or "").strip())]
        if col == "search_tag":   # 네이버 실효값(앞 10개/100자) 기준 비교
            diff = [(i, a, b) for i, a, b in diff if _tags(a.split("|") if a else []) != b]
        rate = 100 * (len(common) - len(diff)) / max(1, len(common))
        ex = f"#{diff[0][0]} '{diff[0][1][:18]}'→'{diff[0][2][:18]}'" if diff else ""
        mark = "  " if rate >= 99 else ("△ " if rate >= 90 else "✗ ")
        print(f"{mark}{col:28s} {rate:5.1f}%   {ex}")


# ═════════════════════════════════════════════════════════════════════════════
def _head(url):
    try:
        req = urllib.request.Request(url, method="HEAD", headers={"User-Agent": "Mozilla/5.0 orom-ep"})
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.status
    except Exception as e:
        return getattr(e, "code", 0)


def fix_images(rows):
    """Cafe24 _dburl 사본 이미지가 없는 상품(404) → 원본 목록이미지로 대체. 둘 다 없으면 오류로 남김."""
    from concurrent.futures import ThreadPoolExecutor
    urls = [r.get("image_link", "") for r in rows]
    with ThreadPoolExecutor(12) as ex:
        codes = list(ex.map(_head, urls))
    fixed, dead = 0, 0
    for r, c in zip(rows, codes):
        if c == 200:
            continue
        alt = r["image_link"].replace("_dburl", "")
        if alt != r["image_link"] and _head(alt) == 200:
            r["image_link"] = alt; fixed += 1
        else:
            r["_image_dead"] = c; dead += 1
    log(f"  이미지 응답 점검 {len(rows)}건 — 대체 {fixed}, 응답없음 {dead}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--mirror-only", action="store_true")
    ap.add_argument("--compare", metavar="SNAPSHOT")
    ap.add_argument("--init-override", action="store_true")
    a = ap.parse_args()

    log("=" * 64)
    log("네이버 EP 생성 v2 — 소스: " + ("구글시트" if SHEET_PRODUCTS_CSV else "Cafe24 API"))
    products, catinfo = load_products()
    log(f"  상품 {len(products)}")

    rows, excl = [], {}
    for p in sorted(products, key=lambda x: x["product_no"]):
        ci = catinfo.get(p["product_no"], {})
        why = exclude_reason(p, ci)
        if why:
            excl[why] = excl.get(why, 0) + 1
            continue
        rows.append(mirror_record(p, ci))
        if a.limit and len(rows) >= a.limit:
            break
    log(f"  미러 레코드 {len(rows)} / 제외 {sum(excl.values())}  " +
        ", ".join(f"{k} {v}" for k, v in sorted(excl.items(), key=lambda x: -x[1])))

    if a.init_override:
        from openpyxl import Workbook
        wb = Workbook(); ws = wb.active; ws.title = "overrides"
        cols = ["id", "_현재상품명"] + [c for c in COLUMNS if c not in LOCKED and c != "id"]
        ws.append(cols)
        for r in rows:
            ws.append([int(r["id"]), r["title"]] + [""] * (len(cols) - 2))
        wb.save(os.path.join(BASE, OVERRIDE_XLSX))
        log(f"  오버라이드 양식 생성 ({len(rows)}행)")
        return

    # 보강: Cafe24 EP 사본
    snap, _ = load_snapshot()
    filled = 0
    for r in rows:
        s = snap.get(r["id"])
        if s:
            for f in FROM_SNAPSHOT:
                if s.get(f):
                    r[f] = s[f]
            filled += 1
    if snap:
        log(f"  사본 보강 {filled}/{len(rows)}건 (리뷰수·주문제작·판매방식·이벤트문구)")

    # 사본에 없는 상품 → 추정 규칙
    inferred = 0
    for r in rows:
        if r["id"] in snap:
            continue
        cns = catinfo.get(int(r["id"]), {}).get("cns", [])
        for rule in INFER_RULES:
            if rule["category_no"] in cns:
                for k, v in rule["set"].items():
                    if not r.get(k):
                        r[k] = v
                inferred += 1
    if inferred:
        log(f"  추정 규칙 적용 {inferred}건 (사본 이후 신상품: 기업/단체 분류 → 주문제작·도매·이벤트문구)")

    if CHECK_IMAGES:
        fix_images(rows)
        dead = [r["id"] for r in rows if r.get("_image_dead")]
        if dead:
            log(f"  [제외] 대표이미지 응답없음 {dead[:10]} — 네이버 수집 오류 방지")
            rows[:] = [r for r in rows if not r.get("_image_dead")]

    if not a.mirror_only:
        ov = load_overrides()
        n = 0
        for r in rows:
            o = ov.get(r["id"])
            if o:
                r.update(o); n += 1
        if ov:
            log(f"  오버라이드 적용 {n}건")

    if a.compare:
        compare(rows, a.compare)
        return

    errs, warns = validate(rows)
    for w in warns[:15]: log(f"  [경고] {w}")
    if len(warns) > 15: log(f"  … 경고 {len(warns) - 15}건 더")
    for e in errs[:20]: log(f"  [오류] {e}")

    if os.path.exists(OUT_TSV):
        prev = sum(1 for _ in open(OUT_TSV, encoding="utf-8")) - 1
        if prev > 0 and abs(len(rows) - prev) / prev > 0.20 and not a.limit:
            raise SystemExit(f"[중단] 상품 수 급변 {prev}→{len(rows)} (±20%)")
    if errs:
        raise SystemExit(f"[중단] 오류 {len(errs)}건 — 기존 파일 유지")
    if a.dry_run:
        log(f"[dry-run] {len(rows)}건 검증 통과"); return

    clean = lambda v: re.sub(r"[\t\r\n\x00-\x1f]", " ", str(v)).strip()
    tmp = OUT_TSV + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:      # UTF-8 no-BOM, LF
        f.write("\t".join(COLUMNS) + "\n")
        for r in rows:
            f.write("\t".join(clean(r.get(k, "")) for k in COLUMNS) + "\n")
    os.replace(tmp, OUT_TSV)
    log(f"완료 {OUT_TSV} ({len(rows)}건, 경고 {len(warns)})")


if __name__ == "__main__":
    main()
