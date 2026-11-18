import os
import re
import time
from pathlib import Path
from difflib import SequenceMatcher

import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill


# ============================================================
# USER SETTINGS
# ============================================================
INPUT_FILE = "OpenAlex_Input.xlsx"
OUTPUT_FILE = "OpenAlex_Output.xlsx"
SHEET_NAME = "DOI 찾기"
HEADER_ROW = 1

OPENALEX_API_URL = "https://api.openalex.org/works"
API_KEY = os.getenv("OPENALEX_API_KEY", "").strip()  # optional

BATCH_SIZE = 100
TIMEOUT = 30
MAX_RETRIES = 4
REQUEST_DELAY = 0.25

# Cross-check thresholds. These are review aids, not bibliographic truth rules.
TITLE_OK_THRESHOLD = 0.90
TITLE_REVIEW_THRESHOLD = 0.75
JOURNAL_OK_THRESHOLD = 0.85
JOURNAL_REVIEW_THRESHOLD = 0.65


# ============================================================
# INPUT / OUTPUT HEADERS
# ============================================================
H_TITLE = "논문명"
H_AUTHOR = "주저자명"
H_JOURNAL = "학술지명"
H_DATE = "게재일"
H_DOI = "DOI"

OUTPUT_HEADERS = [
    "OA 논문명",
    "OA 주저자",
    "OA 학술지명",
    "OA 게재일",
    "Cited by",
    "FWCI",
    "OpenAlex ID",
    "OA DOI",
    "OpenAlex Updated",
    "Cross-check",
]


# ============================================================
# NORMALIZATION
# ============================================================
def normalize_doi(value):
    if value is None:
        return None

    doi = str(value).strip()
    if not doi:
        return None

    doi = re.sub(r"^doi:\s*", "", doi, flags=re.IGNORECASE)
    doi = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", doi, flags=re.IGNORECASE)
    doi = doi.strip().strip("<>[]{}()")
    doi = doi.rstrip(".,;:")

    if not doi.lower().startswith("10.") or "/" not in doi:
        return None

    return doi.lower()


def normalize_text(value):
    if value is None:
        return ""

    text = str(value).lower().strip()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9가-힣]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def compact_person_name(value):
    """Loose author-name comparison key: letters/digits only, no punctuation/spaces."""
    return re.sub(r"[^a-z0-9가-힣]", "", normalize_text(value))


def similarity(a, b):
    a = normalize_text(a)
    b = normalize_text(b)
    if not a or not b:
        return None
    return SequenceMatcher(None, a, b).ratio()


def get_year(value):
    if value is None:
        return None

    if hasattr(value, "year"):
        return value.year

    match = re.search(r"\b(19|20)\d{2}\b", str(value))
    return int(match.group(0)) if match else None


# ============================================================
# OPENALEX FIELD HELPERS
# ============================================================
def get_first_author(work):
    authorships = work.get("authorships") or []
    if not authorships:
        return ""
    return ((authorships[0].get("author") or {}).get("display_name") or "")


def get_journal(work):
    primary_location = work.get("primary_location") or {}
    source = primary_location.get("source") or {}
    return source.get("display_name") or ""


# ============================================================
# CROSS-CHECK
# ============================================================
def author_matches(local_authors, oa_first_author):
    """
    The legacy author cell may contain all authors in inconsistent formats.
    We therefore test whether the OpenAlex first-author tokens are reasonably
    represented in the local author string instead of exact-matching the cell.
    """
    if not local_authors or not oa_first_author:
        return None

    local_compact = compact_person_name(local_authors)
    oa_compact = compact_person_name(oa_first_author)

    if oa_compact and oa_compact in local_compact:
        return True

    # Also compare name tokens to tolerate 'Oh, Sangjun' vs 'Sangjun Oh'.
    oa_tokens = [t for t in normalize_text(oa_first_author).split() if len(t) > 1]
    if not oa_tokens:
        return None

    local_text = normalize_text(local_authors)
    hit_count = sum(1 for token in oa_tokens if token in local_text)
    return hit_count == len(oa_tokens)


def make_cross_check(local_title, local_authors, local_journal, local_date,
                     oa_title, oa_first_author, oa_journal, oa_date):
    flags = []

    title_score = similarity(local_title, oa_title)
    if title_score is not None:
        if title_score < TITLE_REVIEW_THRESHOLD:
            flags.append(f"TITLE {title_score:.0%}")
        elif title_score < TITLE_OK_THRESHOLD:
            flags.append(f"TITLE? {title_score:.0%}")

    author_ok = author_matches(local_authors, oa_first_author)
    if author_ok is False:
        flags.append("AUTHOR")

    journal_score = similarity(local_journal, oa_journal)
    if journal_score is not None:
        if journal_score < JOURNAL_REVIEW_THRESHOLD:
            flags.append(f"JOURNAL {journal_score:.0%}")
        elif journal_score < JOURNAL_OK_THRESHOLD:
            flags.append(f"JOURNAL? {journal_score:.0%}")

    local_year = get_year(local_date)
    oa_year = get_year(oa_date)
    if local_year and oa_year and local_year != oa_year:
        flags.append(f"YEAR {local_year}/{oa_year}")

    return "OK" if not flags else "CHECK: " + " / ".join(flags)


# ============================================================
# EXCEL HELPERS
# ============================================================
def header_map(ws):
    result = {}
    for cell in ws[HEADER_ROW]:
        if cell.value is not None and str(cell.value).strip():
            result[str(cell.value).strip()] = cell.column
    return result


def ensure_output_headers(ws, headers):
    """Append output columns after the existing used columns, preserving A:I."""
    hmap = header_map(ws)
    next_col = ws.max_column + 1

    for name in headers:
        if name not in hmap:
            ws.cell(HEADER_ROW, next_col).value = name
            hmap[name] = next_col
            next_col += 1

    return hmap


def copy_header_style(ws, source_col, target_cols):
    source = ws.cell(HEADER_ROW, source_col)
    for col in target_cols:
        target = ws.cell(HEADER_ROW, col)
        if source.has_style:
            target._style = source._style
        else:
            target.font = Font(bold=True)
            target.fill = PatternFill("solid", fgColor="D9EAF7")


# ============================================================
# OPENALEX API
# ============================================================
def query_batch(session, dois):
    if not dois:
        return {}

    # OpenAlex supports OR filters with up to 100 values.
    doi_filter = "|".join(f"https://doi.org/{doi}" for doi in dois)

    params = {
        "filter": f"doi:{doi_filter}",
        "per_page": 100,
        "select": (
            "id,doi,display_name,publication_date,authorships,"
            "primary_location,cited_by_count,fwci,updated_date"
        ),
    }

    if API_KEY:
        params["api_key"] = API_KEY

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(OPENALEX_API_URL, params=params, timeout=TIMEOUT)

            if response.status_code in {429, 500, 502, 503, 504}:
                wait = min(2 ** attempt, 20)
                print(f"  HTTP {response.status_code}; {wait}s 후 재시도")
                time.sleep(wait)
                continue

            response.raise_for_status()
            payload = response.json()

            found = {}
            for work in payload.get("results", []):
                doi = normalize_doi(work.get("doi"))
                if doi:
                    found[doi] = work
            return found

        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            if attempt < MAX_RETRIES:
                wait = min(2 ** attempt, 20)
                print(f"  요청 오류: {exc}; {wait}s 후 재시도")
                time.sleep(wait)
            else:
                raise RuntimeError(f"OpenAlex API 요청 실패: {last_error}") from exc

    return {}


# ============================================================
# MAIN
# ============================================================
def main():
    input_path = Path(INPUT_FILE)
    output_path = Path(OUTPUT_FILE)

    if not input_path.exists():
        raise FileNotFoundError(f"입력 파일이 없습니다: {input_path.resolve()}")

    wb = load_workbook(input_path)

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"시트 '{SHEET_NAME}'를 찾을 수 없습니다. 실제 시트: {wb.sheetnames}")

    ws = wb[SHEET_NAME]
    hmap = header_map(ws)

    for required in (H_TITLE, H_AUTHOR, H_JOURNAL, H_DATE, H_DOI):
        if required not in hmap:
            raise ValueError(f"필수 열이 없습니다: {required}")

    # Append output headers only. Existing input columns are preserved.
    hmap = ensure_output_headers(ws, OUTPUT_HEADERS)
    copy_header_style(ws, hmap[H_DOI], [hmap[h] for h in OUTPUT_HEADERS])

    doi_rows = {}
    invalid_doi_rows = []

    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        raw = ws.cell(row, hmap[H_DOI]).value
        if raw is None or not str(raw).strip():
            continue

        doi = normalize_doi(raw)
        if not doi:
            invalid_doi_rows.append(row)
            ws.cell(row, hmap["Cross-check"]).value = "INVALID DOI"
            continue

        doi_rows.setdefault(doi, []).append(row)

    unique_dois = list(doi_rows)
    print(f"전체 데이터 행: {ws.max_row - HEADER_ROW}")
    print(f"유효 고유 DOI: {len(unique_dois)}")
    print(f"형식 오류 DOI 행: {len(invalid_doi_rows)}")

    session = requests.Session()
    session.headers.update({
        "User-Agent": "OpenAlex-XLSX-Updater/1.0"
    })

    all_results = {}
    total_batches = (len(unique_dois) + BATCH_SIZE - 1) // BATCH_SIZE

    for i in range(0, len(unique_dois), BATCH_SIZE):
        batch_no = i // BATCH_SIZE + 1
        batch = unique_dois[i:i + BATCH_SIZE]
        print(f"[{batch_no}/{total_batches}] OpenAlex {len(batch)}건 조회")
        result = query_batch(session, batch)
        all_results.update(result)
        print(f"  매칭 {len(result)}/{len(batch)}")
        time.sleep(REQUEST_DELAY)

    matched_rows = 0
    not_found_rows = 0
    review_rows = 0

    for doi, rows in doi_rows.items():
        work = all_results.get(doi)

        if work is None:
            for row in rows:
                ws.cell(row, hmap["Cross-check"]).value = "NOT FOUND"
            not_found_rows += len(rows)
            continue

        oa_title = work.get("display_name") or ""
        oa_author = get_first_author(work)
        oa_journal = get_journal(work)
        oa_date = work.get("publication_date")
        oa_doi = normalize_doi(work.get("doi"))

        for row in rows:
            ws.cell(row, hmap["OA 논문명"]).value = oa_title
            ws.cell(row, hmap["OA 주저자"]).value = oa_author
            ws.cell(row, hmap["OA 학술지명"]).value = oa_journal
            ws.cell(row, hmap["OA 게재일"]).value = oa_date
            ws.cell(row, hmap["Cited by"]).value = work.get("cited_by_count")
            ws.cell(row, hmap["FWCI"]).value = work.get("fwci")
            ws.cell(row, hmap["OpenAlex ID"]).value = work.get("id")
            ws.cell(row, hmap["OA DOI"]).value = oa_doi
            ws.cell(row, hmap["OpenAlex Updated"]).value = work.get("updated_date")

            verdict = make_cross_check(
                ws.cell(row, hmap[H_TITLE]).value,
                ws.cell(row, hmap[H_AUTHOR]).value,
                ws.cell(row, hmap[H_JOURNAL]).value,
                ws.cell(row, hmap[H_DATE]).value,
                oa_title,
                oa_author,
                oa_journal,
                oa_date,
            )
            ws.cell(row, hmap["Cross-check"]).value = verdict

            if verdict.startswith("CHECK"):
                review_rows += 1
            matched_rows += 1

    # Basic formatting for appended columns.
    widths = {
        "OA 논문명": 60,
        "OA 주저자": 24,
        "OA 학술지명": 42,
        "OA 게재일": 14,
        "Cited by": 12,
        "FWCI": 12,
        "OpenAlex ID": 28,
        "OA DOI": 38,
        "OpenAlex Updated": 18,
        "Cross-check": 36,
    }
    for name, width in widths.items():
        ws.column_dimensions[ws.cell(HEADER_ROW, hmap[name]).column_letter].width = width

    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        ws.cell(row, hmap["FWCI"]).number_format = "0.00"
        ws.cell(row, hmap["Cited by"]).number_format = "0"

    wb.save(output_path)

    print("\n==============================")
    print("완료")
    print("==============================")
    print(f"OpenAlex 매칭 행 : {matched_rows}")
    print(f"NOT FOUND       : {not_found_rows}")
    print(f"CHECK 필요      : {review_rows}")
    print(f"INVALID DOI     : {len(invalid_doi_rows)}")
    print(f"출력 파일       : {output_path.resolve()}")


if __name__ == "__main__":
    main()
