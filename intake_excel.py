"""
intake_excel.py

상점마다 제각각인 엑셀 양식(컬럼명, 헤더 위치, 빈 행, 요약 행)을 자동으로 감지하고
표준 스키마(merchant_id, request_type, submitted_at, note)로 정형화한다.

실무에서는 카탈로그 수정 요청이 시스템이 아니라 상점 사장님이 보낸 엑셀로도 들어오는 경우가 많다.
이 스크립트는 그런 비정형 입력을 기존 catalog_requests 파이프라인에 태울 수 있게 전처리하는 역할을 한다.
"""
import re
import pandas as pd
import openpyxl

# 실제로 접할 법한 컬럼명 변형들을 표준 컬럼명으로 매핑
COLUMN_ALIASES = {
    "요청타입": "request_type", "구분": "request_type", "type": "request_type",
    "상점코드": "merchant_id", "점포id": "merchant_id", "store_id": "merchant_id",
    "요청일": "submitted_at", "등록일자": "submitted_at", "date": "submitted_at",
    "내용": "note", "상세내용": "note", "note": "note",
}

REQUEST_TYPE_ALIASES = {
    "가격변경": "가격_수정", "메뉴추가": "신규_메뉴_등록", "신규메뉴등록": "신규_메뉴_등록",
    "영업시간변경": "영업시간_휴무일_변경", "이미지교체": "이미지_교체",
    "메뉴설명수정": "메뉴_설명_수정",
}

REQUIRED_COLS = {"request_type", "merchant_id", "submitted_at"}


def find_header_row(path: str) -> int:
    """표준 컬럼명 중 2개 이상이 등장하는 첫 행을 헤더로 판단한다."""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        cells = [str(c).strip().lower() for c in row if c is not None]
        hits = sum(1 for c in cells if c in COLUMN_ALIASES)
        if hits >= 2:
            return i
    raise ValueError(f"헤더 행을 찾지 못했습니다: {path}")


def parse_merchant_excel(path: str) -> pd.DataFrame:
    header_row = find_header_row(path)
    df = pd.read_excel(path, header=header_row)

    # 컬럼명 표준화 (소문자 매칭)
    rename_map = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[key]
    df = df.rename(columns=rename_map)

    # 표준 컬럼이 없으면 스킵
    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        raise ValueError(f"{path}: 필수 컬럼 누락 {missing}")

    df = df[list(REQUIRED_COLS) + (["note"] if "note" in df.columns else [])]

    # 빈 행 / 요약 행 제거: merchant_id가 'M'으로 시작하는 실제 데이터만 남김
    df = df[df["merchant_id"].astype(str).str.match(r"^M\d+$", na=False)].copy()

    # 요청 유형 표준화
    df["request_type"] = df["request_type"].map(REQUEST_TYPE_ALIASES).fillna(df["request_type"])

    # 날짜 표준화
    df["submitted_at"] = pd.to_datetime(df["submitted_at"], errors="coerce")

    return df.reset_index(drop=True)


if __name__ == "__main__":
    import glob
    all_rows = []
    for path in glob.glob("merchant_*.xlsx"):
        try:
            parsed = parse_merchant_excel(path)
            parsed["source_file"] = path
            all_rows.append(parsed)
            print(f"{path}: {len(parsed)}건 정형화 완료")
        except Exception as e:
            print(f"{path}: 파싱 실패 - {e}")

    result = pd.concat(all_rows, ignore_index=True)
    print(f"\n총 {len(result)}건 표준화 완료")
    print(result.to_string(index=False))
    result.to_csv("intake_standardized.csv", index=False)
