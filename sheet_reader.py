"""
sheet_reader.py
구글 시트에서 데이터를 읽어오는 모듈
"""

import re
import pandas as pd


def parse_category(cat_str: str) -> tuple[str, str]:
    """
    "101645 - Beauty/Makeup/Lips/Lip Tint & Stain"
    → ("101645", "Beauty/Makeup/Lips/Lip Tint & Stain")
    """
    cat_str = str(cat_str).strip()
    match = re.match(r"^(\d+)\s*[-–]\s*(.+)$", cat_str)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    if cat_str.isdigit():
        return cat_str, ""
    return "", cat_str


def extract_spreadsheet_id(url: str) -> str:
    """URL에서 스프레드시트 ID 추출"""
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        raise ValueError("올바른 구글 시트 URL이 아닙니다.")
    return match.group(1)


def read_google_sheet(url: str, tab_name: str = "Collection") -> pd.DataFrame:
    """
    구글 시트 URL에서 탭 이름으로 DataFrame 읽기
    gviz API를 사용하므로 gid 없이 탭 이름만으로 접근 가능

    Parameters
    ----------
    url : str
        구글 시트 URL (공개 설정 필요)
    tab_name : str
        불러올 탭 이름 (기본값: "Collection")
    """
    spreadsheet_id = extract_spreadsheet_id(url)

    # gviz API: 탭 이름으로 직접 CSV 추출 (gid 불필요)
    # URL 인코딩 처리
    from urllib.parse import quote
    encoded_tab = quote(tab_name)
    csv_url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
        f"/gviz/tq?tqx=out:csv&sheet={encoded_tab}"
    )

    try:
        df = pd.read_csv(csv_url, dtype=str)
        df = df.fillna("")
        # 완전히 빈 행 제거
        df = df[df.apply(lambda row: row.str.strip().ne("").any(), axis=1)].reset_index(drop=True)

        if df.empty:
            raise ValueError(
                f"'{tab_name}' 탭이 비어있거나 존재하지 않습니다.\n\n"
                f"다음을 확인해주세요:\n"
                f"1. 구글 시트에 '{tab_name}' 탭이 있는지 확인\n"
                f"2. 탭 이름 대소문자가 정확한지 확인 (예: 'Collection' vs 'collection')\n"
                f"3. 시트가 '링크가 있는 모든 사용자 - 뷰어'로 공개 설정되어 있는지 확인"
            )

        return df

    except Exception as e:
        if "탭이 비어있거나" in str(e):
            raise
        raise ValueError(
            f"'{tab_name}' 탭을 불러올 수 없습니다: {e}\n\n"
            f"다음을 확인해주세요:\n"
            f"1. 구글 시트에 '{tab_name}' 탭이 있는지 확인\n"
            f"2. 시트가 '링크가 있는 모든 사용자 - 뷰어'로 공개 설정되어 있는지 확인"
        )


def validate_dataframe(df: pd.DataFrame) -> list[str]:
    """필수 컬럼 존재 여부 확인"""
    required_cols = [
        "Category", "Product Name", "Global SKU Price",
        "Stock", "Cover image", "Weight", "Days to ship", "Brand"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    warnings = []
    if missing:
        warnings.append(f"⚠️ 필수 컬럼 누락: {', '.join(missing)}")
    return warnings


def group_by_category(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """데이터프레임을 소카테고리 ID별로 그룹핑"""
    groups = {}
    for _, row in df.iterrows():
        cat_str = row.get("Category", "")
        cat_id, _ = parse_category(cat_str)
        if not cat_id:
            cat_id = "unknown"
        if cat_id not in groups:
            groups[cat_id] = []
        groups[cat_id].append(row)

    return {
        cat_id: pd.DataFrame(rows).reset_index(drop=True)
        for cat_id, rows in groups.items()
    }
