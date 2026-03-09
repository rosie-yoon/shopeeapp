"""
app.py - 메인 페이지 (프로필 기반 구글 시트 관리)
구글 시트 링크 입력 → Collection 탭 자동 읽기 → 파일 생성 → 다운로드
"""

import streamlit as st
import zipfile
import io
import pandas as pd
from pathlib import Path
import sys
import time
import re
import gc

sys.path.insert(0, str(Path(__file__).parent))

from sheet_reader import read_google_sheet, validate_dataframe, group_by_category, parse_category
from file_builder import build_all_files
from template_analyzer import load_auto_rules, load_global_rules
from gdrive_manager import get_gdrive_manager

st.set_page_config(
    page_title="쇼피 대량등록 도우미",
    page_icon="🛍️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════
# 프로필 관리 모달 (st.dialog 사용)
# ══════════════════════════════════════════════════════════════
@st.dialog("⚙️ 구글 시트 프로필 설정")
def manage_profiles_dialog():
    st.caption("자주 사용하는 구글 시트를 한글 이름으로 저장하고 팀과 공유하세요.")

    try:
        gdrive = get_gdrive_manager()
        profiles = gdrive.load_config_json("sheet_profiles.json") or {}
        use_gdrive = True
    except Exception:
        profiles = {}
        use_gdrive = False
        st.warning("⚠️ 구글 드라이브 연결 실패. 로컬 저장을 사용합니다.")

    # ── 새 프로필 추가 ──
    with st.form("add_profile_form", clear_on_submit=True):
        st.subheader("➕ 새 프로필 추가")

        col1, col2 = st.columns([1, 2])
        with col1:
            new_name = st.text_input(
                "프로필 이름",
            )
        with col2:
            new_url = st.text_input(
                "구글 시트 URL",
                placeholder="https://docs.google.com/spreadsheets/d/...",
                help="전체 URL을 복사해서 붙여넣으세요"
            )

        submitted = st.form_submit_button("✅ 프로필 저장", type="primary")

        if submitted:
            # 입력 검증
            if not new_name or not new_url:
                st.error("❌ 프로필 이름과 URL을 모두 입력해주세요.")
            elif not re.match(r"https://docs\.google\.com/spreadsheets/d/[a-zA-Z0-9_-]+", new_url):
                st.error("❌ 올바른 구글 시트 URL 형식이 아닙니다.")
            elif new_name in profiles:
                st.error(f"❌ '{new_name}' 프로필이 이미 존재합니다.")
            else:
                # 프로필 저장
                profiles[new_name] = new_url

                if use_gdrive:
                    try:
                        gdrive.save_config_json("sheet_profiles.json", profiles)
                        st.success(f"✅ '{new_name}' 프로필이 구글 드라이브에 저장되었습니다!")
                    except Exception as e:
                        st.error(f"❌ 저장 실패: {e}")
                else:
                    # 로컬 저장 fallback
                    local_path = Path(__file__).parent / "config" / "sheet_profiles.json"
                    local_path.parent.mkdir(parents=True, exist_ok=True)
                    import json
                    with open(local_path, "w", encoding="utf-8") as f:
                        json.dump(profiles, f, ensure_ascii=False, indent=2)
                    st.success(f"✅ '{new_name}' 프로필이 로컬에 저장되었습니다!")

                time.sleep(1)
                st.rerun()

    st.divider()

    # ── 기존 프로필 관리 ──
    st.subheader("📋 저장된 프로필 관리")

    if not profiles:
        st.info("저장된 프로필이 없습니다. 위에서 새 프로필을 추가해보세요.")
    else:
        # 프로필 목록을 데이터프레임으로 표시
        df_profiles = pd.DataFrame([
            {"프로필 이름": name, "구글 시트 URL": url[:50] + "..." if len(url) > 50 else url}
            for name, url in profiles.items()
        ])
        st.dataframe(df, hide_index=True, width="stretch")

        # 프로필 삭제 기능
        col1, col2 = st.columns([3, 1])
        with col1:
            if profiles:
                delete_target = st.selectbox(
                    "삭제할 프로필 선택",
                    options=list(profiles.keys()),
                    key="delete_profile_select"
                )
        with col2:
            st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("🗑️ 삭제", type="secondary", key="delete_profile_btn"):
                if delete_target and delete_target in profiles:
                    del profiles[delete_target]

                    # 저장
                    if use_gdrive:
                        try:
                            gdrive.save_config_json("sheet_profiles.json", profiles)
                            st.success(f"✅ '{delete_target}' 프로필이 삭제되었습니다!")
                        except Exception as e:
                            st.error(f"❌ 삭제 실패: {e}")
                    else:
                        local_path = Path(__file__).parent / "config" / "sheet_profiles.json"
                        import json
                        with open(local_path, "w", encoding="utf-8") as f:
                            json.dump(profiles, f, ensure_ascii=False, indent=2)
                        st.success(f"✅ '{delete_target}' 프로필이 삭제되었습니다!")

                    time.sleep(1)
                    st.rerun()

# ══════════════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/f/fe/Shopee.svg/200px-Shopee.svg.png", width=120)

    # 구글 드라이브 연결 상태 표시
    try:
        gdrive = get_gdrive_manager()
        st.success("✅ 구글 드라이브 연결됨 (OAuth 2.0)")
        st.markdown(f"📁 [공유 폴더 열기]({gdrive.get_folder_link()})")

        # 프로필 데이터 로드
        sheet_profiles = gdrive.load_config_json("sheet_profiles.json") or {}
    except Exception as e:
        st.warning("⚠️ 구글 드라이브 연결 실패")
        st.caption("로컬 모드로 실행 중")

        # 로컬 프로필 로드
        local_path = Path(__file__).parent / "config" / "sheet_profiles.json"
        if local_path.exists():
            import json
            with open(local_path, "r", encoding="utf-8") as f:
                sheet_profiles = json.load(f)
        else:
            sheet_profiles = {}

    st.divider()

    st.markdown("### 사용 방법")
    st.markdown("""
    1. **프로필 선택** 또는 직접 입력
    2. **[데이터 불러오기]** 클릭
    3. 데이터 확인 후 파일 생성
    4. 다운로드
    """)

    auto_rules = load_auto_rules()
    global_rules = load_global_rules()

    st.divider()
    st.caption(f"등록된 카테고리: **{len(auto_rules)}개**")
    st.caption(f"저장된 프로필: **{len(sheet_profiles)}개**")

    # ── 프로필 설정 버튼 (사이드바 하단) ──
    st.divider()
    if st.button("⚙️ 프로필 설정", use_container_width=True, help="구글 시트 프로필 관리"):
        manage_profiles_dialog()

# ══════════════════════════════════════════════════════════════
# 메인 페이지
# ══════════════════════════════════════════════════════════════
st.title("🛍️ 쇼피 대량등록 파일 생성")
st.markdown("구글 시트의 **Collection 탭** 데이터를 쇼피 업로드 템플릿으로 변환합니다.")
st.divider()

# ════════════════════════════════════════════════════════════
# Step 1: 프로필 선택 및 구글 시트 URL 입력
# ════════════════════════════════════════════════════════════
st.subheader("① 구글 시트 선택")

st.info(
    "**시트 공개 설정 필요**  \n"
    "구글 시트 우측 상단 **[공유]** → '링크가 있는 모든 사용자' → **뷰어** 로 설정해주세요."
)

col_profile, col_url = st.columns([1, 2])

# 프로필 선택 드롭다운
with col_profile:
    profile_options = ["직접 입력"] + list(sheet_profiles.keys())
    selected_profile = st.selectbox(
        "프로필 선택",
        options=profile_options,
        index=0,
        help="저장된 프로필을 선택하거나 '직접 입력'을 선택하세요"
    )

    if len(sheet_profiles) == 0:
        st.caption("💡 사이드바의 **⚙️ 프로필 설정**에서 자주 사용하는 시트를 저장하세요")

# URL 입력창 (프로필 선택 시 자동 입력)
with col_url:
    # 프로필 선택에 따른 초기값 설정
    if selected_profile != "직접 입력" and selected_profile in sheet_profiles:
        initial_url = sheet_profiles[selected_profile]
        # session_state 업데이트
        st.session_state["sheet_url"] = initial_url
    else:
        initial_url = st.session_state.get("sheet_url", "")

    sheet_url = st.text_input(
        "구글 시트 URL",
        placeholder="https://docs.google.com/spreadsheets/d/...",
        key="sheet_url_input",
        value=initial_url,
    )

load_btn = st.button("📥 데이터 불러오기", type="primary", disabled=not sheet_url)

if load_btn and sheet_url:
    st.session_state["sheet_url"] = sheet_url
    st.session_state.pop("df", None)

    with st.spinner("Collection 탭 데이터 불러오는 중..."):
        try:
            df = read_google_sheet(sheet_url, tab_name="Collection")
            st.session_state["df"] = df
            st.success(f"✅ Collection 탭에서 **{len(df)}행** 불러오기 완료!")
        except ValueError as e:
            st.error(str(e))

# ════════════════════════════════════════════════════════════
# Step 2: 데이터 확인 (기존 코드 유지)
# ════════════════════════════════════════════════════════════
if "df" in st.session_state:
    df = st.session_state["df"]

    st.divider()
    st.subheader("② 데이터 확인")

    warnings = validate_dataframe(df)
    for w in warnings:
        st.warning(w)

    groups = group_by_category(df)

    col1, col2 = st.columns([1, 2])

    with col1:
        st.markdown("**📊 카테고리별 상품 수**")
        summary_data = []
        for cat_id, group_df in groups.items():
            cat_info = auto_rules.get(cat_id)
            if cat_info:
                cat_name = cat_info["category_path"]
                status = "✅ 템플릿 있음"
            else:
                sample_cat = group_df.iloc[0].get("Category", cat_id)
                _, cat_name = parse_category(str(sample_cat))
                status = "⚠️ 템플릿 없음"
            summary_data.append({
                "카테고리 ID": cat_id,
                "카테고리명": cat_name[:35] + "..." if len(cat_name) > 35 else cat_name,
                "상품 수": len(group_df),
                "상태": status,
            })
        st.dataframe(pd.DataFrame(summary_data), hide_index=True, use_container_width=True)

    with col2:
        st.markdown("**📋 데이터 미리보기** (처음 5행)")
        preview_cols = ["Category", "Product Name", "Global SKU Price", "Stock", "Brand"]
        preview_cols = [c for c in preview_cols if c in df.columns]
        st.dataframe(
            df[preview_cols].head(5) if preview_cols else df.head(5),
            hide_index=True, use_container_width=True
        )

    # ════════════════════════════════════════════════════════════
    # Step 3: 파일 생성 (기존 코드 유지)
    # ════════════════════════════════════════════════════════════
    st.divider()
    st.subheader("③ 파일 생성 및 다운로드")

    missing_templates = [
        cat_id for cat_id in groups
        if cat_id not in auto_rules and cat_id != "unknown"
    ]
    if missing_templates:
        st.warning(
            f"⚠️ 아래 카테고리의 템플릿이 없습니다: `{'`, `'.join(missing_templates)}`  \n"
            "왼쪽 메뉴 **📋 템플릿 관리** 에서 해당 템플릿을 업로드해주세요."
        )

    processable = [cat_id for cat_id in groups if cat_id in auto_rules]

    if not processable:
        st.error("처리 가능한 카테고리가 없습니다. 먼저 템플릿을 업로드해주세요.")
    else:
        st.markdown(f"**{len(processable)}개 카테고리**, 총 **{len(df)}개 상품** 파일을 생성합니다.")

        # 파일 생성 버튼 로직에서 수정
        if st.button("🚀 파일 생성하기", type="primary"):
            with st.spinner("파일 생성 중..."):
                try:
                    results, skipped = build_all_files(groups, auto_rules, global_rules)

                    # ── [추가] 메모리 명시적 정리 ──
                    for msg in skipped:
                        st.warning(f"건너뜸: {msg}")

                    if results:
                        # ... (기존 다운로드 버튼 코드 유지) ...

                        # ── [추가] 처리 완료 후 메모리 정리 ──
                        del results, skipped
                        gc.collect()
                    else:
                        st.error("생성된 파일이 없습니다.")

                except Exception as e:
                    st.error(f"❌ 파일 생성 중 오류: {e}")
                    # ── [추가] 오류 시에도 메모리 정리 ──
                    gc.collect()
