"""
app.py - 메인 페이지
구글 시트 링크 입력 → Collection 탭 자동 읽기 → 파일 생성 → 다운로드
"""

import streamlit as st
import zipfile
import io
import pandas as pd
from pathlib import Path
import sys
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
# 사이드바
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/f/fe/Shopee.svg/200px-Shopee.svg.png", width=120)

    # 구글 드라이브 연결 상태 표시
    try:
        gdrive = get_gdrive_manager()
        st.success("✅ 구글 드라이브 연결됨 (OAuth 2.0)")
        st.markdown(f"📁 [공유 폴더 열기]({gdrive.get_folder_link()})")
    except Exception as e:
        st.warning("⚠️ 구글 드라이브 연결 실패")
        st.caption("로컬 모드로 실행 중")
        with st.expander("연결 실패 상세 정보"):
            st.error(str(e))
            st.markdown("""
            **OAuth 2.0 인증 해결 방법:**
            1. `credentials.json` 파일이 프로젝트 루트에 있는지 확인
            2. Google Cloud Console에서 Drive API 활성화 확인
            3. OAuth 동의 화면에서 테스트 사용자 등록 확인
            4. 최초 실행 시 브라우저에서 구글 계정 인증 필요
            5. 토큰 만료 시: `rm token.json` 후 재실행
            """)

    st.divider()

    st.markdown("### 사용 방법")
    st.markdown("""
    1. 구글 시트 링크 붙여넣기
    2. **[데이터 불러오기]** 클릭
       → **Collection 탭** 자동으로 읽어옴
    3. 데이터 확인 후 파일 생성
    4. 다운로드
    
    ---
    ⚙️ 관리 메뉴:
    - **📋 템플릿 관리**: 새 템플릿 업로드
    - **⚙️ 필수값 관리**: FDA번호 등 고정값 설정
    """)

    auto_rules = load_auto_rules()
    global_rules = load_global_rules()

    st.divider()
    st.caption(f"등록된 카테고리: **{len(auto_rules)}개**")
    st.caption(f"설정된 고정값: **{len(global_rules)}개**")

# ══════════════════════════════════════════════════════════════
# 메인 페이지
# ══════════════════════════════════════════════════════════════
st.title("🛍️ 쇼피 대량등록 파일 생성")
st.markdown("구글 시트의 **Collection 탭** 데이터를 쇼피 업로드 템플릿으로 변환합니다.")
st.divider()

# ════════════════════════════════════════════════════════════
# Step 1: 구글 시트 URL 입력
# ════════════════════════════════════════════════════════════
st.subheader("① 구글 시트 링크 입력")

st.info(
    "**시트 공개 설정 필요**  \n"
    "구글 시트 우측 상단 **[공유]** → '링크가 있는 모든 사용자' → **뷰어** 로 설정해주세요."
)

sheet_url = st.text_input(
    "구글 시트 URL",
    placeholder="https://docs.google.com/spreadsheets/d/...",
    key="sheet_url_input",
    value=st.session_state.get("sheet_url", ""),
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
# Step 2: 데이터 확인
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
    # Step 3: 파일 생성
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

        if st.button("🚀 파일 생성하기", type="primary"):
            with st.spinner("파일 생성 중..."):
                try:
                    results, skipped = build_all_files(groups, auto_rules, global_rules)

                    for msg in skipped:
                        st.warning(f"건너뜀: {msg}")

                    if results:
                        if len(results) == 1:
                            filename, file_bytes = next(iter(results.items()))
                            st.success("✅ 파일 생성 완료!")
                            st.download_button(
                                label=f"⬇️ {filename} 다운로드",
                                data=file_bytes,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                for fname, fbytes in results.items():
                                    zf.writestr(fname, fbytes)
                            zip_buffer.seek(0)

                            st.success(f"✅ {len(results)}개 파일 생성 완료!")
                            for fname in results:
                                st.markdown(f"  📄 `{fname}`")

                            st.download_button(
                                label=f"📦 ZIP 다운로드 ({len(results)}개 파일)",
                                data=zip_buffer.getvalue(),
                                file_name="shopee_upload_files.zip",
                                mime="application/zip",
                                type="primary",
                            )
                    else:
                        st.error("생성된 파일이 없습니다.")

                except Exception as e:
                    st.error(f"❌ 파일 생성 중 오류: {e}")
                    import traceback
                    st.code(traceback.format_exc())
