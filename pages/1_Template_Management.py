"""
pages/1_Template_Management.py
쇼피 대량등록 템플릿 업로드 및 분석 페이지 (구글 드라이브 연동)
"""

import streamlit as st
from pathlib import Path
import sys

# 부모 디렉토리를 import 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from template_analyzer import (
    analyze_template,
    save_auto_rules,
    load_auto_rules,
    extract_template_info
)
from gdrive_manager import get_gdrive_manager

# ════════════════════════════════════════
# 페이지 기본 설정
# ════════════════════════════════════════
st.set_page_config(
    page_title="📋 템플릿 관리",
    page_icon="📋",
    layout="wide"
)

st.title("📋 템플릿 관리")
st.caption("쇼피 대량등록 템플릿을 업로드하면 자동으로 분석하여 구글 드라이브에 저장합니다.")

# ════════════════════════════════════════
# 구글 드라이브 연결 상태 확인
# ════════════════════════════════════════
try:
    gdrive = get_gdrive_manager()
    st.success(f"✅ 구글 드라이브 연결됨")
    st.markdown(f"📁 [공유 폴더 열기]({gdrive.get_folder_link()})")
    USE_GDRIVE = True
except Exception as e:
    st.warning("⚠️ 구글 드라이브 연결에 실패했습니다. 로컬 모드로 동작합니다.")
    with st.expander("연결 실패 상세 정보"):
        st.error(str(e))
        st.markdown("""
        **해결 방법:**
        1. `service_account.json` 파일이 프로젝트 루트에 있는지 확인
        2. Google Cloud Console에서 Drive API가 활성화되어 있는지 확인
        3. 서비스 계정에 적절한 권한이 부여되어 있는지 확인
        """)
    USE_GDRIVE = False

st.divider()

# ════════════════════════════════════════
# 등록된 템플릿 목록
# ════════════════════════════════════════
auto_rules = load_auto_rules()

st.subheader("📁 등록된 템플릿")

# 저장 위치에 따른 템플릿 목록 조회
if USE_GDRIVE:
    with st.spinner("구글 드라이브에서 템플릿 목록을 불러오는 중..."):
        saved_templates = gdrive.list_templates()
    storage_info = "🌐 구글 드라이브"
else:
    # 로컬 fallback
    from file_builder import TEMPLATES_DIR
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    saved_templates = [f.name for f in sorted(TEMPLATES_DIR.glob("*.xlsx"))]
    storage_info = "💻 로컬 (templates/)"

st.caption(f"저장 위치: {storage_info}")

# 분석 완료된 템플릿 파일명 목록
analyzed_files = set(v.get("template_file", "") for v in auto_rules.values())

if saved_templates:
    # 템플릿 목록을 2열로 표시
    col1, col2 = st.columns(2)
    for idx, filename in enumerate(saved_templates):
        target_col = col1 if idx % 2 == 0 else col2

        if filename in analyzed_files:
            target_col.markdown(f"- ✅ `{filename}` (분석 완료)")
        else:
            target_col.markdown(f"- ⚠️ `{filename}` (분석 결과 없음)")
else:
    st.info("아직 등록된 템플릿이 없습니다.")

st.divider()

# ════════════════════════════════════════
# 템플릿 업로드
# ════════════════════════════════════════
st.subheader("⬆️ 템플릿 업로드")

st.markdown("""
**업로드하면 자동으로:**
1. `Template` 시트 **C2 코드**로 동일 템플릿 여부 판별
2. `Pre-order DTS Range` 시트에서 **대카테고리 + 중카테고리명** 추출 → 파일명 자동 결정
3. `HiddenCatProps` 시트 분석 → 소카테고리별 필수 입력 항목 추출
4. 같은 C2 코드의 **기존 데이터 교체**
""")

if USE_GDRIVE:
    st.info("📤 업로드된 파일은 구글 드라이브에 저장되어 팀원들과 자동으로 공유됩니다.")
else:
    st.warning("⚠️ 구글 드라이브 연결 실패로 로컬 폴더에 저장됩니다.")

uploaded_file = st.file_uploader(
    "쇼피 대량등록 템플릿 xlsx 파일",
    type=["xlsx"],
    key="template_upload",
)

# ════════════════════════════════════════
# 업로드 파일 처리
# ════════════════════════════════════════
if uploaded_file:
    try:
        # 파일을 bytes로 읽기
        file_bytes = uploaded_file.getvalue()

        # 템플릿 기본 정보 추출
        template_code, top_category, mid_category, template_filename = extract_template_info(file_bytes)

        # 기존 템플릿 존재 여부 확인
        existing_code_match = any(
            info.get("template_code") == template_code
            for info in auto_rules.values()
        )
        is_replace = template_filename in saved_templates or existing_code_match

        # 업로드 정보 미리보기
        st.info(
            f"🔑 **식별 코드 (C2):** `{template_code}`  \n"
            f"📂 **카테고리:** {top_category} > {mid_category}  \n"
            f"💾 **저장 파일명:** `{template_filename}`  \n"
            f"{'🔄 기존 템플릿을 교체합니다.' if is_replace else '🆕 새 템플릿으로 등록됩니다.'}"
        )

        # 분석 및 저장 버튼
        if st.button("✅ 분석 및 저장", type="primary"):
            with st.spinner("템플릿 분석 및 저장 중..."):
                try:
                    # ─────────────────────────────
                    # 1단계: 템플릿 상세 분석
                    # ─────────────────────────────
                    new_rules, t_code, t_top, t_mid, t_file = analyze_template(file_bytes)

                    # ─────────────────────────────
                    # 2단계: 파일 저장
                    # ─────────────────────────────
                    if USE_GDRIVE:
                        # 구글 드라이브에 업로드
                        gdrive.upload_template(t_file, file_bytes)
                        storage_msg = "구글 드라이브"
                    else:
                        # 로컬 폴더에 저장 (fallback)
                        from file_builder import TEMPLATES_DIR
                        save_path = TEMPLATES_DIR / t_file
                        save_path.write_bytes(file_bytes)
                        storage_msg = "로컬 (templates/)"

                    # ─────────────────────────────
                    # 3단계: 분석 결과 저장
                    # ─────────────────────────────
                    save_auto_rules(new_rules, t_code, t_mid)

                    # 성공 메시지
                    st.success(
                        f"✅ **완료!** {t_top} > {t_mid} 템플릿 저장  \n"
                        f"💾 저장 위치: {storage_msg}  \n"
                        f"📄 파일명: `{t_file}`  \n"
                        f"📊 소카테고리: **{len(new_rules)}개**"
                    )

                    # 분석 결과 상세 보기
                    with st.expander("📊 분석 결과 상세 보기", expanded=False):
                        st.markdown(f"**총 {len(new_rules)}개 소카테고리 규칙이 등록되었습니다.**")

                        for cat_id, info in new_rules.items():
                            mandatory_attrs = info.get("mandatory_attrs", {})

                            # 필수 속성 요약 (처음 3개만 표시)
                            if mandatory_attrs:
                                attr_list = list(mandatory_attrs.values())
                                attr_summary = ", ".join(
                                    f"{a['display']}={a['auto_value']}"
                                    for a in attr_list[:3]
                                )
                                if len(mandatory_attrs) > 3:
                                    attr_summary += f" 외 {len(mandatory_attrs) - 3}개"
                            else:
                                attr_summary = "자동입력 항목 없음"

                            # 소카테고리 경로의 마지막 부분만 표시
                            cat_name = info['category_path'].split('/')[-1]

                            st.markdown(
                                f"- **`{cat_id}`** {cat_name}  \n"
                                f"  └ 필수 속성: {attr_summary}"
                            )

                    # 페이지 새로고침
                    st.rerun()

                except Exception as e:
                    st.error("❌ 분석 또는 저장 중 오류가 발생했습니다.")
                    with st.expander("오류 상세 보기"):
                        st.code(str(e))
                        import traceback
                        st.code(traceback.format_exc())

    except Exception as e:
        st.error("❌ 파일 읽기 중 오류가 발생했습니다.")
        with st.expander("오류 상세 보기"):
            st.code(str(e))
            import traceback
            st.code(traceback.format_exc())

# ════════════════════════════════════════
# 사용 가이드
# ════════════════════════════════════════
st.divider()

with st.expander("💡 사용 가이드"):
    st.markdown("""
    **템플릿 관리 가이드:**
    
    **🔍 템플릿 분석 항목**
    - **C2 코드**: 템플릿 고유 식별자 (중복 방지)
    - **카테고리 구조**: 대/중/소 카테고리 자동 추출
    - **필수 속성**: MANDATORY 항목 자동 감지 및 기본값 설정
    
    **👥 팀 협업**
    - 구글 드라이브 연결 시 모든 팀원이 동일한 템플릿 사용
    - 템플릿 업데이트 시 자동으로 모든 사용자에게 반영
    - 설정 변경사항도 실시간 동기화
    
    **🔧 문제 해결**
    - "분석 결과 없음" 표시 시 → 해당 템플릿 재업로드
    - 구글 드라이브 연결 실패 시 → 로컬 모드로 자동 전환
    - 업로드 오류 시 → 파일 형식 및 권한 확인
    """)

st.caption("📌 템플릿 업로드 후 메인 페이지에서 바로 사용할 수 있습니다.")
