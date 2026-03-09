"""
pages/1_Template_Management.py
쇼피 대량등록 템플릿 업로드 및 분석 페이지 (UI 완전 개선 버전)
"""

import streamlit as st
from pathlib import Path
import sys
from collections import defaultdict

# 부모 디렉토리를 import 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

# 안전한 import with 예외 처리
try:
    from template_analyzer import (
        analyze_template,
        save_auto_rules,
        load_auto_rules,
        extract_template_info
    )
    from gdrive_manager import get_gdrive_manager
except ImportError as e:
    st.error(f"❌ 모듈 import 오류: {e}")
    st.info("앱을 재시작하거나 관리자에게 문의하세요.")
    st.stop()
except Exception as e:
    st.error(f"❌ 초기화 오류: {e}")
    st.info("구글 드라이브 인증 설정을 확인해주세요.")
    st.stop()

# ════════════════════════════════════════════════════════════
# 페이지 기본 설정 및 세션 상태 초기화
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="템플릿 관리",
    page_icon="📋",
    layout="wide"
)

# 세션 상태 초기화
if "expand_all_templates" not in st.session_state:
    st.session_state.expand_all_templates = False

# ════════════════════════════════════════════════════════════
# 헤더 및 연결 상태
# ════════════════════════════════════════════════════════════
st.title("📋 템플릿 관리")
st.caption("쇼피 대량등록 템플릿을 체계적으로 관리하고 분석합니다.")

# 구글 드라이브 연결 상태 확인
try:
    gdrive = get_gdrive_manager()
    st.success("✅ 구글 드라이브 연결됨 (OAuth 2.0)")
    st.markdown(f"📁 [공유 폴더 열기]({gdrive.get_folder_link()})")
    USE_GDRIVE = True
except Exception as e:
    st.warning("⚠️ 구글 드라이브 연결에 실패했습니다. 로컬 모드로 동작합니다.")
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
    USE_GDRIVE = False

st.divider()

# ════════════════════════════════════════════════════════════
# 1. 템플릿 업로드 (최상단 배치)
# ════════════════════════════════════════════════════════════
st.subheader("⬆️ 새 템플릿 업로드")

with st.container(border=True):
    col_info, col_upload = st.columns([2, 3])

    with col_info:
        st.markdown("""
        **자동 처리 프로세스:**
        1. 📄 `Template` 시트 **C2 코드** 기반 중복 검사
        2. 📂 `Pre-order DTS Range`에서 카테고리 정보 추출
        3. 🔍 `HiddenCatProps` 분석으로 필수 속성 매핑
        4. 🔄 동일 C2 코드 기존 데이터 자동 교체
        """)

        if USE_GDRIVE:
            st.info("📤 OAuth 2.0 인증을 통해 개인 구글 드라이브 공간을 활용합니다.")
        else:
            st.warning("⚠️ 구글 드라이브 연결 실패로 로컬 저장됩니다.")

    with col_upload:
        uploaded_file = st.file_uploader(
            "쇼피 대량등록 템플릿 xlsx 파일",
            type=["xlsx"],
            key="template_upload",
            help="쇼피에서 다운로드한 대량등록 템플릿 파일을 업로드하세요"
        )

        if uploaded_file:
            try:
                file_bytes = uploaded_file.getvalue()
                template_code, top_category, mid_category, template_filename = extract_template_info(file_bytes)

                # 기존 템플릿 확인
                auto_rules = load_auto_rules()
                existing_code_match = any(
                    info.get("template_code") == template_code
                    for info in auto_rules.values()
                )

                # 저장된 템플릿 목록 확인
                if USE_GDRIVE:
                    saved_templates = gdrive.list_templates()
                else:
                    from file_builder import TEMPLATES_DIR
                    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
                    saved_templates = [f.name for f in sorted(TEMPLATES_DIR.glob("*.xlsx"))]

                is_replace = template_filename in saved_templates or existing_code_match

                st.info(
                    f"🔑 **템플릿 코드:** `{template_code}`  \n"
                    f"📂 **카테고리:** {top_category} > {mid_category}  \n"
                    f"💾 **파일명:** `{template_filename}`  \n"
                    f"{'🔄 기존 템플릿 교체' if is_replace else '🆕 새 템플릿 등록'}"
                )

                if st.button("✅ 분석 및 저장", type="primary", use_container_width=True):
                    with st.spinner("템플릿 정밀 분석 및 저장 중..."):
                        try:
                            # 1단계: 템플릿 상세 분석
                            new_rules, t_code, t_top, t_mid, t_file = analyze_template(file_bytes)

                            # 2단계: 파일 저장
                            if USE_GDRIVE:
                                gdrive.upload_template(t_file, file_bytes)
                                storage_msg = "구글 드라이브 (OAuth 2.0)"
                            else:
                                from file_builder import TEMPLATES_DIR
                                save_path = TEMPLATES_DIR / t_file
                                save_path.write_bytes(file_bytes)
                                storage_msg = "로컬 (templates/)"

                            # 3단계: 분석 결과 저장
                            save_auto_rules(new_rules, t_code, t_mid)

                            st.success(
                                f"✅ **완료!** {t_top} > {t_mid} 템플릿 저장  \n"
                                f"💾 위치: {storage_msg}  \n"
                                f"📄 파일: `{t_file}`  \n"
                                f"📊 소카테고리: **{len(new_rules)}개**"
                            )

                            with st.expander("📊 분석 결과 상세", expanded=False):
                                st.markdown(f"**{len(new_rules)}개 소카테고리 규칙 등록 완료**")
                                for cat_id, info in new_rules.items():
                                    mandatory_attrs = info.get("mandatory_attrs", {})
                                    cat_name = info['category_path'].split('/')[-1]
                                    st.markdown(f"- **`{cat_id}`** {cat_name} → 필수 속성 {len(mandatory_attrs)}개")

                            st.rerun()

                        except Exception as e:
                            st.error("❌ 처리 중 오류 발생")
                            with st.expander("오류 상세"):
                                st.code(str(e))
                                import traceback
                                st.code(traceback.format_exc())

            except Exception as e:
                st.error("❌ 파일 읽기 오류")
                with st.expander("오류 상세"):
                    st.code(str(e))

st.divider()

# ════════════════════════════════════════════════════════════
# 2. 등록된 템플릿 (계층형 트리 + 검색)
# ════════════════════════════════════════════════════════════
st.subheader("📂 등록된 템플릿")

# 데이터 로드
auto_rules = load_auto_rules()

# 템플릿 파일 목록 조회
if USE_GDRIVE:
    with st.spinner("구글 드라이브에서 템플릿 목록 로딩..."):
        saved_templates = gdrive.list_templates()
    storage_info = "🌐 구글 드라이브 (OAuth 2.0)"
else:
    from file_builder import TEMPLATES_DIR
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    saved_templates = [f.name for f in sorted(TEMPLATES_DIR.glob("*.xlsx"))]
    storage_info = "💻 로컬 (templates/)"

# ── 템플릿 데이터 구조화 (template_file 기준으로 그룹핑) ──
template_nodes = {}  # {template_file: {top, mid, cat_count, sample_paths, code, status}}
analyzed_files = set()

for cat_id, info in auto_rules.items():
    t_file = info.get("template_file", "UNKNOWN.xlsx")
    if t_file != "UNKNOWN.xlsx":
        analyzed_files.add(t_file)

    top = info.get("top_category", "Unknown")
    mid = info.get("mid_category", "Unknown")
    cat_path = info.get("category_path", "")
    template_code = info.get("template_code", "")

    if t_file not in template_nodes:
        template_nodes[t_file] = {
            "top": top,
            "mid": mid,
            "template_code": template_code,
            "cat_ids": set(),
            "sample_paths": set(),
        }

    template_nodes[t_file]["cat_ids"].add(cat_id)
    if cat_path:
        template_nodes[t_file]["sample_paths"].add(cat_path)

# 각 템플릿의 상태 결정
for t_file in template_nodes:
    file_exists = t_file in saved_templates
    is_analyzed = t_file in analyzed_files

    if file_exists and is_analyzed:
        template_nodes[t_file]["status"] = "✅ 정상"
        template_nodes[t_file]["status_priority"] = 1
    elif is_analyzed and not file_exists:
        template_nodes[t_file]["status"] = "⚠️ 파일 없음"
        template_nodes[t_file]["status_priority"] = 2
    elif file_exists and not is_analyzed:
        template_nodes[t_file]["status"] = "⚠️ 재분석 필요"
        template_nodes[t_file]["status_priority"] = 3
    else:
        template_nodes[t_file]["status"] = "❌ 오류"
        template_nodes[t_file]["status_priority"] = 4

# 파일만 있고 분석되지 않은 템플릿 추가
for t_file in saved_templates:
    if t_file not in template_nodes:
        template_nodes[t_file] = {
            "top": "Unknown",
            "mid": "Unknown",
            "template_code": "",
            "cat_ids": set(),
            "sample_paths": set(),
            "status": "⚠️ 재분석 필요",
            "status_priority": 3
        }

# top > mid > [template_files] 트리 구조 생성
tree = defaultdict(lambda: defaultdict(list))
for t_file, meta in template_nodes.items():
    top = meta["top"]
    mid = meta["mid"]
    tree[top][mid].append((t_file, meta))

# 통계 정보
total_templates = len(template_nodes)
total_rules = len(auto_rules)
total_saved_files = len(saved_templates)

# ── 상단 통계 및 제어 UI ──
col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
with col_stat1:
    st.metric("분석된 규칙", total_rules)
with col_stat2:
    st.metric("템플릿 파일", total_templates)
with col_stat3:
    st.metric("저장된 파일", total_saved_files)
with col_stat4:
    st.caption(f"저장 위치: {storage_info}")

# ── 검색 및 제어 바 ──
col_search, col_sort, col_expand, col_collapse = st.columns([3, 1, 1, 1])

with col_search:
    search_query = st.text_input(
        "🔍 빠른 검색",
        placeholder="카테고리명, 파일명, 템플릿 코드로 검색...",
        key="template_search",
        label_visibility="collapsed"
    )

with col_sort:
    sort_option = st.selectbox(
        "정렬",
        ["이름순", "카테고리순", "상태순"],
        key="sort_templates",
        label_visibility="collapsed"
    )

with col_expand:
    if st.button("📂 모두 펼치기", use_container_width=True):
        st.session_state.expand_all_templates = True
        st.rerun()

with col_collapse:
    if st.button("📁 모두 접기", use_container_width=True):
        st.session_state.expand_all_templates = False
        st.rerun()

# ── 검색 필터링 함수 ──
def template_matches_search(t_file: str, meta: dict, query: str) -> bool:
    if not query:
        return True
    query = query.lower()

    search_targets = [
        t_file.lower(),
        meta.get("top", "").lower(),
        meta.get("mid", "").lower(),
        meta.get("template_code", "").lower(),
    ]

    # 샘플 카테고리 경로와 cat_id도 검색 대상에 포함
    for path in meta.get("sample_paths", []):
        search_targets.append(path.lower())
    for cat_id in meta.get("cat_ids", []):
        search_targets.append(str(cat_id).lower())

    return any(query in target for target in search_targets if target)

# 카테고리 아이콘 매핑
CATEGORY_ICONS = {
    "Beauty": "🎨",
    "Fashion": "👗",
    "Electronics": "📱",
    "Home": "🏠",
    "Health": "💊",
    "Sports": "⚽",
    "Food": "🍔",
    "Baby": "👶",
    "Pet": "🐾",
    "Books": "📚",
    "Unknown": "📦"
}

st.markdown("---")

# ── 트리 구조 렌더링 ──
if not tree:
    st.info("아직 등록된 템플릿이 없습니다. 위에서 새 템플릿을 업로드해보세요.")
else:
    # 검색 결과 카운터
    match_count = 0

    # 대카테고리별 표시
    sorted_top_keys = sorted(tree.keys())

    for top in sorted_top_keys:
        mid_map = tree[top]

        # 검색 필터링: 해당 대카테고리에 매칭되는 템플릿이 있는지 확인
        top_has_match = False
        for mid, t_list in mid_map.items():
            for t_file, meta in t_list:
                if template_matches_search(t_file, meta, search_query):
                    top_has_match = True
                    break
            if top_has_match:
                break

        if not top_has_match:
            continue

        # 대카테고리 아이콘
        icon = CATEGORY_ICONS.get(top, "📦")
        total_in_top = sum(len(v) for v in mid_map.values())

        with st.expander(
            f"{icon} **{top}** ({total_in_top}개 템플릿)",
            expanded=st.session_state.expand_all_templates
        ):
            # 중카테고리별 표시
            # 중카테고리별 표시
            for mid in sorted(mid_map.keys()):
                t_list = mid_map[mid]

                # 중카테고리 레벨에서 검색 필터링
                filtered_templates = [
                    (t_file, meta) for t_file, meta in t_list
                    if template_matches_search(t_file, meta, search_query)
                ]

                if not filtered_templates:
                    continue

                st.markdown(f"#### 📁 {mid} ({len(filtered_templates)}개)")

                # 정렬 적용
                if sort_option == "이름순":
                    filtered_templates.sort(key=lambda x: x[0].lower())
                elif sort_option == "상태순":
                    filtered_templates.sort(key=lambda x: x[1].get("status_priority", 5))
                elif sort_option == "카테고리순":
                    filtered_templates.sort(key=lambda x: (x[1].get("top", ""), x[1].get("mid", ""), x[0].lower()))

                # ─────────────────────────────────────
                # 테이블형 리스트 UI
                # ─────────────────────────────────────
                col_ratios = [1.0, 1.5, 3.0, 4.0]  # 상태, 코드, 카테고리, 파일명
                h1, h2, h3, h4 = st.columns(col_ratios)

                with h1:
                    st.markdown("**상태**")
                with h2:
                    st.markdown("**템플릿 코드**")
                with h3:
                    st.markdown("**카테고리**")
                with h4:
                    st.markdown("**파일명**")

                st.markdown("---")  # 헤더 구분선

                # 템플릿 데이터 행들
                for t_file, meta in filtered_templates:
                    template_code = meta.get("template_code", "") or "미확인"
                    status = meta.get("status", "❓ 알 수 없음")
                    top_cat = meta.get("top", "Unknown")
                    mid_cat = meta.get("mid", "Unknown")
                    category_path = f"{top_cat} > {mid_cat}"

                    # 상태별 색상 및 아이콘
                    if status.startswith("✅"):
                        status_display = "✅"
                        status_color = "normal"
                    elif status.startswith("⚠️"):
                        status_display = "⚠️"
                        status_color = "warning"
                    else:
                        status_display = "❌"
                        status_color = "error"

                    # 데이터 행 출력
                    c1, c2, c3, c4 = st.columns(col_ratios)

                    with c1:
                        if status_color == "normal":
                            st.success(status_display, icon="✅")
                        elif status_color == "warning":
                            st.warning(status_display, icon="⚠️")
                        else:
                            st.error(status_display, icon="❌")

                    with c2:
                        st.code(template_code, language=None)

                    with c3:
                        st.markdown(category_path)

                    with c4:
                        st.markdown(f"**{t_file}**")

                    match_count += 1

                # 중카테고리 간 간격
                st.markdown("")

    # 검색 결과 요약
    if search_query:
        st.caption(f"🔍 검색어 '{search_query}'에 매칭되는 템플릿: **{match_count}개**")
        if match_count == 0:
            st.info("검색 조건에 맞는 템플릿이 없습니다. 다른 키워드로 검색해보세요.")

st.divider()

# ════════════════════════════════════════════════════════════
# 도움말
# ════════════════════════════════════════════════════════════
with st.expander("💡 템플릿 관리 가이드"):
    st.markdown("""
    **📂 계층 구조**
    - **대카테고리**: Beauty, Fashion, Electronics 등 최상위 분류
    - **중카테고리**: Men's Care, Women's Bags 등 중간 분류  
    - **템플릿 파일**: 각 중카테고리별 업로드된 쇼피 템플릿
    - **소카테고리 규칙**: 실제 상품 등록 시 적용되는 세부 규칙
    
    **🔍 검색 기능**
    - **카테고리명**: "Beauty", "Makeup", "Men's Care"
    - **파일명**: "Beauty_Mens_Care.xlsx"  
    - **템플릿 코드**: "100660", "100875"
    - **카테고리 경로**: "Beauty/Makeup/Lips"
    - **소카테고리 ID**: "101645"
    
    **📊 상태 표시**
    - **✅ 정상**: 파일 존재 + 분석 완료 (즉시 사용 가능)
    - **⚠️ 파일 없음**: 분석은 됐지만 실제 파일이 구글 드라이브에 없음
    - **⚠️ 재분석 필요**: 파일은 있지만 아직 분석되지 않음 (재업로드 권장)
    - **❌ 오류**: 파일도 없고 분석도 안 된 상태
    
    **👥 팀 협업**
    - 구글 드라이브 연동 시 모든 팀원이 동일한 템플릿 환경 공유
    - 템플릿 업데이트 시 실시간으로 모든 사용자에게 반영
    - 검색 및 필터링 설정도 개인별 세션에서 유지
    
    **🔧 문제 해결**  
    - "재분석 필요" 표시 → 해당 템플릿을 다시 업로드
    - "파일 없음" 표시 → 구글 드라이브 연결 상태 확인
    - 검색 결과가 없음 → 키워드를 바꾸거나 "모두 펼치기"로 전체 확인
    """)

st.caption("📌 템플릿 업로드 후 메인 페이지에서 바로 사용할 수 있습니다.")
