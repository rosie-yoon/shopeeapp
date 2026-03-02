"""
pages/1_Template_Management.py
쇼피 대량등록 템플릿 업로드 및 분석 페이지
"""

import streamlit as st
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from template_analyzer import analyze_template, save_auto_rules, load_auto_rules, extract_template_info
from file_builder import TEMPLATES_DIR

st.set_page_config(page_title="📋 템플릿 관리", page_icon="📋", layout="wide")

st.title("📋 템플릿 관리")
st.caption("쇼피 대량등록 템플릿을 업로드하면 자동으로 분석하여 저장합니다.")

auto_rules = load_auto_rules()

# ── 등록된 템플릿 목록 ──────────────────────────────
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
saved_templates = sorted(TEMPLATES_DIR.glob("*.xlsx"))

# auto_rules에서 등록된 template_file 목록
analyzed_files = set(v.get("template_file", "") for v in auto_rules.values())

st.subheader("📁 등록된 템플릿")
if saved_templates:
    for t in saved_templates:
        if t.name in analyzed_files:
            st.markdown(f"- ✅ `{t.name}`")
        else:
            # 파일은 있지만 분석 결과가 없음 → 재업로드 필요
            st.markdown(f"- ⚠️ `{t.name}` — 분석 결과 없음, 아래에서 다시 업로드해주세요")
else:
    st.info("아직 등록된 템플릿이 없습니다.")

st.divider()

# ── 템플릿 업로드 ───────────────────────────────────
st.subheader("⬆️ 템플릿 업로드")
st.markdown("""
업로드하면 자동으로:
1. `Template` 시트 **C2 코드**로 동일 템플릿 여부 판별
2. `Pre-order DTS Range` 시트에서 **대카테고리 + 중카테고리명** 추출 → 파일명 자동 결정
3. `HiddenCatProps` 시트 분석 → 소카테고리별 필수 입력 항목 추출
4. 같은 C2 코드의 **기존 데이터 교체**
""")

uploaded_file = st.file_uploader(
    "쇼피 대량등록 템플릿 xlsx 파일",
    type=["xlsx"],
    key="template_upload",
)

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()  # bytes로 한 번만 읽기

        template_code, top_category, mid_category, template_filename = extract_template_info(file_bytes)

        save_path = TEMPLATES_DIR / template_filename
        existing_code_match = any(
            info.get("template_code") == template_code
            for info in auto_rules.values()
        )
        is_replace = save_path.exists() or existing_code_match

        st.info(
            f"🔑 식별 코드 (C2): `{template_code}`  \n"
            f"📂 카테고리: **{top_category} > {mid_category}**  \n"
            f"💾 저장 파일명: `{template_filename}`  \n"
            f"{'🔄 기존 템플릿을 교체합니다.' if is_replace else '🆕 새 템플릿으로 등록됩니다.'}"
        )

        if st.button("✅ 분석 및 저장", type="primary"):
            with st.spinner("템플릿 분석 중..."):
                try:
                    new_rules, t_code, t_top, t_mid, t_file = analyze_template(file_bytes)
                    save_auto_rules(new_rules, t_code, t_mid)

                    # templates/ 폴더에 원본 bytes 저장
                    save_path.write_bytes(file_bytes)

                    st.success(
                        f"✅ 완료! **{t_top} > {t_mid}** 템플릿 저장 → "
                        f"`{t_file}` ({len(new_rules)}개 소카테고리)"
                    )

                    with st.expander("📊 분석 결과 보기", expanded=False):
                        for cat_id, info in new_rules.items():
                            mandatory_attrs = info.get("mandatory_attrs", {})
                            attr_summary = ", ".join(
                                f"{a['display']}={a['auto_value']}"
                                for a in mandatory_attrs.values()
                            ) if mandatory_attrs else "자동입력 항목 없음"
                            st.markdown(
                                f"- `{cat_id}` **{info['category_path'].split('/')[-1]}** "
                                f"→ {attr_summary}"
                            )

                    st.rerun()

                except Exception as e:
                    st.error(f"❌ 분석 오류: {e}")
                    import traceback
                    st.code(traceback.format_exc())

    except Exception as e:
        st.error(f"❌ 파일 읽기 오류: {e}")
        import traceback
        st.code(traceback.format_exc())
