"""
pages/2_⚙️_필수값_관리.py
global_rules (전체 공통 고정값) 관리 페이지
"""

import streamlit as st
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from template_analyzer import load_auto_rules, load_global_rules, save_global_rules

st.set_page_config(page_title="⚙필수값 관리", page_icon="⚙️", layout="wide")

st.title("⚙️ 필수값 관리")
st.caption("템플릿에 항상 고정으로 들어가야 하는 값들을 관리합니다.")

st.info("""
**두 가지 유형의 값을 관리합니다:**
- 🔵 **항상 입력** (`exists`): 해당 컬럼이 템플릿에 존재하면 무조건 입력 (예: FDA번호)
- 🟡 **필수일 때만** (`mandatory`): 소카테고리에서 MANDATORY로 지정된 경우만 입력 (예: 주관식 항목)
""")

auto_rules = load_auto_rules()
global_rules = load_global_rules()

# ── 현재 저장된 global_rules 표시 및 편집 ──
st.subheader("📝 저장된 고정값 목록")

if not global_rules:
    st.info("저장된 고정값이 없습니다. 아래에서 추가해주세요.")
else:
    rules_to_delete = []
    updated_rules = {}

    for internal_key, rule in global_rules.items():
        with st.container(border=True):
            col1, col2, col3, col4 = st.columns([3, 3, 2, 1])

            with col1:
                badge = "🔵" if rule.get("apply_when") == "exists" else "🟡"
                st.markdown(f"{badge} **{rule.get('display', internal_key)}**")
                st.caption(f"`{internal_key}`")

            with col2:
                new_val = st.text_input(
                    "값",
                    value=rule.get("value", ""),
                    key=f"val_{internal_key}",
                    label_visibility="collapsed",
                )

            with col3:
                apply_options = {"exists": "🔵 항상 입력", "mandatory": "🟡 필수일 때만"}
                current_apply = rule.get("apply_when", "exists")
                new_apply = st.selectbox(
                    "적용 조건",
                    options=list(apply_options.keys()),
                    format_func=lambda x: apply_options[x],
                    index=list(apply_options.keys()).index(current_apply),
                    key=f"apply_{internal_key}",
                    label_visibility="collapsed",
                )

            with col4:
                if st.button("🗑️", key=f"del_{internal_key}", help="삭제"):
                    rules_to_delete.append(internal_key)
                else:
                    updated_rules[internal_key] = {
                        "display": rule.get("display", internal_key),
                        "value": new_val,
                        "apply_when": new_apply,
                    }

    # 저장 버튼
    if st.button("💾 변경사항 저장", type="primary"):
        for k in rules_to_delete:
            updated_rules.pop(k, None)
        save_global_rules(updated_rules)
        st.success("✅ 저장되었습니다!")
        st.rerun()

st.divider()

# ── 새 고정값 추가 ──
st.subheader("➕ 새 고정값 추가")

st.markdown("아래에서 **속성 이름으로 검색**해서 추가하거나, 직접 입력할 수 있습니다.")

# auto_rules에서 모든 속성 목록 수집 (internal_key → display_name)
all_attrs = {}
for cat_id, info in auto_rules.items():
    for key, attr in info.get("mandatory_attrs", {}).items():
        if key not in all_attrs:
            all_attrs[key] = attr.get("display", key)

# 템플릿의 전체 속성도 포함 (고정값은 mandatory가 아닌 것도 설정 가능)
with st.expander("🔍 알려진 속성에서 선택"):
    if all_attrs:
        search = st.text_input("속성명 검색", placeholder="예: FDA, Ingredient, Skin Type...")
        filtered = {k: v for k, v in all_attrs.items()
                    if search.lower() in v.lower() or search.lower() in k.lower()} if search else all_attrs

        if filtered:
            selected_key = st.selectbox(
                "속성 선택",
                options=list(filtered.keys()),
                format_func=lambda k: f"{filtered[k]} ({k})",
            )

            col1, col2, col3 = st.columns([3, 2, 1])
            with col1:
                new_value = st.text_input("입력할 고정값", key="new_attr_value",
                                          placeholder="예: FDA-2024-XXXXX")
            with col2:
                apply_options = {"exists": "🔵 항상 입력", "mandatory": "🟡 필수일 때만"}
                new_apply = st.selectbox("적용 조건", options=list(apply_options.keys()),
                                         format_func=lambda x: apply_options[x],
                                         key="new_attr_apply")
            with col3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("추가", key="add_from_search"):
                    if new_value:
                        global_rules[selected_key] = {
                            "display": filtered[selected_key],
                            "value": new_value,
                            "apply_when": new_apply,
                        }
                        save_global_rules(global_rules)
                        st.success(f"✅ '{filtered[selected_key]}' 추가 완료!")
                        st.rerun()
                    else:
                        st.error("값을 입력해주세요.")
    else:
        st.info("먼저 템플릿을 업로드해주세요.")

with st.expander("✏️ 직접 입력"):
    col1, col2, col3, col4 = st.columns([2, 2, 2, 1])
    with col1:
        direct_display = st.text_input("속성 표시명", placeholder="예: FDA Registration No.")
    with col2:
        direct_key = st.text_input("내부 key", placeholder="예: ps_product_global_attribute.100963")
    with col3:
        direct_value = st.text_input("고정값", placeholder="예: FDA-2024-XXXXX")
    with col4:
        st.markdown("<br>", unsafe_allow_html=True)
        direct_apply = st.selectbox(
            "조건",
            options=["exists", "mandatory"],
            format_func=lambda x: "🔵 항상" if x == "exists" else "🟡 필수만",
            key="direct_apply",
        )

    if st.button("직접 추가", key="add_direct"):
        if direct_display and direct_key and direct_value:
            global_rules[direct_key] = {
                "display": direct_display,
                "value": direct_value,
                "apply_when": direct_apply,
            }
            save_global_rules(global_rules)
            st.success(f"✅ '{direct_display}' 추가 완료!")
            st.rerun()
        else:
            st.error("모든 항목을 입력해주세요.")
