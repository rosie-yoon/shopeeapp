"""
pages/3_🎨_Cover_Maker.py
Cover Maker 페이지 - 로그인 없는 통합 버전
"""

import streamlit as st
import io
import zipfile
import re
from pathlib import Path
from PIL import Image as PILImage
from datetime import datetime
import sys

# 상위 폴더 모듈 import를 위한 경로 설정
sys.path.insert(0, str(Path(__file__).parent.parent))

from composer_utils import (
    compose_one_bytes,
    SHADOW_PRESETS,
    has_useful_alpha,
    ensure_rgba,
)

# 페이지 설정
st.set_page_config(
    page_title="Cover Maker",
    page_icon="🎨",
    layout="wide"
)

# 설정값 (인증 관련 제거)
CONFIG = {
    "APP_TITLE": "Cover Maker",
    "APP_VERSION": "v1.3",
    "MAX_PREVIEW_COUNT": 50,
    "GALLERY_COLS": 6,
    "OUTPUT_FORMAT": "JPEG",
    "JPEG_QUALITY": 95,
}


@st.dialog("📖 사용 가이드")
def show_manual():
    st.markdown(f"""
    ### 🎨 Cover Maker 사용법

    **핵심 기능**
    - 상품 이미지(누끼)와 샵 템플릿을 자동 합성
    - **파일명 규칙**: `(상품명)_C_(템플릿명).jpg`

    **이미지 준비**
    - **상품 이미지**: 투명 배경 PNG 권장 (Remove.bg 사용)
    - **템플릿 이미지**: 
        - JPG = 배경형 (그림자 효과 가능)
        - PNG = 액자형 (투명한 프레임)

    **사용 순서**
    1. 왼쪽에서 이미지들을 업로드
    2. 오른쪽에서 위치/크기/그림자 조정  
    3. 미리보기 확인 후 ZIP 다운로드

    **파일명 제한**
    - 영문, 숫자, `_`, `-` 만 사용 가능
    - 한글, 공백, 특수문자 사용 금지
    """)


# 헤더 (로그인 체크 완전 제거)
col1, col2 = st.columns([5, 1])
with col1:
    st.title("🎨 Cover Maker")
    st.caption("상품 이미지와 템플릿을 합성하여 썸네일을 대량 제작합니다.")
with col2:
    if st.button("📖 사용법", use_container_width=True):
        show_manual()

st.divider()


# 유틸리티 함수들
def validate_template_names(files):
    """템플릿 파일명 유효성 검사"""
    if not files:
        return True, []
    seen_stems = set()
    errors = []
    pattern = re.compile(r'^[a-zA-Z0-9_-]+$')

    for f in files:
        stem = Path(f.name).stem
        if not pattern.match(stem):
            errors.append(f"'{f.name}' - 영문, 숫자, _, - 만 사용 가능")
            continue
        if stem in seen_stems:
            errors.append(f"'{stem}' - 중복된 템플릿명")
        else:
            seen_stems.add(stem)

    return (False, errors) if errors else (True, [])


def analyze_combinations(item_files, template_files):
    """이미지 조합 가능성 분석"""
    valid_combinations = []
    invalid_combinations = []

    for item_file in item_files:
        try:
            item_file.seek(0)
            item_img = PILImage.open(item_file)
            has_alpha = has_useful_alpha(ensure_rgba(item_img))
        except:
            continue

        for template_file in template_files:
            template_ext = Path(template_file.name).suffix.lower()
            is_png_template = (template_ext == '.png')

            if has_alpha:
                mode = 'frame' if is_png_template else 'normal'
                valid_combinations.append((item_file, template_file, mode))
            else:
                if is_png_template:
                    valid_combinations.append((item_file, template_file, 'frame'))
                else:
                    invalid_combinations.append((item_file, template_file))

    return {
        'valid_combinations': valid_combinations,
        'invalid_combinations': invalid_combinations,
        'summary': {
            'valid': len(valid_combinations),
            'invalid': len(invalid_combinations)
        }
    }


# 세션 상태 초기화 (Cover Maker 전용 네임스페이스)
ss = st.session_state
cm_defaults = {
    "cm_anchor": "center",
    "cm_resize_ratio": 1.0,
    "cm_shadow_preset": "off",
    "cm_preview_list": [],
    "cm_zip_cache": None,
    "cm_item_uploader_key": 0,
    "cm_template_uploader_key": 0,
    "cm_cached_analysis": None,
    "cm_last_file_sig": None,
    "cm_last_settings_sig": None,
    "cm_needs_preview_regen": False,
}
for k, v in cm_defaults.items():
    ss.setdefault(k, v)

# 메인 UI 레이아웃
left_col, right_col = st.columns([1, 1.2])

# 왼쪽: 파일 업로드
with left_col:
    st.subheader("📤 이미지 업로드")

    item_files = st.file_uploader(
        "1️⃣ 상품 이미지 (투명 배경 PNG 권장)",
        type=["png", "webp", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"cm_item_uploader_{ss.cm_item_uploader_key}",
        help="Remove.bg로 배경을 제거한 PNG 파일이 가장 좋습니다"
    )

    if st.button("🗑️ 상품 이미지 비우기",
                 key="cm_clear_items",
                 disabled=not bool(item_files)):
        ss.cm_item_uploader_key += 1
        ss.cm_preview_list = []
        ss.cm_zip_cache = None
        ss.cm_cached_analysis = None
        ss.cm_last_file_sig = None
        ss.cm_needs_preview_regen = False
        st.rerun()

    st.markdown("---")

    template_files = st.file_uploader(
        "2️⃣ 템플릿 이미지 (파일명=샵코드)",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key=f"cm_template_uploader_{ss.cm_template_uploader_key}",
        help="JPG=배경형(그림자 가능), PNG=액자형"
    )

    if st.button("🗑️ 템플릿 이미지 비우기",
                 key="cm_clear_templates",
                 disabled=not bool(template_files)):
        ss.cm_template_uploader_key += 1
        ss.cm_preview_list = []
        ss.cm_zip_cache = None
        ss.cm_cached_analysis = None
        ss.cm_last_file_sig = None
        ss.cm_needs_preview_regen = False
        st.rerun()

    # 파일 유효성 검사
    is_valid_tpl, tpl_errors = validate_template_names(template_files)
    if template_files and not is_valid_tpl:
        st.error("🚨 템플릿 파일명 오류")
        for err in tpl_errors:
            st.write(f"❌ {err}")
        st.info("💡 파일명을 수정한 후 다시 업로드해주세요.")

    # 조합 분석
    if item_files and template_files and is_valid_tpl:
        current_file_sig = (
            tuple(f.name for f in item_files),
            tuple(f.name for f in template_files),
            len(item_files), len(template_files)
        )

        if ss.cm_last_file_sig != current_file_sig or ss.cm_cached_analysis is None:
            with st.spinner("이미지 분석 중..."):
                ss.cm_cached_analysis = analyze_combinations(item_files, template_files)
                ss.cm_last_file_sig = current_file_sig
                ss.cm_needs_preview_regen = True

        analysis = ss.cm_cached_analysis
        if analysis:
            summary = analysis['summary']
            if summary['invalid'] > 0:
                st.warning(f"""
                ⚠️ **조합 분석 결과**
                - ✅ 생성 가능: **{summary['valid']}개**
                - ❌ 자동 제외: **{summary['invalid']}개** (투명배경 없음 + JPG 템플릿)
                """)
            else:
                st.success(f"✅ 모든 조합 생성 가능 ({summary['valid']}개)")

# 오른쪽: 설정 및 미리보기
with right_col:
    st.subheader("⚙️ 설정 및 미리보기")

    c1, c2, c3 = st.columns(3)

    ss.cm_anchor = c1.selectbox(
        "📍 배치 위치",
        ["center", "top", "bottom", "left", "right",
         "top-left", "top-right", "bottom-left", "bottom-right"],
        index=0,
        key="cm_anchor_select"
    )

    resize_options = [1.2, 1.15, 1.1, 1.05, 1.0, 0.95, 0.9, 0.85, 0.8]
    ss.cm_resize_ratio = c2.selectbox(
        "📏 크기 조정",
        resize_options,
        index=resize_options.index(1.0),
        format_func=lambda x: f"{int(round(x * 100))}%",
        key="cm_resize_select"
    )

    ss.cm_shadow_preset = c3.selectbox(
        "🌑 그림자",
        list(SHADOW_PRESETS.keys()),
        index=0,
        help="JPG 템플릿 + 투명 배경 상품에만 적용",
        key="cm_shadow_select"
    )

    st.divider()

    # 설정 변경 감지
    current_settings_sig = (ss.cm_anchor, ss.cm_resize_ratio, ss.cm_shadow_preset)
    if ss.cm_last_settings_sig != current_settings_sig:
        ss.cm_needs_preview_regen = True
        ss.cm_last_settings_sig = current_settings_sig

    # 미리보기 및 ZIP 생성
    if item_files and template_files and is_valid_tpl and ss.cm_cached_analysis:
        if ss.cm_needs_preview_regen:
            ss.cm_preview_list = []
            ss.cm_zip_cache = None

            valid_combinations = ss.cm_cached_analysis['valid_combinations']
            preview_combinations = valid_combinations[:CONFIG["MAX_PREVIEW_COUNT"]]

            with st.spinner("미리보기 및 다운로드 파일 생성 중..."):
                # 미리보기 생성
                for item_file, template_file, mode in preview_combinations:
                    try:
                        item_file.seek(0)
                        template_file.seek(0)

                        item_img = PILImage.open(item_file)
                        template_img = PILImage.open(template_file)

                        template_ext = Path(template_file.name).suffix.lower()
                        composition_mode = "frame" if template_ext == ".png" else "normal"
                        shadow_preset = ss.cm_shadow_preset if composition_mode == "normal" else "off"

                        opts = {
                            "anchor": ss.cm_anchor,
                            "resize_ratio": ss.cm_resize_ratio,
                            "shadow_preset": shadow_preset,
                            "out_format": "PNG",
                            "composition_mode": composition_mode,
                        }

                        result = compose_one_bytes(item_img, template_img, **opts)
                        if result:
                            ss.cm_preview_list.append(result[0].getvalue())
                    except Exception:
                        pass

                # ZIP 파일 생성
                if valid_combinations:
                    zip_buf = io.BytesIO()
                    count = 0

                    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for item_file, template_file, mode in valid_combinations:
                            try:
                                item_file.seek(0)
                                template_file.seek(0)

                                item_img = PILImage.open(item_file)
                                template_img = PILImage.open(template_file)

                                template_ext = Path(template_file.name).suffix.lower()
                                composition_mode = "frame" if template_ext == ".png" else "normal"
                                shadow_preset = ss.cm_shadow_preset if composition_mode == "normal" else "off"

                                opts = {
                                    "anchor": ss.cm_anchor,
                                    "resize_ratio": ss.cm_resize_ratio,
                                    "shadow_preset": shadow_preset,
                                    "out_format": CONFIG["OUTPUT_FORMAT"],
                                    "quality": CONFIG["JPEG_QUALITY"],
                                    "composition_mode": composition_mode,
                                }

                                result = compose_one_bytes(item_img, template_img, **opts)
                                if result:
                                    img_buf, ext = result
                                    item_name = Path(item_file.name).stem
                                    template_code = Path(template_file.name).stem
                                    filename = f"{item_name}_C_{template_code}.{ext}"
                                    zf.writestr(filename, img_buf.getvalue())
                                    count += 1
                            except:
                                pass

                    zip_buf.seek(0)
                    ss.cm_zip_cache = (zip_buf.getvalue(), count, len(valid_combinations) - count)

            ss.cm_needs_preview_regen = False

        # 갤러리 표시
        if ss.cm_preview_list:
            total_count = len(ss.cm_preview_list)
            st.markdown(f"**미리보기** ({total_count}개)")

            st.markdown("""
            <style>
            .stImage > img {
                border: 1px solid #e6e6e6;
                border-radius: 4px;
                transition: transform 0.2s;
            }
            .stImage > img:hover {
                transform: scale(1.05);
            }
            </style>
            """, unsafe_allow_html=True)

            cols_per_row = CONFIG["GALLERY_COLS"]
            for i in range(0, total_count, cols_per_row):
                cols = st.columns(cols_per_row)
                for j, col in enumerate(cols):
                    idx = i + j
                    if idx < total_count:
                        with col:
                            st.image(ss.cm_preview_list[idx], use_column_width=True)
        else:
            st.info("조합 가능한 이미지가 없습니다.")
    else:
        st.info("👈 왼쪽에서 이미지를 업로드하면 미리보기가 표시됩니다.")

st.divider()

# 다운로드 버튼
if ss.cm_zip_cache:
    zip_data, success_count, invalid_count = ss.cm_zip_cache

    if success_count > 0:
        st.success(f"✅ 총 {success_count}장 생성 완료!")
        if invalid_count > 0:
            st.info(f"ℹ️ {invalid_count}개 조합은 자동으로 제외되었습니다.")

        now = datetime.now()
        date_time_str = now.strftime("%y%m%d_%H%M")
        zip_filename = f"CoverMaker_{date_time_str}.zip"

        download_key = f"cm_download_zip_{len(zip_data)}_{success_count}"

        st.download_button(
            label=f"📦 {zip_filename} 다운로드 ({success_count}장)",
            data=zip_data,
            file_name=zip_filename,
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key=download_key,
        )
    else:
        st.error("생성된 이미지가 없습니다. 조합을 확인해주세요.")
elif item_files and template_files and is_valid_tpl:
    st.info("설정을 조정하면 다운로드 파일이 자동으로 생성됩니다.")
else:
    st.info("이미지를 업로드하고 파일명을 확인해주세요.")

st.divider()
st.caption(f"Cover Maker {CONFIG['APP_VERSION']} - 쇼피 대량등록 도우미 통합 버전")
