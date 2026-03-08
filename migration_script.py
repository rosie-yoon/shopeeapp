"""
migration_script.py
기존 로컬 파일들을 구글 드라이브로 일괄 마이그레이션하는 스크립트
"""

from gdrive_manager import get_gdrive_manager
from pathlib import Path
import json


def migrate_existing_data():
    """기존 templates/ 및 config/ 폴더 데이터를 구글 드라이브로 마이그레이션"""
    print("🚀 구글 드라이브 마이그레이션 시작...\n")

    try:
        gdrive = get_gdrive_manager()
        print(f"✅ 구글 드라이브 연결 성공")
        print(f"📁 폴더 링크: {gdrive.get_folder_link()}\n")
    except Exception as e:
        print(f"❌ 구글 드라이브 연결 실패: {e}")
        return

    # 1. templates/ 폴더의 모든 xlsx 파일 업로드
    templates_dir = Path("templates")
    uploaded_count = 0

    if templates_dir.exists():
        print("📤 템플릿 파일 업로드 중...")
        for xlsx_file in templates_dir.glob("*.xlsx"):
            try:
                with open(xlsx_file, "rb") as f:
                    gdrive.upload_template(xlsx_file.name, f.read())
                print(f"  ✅ {xlsx_file.name}")
                uploaded_count += 1
            except Exception as e:
                print(f"  ❌ {xlsx_file.name} 실패: {e}")
        print(f"📊 템플릿 파일 {uploaded_count}개 업로드 완료\n")
    else:
        print("⚠️  templates/ 폴더가 없습니다.\n")

    # 2. config/*.json 파일들 업로드
    config_dir = Path("config")
    config_files = ["auto_rules.json", "global_rules.json"]
    config_count = 0

    if config_dir.exists():
        print("📤 설정 파일 업로드 중...")
        for json_file in config_files:
            json_path = config_dir / json_file
            if json_path.exists():
                try:
                    with open(json_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    gdrive.save_config_json(json_file, data)
                    print(f"  ✅ {json_file} ({len(data)} 항목)")
                    config_count += 1
                except Exception as e:
                    print(f"  ❌ {json_file} 실패: {e}")
            else:
                print(f"  ⚠️  {json_file} 파일 없음")
        print(f"📊 설정 파일 {config_count}개 업로드 완료\n")
    else:
        print("⚠️  config/ 폴더가 없습니다.\n")

    print("✨ 마이그레이션 완료!")
    print(f"📁 구글 드라이브에서 확인: {gdrive.get_folder_link()}")


if __name__ == "__main__":
    # Streamlit 없이 실행하기 위한 임시 설정
    import streamlit as st

    if not hasattr(st, 'secrets'):
        st.secrets = {}

    migrate_existing_data()