"""
gdrive_manager.py
구글 드라이브 연동을 위한 통합 관리 모듈 (OAuth 2.0 사용자 인증)
"""

import io
import json
import os
from pathlib import Path
from typing import Optional, Dict, List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# 권한 범위
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# 드라이브 폴더 구조
ROOT_FOLDER_NAME = "Shopee_Upload_Helper"
TEMPLATES_SUBFOLDER = "templates"


class GDriveManager:
    """구글 드라이브 파일 관리 통합 클래스 (OAuth 2.0)"""

    def __init__(self):
        self.service = None
        self.root_folder_id = None
        self.templates_folder_id = None
        self._authenticate()
        self._ensure_folder_structure()

    def _authenticate(self):

    def _ensure_folder_structure(self):
        """필요한 폴더 구조 자동 생성"""
        self.root_folder_id = self._get_or_create_folder(ROOT_FOLDER_NAME)
        self.templates_folder_id = self._get_or_create_folder(
            TEMPLATES_SUBFOLDER, parent_id=self.root_folder_id
        )

    def _get_or_create_folder(self, folder_name: str, parent_id: str = None) -> str:
        """폴더 찾기 또는 생성"""
        query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        if parent_id:
            query += f" and '{parent_id}' in parents"

        try:
            results = self.service.files().list(
                q=query, spaces='drive', fields='files(id, name)'
            ).execute()
            folders = results.get('files', [])

            if folders:
                return folders[0]['id']

            # 폴더 생성
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            if parent_id:
                folder_metadata['parents'] = [parent_id]

            folder = self.service.files().create(
                body=folder_metadata, fields='id'
            ).execute()
            return folder['id']
        except Exception as e:
            raise Exception(f"폴더 생성/검색 실패 ({folder_name}): {e}")

    def upload_template(self, filename: str, file_bytes: bytes) -> str:
        """템플릿 파일 업로드 (덮어쓰기 지원)"""
        return self._upload_file(
            filename, file_bytes, self.templates_folder_id,
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    def download_template(self, filename: str) -> Optional[bytes]:
        """템플릿 파일 다운로드"""
        return self._download_file(filename, self.templates_folder_id)

    def list_templates(self) -> List[str]:
        """템플릿 목록 조회"""
        return self._list_files_in_folder(self.templates_folder_id)

    def save_config_json(self, filename: str, data: Dict):
        """설정 JSON 파일 저장 (루트 폴더)"""
        json_bytes = json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
        self._upload_file(filename, json_bytes, self.root_folder_id, 'application/json')

    def load_config_json(self, filename: str) -> Dict:
        """설정 JSON 파일 로드"""
        file_bytes = self._download_file(filename, self.root_folder_id)
        if not file_bytes:
            return {}
        try:
            return json.loads(file_bytes.decode('utf-8'))
        except json.JSONDecodeError as e:
            print(f"JSON 파싱 오류 ({filename}): {e}")
            return {}

    def _upload_file(self, filename: str, file_bytes: bytes, folder_id: str, mime_type: str) -> str:
        """파일 업로드 (내부 메서드)"""
        try:
            # 기존 파일 검색
            query = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
            results = self.service.files().list(q=query, fields='files(id)').execute()
            existing = results.get('files', [])

            media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type)

            if existing:
                # 덮어쓰기
                file_id = existing[0]['id']
                self.service.files().update(fileId=file_id, media_body=media).execute()
                return file_id
            else:
                # 새 파일 생성
                file_metadata = {'name': filename, 'parents': [folder_id]}
                file = self.service.files().create(
                    body=file_metadata, media_body=media, fields='id'
                ).execute()
                return file['id']
        except Exception as e:
            raise Exception(f"파일 업로드 실패 ({filename}): {e}")

    def _download_file(self, filename: str, folder_id: str) -> Optional[bytes]:
        """파일 다운로드 (내부 메서드)"""
        try:
            query = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
            results = self.service.files().list(q=query, fields='files(id)').execute()
            files = results.get('files', [])

            if not files:
                return None

            file_id = files[0]['id']
            request = self.service.files().get_media(fileId=file_id)
            buffer = io.BytesIO()
            downloader = MediaIoBaseDownload(buffer, request)

            done = False
            while not done:
                status, done = downloader.next_chunk()

            buffer.seek(0)
            return buffer.read()
        except Exception as e:
            print(f"파일 다운로드 실패 ({filename}): {e}")
            return None

    def _list_files_in_folder(self, folder_id: str) -> List[str]:
        """폴더 내 파일 목록 (내부 메서드)"""
        try:
            query = f"'{folder_id}' in parents and trashed=false"
            results = self.service.files().list(
                q=query, fields='files(name)', orderBy='name'
            ).execute()
            files = results.get('files', [])
            return [f['name'] for f in files]
        except Exception as e:
            print(f"파일 목록 조회 실패: {e}")
            return []

    def get_folder_link(self) -> str:
        """루트 폴더 공유 링크"""
        return f"https://drive.google.com/drive/folders/{self.root_folder_id}"


# 싱글톤 인스턴스
try:
    import streamlit as st
    @st.cache_resource
    def get_gdrive_manager() -> GDriveManager:
        """캐시된 GDriveManager 인스턴스 반환"""
        return GDriveManager()
except ImportError:
    # Streamlit 없는 환경
    _gdrive_manager_instance = None
    def get_gdrive_manager() -> GDriveManager:
        global _gdrive_manager_instance
        if _gdrive_manager_instance is None:
            _gdrive_manager_instance = GDriveManager()
        return _gdrive_manager_instance
