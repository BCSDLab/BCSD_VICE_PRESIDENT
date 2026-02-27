import io
import os
import re
import tempfile
from urllib.parse import parse_qs, urlparse


# ============================================================================
# URL parsing
# ============================================================================

def _extract_sheet_id(url):
    """Google Sheets URL에서 스프레드시트 ID 추출."""
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not match:
        raise ValueError(f'Google Sheets URL에서 ID를 파싱할 수 없습니다: {url}')
    return match.group(1)


def _extract_drive_folder_id(url):
    """Google Drive 폴더 URL에서 폴더 ID 추출."""
    if '/folders/' in url:
        return url.split('/folders/')[1].split('/')[0].split('?')[0]
    return parse_qs(urlparse(url).query).get('id', [None])[0]


def _extract_drive_file_id(url):
    """Google Drive 파일 URL에서 파일 ID 추출."""
    if '/d/' in url:
        return url.split('/d/')[1].split('/')[0]
    if 'id=' in url:
        return parse_qs(urlparse(url).query).get('id', [None])[0]
    return None


# ============================================================================
# Download
# ============================================================================

def _download_request_to_tempfile(request, suffix='.xlsx'):
    """googleapiclient request를 임시 파일로 다운로드."""
    from googleapiclient.http import MediaIoBaseDownload

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        tmp.write(buf.getvalue())
        tmp.close()
    except Exception:
        tmp.close()
        os.unlink(tmp.name)
        raise

    return tmp.name
