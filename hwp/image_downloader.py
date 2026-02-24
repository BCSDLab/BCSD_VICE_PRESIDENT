import os
import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from urllib.parse import urlparse, parse_qs

load_dotenv()
_GOOGLE_SECRET_JSON = os.getenv('GOOGLE_SECRET_JSON')

_credentials = service_account.Credentials.from_service_account_file(
    _GOOGLE_SECRET_JSON,
    scopes=[
        'https://www.googleapis.com/auth/documents.readonly',
        'https://www.googleapis.com/auth/drive.readonly'
    ]
)
_docs_service = build('docs', 'v1', credentials=_credentials)
_drive_service = build('drive', 'v3', credentials=_credentials)

def _extract_id(url):
    """URL에서 Drive 파일 ID 추출"""
    if '/d/' in url:
        return url.split('/d/')[1].split('/')[0]
    if 'id=' in url:
        return parse_qs(urlparse(url).query).get('id', [None])[0]
    return None


def _get_urls_from_doc(doc_id):
    """Google Docs API로 문서에서 이미지 URL 추출"""
    doc = _docs_service.documents().get(documentId=doc_id).execute()
    urls = []

    # 인라인 객체에서 이미지 추출 (Drive 링크 삽입은 sourceUri, 직접 업로드는 contentUri)
    for obj in doc.get('inlineObjects', {}).values():
        props = (obj.get('inlineObjectProperties', {})
                    .get('embeddedObject', {})
                    .get('imageProperties', {}))
        source_uri = props.get('sourceUri')
        content_uri = props.get('contentUri')
        if source_uri and 'drive.google.com' in source_uri:
            urls.append(source_uri)
        elif content_uri:
            urls.append(content_uri)

    # 텍스트 링크
    for elem in doc.get('body', {}).get('content', []):
        for para_elem in elem.get('paragraph', {}).get('elements', []):
            link = para_elem.get('textRun', {}).get('textStyle', {}).get('link', {}).get('url')
            if link and 'drive.google.com' in link:
                urls.append(link)

    return urls


def _get_urls(url):
    if not url: return []

    file_id = _extract_id(url)

    # Drive 직접 링크인 경우 MIME 타입 확인
    if file_id and 'docs.google.com' not in url:
        mime = _drive_service.files().get(fileId=file_id, fields='mimeType').execute().get('mimeType', '')
        if mime == 'application/vnd.google-apps.document':
            return _get_urls_from_doc(file_id)
        return [url]  # 이미지 등 바이너리 파일

    if not file_id: return []
    return _get_urls_from_doc(file_id)

def _download(urls, dir, prefix):
    """Google Drive API로 고화질 다운로드"""
    paths = []
    for i, url in enumerate(urls):
        # Drive 파일 ID 추출
        file_id = None
        if '/d/' in url:
            file_id = url.split('/d/')[1].split('/')[0]
        elif 'id=' in url:
            file_id = parse_qs(urlparse(url).query).get('id', [None])[0]
        
        ext, content = None, None
        if file_id:
            file_info = _drive_service.files().get(fileId=file_id).execute()
            file_name = file_info.get('name', 'image.png')
            ext = file_name.split('.')[-1]
            
            content = _drive_service.files().get_media(fileId=file_id).execute()
        else:
            res = requests.get(url, allow_redirects=True)
            if res.status_code != 200: continue
            ext = res.headers.get('Content-Type', 'image/png').split('/')[-1]
            content = res.content
        
        if not ext or not content: continue

        path = os.path.join(dir, f'{prefix}_{i}.{ext}')
        
        with open(path, 'wb') as f:
            f.write(content)
        paths.append(path)
    
    return paths

def run(data, img_dir):
    img_paths_list = []

    if os.path.exists(img_dir):
        all_files = os.listdir(img_dir)
        for idx, row in data.iterrows():
            prefix = f"row_{idx + 1}_"
            row_imgs = sorted([os.path.join(img_dir, f) for f in all_files if f.startswith(prefix)])
            img_paths_list.append(row_imgs)
    else:
        os.makedirs(img_dir, exist_ok=True)
        for idx, row in data.iterrows():
            urls = _get_urls(row['링크'])
            paths = _download(urls, img_dir, prefix=f"row_{idx + 1}")
            img_paths_list.append(paths)
    
    data['img_paths'] = img_paths_list
    return data
