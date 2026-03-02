import os
import re
import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from common.google_drive import _extract_drive_file_id

load_dotenv()
_GOOGLE_SECRET_JSON = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON') or os.getenv('GOOGLE_SECRET_JSON')
if not _GOOGLE_SECRET_JSON:
    raise RuntimeError(
        "GOOGLE_SERVICE_ACCOUNT_JSON 환경변수가 설정되지 않았습니다. "
        ".env 파일에 서비스 계정 키 경로를 지정해주세요."
    )

_credentials = service_account.Credentials.from_service_account_file(
    _GOOGLE_SECRET_JSON,
    scopes=[
        'https://www.googleapis.com/auth/documents.readonly',
        'https://www.googleapis.com/auth/drive.readonly'
    ]
)
_docs_service = build('docs', 'v1', credentials=_credentials)
_drive_service = build('drive', 'v3', credentials=_credentials)

def _get_urls_from_doc(doc_id, tx_amount=None):
    """Google Docs API로 문서에서 이미지 URL 추출 및 이체 수수료 여부 확인.

    tx_amount: 거래 금액 절댓값 (원) — 영수증 금액 비교를 통한 수수료 감지용
    반환: (urls, has_transfer_fee)
    """
    doc = _docs_service.documents().get(documentId=doc_id).execute()
    urls = []
    has_transfer_fee = False
    max_krw = None

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

    for elem in doc.get('body', {}).get('content', []):
        para = elem.get('paragraph', {})
        if not para:
            continue
        elements = para.get('elements', [])
        # 단락 전체 텍스트 합산 (textRun 분할 여부와 무관하게 검사)
        para_text = ''.join(pe.get('textRun', {}).get('content', '') for pe in elements)

        # 텍스트 기반 이체 수수료 감지
        if '이체' in para_text and '수수료' in para_text:
            has_transfer_fee = True

        # 금액 기반 감지용 — 단락에서 원화 금액 추출
        if tx_amount is not None:
            for m in re.finditer(r'([\d,]+)원', para_text):
                try:
                    amt = int(m.group(1).replace(',', ''))
                    if max_krw is None or amt > max_krw:
                        max_krw = amt
                except ValueError:
                    pass

        # 텍스트 링크
        for pe in elements:
            link = pe.get('textRun', {}).get('textStyle', {}).get('link', {}).get('url')
            if link and 'drive.google.com' in link:
                urls.append(link)

    # 금액 기반 이체 수수료 감지 (영수증에 문구가 없는 경우)
    if not has_transfer_fee and tx_amount is not None and max_krw is not None:
        if tx_amount - max_krw == 500:
            has_transfer_fee = True

    return urls, has_transfer_fee


def _get_urls(url, tx_amount=None):
    """이미지 URL 목록과 이체 수수료 포함 여부 반환: (urls, has_transfer_fee)"""
    if not url: return [], False

    file_id = _extract_drive_file_id(url)

    # Drive 직접 링크인 경우 MIME 타입 확인
    if file_id and 'docs.google.com' not in url:
        mime = _drive_service.files().get(fileId=file_id, fields='mimeType').execute().get('mimeType', '')
        if mime == 'application/vnd.google-apps.document':
            return _get_urls_from_doc(file_id, tx_amount=tx_amount)
        return [url], False  # 이미지 등 바이너리 파일

    if not file_id: return [], False
    return _get_urls_from_doc(file_id, tx_amount=tx_amount)

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
    os.makedirs(img_dir, exist_ok=True)
    all_files = os.listdir(img_dir)
    img_paths_list = []
    has_fee_list = []

    for idx, row in data.iterrows():
        prefix = f"row_{idx + 1}_"
        fee_filename = f'{prefix}fee'
        fee_file = os.path.join(img_dir, fee_filename)

        # 이미지 파일만 캐시 목록에 포함 (fee 파일 제외)
        cached = sorted([
            os.path.join(img_dir, f) for f in all_files
            if f.startswith(prefix) and f != fee_filename
        ])

        if cached and os.path.exists(fee_file):
            # 이미지 + 이체 수수료 모두 캐시된 경우
            img_paths_list.append(cached)
            with open(fee_file) as f:
                has_fee = f.read().strip() == '1'
        else:
            link = row['링크']
            tx_amount = abs(int(row['입/출']))
            urls, has_fee = _get_urls(link, tx_amount=tx_amount) if isinstance(link, str) and link.strip() else ([], False)

            if cached:
                img_paths_list.append(cached)
            else:
                paths = _download(urls, img_dir, prefix=prefix)
                img_paths_list.append(paths)

            with open(fee_file, 'w') as f:
                f.write('1' if has_fee else '0')

        has_fee_list.append(has_fee)

    data['img_paths'] = img_paths_list
    data['이체수수료'] = has_fee_list
    return data
