import re

DRIVE_URL_PATTERN = re.compile(
    r"^https://drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)"
)

DOC_URL_PATTERN = re.compile(
    r"^https://docs\.google\.com/document/d/([a-zA-Z0-9_-]+)/edit"
)

TARGET_FORMAT = "https://docs.google.com/document/d/{}/edit"


def convert_to_doc_url(url: str) -> str:
    open_match = DRIVE_URL_PATTERN.match(url)
    if open_match:
        file_id = open_match.group(1)
        return TARGET_FORMAT.format(file_id)

    doc_match = DOC_URL_PATTERN.match(url)
    if doc_match:
        file_id = doc_match.group(1)
        return TARGET_FORMAT.format(file_id)

    raise ValueError("유효한 Google Docs URL 형식이 아닙니다.")
