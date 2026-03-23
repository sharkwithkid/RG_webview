import unicodedata
from typing import Optional


def normalize_text(s: Optional[str]) -> str:
    """
    OS 독립적 문자열 비교를 위한 정규화

    1) NFC 정규화 (Mac/Windows 차이 제거)
    2) 앞뒤 공백 제거
    3) 내부 공백 제거
    4) 언더바 제거
    5) 소문자 변환 (영문 대비)
    """
    if s is None:
        return ""
    s = unicodedata.normalize("NFC", str(s))
    s = s.strip()
    s = s.replace(" ", "")
    s = s.replace("_", "")
    s = s.lower()
    return s


def text_eq(a: Optional[str], b: Optional[str]) -> bool:
    return normalize_text(a) == normalize_text(b)


def text_contains(text: Optional[str], keyword: Optional[str]) -> bool:
    return normalize_text(keyword) in normalize_text(text)
