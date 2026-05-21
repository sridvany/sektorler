"""TCMB Sektör Bilançoları XLSX indirici."""
import requests
import streamlit as st

BASE_URL = "https://www3.tcmb.gov.tr/sektor/dosyalar/tr"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www3.tcmb.gov.tr/sektor/",
    "Accept": "*/*",
}

FILE_TYPES = {
    "SEKKIM": "Sektör Kimliği",
    "BIL": "Bilanço",
    "GEL": "Gelir Tablosu",
    "YAPISAL": "Yapısal Analiz",
    "KARTIL": "Oranlar",
    "RISK": "Sektör Riski",
}

SCALES = {
    "GENEL": "Genel (Tüm Ölçekler)",
    "MIKRO": "Mikro",
    "KUCUK": "Küçük",
    "ORTA": "Orta",
    "BUYUK": "Büyük",
}


class FetchError(Exception):
    """TCMB indirme başarısız (cache'lenmemesi için exception kullanılır)."""


def build_url(sector_code: str, file_type: str, scale: str) -> str:
    return f"{BASE_URL}/{sector_code}_{file_type}_{scale}.xlsx"


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def fetch_xlsx_bytes(sector_code: str, file_type: str, scale: str) -> bytes:
    """XLSX'i indir ve bytes döner. Hata durumunda FetchError fırlatır
    (None döndürmeyiz, çünkü Streamlit None'u 24 saat cache'ler)."""
    url = build_url(sector_code, file_type, scale)
    try:
        r = requests.get(url, headers=HEADERS, timeout=30, allow_redirects=True)
    except requests.RequestException as e:
        raise FetchError(f"Bağlantı hatası: {e}")

    if r.status_code != 200:
        raise FetchError(f"HTTP {r.status_code}")

    # HTML hata sayfası mı geldi?
    head = r.content[:300].lower()
    if b"<html" in head or b"<!doc" in head:
        raise FetchError("Dosya yok (sunucu HTML hata sayfası döndü)")

    if len(r.content) < 500:
        raise FetchError(f"Dosya çok küçük ({len(r.content)} bayt)")

    return r.content
