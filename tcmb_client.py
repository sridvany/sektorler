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


def build_url(sector_code: str, file_type: str, scale: str) -> str:
    return f"{BASE_URL}/{sector_code}_{file_type}_{scale}.xlsx"


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def fetch_xlsx_bytes(sector_code: str, file_type: str, scale: str):
    """XLSX'i indir; başarılıysa bytes, yoksa None döner.
    Cache 24 saat. TCMB sayfası User-Agent + Referer kontrolü yapıyor."""
    url = build_url(sector_code, file_type, scale)
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        # XLSX magic bytes ile doğrula (HTML hata sayfası gelirse ayırt edelim)
        if r.status_code == 200 and r.content[:4] == b"PK\x03\x04":
            return r.content
        return None
    except requests.RequestException:
        return None
