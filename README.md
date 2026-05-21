# TCMB Sektör Bilançoları

[TCMB Sektör Bilançoları](https://www3.tcmb.gov.tr/sektor/) sayfasındaki verileri
görüntüleyen ve farklı sektör/ölçekleri yan yana karşılaştıran bir Streamlit uygulaması.

## Özellikler

- 21 ana sektör (NACE Rev.2) ve tüm 2'li / 3'lü alt kırılımları
- 5 ölçek: Genel, Mikro, Küçük, Orta, Büyük
- 6 dosya tipi: Sektör Kimliği, Bilanço, Gelir Tablosu, Yapısal Analiz, Oranlar, Sektör Riski
- Karşılaştırma modu: iki sektör/ölçeği yan yana göster
- Excel olarak indir

## Çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

1. Repoyu GitHub'a push et
2. https://share.streamlit.io adresinde repoyu bağla
3. Ana dosya: `app.py`

## Veri Notu

TCMB her yıl yalnızca son 2 yılın verisini yayınlar (örn. 2025 yayını = 2023-2024).
Önceki yıllar için TCMB EVDS sistemi kullanılabilir.

Gizlilik kuralı gereği bazı sektör+ölçek kombinasyonları yayınlanmaz
(firma sayısı 12'den az veya tek firma %80+ paya sahipse).

## Dosya yapısı

```
.
├── app.py              # Streamlit UI
├── tcmb_client.py      # URL builder + cache'li indirici
├── parsers.py          # 6 dosya tipi için XLSX parser'ları
├── nace.py             # NACE Rev.2 hiyerarşi parser
├── naceler.txt         # NACE kod listesi (statik)
└── requirements.txt
```
