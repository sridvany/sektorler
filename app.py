"""TCMB Sektör Bilançoları - Streamlit Görüntüleyici & Karşılaştırıcı."""
import streamlit as st
import pandas as pd

from nace import load_nace_tree, get_subsectors
from tcmb_client import fetch_xlsx_bytes, build_url, FILE_TYPES, SCALES
from parsers import parse


st.set_page_config(
    page_title="TCMB Sektör Bilançoları",
    page_icon="📊",
    layout="wide",
)

st.title("📊 TCMB Sektör Bilançoları")
st.caption(
    "Kaynak: https://www3.tcmb.gov.tr/sektor/ · "
    "Her yayın yalnızca son 2 yılı içerir (en güncel: 2023-2024)."
)


@st.cache_data
def _tree():
    return load_nace_tree()


tree = _tree()
MAIN_LETTERS = sorted(tree.keys())
MAIN_LABELS = {ltr: f"{ltr} - {tree[ltr]['name']}" for ltr in MAIN_LETTERS}


def sector_picker(prefix: str = ""):
    """Sidebar'da bir sektör seçici grubu çizer; (sector_code, scale, label) döner."""
    main_letter = st.sidebar.selectbox(
        "Ana Sektör",
        options=MAIN_LETTERS,
        format_func=lambda l: MAIN_LABELS[l],
        key=f"{prefix}main",
    )
    sub_opts = get_subsectors(tree, main_letter)
    sub_choices = [("__main__", f"Tüm {main_letter} sektörü")] + sub_opts
    sub_labels = dict(sub_choices)
    selected_sub = st.sidebar.selectbox(
        "Alt Sektör",
        options=[c for c, _ in sub_choices],
        format_func=lambda c: sub_labels[c],
        key=f"{prefix}sub",
    )
    sector_code = main_letter if selected_sub == "__main__" else selected_sub

    scale = st.sidebar.selectbox(
        "Ölçek",
        options=list(SCALES.keys()),
        format_func=lambda s: SCALES[s],
        key=f"{prefix}scale",
    )

    label = f"{sector_code} · {SCALES[scale]}"
    return sector_code, scale, label


# ---------- SIDEBAR ----------
st.sidebar.header("Birinci Sektör")
sector_code, scale, label_1 = sector_picker(prefix="a_")

st.sidebar.divider()
compare_on = st.sidebar.toggle("🔀 Karşılaştırma Modu", value=False)

sector_code_2 = scale_2 = label_2 = None
if compare_on:
    st.sidebar.header("İkinci Sektör")
    sector_code_2, scale_2, label_2 = sector_picker(prefix="b_")

st.sidebar.divider()
st.sidebar.caption(
    "ℹ️ Bazı sektör+ölçek kombinasyonlarında veri bulunmayabilir "
    "(gizlilik kuralı: firma sayısı 12'den az veya tek firma %80+ paya sahipse "
    "TCMB yayın yapmaz)."
)


# ---------- RENDER ----------
def fmt_thousands(df: pd.DataFrame) -> pd.DataFrame:
    """Sayısal sütunları binlik ayraçla görüntüle (kopya yapar)."""
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                if pd.notna(x) and isinstance(x, (int, float))
                else x
            )
    return out


def render_one(sector_code: str, scale: str, file_type: str):
    """Tek sektör + ölçek için belirli bir dosya tipini göster."""
    with st.spinner(f"{sector_code}_{file_type}_{scale}.xlsx indiriliyor..."):
        content = fetch_xlsx_bytes(sector_code, file_type, scale)

    if content is None:
        st.warning(
            f"**Veri bulunamadı:** `{sector_code}_{file_type}_{scale}.xlsx`\n\n"
            "Olası nedenler: bu sektör+ölçek kombinasyonu için TCMB'de dosya yok, "
            "veya gizlilik nedeniyle yayınlanmamış.\n\n"
            f"Denenen URL: {build_url(sector_code, file_type, scale)}"
        )
        return

    try:
        result, meta = parse(file_type, content)
    except Exception as e:
        st.error(f"Dosya okunamadı: {e}")
        return

    if meta.get("sector_name"):
        st.caption(f"**{meta['sector_name']}** · {meta.get('title', '')}")

    if file_type == "SEKKIM":
        blocks = result
        st.markdown("**📋 Genel Bilgi**")
        st.dataframe(blocks["info"], hide_index=True, use_container_width=True)

        st.markdown("**⚖️ Hukuki Yapı**")
        st.dataframe(fmt_thousands(blocks["legal"]), hide_index=True, use_container_width=True)

        st.markdown("**📏 Ölçek Dağılımı**")
        st.dataframe(fmt_thousands(blocks["scale"]), hide_index=True, use_container_width=True)

        st.markdown("**💳 Sektör Riski (Bin TL)**")
        st.dataframe(fmt_thousands(blocks["risk"]), hide_index=True, use_container_width=True)

        st.markdown("**📈 Bilanço Dönemi Sonuçları**")
        st.dataframe(fmt_thousands(blocks["period"]), hide_index=True, use_container_width=True)
    else:
        st.dataframe(
            fmt_thousands(result),
            hide_index=True,
            use_container_width=True,
            height=600,
        )

    st.download_button(
        f"⬇️ Excel olarak indir",
        data=content,
        file_name=f"{sector_code}_{file_type}_{scale}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_{sector_code}_{file_type}_{scale}",
    )


# ---------- TABS ----------
tab_labels = [FILE_TYPES[k] for k in FILE_TYPES]
tabs = st.tabs(tab_labels)

KARTIL_HELP = """
**Bu tablo, sektördeki firmaların oran dağılımını gösterir.** Toplulaştırılmış tutarlar değil — firma bazında hesaplanmış oranların istatistiksel özetidir. Her yıl için 5 sütun var:

- **Firma**: O oranın pay/paydası sıfır olmayan firma sayısı (uç değer analizi sonrası).
- **Q (Ağr.)**: Aktif büyüklüğüne göre **ağırlıklı ortalama**. Büyük firmalar daha fazla ağırlık taşır.
- **Q1**: Alt çeyrek. Firmaların **%25'i** bu değerin altında.
- **Q2 (Med.)**: **Medyan**. Sıralandığında tam ortadaki firmanın değeri — "tipik firma" budur.
- **Q3**: Üst çeyrek. Firmaların **%75'i** bu değerin altında.

**Örnek — Cari Oran:** Q1=82, Q2=110, Q3=168, Q (Ağr.)=143

- Tipik (medyan) firmanın cari oranı **110**.
- Sektörün alt %25'i 82'nin altında çalışıyor (likidite zayıf).
- Üst %25'i 168'in üstünde (rahat).
- Ağırlıklı ortalama (143) medyandan (110) yüksek → birkaç büyük firma sektörü yukarı çekiyor.

**Kendi firmanla karşılaştırma:** Cari oranın 95 ise → Q1'in üstünde ama medyanın altında, **alt yarıda** sıralanıyorsun.

**Dağınıklık ölçüsü:** Q3 − Q1 farkı geniş ise sektörde firma farklılıkları büyük, dar ise homojen.

ℹ️ TCMB uç değerleri Tukey's Hinges yöntemiyle dışlar; "Firma" sütunu analizdeki firma sayısını verir.
"""

for tab, (file_type, _) in zip(tabs, FILE_TYPES.items()):
    with tab:
        if file_type == "KARTIL":
            with st.expander("ℹ️ Q, Q1, Q2, Q3 ne anlama geliyor? (örnekli açıklama)"):
                st.markdown(KARTIL_HELP)
        if compare_on and sector_code_2:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader(label_1)
                render_one(sector_code, scale, file_type)
            with col2:
                st.subheader(label_2)
                render_one(sector_code_2, scale_2, file_type)
        else:
            render_one(sector_code, scale, file_type)
