# -*- coding: utf-8 -*-
"""
Import fiche descriptive (.docx) -> HTML + nettoyage + forçages stricts
"""

import base64
import re
from io import BytesIO
from typing import Any, Dict, List, Tuple

import mammoth
import streamlit as st
from bs4 import BeautifulSoup, Tag, NavigableString

# -----------------------------------------------------------------------------
# CONFIG
# -----------------------------------------------------------------------------

st.set_page_config(page_title="Import fiche descriptive", layout="wide")

CSS = """
<style>
section.main > div { max-width: 1400px; }
h2,h3 { margin: .6rem 0 .35rem; }
hr { margin: 1rem 0; }
p { margin: .25rem 0 .45rem; line-height: 1.45; }
li { margin: .15rem 0; line-height: 1.45; }
.card { border:1px solid #e8e8e8; border-radius:8px; padding:12px 14px; margin:12px 0 20px; background:#fff;}
.card h3 { margin:0 0 8px; }
.anchor { margin-left:6px; font-size:.9rem; opacity:.5; }
.img-placeholder { display:inline-block; padding:6px 8px; margin:4px 0; background:#fff7e6; border:1px dashed #f0ad4e; color:#9a6c00; border-radius:5px; font-size:.9rem; }
ol, ul { margin-left: 20px; }
h1 { font-size: 1.15rem; font-weight: 600; }
h2 { font-size: 1.10rem; font-weight: 600; }
h3 { font-size: 1.05rem; font-weight: 600; }
strong { font-weight:600; }
</style>
"""

EXPECTED_HEADINGS = [
    "Description",
    "Contexte et usage des fonds",
    "Les points d'attention",
    "Les bonnes raisons d'investir",
    "Présentation de l'opération",
    "Localisation",
    "Planning",
    "Marché et références",
    "Budget",
    "Présentation de l'opérateur",
    "Actionnariat",
    "Finances",
]

CRM_KEY_BY_LABEL = {
    "Description": "description_fr",
    "Contexte et usage des fonds": "contexte_fonds_fr",
    "Les points d'attention": "points_attention_fr",
    "Les bonnes raisons d'investir": "bonnes_raisons_fr",
    "Présentation de l'opération": "operation_presentation_fr",
    "Localisation": "localisation_fr",
    "Planning": "planning_fr",
    "Marché et références": "marche_references_fr",
    "Budget": "budget_fr",
    "Présentation de l'opérateur": "operateur_presentation_fr",
    "Actionnariat": "actionnariat_fr",
    "Finances": "finances_fr",
}

# Libellés pour forçages
RISKS_LABELS = [
    "risque lié au projet",
    "risque lié au secteur",
    "risque de défaut",
]
BONNES_RAISONS_ASSURANCE = "une assurance sur 100% du capital investi"
BONNES_RAISONS_FIDUCIE = "une fiducie-sûreté sur l'actif"
BUDGET_TITLES = [
    "prix de revient",
    "financement et ratios",
    "revenus et marges",
    "couverture des intérêts",
    "stress test",
]

# -----------------------------------------------------------------------------
# OUTILS
# -----------------------------------------------------------------------------

def docx_to_html(file_bytes: bytes) -> str:
    """DOCX -> HTML via Mammoth."""
    result = mammoth.convert_to_html(
        BytesIO(file_bytes),
        include_default_style_map=True,
        style_map="",
        convert_image=mammoth.images.inline(lambda image: {"src": _img_to_data_uri(image)}),
    )
    return result.value or ""

def _img_to_data_uri(image: Any) -> str:
    """Mammoth Image -> data URI (utilise image.open())."""
    content_type = getattr(image, "content_type", None) or "application/octet-stream"
    # `open()` donne un flux binaire
    try:
        with image.open() as fp:
            data = fp.read()
    except Exception:
        # fallback (au cas où)
        data = b""
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{content_type};base64,{b64}"

def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip().lower()

def strip_leading_numbering(txt: str) -> str:
    if not txt:
        return ""
    return re.sub(r"^\s*([0-9ivxlcdmIVXLCDM]+[\.\)\-:])\s*", "", txt).strip()

def _drop_stray_img_alt_strings(soup: BeautifulSoup) -> None:
    pattern = re.compile(r"^\s*<img\s+alt=.*?>\s*/?>\s*$", re.I)
    for node in list(soup.find_all(text=True)):
        if isinstance(node, NavigableString) and pattern.match(str(node)):
            node.extract()

def prepare_section_html(raw_html: str) -> Tuple[str, List[Tuple[str, bytes]]]:
    """Nettoyage de base d'une section + collecte d'EMF/WMF."""
    soup = BeautifulSoup(raw_html, "html.parser")
    _drop_stray_img_alt_strings(soup)

    downloads: List[Tuple[str, bytes]] = []
    for img in list(soup.find_all("img")):
        src = img.get("src", "")
        if not src.startswith("data:image/"):
            continue
        m = re.match(r"^data:(image/[^;]+);base64,(.*)$", src, re.I)
        if not m:
            continue
        mime = m.group(1).lower()
        b64 = m.group(2)
        if mime in ("image/emf", "image/x-emf", "image/wmf", "image/x-wmf"):
            try:
                blob = base64.b64decode(b64)
                fname = f"{mime.split('/')[-1]}_image.emf"
                downloads.append((fname, blob))
                ph = soup.new_tag("span", **{"class": "img-placeholder"})
                ph.string = f"[Image {mime} non affichable - disponible en téléchargement]"
                img.replace_with(ph)
            except Exception:
                pass

    return str(soup), downloads

# -----------------------------------------------------------------------------
# FORÇAGES STRICTS
# -----------------------------------------------------------------------------

def _li_starts_with(li: Tag, label: str) -> bool:
    txt = strip_leading_numbering(li.get_text(" ", strip=True))
    return _normalize_text(txt).startswith(_normalize_text(label))

def _find_items_by_labels_in_soup(soup: BeautifulSoup, labels: List[str]) -> Dict[str, Tag]:
    found: Dict[str, Tag] = {}
    for li in soup.find_all("li"):
        for lab in labels:
            if lab not in found and _li_starts_with(li, lab):
                found[lab] = li
    return found

def _convert_other_ol_to_ul(soup: BeautifulSoup, keep: Tag) -> None:
    for ol in soup.find_all("ol"):
        if ol is not keep:
            ol.name = "ul"

def _wrap_whole_tag_with_em_u(tag: Tag) -> None:
    if not isinstance(tag, Tag) or not hasattr(tag, "new_tag"):
        return
    children = list(tag.contents)
    tag.clear()
    em = tag.new_tag("em"); u = tag.new_tag("u")
    tag.append(em); em.append(u)
    for c in children: u.append(c)

def _postprocess_risks(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    base_ol = soup.find("ol")
    found = _find_items_by_labels_in_soup(soup, RISKS_LABELS)
    if not found:
        # Rien à forcer → convertir toutes les <ol> en <ul> pour éviter 1..16
        for ol in soup.find_all("ol"):
            ol.name = "ul"
        return str(soup)

    new_ol = soup.new_tag("ol")
    for lab in RISKS_LABELS:
        if lab in found:
            new_ol.append(found[lab].extract())

    if base_ol:
        base_ol.replace_with(new_ol)
    else:
        soup.insert(0, new_ol)

    _convert_other_ol_to_ul(soup, new_ol)  # le reste en puces
    return str(soup)

def _postprocess_bonnes_raisons(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    base_ol = soup.find("ol")
    wanted = []
    found = _find_items_by_labels_in_soup(
        soup, [BONNES_RAISONS_ASSURANCE, BONNES_RAISONS_FIDUCIE]
    )
    if BONNES_RAISONS_ASSURANCE in found:
        wanted = [BONNES_RAISONS_ASSURANCE, BONNES_RAISONS_FIDUCIE]
    elif BONNES_RAISONS_FIDUCIE in found:
        wanted = [BONNES_RAISONS_FIDUCIE]
    else:
        # Rien à forcer → convertir toutes les <ol> en <ul> (sécurité)
        for ol in soup.find_all("ol"):
            ol.name = "ul"
        return str(soup)

    new_ol = soup.new_tag("ol")
    for lab in wanted:
        if lab in found:
            new_ol.append(found[lab].extract())

    if base_ol:
        base_ol.replace_with(new_ol)
    else:
        soup.insert(0, new_ol)

    _convert_other_ol_to_ul(soup, new_ol)
    return str(soup)

def _postprocess_budget(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    targets = soup.find_all(["p", "li", "h1", "h2", "h3", "strong"])
    titles_norm = [_normalize_text(t) for t in BUDGET_TITLES]
    for t in targets:
        txt = strip_leading_numbering(t.get_text(" ", strip=True))
        if any(_normalize_text(txt).startswith(n) for n in titles_norm):
            _wrap_whole_tag_with_em_u(t)
    return str(soup)

def postprocess_domain_section(section_label: str, clean_html: str) -> str:
    label = _normalize_text(section_label or "")
    if label == _normalize_text("Les points d'attention"):
        return _postprocess_risks(clean_html)
    if label == _normalize_text("Les bonnes raisons d'investir"):
        return _postprocess_bonnes_raisons(clean_html)
    if label == _normalize_text("Budget"):
        return _postprocess_budget(clean_html)
    return clean_html

# -----------------------------------------------------------------------------
# SPLIT SECTIONS
# -----------------------------------------------------------------------------

HDR_RE = re.compile(r"^h[1-4]$", re.I)

def split_into_sections(full_html: str) -> Dict[str, str]:
    soup = BeautifulSoup(full_html, "html.parser")
    sections: Dict[str, str] = {}

    headers = soup.find_all(HDR_RE)
    if not headers:
        sections["Description"] = str(soup)
        return sections

    for i, h in enumerate(headers):
        title = h.get_text(" ", strip=True)
        content_nodes: List[Tag] = []
        cur = h.next_sibling
        while cur and not (isinstance(cur, Tag) and HDR_RE.match(cur.name or "")):
            if isinstance(cur, Tag):
                content_nodes.append(cur)
            cur = cur.next_sibling
        sections[title] = "".join(str(n) for n in content_nodes)

    return sections

# -----------------------------------------------------------------------------
# UI
# -----------------------------------------------------------------------------

def main():
    st.markdown(CSS, unsafe_allow_html=True)
    st.title("Import fiche descriptive (.docx)")

    file = st.file_uploader("Choisir un fichier Word (.docx)", type=["docx"])
    if not file:
        st.info("Charge un .docx pour commencer.")
        return

    with st.spinner("Conversion Word → HTML…"):
        full_html = docx_to_html(file.read())

    raw_sections = split_into_sections(full_html)

    # Mapping
    detected = list(raw_sections.keys())
    rows = []
    for expected in EXPECTED_HEADINGS:
        present, matched_title = "Non", ""
        for d in detected:
            if _normalize_text(d) == _normalize_text(expected):
                present, matched_title = "Oui", d
                break
        rows.append((expected, present, matched_title, CRM_KEY_BY_LABEL.get(expected, "")))

    st.subheader("Résultat du mapping automatique")
    st.table({
        "Word heading attendu": [r[0] for r in rows],
        "Dans le .docx ?": [r[1] for r in rows],
        "PDF/CRM heading": [r[2] for r in rows],
        "CRM key": [r[3] for r in rows]
    })

    st.subheader("Aperçu des sections (mise en forme préservée)")

    for expected in EXPECTED_HEADINGS:
        the_key = None
        for k in raw_sections.keys():
            if _normalize_text(k) == _normalize_text(expected):
                the_key = k
                break
        raw = raw_sections.get(the_key or "", "")

        st.markdown(f"### {expected} <span class='anchor'>↪</span>", unsafe_allow_html=True)

        if not raw.strip():
            st.caption("(vide)")
            st.markdown("<hr/>", unsafe_allow_html=True)
            continue

        clean_html, downloads = prepare_section_html(raw)
        clean_html = postprocess_domain_section(expected, clean_html)

        st.markdown(f"<div class='card'>{clean_html}</div>", unsafe_allow_html=True)

        if downloads:
            st.info("Certaines images (EMF/WMF) ne peuvent pas s’afficher dans le navigateur. Télécharge-les :")
            cols = st.columns(max(1, min(3, len(downloads))))
            for i, (fname, blob) in enumerate(downloads):
                with cols[i % len(cols)]:
                    st.download_button(
                        label=f"Télécharger {fname}",
                        data=blob,
                        file_name=fname,
                        mime="application/octet-stream",
                    )

        st.markdown("<hr/>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
