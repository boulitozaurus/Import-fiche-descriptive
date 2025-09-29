import os
import io
import csv
import json
import yaml
import requests
import streamlit as st
from pathlib import Path
from typing import Dict, List
from streamlit.components.v1 import html as st_html
import mammoth
from bs4 import BeautifulSoup, Tag, NavigableString
import uuid
import re
from docx import Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
import base64
from docx.text.run import Run
from docx.oxml.ns import qn

# ================= CONFIGURATION =================
HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}
NS_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
NS_A = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

HEADING_WORDS = (
    "introduction|contexte|les points d'attention|les bonnes raisons|"
    "projet|localisation|planning|budget|opérateur|opérateur|actionnariat|finances?"
)

_HEADING_RE = re.compile(rf"^\s*(?:{HEADING_WORDS})\b", re.I)
NUM_PREFIX_RE = re.compile(r"^\s*(\d+)[\.\)]\s+")
NBSP = "\u00A0"
_BULLETS = "•◦◘▪▫·—–-o" + NBSP
_BULLET_CLASS = "".join(re.escape(ch) for ch in _BULLETS)
BULLET_ONLY_RE = re.compile(r'^[\s' + _BULLET_CLASS + r']+$', re.I)

SCHEMA_PATH = Path("crm_schema.yaml")
MAP_PATH = Path("heading_map.yaml")

DEFAULT_HEADING_MAP = {
    "Introduction": "Description",
    "Contexte et usage des fonds": "Contexte et usage des fonds",
    "Facteurs de risque": "Les points d'attention",
    "Les bonnes raisons d'investir": "Les bonnes raisons d'investir",
    "Projet": "Présentation de l'opération",
    "Localisation": "Localisation",
    "Administratif et timing": "Planning",
    "Marché et références": "Marché et références",
    "Budget de l'opération": "Budget",
    "L'opérateur": "Présentation de l'opérateur",
    "Track record et opérations en cours": "Track record",
    "Structure et Management": "Structure et Management",
    "Actionnariat et structure de l'opération": "Actionnariat",
    "Finances": "Finances",
}

DEFAULT_SCHEMA = {
    "fields": [
      {"key": "description_fr", "nl_key": "description_nl", "label": "Description"},
      {"key": "contexte_fonds_fr", "nl_key": "contexte_fonds_nl", "label": "Contexte et usage des fonds"},
      {"key": "points_attention_fr", "nl_key": "points_attention_nl", "label": "Les points d'attention"},
      {"key": "bonnes_raisons_fr", "nl_key": "bonnes_raisons_nl", "label": "Les bonnes raisons d'investir"},
      {"key": "operation_presentation_fr", "nl_key": "operation_presentation_nl", "label": "Présentation de l'opération"},
      {"key": "localisation_fr", "nl_key": "localisation_nl", "label": "Localisation"},
      {"key": "planning_fr", "nl_key": "planning_nl", "label": "Planning"},
      {"key": "marche_references_fr", "nl_key": "marche_references_nl", "label": "Marché et références"},
      {"key": "budget_fr", "nl_key": "budget_nl", "label": "Budget"},
      {"key": "operateur_presentation_fr", "nl_key": "operateur_presentation_nl", "label": "Présentation de l'opérateur"},
      {"key": "track_record_fr", "nl_key": "track_record_nl", "label": "Track record"},
      {"key": "structure_management_fr", "nl_key": "structure_management_nl", "label": "Structure et Management"},
      {"key": "actionnariat_fr", "nl_key": "actionnariat_nl", "label": "Actionnariat"},
      {"key": "finances_fr", "nl_key": "finances_nl", "label": "Finances"},
    ]
}

# ================= FONCTIONS UTILITAIRES =================

def _strip_accents(x: str) -> str:
    if x is None: return ""
    try:
        import unicodedata
        nfkd = unicodedata.normalize("NFKD", x)
        return "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    except Exception:
        return x

def _norm(s: str) -> str:
    return " ".join(_strip_accents((s or "")).lower().replace("'","'").split())

def _strip_leading_numbering(s: str) -> str:
    return re.sub(r'^\s*(?:[\(\[]?\d+(?:\.\d+)*[\)\.]?|[ivxlcdm]+[\)\.]|[A-Z]\)|•|—|-)\s*', '', s or '', flags=re.I)

def _is_bullet_only_text(text: str) -> bool:
    t = (text or "").replace(NBSP, " ").strip()
    return bool(BULLET_ONLY_RE.match(t))

def _is_section_heading_p(p: Tag) -> bool:
    if p.name != "p":
        return False
    txt = p.get_text(" ", strip=True)
    if not txt:
        return False
    if len(txt) <= 90 and (p.find(["strong", "b"]) or _HEADING_RE.match(txt)):
        return True
    return False

def _html_escape(s: str) -> str:
    return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

# ================= GESTION DES IMAGES =================

def _image_handler(image) -> dict:
    ctype = (getattr(image, "content_type", None) or "application/octet-stream").lower()
    try:
        with image.open() as f:
            data = f.read()
    except Exception:
        data = b""

    if ctype in ("image/x-emf", "image/emf", "image/x-wmf", "image/wmf", "application/octet-stream"):
        uid = uuid.uuid4().hex
        fname = f"{uid}.{'emf' if 'emf' in ctype else ('wmf' if 'wmf' in ctype else 'bin')}"
        st.session_state.setdefault("img_store", {})[uid] = (fname, data, ctype)
        return {
            "src": "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==",
            "alt": "",
            "data-unsupported": "1",
            "data-uid": uid,
        }

    b64 = base64.b64encode(data).decode("ascii")
    return {"src": f"data:{ctype};base64,{b64}", "alt": ""}

# ================= CONVERSION DOCX -> HTML =================

def docx_to_html(path: str) -> str:
    style_map = """
p[style-name='Heading 1'] => h1:fresh
p[style-name='Heading 2'] => h2:fresh
p[style-name='Heading 3'] => h3:fresh
p[style-name='Titre 1']   => h1:fresh
p[style-name='Titre 2']   => h2:fresh
p[style-name='Titre 3']   => h3:fresh
"""
    with open(path, "rb") as f:
        result = mammoth.convert_to_html(
            f,
            convert_image=mammoth.images.inline(_image_handler),
            style_map=style_map
        )
    return result.value

# ================= DÉCOUPAGE PAR SECTIONS =================

def build_heading_index(expected_headings: list[str], word_to_pdf: dict[str, str]) -> dict[str, str]:
    idx: dict[str, str] = {}
    for wh in expected_headings:
        for alias in {wh, word_to_pdf.get(wh, ''), _strip_leading_numbering(wh)}:
            if alias:
                idx[_norm(alias)] = wh
    
    idx[_norm("description")] = "Introduction"
    idx[_norm("contexte & usage des fonds")] = "Contexte et usage des fonds"
    return idx

def split_sections_by_headings(html: str, heading_index: dict[str, str]) -> dict[str, str]:
    def nrm(s: str) -> str:
        return _norm(_strip_leading_numbering((s or "").rstrip(" :")))

    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    out = {v: "" for v in set(heading_index.values())}
    current = None

    for el in soup.div.children:
        if not hasattr(el, "get_text"):
            continue

        key = None
        if getattr(el, "name", None) in {"h1","h2","h3","h4","h5","h6","p"}:
            key = heading_index.get(nrm(el.get_text()))

        if key:
            current = key
            continue

        if current is not None:
            out[current] += str(el)

    return out

# ================= NETTOYAGE DES LISTES =================

def _convert_numbered_paragraphs_to_ol(parent: Tag) -> bool:
    changed = False
    children = list(parent.children)
    i = 0
    while i < len(children):
        node = children[i]
        if getattr(node, "name", None) == "p":
            text = node.get_text(" ", strip=True)
            m = NUM_PREFIX_RE.match(text or "")
            if m:
                ol = parent.new_tag("ol")
                parent.insert_before(ol, node)
                last_li = None

                j = i
                while j < len(children):
                    cur = children[j]
                    name = getattr(cur, "name", None)

                    if name == "p":
                        t = cur.get_text(" ", strip=True)
                        m2 = NUM_PREFIX_RE.match(t or "")
                        if m2:
                            li = parent.new_tag("li")
                            first_txt = None
                            for c in cur.contents:
                                if isinstance(c, NavigableString):
                                    first_txt = c
                                    break
                            if first_txt:
                                m3 = NUM_PREFIX_RE.match(str(first_txt))
                                if m3:
                                    first_txt.replace_with(first_txt[m3.end():])
                            for c in list(cur.contents):
                                li.append(c.extract())
                            cur.decompose()
                            ol.append(li)
                            last_li = li
                            j += 1
                            changed = True
                            continue

                        if last_li and t and not _is_bullet_only_text(t) and not _is_section_heading_p(cur):
                            last_li.append(cur.extract())
                            children.pop(j)
                            changed = True
                            continue

                        break

                    elif name in ("ul", "ol") and last_li:
                        last_li.append(cur.extract())
                        children.pop(j)
                        changed = True
                        continue
                    else:
                        break

                children = list(parent.children)
                i = list(parent.children).index(ol) + 1
                continue
        i += 1
    return changed

def _merge_split_ol_blocks(soup: BeautifulSoup) -> bool:
    changed = False
    for ol in list(soup.find_all("ol")):
        sib = ol.find_next_sibling()
        if sib is None:
            continue

        trail = []
        cur = sib
        while cur is not None and getattr(cur, "name", None) == "p" and not _is_section_heading_p(cur):
            trail.append(cur)
            cur = cur.find_next_sibling()

        if cur is not None and getattr(cur, "name", None) == "ol":
            lis = ol.find_all("li", recursive=False)
            last_li = lis[-1] if lis else None

            if trail:
                if last_li is None:
                    last_li = soup.new_tag("li")
                    ol.append(last_li)
                for p in trail:
                    last_li.append(p.extract())

            for li in list(cur.find_all("li", recursive=False)):
                ol.append(li.extract())

            cur.decompose()
            changed = True

    return changed

def _promote_nested_ol_to_siblings(soup: BeautifulSoup) -> bool:
    changed = False
    for li in list(soup.find_all("li")):
        parent = li.parent
        if getattr(parent, "name", None) != "ol":
            continue

        inner_ols = [c for c in li.contents if getattr(c, "name", None) == "ol"]
        if len(inner_ols) != 1:
            continue

        inner = inner_ols[0]
        sub_items = inner.find_all("li", recursive=False)
        if not (2 <= len(sub_items) <= 6):
            continue

        direct_text = "".join(t for t in li.find_all(string=True, recursive=False)).strip()
        if len(direct_text) < 40:
            continue

        long_items = sum(1 for s in sub_items if len(s.get_text(" ", strip=True)) >= 30)
        if long_items < len(sub_items) // 2:
            continue

        anchor = li
        for sub in list(sub_items):
            anchor.insert_after(sub.extract())
            anchor = sub

        inner.decompose()
        changed = True

    return changed

def _fix_lists_in_soup(soup):
    changed = True
    while changed:
        changed = False

        for p in list(soup.find_all("p")):
            if _is_bullet_only_text(p.get_text(" ", strip=True)):
                p.decompose()
                changed = True

        for li in list(soup.find_all("li")):
            txt = li.get_text(" ", strip=True)
            if not txt or _is_bullet_only_text(txt):
                li.decompose()
                changed = True

        for li in list(soup.find_all("li")):
            for p in list(li.find_all("p", recursive=False)):
                if not p.get_text(strip=True):
                    p.decompose()
                    changed = True

        for li in list(soup.find_all("li")):
            direct_text = "".join(t for t in li.find_all(string=True, recursive=False)).strip()
            child_lists = [c for c in li.contents if getattr(c, "name", None) in ("ol", "ul")]
            if direct_text == "" and len(child_lists) == 1:
                inner = child_lists[0]
                parent = li.parent
                if getattr(parent, "name", None) not in ("ul", "ol"):
                    continue

                if parent.name == "ul" and inner.name == "ol":
                    siblings = parent.find_all("li", recursive=False)
                    if len(siblings) == 1 and siblings[0] is li:
                        parent.replace_with(inner)
                    else:
                        new_ol = soup.new_tag("ol")
                        for sub_li in inner.find_all("li", recursive=False):
                            new_ol.append(sub_li)
                        parent.insert_before(new_ol)
                        li.decompose()
                    changed = True
                else:
                    if parent.name == inner.name:
                        for sub_li in inner.find_all("li", recursive=False):
                            li.insert_before(sub_li)
                        li.decompose()
                        changed = True
                    else:
                        li.replace_with(inner)
                        changed = True

        for ul in list(soup.find_all("ul")):
            lis = ul.find_all("li", recursive=False)
            if len(lis) == 0:
                only_lists = [c for c in ul.contents if getattr(c, "name", None) in ("ol", "ul")]
                if len(only_lists) == 1:
                    ul.replace_with(only_lists[0])
                    changed = True

        if _merge_split_ol_blocks(soup):
            changed = True

        if _promote_nested_ol_to_siblings(soup):
            changed = True

        for ol in list(soup.find_all("ol")):
            nxt = ol.find_next_sibling()
            if getattr(nxt, "name", None) == "ol":
                for li in list(nxt.find_all("li", recursive=False)):
                    ol.append(li)
                nxt.decompose()
                changed = True

    return soup

# ================= PRÉPARATION FINALE DES SECTIONS =================

def prepare_section_html(html: str):
    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    downloads = []

    IA_PAT = re.compile(r"le contenu\s+g[éè]n[éè]r[éè]\s+par l[''' ]?ia\s+peut\s+être\s+incorrect\.?", re.I)
    for p in list(soup.find_all("p")):
        if IA_PAT.search(p.get_text(" ", strip=True)):
            p.decompose()

    for img in list(soup.find_all("img")):
        if img.get("data-unsupported") != "1" and img.has_attr("alt"):
            del img["alt"]

    for p in list(soup.find_all("p")):
        if len(p.contents) == 1 and getattr(p.contents[0], "name", None) in ("strong", "b"):
            p.contents[0].unwrap()

    _convert_numbered_paragraphs_to_ol(soup)
    for cont in soup.find_all(["div", "section"]):
        _convert_numbered_paragraphs_to_ol(cont)

    _fix_lists_in_soup(soup)

    cleaned = soup.div.decode_contents()
    cleaned = re.sub(r'(?is)le contenu\s+g[éè]n[éè]r[éè]\s+par l[''' ]?ia\s+peut\s+être\s+incorrect\.?', '', cleaned)

    return cleaned, downloads

# ================= CORRECTION DE LA NUMÉROTATION =================

def fix_section_numbering(html: str, section_key: str) -> str:
    """Force la numérotation correcte pour les sections spécifiques."""
    if not html or not html.strip():
        return html
    
    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    
    EXPECTED_TITLES = {
        'points_attention_fr': [
            "Risque lié au projet",
            "Risque lié au secteur", 
            "Risque de défaut"
        ],
        'bonnes_raisons_fr': [
            "Une assurance sur 100% du capital investi",
            "Une fiducie-sûreté sur l'actif"
        ],
        'budget_fr': [
            "Prix de revient",
            "Financement et ratios",
            "Revenus et marges",
            "Couverture des intérêts",
            "Stress test"
        ]
    }
    
    if section_key not in EXPECTED_TITLES:
        return html
    
    expected = EXPECTED_TITLES[section_key]
    
    def normalize(text):
        text = text.lower().strip()
        text = re.sub(r'^\s*\d+[\.\)]\s*', '', text)
        text = re.sub(r'^[\•\-\*]\s*', '', text)
        text = _strip_accents(text)
        return ' '.join(text.split())
    
    expected_normalized = {normalize(t): t for t in expected}
    
    potential_titles = []
    for elem in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'strong', 'b']):
        text = elem.get_text(strip=True)
        if not text:
            continue
            
        norm_text = normalize(text)
        if norm_text in expected_normalized:
            potential_titles.append((elem, norm_text))
    
    if potential_titles:
        if section_key == 'bonnes_raisons_fr':
            found_titles = [norm for _, norm in potential_titles]
            has_assurance = normalize(expected[0]) in found_titles
            
            if not has_assurance:
                expected = expected[1:]
                expected_normalized = {normalize(t): t for t in expected}
        
        for elem, norm_text in potential_titles:
            if norm_text in expected_normalized:
                original_title = expected_normalized[norm_text]
                try:
                    idx = expected.index(original_title)
                    num = idx + 1
                except ValueError:
                    continue
                
                new_title = f"{num}. {original_title}"
                
                if elem.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    elem.string = new_title
                elif elem.name in ['strong', 'b']:
                    elem.string = new_title
                elif elem.name == 'p':
                    strong_elem = elem.find(['strong', 'b'])
                    if strong_elem and len(list(elem.children)) == 1:
                        strong_elem.string = new_title
                    else:
                        elem.string = new_title
    
    result = soup.div.decode_contents()
    return result

def apply_fixed_numbering(fr_payload: dict) -> dict:
    """Applique la numérotation fixe aux sections concernées."""
    sections_to_fix = ['points_attention_fr', 'bonnes_raisons_fr', 'budget_fr']
    result = fr_payload.copy()
    
    for key in sections_to_fix:
        if key in result:
            result[key] = fix_section_numbering(result[key], key)
    
    return result

# ================= CHARGEMENT DE LA CONFIGURATION =================

def load_schema() -> Dict:
    if SCHEMA_PATH.exists():
        return yaml.safe_load(SCHEMA_PATH.read_text(encoding="utf-8")) or DEFAULT_SCHEMA
    return DEFAULT_SCHEMA

def load_heading_map() -> Dict[str, str]:
    if MAP_PATH.exists():
        cfg = yaml.safe_load(MAP_PATH.read_text(encoding="utf-8")) or {}
        m = cfg.get("word_to_pdf")
        if isinstance(m, dict) and m:
            return m
    return DEFAULT_HEADING_MAP

def inject_css():
    st.markdown("""
    <style>
      .sect h1, .sect h2, .sect h3 {
        font-size: 1.15rem;
        line-height: 1.5;
        font-weight: 600;
        margin: .35rem 0 .35rem;
      }
      .sect h4, .sect h5, .sect h6 {
        font-size: 1.05rem;
        line-height: 1.45;
        font-weight: 600;
        margin: .30rem 0 .30rem;
      }
      .sect p { margin: .30rem 0; }
      .sect ol, .sect ul { margin: .40rem 0 .60rem 1.4rem; padding-left: 1.2rem; list-style-position: outside; }
      .sect ol { list-style-type: decimal; }
      .sect ul { list-style-type: disc; }
    </style>
    """, unsafe_allow_html=True)

# ================= INTERFACE STREAMLIT =================

st.set_page_config(page_title="Auto-Mapping Word", layout="wide")
st.title("Auto-Mapping Word")
st.caption("Déposez votre fiche .docx : mapping fixe Word→PDF/CRM avec numérotation corrigée")

if "unsupported_images" not in st.session_state:
    st.session_state["unsupported_images"] = []

schema = load_schema()
fields = schema.get("fields", [])
key_by_pdf_label_norm = {_norm(f["label"]): f["key"] for f in fields}
nl_key_by_key = {f["key"]: f.get("nl_key") for f in fields}
word_to_pdf = load_heading_map()
expected_word_headings = list(word_to_pdf.keys())

crm_map = {}
missing = []
for wh, pdf_label in word_to_pdf.items():
    k = key_by_pdf_label_norm.get(_norm(pdf_label))
    if k:
        crm_map[wh] = (pdf_label, k)
    else:
        missing.append(pdf_label)

if missing:
    st.warning("Champs non trouvés dans le schema: " + ", ".join(missing))

st.header("1) Charger la fiche .docx")
uploaded = st.file_uploader("Glissez le .docx ici", type=["docx"])

if uploaded is not None:
    tmp_path = Path("uploaded.docx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())

    html = docx_to_html(str(tmp_path))
    heading_index = build_heading_index(expected_word_headings, word_to_pdf)
    sections = split_sections_by_headings(html, heading_index)

    fr_payload = {}
    rows = []
    
    for w_heading in expected_word_headings:
        crm_label, crm_key = crm_map[w_heading]
        content_html = sections.get(w_heading, "")
        fr_payload[crm_key] = content_html
    
        rows.append({
            "Word heading attendu": w_heading,
            "Dans le .docx ?": "✅ Oui" if content_html.strip() else "❌ Non",
            "PDF/CRM heading": crm_label,
            "CRM key": crm_key,
        })

    # CORRECTION DE LA NUMÉROTATION
    fr_payload = apply_fixed_numbering(fr_payload)

    st.subheader("Résultat du mapping automatique")
    st.dataframe(rows, use_container_width=True)

    st.header("Aperçu des sections (mise en forme préservée)")
    inject_css()
    
    for fdef in fields:
        key = fdef["key"]
        label = fdef["label"]
    
        raw_html = fr_payload.get(key, "")
        clean_html, dls = prepare_section_html(raw_html)
    
        st.subheader(label)
        st.markdown(f"<div class='sect'>{clean_html or '<p><em>(vide)</em></p>'}</div>", unsafe_allow_html=True)
    
        for uid, fname, data, ctype in dls:
            st.download_button(f"Télécharger {fname}", data=data, file_name=fname, mime=ctype, key=f"dl_{uid}")
    
        st.divider()
