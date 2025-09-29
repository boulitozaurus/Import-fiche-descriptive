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
from bs4 import BeautifulSoup
import uuid
# ---------------- Utils: headings + parsing ----------------
from docx import Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
import base64, re
from docx.text.run import Run
from docx.oxml.ns import qn

HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}
NS_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
NS_A = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

def _data_uri(image):
    data = base64.b64encode(image.read()).decode("ascii")
    return {"src": f"data:{image.content_type};base64,{data}"}

def split_sections_by_headings(html: str, expected_headings: list[str]) -> dict[str, str]:
    # normalisation légère
    norm = lambda s: re.sub(r"\s+", " ", s).strip().lower().rstrip(" :")
    wanted = {norm(h): h for h in expected_headings}

    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    out = {h: "" for h in expected_headings}

    current = None
    for el in soup.div.children:
        tag = getattr(el, "name", None)
        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            key = wanted.get(norm(el.get_text()))
            if key:
                current = key
            continue
        if current:
            out[current] += str(el)

    return out

def _strip_accents(x: str) -> str:
    if x is None: return ""
    try:
        import unicodedata
        nfkd = unicodedata.normalize("NFKD", x)
        return "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    except Exception:
        return x

def _norm(s: str) -> str:
    return " ".join(_strip_accents((s or "")).lower().replace("’","'").split())

def _is_heading_style(p: Paragraph) -> bool:
    s = p.style.name if getattr(p, "style", None) else ""
    return (s in HEADING_STYLES) or s.startswith("Heading")

def _looks_like_heading(text: str, p: Paragraph, expected_map: dict) -> bool:
    t = (text or "").strip()
    if not t: return False
    if _norm(t) in expected_map or _norm(t.rstrip(":")) in expected_map:
        return True
    if _is_heading_style(p):
        if len(t) <= 80 and t.count(" ") <= 11 and all(x not in t for x in [".","!","?"]):
            return True
    return False

def _html_escape(s: str) -> str:
    return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def _wrap_styles(run, txt: str) -> str:
    open_tags, close_tags = "", ""
    color = getattr(getattr(run.font, "color", None), "rgb", None)
    if color:
        open_tags += f'<span style="color:#{str(color)}">'; close_tags = "</span>" + close_tags
    if getattr(run, "underline", False):
        open_tags += "<u>"; close_tags = "</u>" + close_tags
    if getattr(run, "italic", False):
        open_tags += "<em>"; close_tags = "</em>" + close_tags
    if getattr(run, "bold", False):
        open_tags += "<strong>"; close_tags = "</strong>" + close_tags
    return f"{open_tags}{txt}{close_tags}"

def _run_image_payload(run):
    """Retourne un dict:
       {"kind":"img","src":data_uri}  ou
       {"kind":"download","mime":mime,"href":data_uri_download}  pour EMF/WMF."""
    try:
        def find_rid(el):
            for child in el.iterchildren():
                tag = child.tag
                if tag.endswith("}blip"):
                    rid = child.get(qn("r:embed"))
                    if rid: return rid
                if tag.endswith("}imagedata"):
                    rid = child.get(qn("r:id"))
                    if rid: return rid
                rid = find_rid(child)
                if rid: return rid
            return None

        rId = find_rid(run._r)
        if not rId:
            return None
        part = run.part.related_parts.get(rId)
        if not part:
            return None
        mime = (getattr(part, "content_type", "application/octet-stream") or "").lower()
        b64  = base64.b64encode(part.blob).decode("ascii")
        # Formats d'image supportés nativement par le navigateur
        ok = {"image/png","image/jpeg","image/jpg","image/gif","image/webp","image/svg+xml","image/bmp","image/tiff"}
        if mime in ok:
            return {"kind":"img","src": f"data:{mime};base64,{b64}"}
        # EMF/WMF -> non supporté : proposer un téléchargement propre
        if "emf" in mime or "wmf" in mime:
            return {"kind":"download","mime": mime, "href": f"data:application/octet-stream;base64,{b64}"}
        # sinon on tente quand même
        return {"kind":"img","src": f"data:{mime};base64,{b64}"}
    except Exception:
        return None

def _run_to_html(run) -> str:
    payload = _run_image_payload(run)
    if payload:
        if payload["kind"] == "img":
            return f'<img src="{payload["src"]}" />'
        else:
            ext = payload["mime"].split("/")[-1]
            return (f'<span class="img-unsupported">[image {ext.upper()} non supportée] '
                    f'<a href="{payload["href"]}" download="image.{ext}">Télécharger</a></span>')
    # sinon texte & sauts de ligne
    frags = []
    for child in run._r.iterchildren():
        if child.tag.endswith("}t"):
            txt = _html_escape(child.text or "")
            if txt: frags.append(_wrap_styles(run, txt))
        elif child.tag.endswith("}br"):
            frags.append("<br/>")
    if not frags:
        txt = _html_escape(run.text or "")
        if txt: frags.append(_wrap_styles(run, txt))
    return "".join(frags)

def _hyperlink_map(p: Paragraph) -> dict:
    """Associe chaque run XML à son URL (<w:hyperlink> et <w:fldSimple instr='HYPERLINK ...'>)."""
    m = {}
    try:
        for el in p._p.iterchildren():
            tag = el.tag
            # Cas 1 : <w:hyperlink r:id="...">...</w:hyperlink>
            if tag.endswith("}hyperlink"):
                r_id = el.get(qn("r:id"))
                url = None
                if r_id:
                    rel = p.part.rels.get(r_id)
                    if rel is not None:
                        url = getattr(rel, "target_ref", None) or getattr(rel, "target_part", None)
                        if hasattr(url, "partname"):
                            url = str(url.partname)
                for r in el.iterchildren():
                    if r.tag.endswith("}r"):
                        m[r] = url
            # Cas 2 : <w:fldSimple w:instr="HYPERLINK \"https://...\"">...</w:fldSimple>
            elif tag.endswith("}fldSimple"):
                instr = el.get(qn("w:instr")) or ""
                m_url = re.search(r'HYPERLINK\s+"([^"]+)"', instr, flags=re.I) or re.search(r'HYPERLINK\s+(\S+)', instr, flags=re.I)
                if m_url:
                    url = m_url.group(1)
                    for r in el.iterchildren():
                        if r.tag.endswith("}r"):
                            m[r] = url
    except Exception:
        pass
    return m

def _autolink_html(s: str) -> str:
    # transforme http(s)://... en lien si aucune balise <a ...> n'est déjà présente
    if "<a " in s: 
        return s
    return re.sub(r'(?<!["\'>])(https?://[^\s<]+)', 
                  r'<a href="\1" target="_blank" rel="noopener noreferrer">\1</a>', 
                  s)

def _image_handler(image) -> dict:
    ctype = (getattr(image, "content_type", None) or "application/octet-stream").lower()
    try:
        with image.open() as f:
            data = f.read()
    except Exception:
        data = b""

    if ctype in ("image/x-emf", "image/emf", "image/x-wmf", "image/wmf", "application/octet-stream"):
        import uuid, base64
        uid = uuid.uuid4().hex
        fname = f"{uid}.{'emf' if 'emf' in ctype else ('wmf' if 'wmf' in ctype else 'bin')}"
        if "img_store" not in st.session_state:
            st.session_state["img_store"] = {}
        st.session_state["img_store"][uid] = (fname, data, ctype)

        # pixel transparent + marqueurs
        return {
            "src": "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==",
            "alt": "Image non affichable (EMF/WMF).",
            "data-unsupported": "1",
            "data-uid": uid,
        }

    import base64
    b64 = base64.b64encode(data).decode("ascii")
    return {"src": f"data:{ctype};base64,{b64}"}

def docx_to_html(path: str) -> str:
    """Convertit le .docx en HTML avec Mammoth en utilisant notre handler d’images."""
    with open(path, "rb") as f:
        result = mammoth.convert_to_html(
            f,
            convert_image=mammoth.images.inline(_image_handler)  # <-- garde bien inline()
        )
    return result.value

def prepare_section_html(html: str):
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    downloads = []

    for img in list(soup.find_all("img")):
        if img.get("data-unsupported") == "1":
            uid = img.get("data-uid")
            if uid and "img_store" in st.session_state and uid in st.session_state["img_store"]:
                fname, data, ctype = st.session_state["img_store"][uid]
                downloads.append((uid, fname, data, ctype))
            img.decompose()  # enlève le pixel transparent

    # On répare au passage les 'src="data:..."' orphelins
    cleaned = _fix_stray_data_uri(soup.div.decode_contents())
    return cleaned, downloads

# ---------------- Load schema + fixed heading map ----------------

SCHEMA_PATH = Path("crm_schema.yaml")
MAP_PATH    = Path("heading_map.yaml")

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
      { "key": "description_fr",            "nl_key": "description_nl",            "label": "Description" },
      { "key": "contexte_fonds_fr",         "nl_key": "contexte_fonds_nl",         "label": "Contexte et usage des fonds" },
      { "key": "points_attention_fr",       "nl_key": "points_attention_nl",       "label": "Les points d'attention" },
      { "key": "bonnes_raisons_fr",         "nl_key": "bonnes_raisons_nl",         "label": "Les bonnes raisons d'investir" },
      { "key": "operation_presentation_fr", "nl_key": "operation_presentation_nl", "label": "Présentation de l'opération" },
      { "key": "localisation_fr",           "nl_key": "localisation_nl",           "label": "Localisation" },
      { "key": "planning_fr",               "nl_key": "planning_nl",               "label": "Planning" },
      { "key": "marche_references_fr",      "nl_key": "marche_references_nl",      "label": "Marché et références" },
      { "key": "budget_fr",                 "nl_key": "budget_nl",                 "label": "Budget" },
      { "key": "operateur_presentation_fr", "nl_key": "operateur_presentation_nl", "label": "Présentation de l'opérateur" },
      { "key": "track_record_fr",           "nl_key": "track_record_nl",           "label": "Track record" },
      { "key": "structure_management_fr",   "nl_key": "structure_management_nl",   "label": "Structure et Management" },
      { "key": "actionnariat_fr",           "nl_key": "actionnariat_nl",           "label": "Actionnariat" },
      { "key": "finances_fr",               "nl_key": "finances_nl",               "label": "Finances" },
    ]
}

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
      .sect p { margin:.35rem 0; }
      .sect ol, .sect ul { margin:.35rem 0 .55rem 1.4rem; padding-left:1.4rem; list-style-position:outside; }
      .sect li { margin:.15rem 0; }
      .sect table { width:100%; border-collapse:collapse; }
      .sect td, .sect th { border:1px solid #ccc; padding:6px; }
    </style>
    """, unsafe_allow_html=True)

def _fix_stray_data_uri(html: str) -> str:
    # transforme une ligne de type:  src="data:image/png;base64,...."
    # en véritable balise <img src="data:..."/>
    return re.sub(r'(?<!<img )src="data:image/[^"]+"', 
                  lambda m: f'<img {m.group(0)} />', html)

# ---------------- UI ----------------

st.set_page_config(page_title="Auto-Mapping Word", layout="wide")
st.title("Auto-Mapping Word")
st.caption("Déposez votre fiche .docx : mapping fixe Word→PDF/CRM")
if "unsupported_images" not in st.session_state:
    st.session_state["unsupported_images"] = []

# Load schema + map
def norm(s: str) -> str:
    try:
        import unicodedata
        nfkd = unicodedata.normalize("NFKD", s or "")
        s = "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    except Exception:
        s = s or ""
    return " ".join(s.lower().replace("’", "'").split())
        
schema = load_schema()
fields = schema.get("fields", [])
key_by_pdf_label_norm = {_norm(f["label"]): f["key"] for f in fields}
nl_key_by_key = {f["key"]: f.get("nl_key") for f in fields}
word_to_pdf = load_heading_map()
expected_word_headings = list(word_to_pdf.keys())

# Build: Word heading -> (libellé PDF/CRM, clé CRM)
crm_map = {}
missing = []
for wh, pdf_label in word_to_pdf.items():
    k = key_by_pdf_label_norm.get(_norm(pdf_label))
    if k:
        crm_map[wh] = (pdf_label, k)
    else:
        missing.append(pdf_label)

# Optionnel: remonter les libellés non résolus
if missing:
    st.warning("Champs non trouvés dans le schema: " + ", ".join(missing))
        
# Upload
st.header("1) Charger la fiche .docx")
uploaded = st.file_uploader("Glissez le .docx ici", type=["docx"])
if uploaded is not None:
    tmp_path = Path("uploaded.docx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())

# 1) Conversion DOCX -> HTML (Mammoth) puis découpe par titres
    html = docx_to_html(str(tmp_path))
    sections = split_sections_by_headings(html, expected_word_headings)
    sections_norm = {_norm(k): v for k, v in sections.items()}


    # 2) Auto-mapping Word -> PDF/CRM (valeurs = HTML)
    fr_payload = {}
    rows = []
    
    for w_heading in expected_word_headings:
        crm_label, crm_key = crm_map[w_heading]
        content_html = sections.get(w_heading, "")  # <-- prend la section brute
        fr_payload[crm_key] = content_html          # <-- pas via sections_norm
    
        rows.append({
            "Word heading attendu": w_heading,
            "Dans le .docx ?": "✅ Oui" if content_html.strip() else "❌ Non",
            "PDF/CRM heading": crm_label,
            "CRM key": crm_key,
        })

    st.subheader("Résultat du mapping automatique")
    st.dataframe(rows, use_container_width=True)

    # 3) Affichage vertical fidèle (HTML)

    st.header("Aperçu des sections (mise en forme préservée)")
    inject_css()
    
    for fdef in fields:
        key   = fdef["key"]
        label = fdef["label"]
    
        raw_html = fr_payload.get(key, "")
        clean_html, dls = prepare_section_html(raw_html)
    
        st.subheader(label)
        st.markdown(f"<div class='sect'>{clean_html or '<p><em>(vide)</em></p>'}</div>", unsafe_allow_html=True)
    
        # Boutons de téléchargement, localisés à la section
        for uid, fname, data, ctype in dls:
            st.download_button(f"Télécharger {fname}", data=data, file_name=fname, mime=ctype, key=f"dl_{uid}")
    
        st.divider()

