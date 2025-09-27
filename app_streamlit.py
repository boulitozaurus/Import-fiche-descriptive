import os
import io
import csv
import json
import yaml
import requests
import streamlit as st
from pathlib import Path
from typing import Dict, List
from docx import Document
from streamlit.components.v1 import html as st_html


# ---------------- Utils: headings + parsing ----------------

HEADING_STYLES = {
    "Heading 1","Heading 2","Heading 3",
    "Titre 1","Titre 2","Titre 3","Title","Subtitle"
}

def _is_heading_style(p) -> bool:
    s = p.style.name if getattr(p, "style", None) else ""
    return (s in HEADING_STYLES) or s.startswith("Heading")

def _looks_like_heading(text: str, paragraph, expected_map: dict) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    # 1) Titre exact attendu (selon ton mapping) → heading
    if norm(t) in expected_map:
        return True
    # 2) Sinon, style "Heading" court et sans ponctuation de phrase → heading
    if _is_heading_style(paragraph):
        if len(t) <= 80 and t.count(" ") <= 11 and all(p not in t for p in [".", "!", "?"]):
            return True
    return False

def parse_docx_sections(path: Path, expected_headings: list = None) -> Dict[str, str]:
    """Return {heading: text}. Detect headings either from expected list or from short 'Heading' styled paras.
       Tables are appended as markdown under the last heading."""
    from docx import Document
    doc = Document(path)

    # map normalisé -> étiquette canonique (ex: "introduction" -> "Introduction")
    expected_map = {norm(h): h for h in (expected_headings or [])}

    sections: Dict[str, str] = {}
    current = None
    buff: List[str] = []

    def flush():
        nonlocal current, buff
        if current and buff:
            text = "\n\n".join([x for x in buff if x.strip()])
            if text.strip():
                sections[current] = (sections.get(current, "") + ("\n\n" if sections.get(current) else "") + text).strip()
        buff = []

    # Parcours des paragraphes
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        if _looks_like_heading(t, p, expected_map):
            flush()
            # étiquette canonique si c'est un titre attendu, sinon le texte brut
            current = expected_map.get(norm(t), t)
        else:
            buff.append(t)
    flush()

    # Ajout des tableaux au dernier heading détecté
    last_heading = None
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t and _looks_like_heading(t, p, expected_map):
            last_heading = expected_map.get(norm(t), t)

    if doc.tables:
        md_chunks = []
        for table in doc.tables:
            rows = []
            for row in table.rows:
                rows.append([c.text.strip() for c in row.cells])
            md = "\n".join("| " + " | ".join(r) + " |" for r in rows if any(x for x in r))
            if md.strip():
                md_chunks.append(md)
        if md_chunks:
            target = last_heading or "TABLES"
            sections[target] = (sections.get(target, "") + ("\n\n" if sections.get(target) else "") + "\n\n".join(md_chunks)).strip()

    return sections

# ---------------- Robust normalization (no Unidecode needed) ----------------

def _strip_accents(x: str) -> str:
    if x is None:
        return ""
    try:
        import unicodedata
        nfkd = unicodedata.normalize("NFKD", x)
        return "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    except Exception:
        return x

def norm(s: str) -> str:
    # lower + remove accents + normalize apostrophes + collapse spaces
    return " ".join(_strip_accents((s or "")).lower().replace("’", "'").split())

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

# ---------------- UI ----------------

st.set_page_config(page_title="Auto-Mapping Word", layout="wide")
st.title("Auto-Mapping Word")
st.caption("Déposez votre fiche .docx : mapping fixe Word→PDF/CRM")

# Load schema + map
schema = load_schema()
fields = schema.get("fields", [])
key_by_pdf_label_norm = {norm(f["label"]): f["key"] for f in fields}
nl_key_by_key = {f["key"]: f.get("nl_key") for f in fields}
word_to_pdf = load_heading_map()
expected_word_headings = list(word_to_pdf.keys())

# Upload
st.header("1) Charger la fiche .docx")
uploaded = st.file_uploader("Glissez le .docx ici", type=["docx"])

if uploaded is not None:
    tmp_path = Path("uploaded.docx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())

    # Parse (HTML par section)
    sections = parse_docx_sections(tmp_path, expected_headings=expected_word_headings)
    sections_norm = {norm(k): v for k, v in sections.items()}
    
    # Auto-map (FR HTML)
    rows = []
    fr_payload = {}
    for word_h, pdf_h in word_to_pdf.items():
        w_norm = norm(word_h)
        target_key = key_by_pdf_label_norm.get(norm(pdf_h))
        found = w_norm in sections_norm or norm(word_h.rstrip(":")) in sections_norm
        fr_html = sections_norm.get(w_norm) or sections_norm.get(norm(word_h.rstrip(":")), "")
        if target_key:
            fr_payload[target_key] = fr_html
        rows.append({
            "Word heading attendu": word_h,
            "Dans le .docx ?": "✅ Oui" if found else "❌ Non",
            "PDF/CRM heading": pdf_h,
            "CRM key": target_key or "(non défini)",
        })
    
    st.subheader("Résultat du mapping automatique")
    st.dataframe(rows)  # affichage simple
    
    # ====== Affichage VERTICAL des sections (HTML) ======
    st.header("Aperçu des sections (mise en forme préservée)")
    CSS = """
    <style>
      .sect { padding: 8px 0; border-bottom: 1px solid rgba(128,128,128,.2); }
      .sect h3 { margin: 0 0 6px 0; }
      .sect ul, .sect ol { margin: 0.4rem 0 0.8rem 1.4rem; }
      .sect table { border-collapse: collapse; width: 100%; margin: .4rem 0; }
      .sect td, .sect th { border: 1px solid #666; padding: 6px; vertical-align: top; }
      .sect img { max-width: 100%; height: auto; }
      body, p, li, td { color: inherit; }
    </style>
    """
    for fdef in fields:
        key = fdef["key"]
        label = fdef["label"]
        html_content = fr_payload.get(key, "")
        block = f"<div class='sect'><h3>{label}</h3>{html_content or '<p><em>(vide)</em></p>'}</div>"
        st_html(CSS + block, height=min(2400, 150 + len(html_content)//3))

