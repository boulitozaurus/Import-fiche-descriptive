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


# ---------------- Utils: headings + parsing ----------------
# === Parser DOCX -> sections HTML fidèles (listes/tableaux/images/formatage) ===
from docx import Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
import base64, re
from docx.oxml.ns import qn

HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}
NS_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
NS_A = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

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

def _run_image_dataurl(run) -> str | None:
    try:
        blips = run._r.xpath(".//a:blip/@r:embed", namespaces={**NS_A})
        if not blips: return None
        part = run.part.related_parts[blips[0]]
        ctype = getattr(part, "content_type", "image/png")
        b64 = base64.b64encode(part.blob).decode("ascii")
        return f"data:{ctype};base64,{b64}"
    except Exception:
        return None

def _run_to_html(run) -> str:
    # images & sauts de ligne <w:br/>
    dataurl = _run_image_dataurl(run)
    if dataurl:
        return f'<img src="{dataurl}" />'
    frags = []
    for child in run._r.iterchildren():
        if child.tag.endswith("}t"):
            txt = _html_escape(child.text or "")
            if txt:
                frags.append(_wrap_styles(run, txt))
        elif child.tag.endswith("}br"):
            frags.append("<br/>")
    if not frags:
        txt = _html_escape(run.text or "")
        if txt: frags.append(_wrap_styles(run, txt))
    return "".join(frags)

def _hyperlink_map(p: Paragraph) -> dict:
    """Mappe chaque run XML à son URL si le run est à l'intérieur d'un <w:hyperlink>."""
    m = {}
    try:
        for hl in p._p.iterchildren():
            if hl.tag.endswith("}hyperlink"):
                r_id = hl.get(qn("r:id"))
                url = None
                if r_id:
                    rel = p.part.rels.get(r_id)
                    if rel is not None:
                        # python-docx: .target_ref pour les liens externes
                        url = getattr(rel, "target_ref", None) or getattr(rel, "target_part", None)
                        if hasattr(url, "partname"):  # cible interne
                            url = str(url.partname)
                # marque les <w:r> enfants de ce hyperlink
                for r in hl.iterchildren():
                    if r.tag.endswith("}r"):
                        m[r] = url
    except Exception:
        pass
    return m

def _para_inner_html(p: Paragraph) -> str:
    """Constitue l'HTML interne d'un paragraphe en conservant <a href> autour des runs dans les hyperliens."""
    link_by_run = _hyperlink_map(p)
    frags = []
    for run in p.runs:
        chunk = _run_to_html(run)
        url = link_by_run.get(run._r)
        if url:
            url_esc = _html_escape(str(url))
            chunk = f'<a href="{url_esc}" target="_blank" rel="noopener noreferrer">{chunk}</a>'
        frags.append(chunk)
    return "".join(frags)

def _para_list_kind(p: Paragraph, text: str) -> str | None:
    """Renvoie 'ul', 'ol' ou None sans xpath."""
    # 1) Numérotation Word native ?
    try:
        pPr = getattr(p._p, "pPr", None)
        numPr = getattr(pPr, "numPr", None) if pPr is not None else None
    except Exception:
        numPr = None
    if numPr is not None:
        sname = (p.style.name if getattr(p, "style", None) else "") or ""
        if "Number" in sname or re.match(r"^\s*\d+([.)]\s|$)", text or ""):
            return "ol"
        return "ul"

    # 2) Styles usuels
    sname = (p.style.name if getattr(p, "style", None) else "") or ""
    if any(k in sname for k in ["List", "Puces", "Bullet"]):
        return "ul"
    if "Number" in sname:
        return "ol"

    # 3) Symbole en début
    if (text or "").lstrip().startswith(("•", "◦", "▪", "-", "–", "—", "*")):
        return "ul"
    return None

def _para_list_info(p: Paragraph, text: str) -> tuple[str | None, int | None]:
    """('ul'/'ol', niveau>=0) ou (None, None). Niveau via numPr.ilvl, sinon heuristique indent."""
    kind = _para_list_kind(p, text)
    if not kind:
        return None, None
    level = 0
    try:
        pPr = getattr(p._p, "pPr", None)
        numPr = getattr(pPr, "numPr", None) if pPr is not None else None
        ilvl = getattr(numPr, "ilvl", None) if numPr is not None else None
        if ilvl is not None and getattr(ilvl, "val", None) is not None:
            level = int(ilvl.val)
        else:
            ind = getattr(pPr, "ind", None) if pPr is not None else None
            left = getattr(ind, "left", None) if ind is not None else None
            if left is not None:
                # 720 twips ~ 0.5", approximons un niveau par 720 twips
                level = max(0, min(6, int(left) // 720))
    except Exception:
        pass
    return kind, level

def _para_to_html(p: Paragraph) -> tuple[str, str]:
    """("p"|"li-ul"|"li-ol", html) — conserve <br>, <img>, styles, liens."""
    inner = _para_inner_html(p) or _html_escape(p.text or "")
    kind, _ = _para_list_info(p, p.text or "")
    if kind == "ol":
        return ("li-ol", f"<li>{inner}</li>")
    if kind == "ul":
        # si le symbole est dans le texte, on le retire
        for b in ("•","◦","▪","-","–","—","*"):
            if inner.startswith(b):
                inner = inner[len(b):].lstrip(); break
        return ("li-ul", f"<li>{inner}</li>")
    return ("p", f"<p>{inner}</p>")


def _iter_blocks(parent):
    """Parcourt Paragraph/Table dans l'ordre d'apparition."""
    # Cas Document : il faut descendre dans le body
    try:
        from docx.document import Document as _Doc
    except Exception:
        _Doc = None

    if _Doc is not None and isinstance(parent, _Doc):
        body = parent.element.body
        for child in body.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)
        return

    # Cas cellule de tableau
    if isinstance(parent, _Cell):
        for child in parent._tc.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)
        return

    # Fallback (autres parents)
    parent_elm = parent._element
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def parse_docx_sections_html(path, expected_headings: list[str]) -> dict[str, str]:
    doc = Document(path)
    exp = {_norm(h): h for h in expected_headings}
    exp.update({_norm(h.rstrip(":")): h for h in expected_headings})

    sections: dict[str, str] = {}
    current = None
    buf = []
    in_list = False
    list_kind = None  # 'ul'/'ol'

    def flush():
        nonlocal buf, in_list, list_kind, current
        if in_list:
            buf.append(f"</{list_kind}>"); in_list=False; list_kind=None
        if current and buf:
            html = "".join(buf).strip()
            if html: sections[current] = (sections.get(current,"") + html)
        buf = []

    for block in _iter_blocks(doc):
            if isinstance(block, Paragraph):
                t = (block.text or "").strip()
                if t and _looks_like_heading(t, block, exp):
                    # On ferme toutes les listes ouvertes avant de changer de section
                    while list_stack:
                        buf.append(f"</{list_stack.pop()}>")
                    flush()
                    current = exp.get(_norm(t), exp.get(_norm(t.rstrip(":")), t))
                    continue
        
                # --- gestion des listes NIVEAU / TYPE ---
                kind, level = _para_list_info(block, block.text or "")
                if kind is None:
                    # on ferme toutes les listes si on n'est plus dans une liste
                    while list_stack:
                        buf.append(f"</{list_stack.pop()}>")
                    # paragraphe simple
                    buf.append(_para_to_html(block)[1])
                    continue
        
                # On veut une profondeur cible = level+1 (car niveau 0 => 1 liste ouverte)
                target_depth = (level or 0) + 1
        
                # Ferme si on est trop profond
                while len(list_stack) > target_depth:
                    buf.append(f"</{list_stack.pop()}>")
        
                # Ouvre si pas assez profond
                while len(list_stack) < target_depth:
                    # si on ouvre le dernier niveau demandé, on respecte le type détecté (ul/ol)
                    to_open = kind if len(list_stack) + 1 == target_depth else "ul"
                    buf.append(f"<{to_open}>")
                    list_stack.append(to_open)
        
                # Ajuste le type au niveau courant si besoin (ul -> ol, etc.)
                if list_stack and list_stack[-1] != kind:
                    buf.append(f"</{list_stack.pop()}>")
                    buf.append(f"<{kind}>")
                    list_stack.append(kind)
        
                # Ajoute l'item
                _, li_html = _para_to_html(block)
                buf.append(li_html)
        
            else:
                # ---------- TABLE ----------
                # On ferme les listes ouvertes avant d'insérer un tableau
                while list_stack:
                    buf.append(f"</{list_stack.pop()}>")
        
                rows = []
                for row in block.rows:
                    tds = []
                    for cell in row.cells:
                        parts = []
                        for pp in cell.paragraphs:
                            k, frag = _para_to_html(pp)
                            parts.append(f"<ul>{frag}</ul>" if k.startswith("li") else frag)
                        tds.append(f"<td>{''.join(parts) or '&nbsp;'}</td>")
                    rows.append(f"<tr>{''.join(tds)}</tr>")
                buf.append(
                    "<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;width:100%'>"
                    + "".join(rows) + "</table>"
                )
    flush()
    return sections

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

# Upload
st.header("1) Charger la fiche .docx")
uploaded = st.file_uploader("Glissez le .docx ici", type=["docx"])
if uploaded is not None:
    tmp_path = Path("uploaded.docx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())

    # 1) Parser HTML robuste
    sections = parse_docx_sections_html(tmp_path, expected_headings=expected_word_headings)
    sections_norm = {_norm(k): v for k, v in sections.items()}

    # 2) Auto-mapping Word -> PDF/CRM (valeurs = HTML)
    fr_payload = {}
    rows = []
    for word_h, pdf_h in word_to_pdf.items():
        w_norm = _norm(word_h); w_norm2 = _norm(word_h.rstrip(":"))
        target_key = key_by_pdf_label_norm.get(_norm(pdf_h))
        found = (w_norm in sections_norm) or (w_norm2 in sections_norm)
        fr_html = sections_norm.get(w_norm) or sections_norm.get(w_norm2, "")
        if target_key: fr_payload[target_key] = fr_html
        rows.append({
            "Word heading attendu": word_h,
            "Dans le .docx ?": "✅ Oui" if found else "❌ Non",
            "PDF/CRM heading": pdf_h,
            "CRM key": target_key or "(non défini)"
        })

    st.subheader("Résultat du mapping automatique")
    st.dataframe(rows)

    # 3) Affichage vertical fidèle (HTML)
    st.header("Aperçu des sections (mise en forme préservée)")
    st.markdown("""
        <style>
          .sect p { margin: 0 0 10px 0; line-height: 1.55; }
          .sect ul, .sect ol { margin: 6px 0 12px 1.4rem; }
          .sect ul ul, .sect ol ol, .sect ul ol, .sect ol ul { margin-left: 1.2rem; }
          .sect table { border-collapse: collapse; width: 100%; margin: 6px 0 12px 0; }
          .sect td, .sect th { border: 1px solid #666; padding: 6px; vertical-align: top; }
          .sect img { max-width: 100%; height: auto; display: inline-block; }
          .sect a { text-decoration: underline; }
        </style>
        """, unsafe_allow_html=True)

    for fdef in fields:
        key = fdef["key"]; label = fdef["label"]
        html_content = fr_payload.get(key, "")
        st.subheader(label)
        st.markdown(f"<div class='sect'>{html_content or '<p><em>(vide)</em></p>'}</div>", unsafe_allow_html=True)
        st.divider()


