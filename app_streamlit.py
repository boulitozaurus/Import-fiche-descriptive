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
from bs4 import BeautifulSoup, Tag, NavigableString, Comment
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
import difflib

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

# Reconnaissance par mots-clés (secours si le titre exact n'est pas trouvé)
KEYWORD_TO_WH = {
    "points d attention": "Facteurs de risque",
    "facteurs de risque": "Facteurs de risque",
    "bonnes raisons": "Les bonnes raisons d'investir",
    "presentation de l operation": "Présentation de l'opération",
    "projet": "Projet",
    "localisation": "Localisation",
    "planning": "Administratif et timing",
    "marche": "Marché et références",
    "references": "Marché et références",
    "budget": "Budget de l'opération",
    "operateur": "L'opérateur",
    "track record": "Track record et opérations en cours",
    "structure": "Structure et Management",
    "management": "Structure et Management",
    "actionnariat": "Actionnariat et structure de l'opération",
    "finances": "Finances",
    "contexte": "Contexte et usage des fonds",
    "description": "Introduction",
    "finance": "Finances",
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
    s = (s or "")
    # Normaliser espaces & ponctuation “exotiques”
    s = s.replace("\u00A0", " ")  # NBSP -> espace
    s = s.translate(str.maketrans({
        "\u2019": "'",  # apostrophe courbe -> '
        "\u2018": "'",
        "\u2032": "'",
        "\u201C": '"',  # guillemets courbes -> "
        "\u201D": '"',
        "\u2013": "-",  # en dash/em dash/minus -> -
        "\u2014": "-",
        "\u2212": "-",
    }))
    s = re.sub(r"\s+", " ", s).strip()
    s = _strip_accents(s).lower()
    return s

def _strip_leading_numbering(s: str) -> str:
    return re.sub(r'^\s*(?:[\(\[]?\d+(?:\.\d+)*[\)\.]?|[ivxlcdm]+[\)\.]|[A-Z]\)|•|—|-)\s*', '', s or '', flags=re.I)

def _is_bullet_only_text(text: str) -> bool:
    t = (text or "").replace(NBSP, " ").strip()
    return bool(BULLET_ONLY_RE.match(t))

def _is_section_heading_p(p: Tag) -> bool:
    # h1–h3 => toujours un titre
    if getattr(p, "name", None) in {"h1","h2","h3"}:
        return True
    if getattr(p, "name", None) != "p":
        return False

    txt = p.get_text(" ", strip=True) or ""
    if not txt or len(txt) > 90:
        return False

    norm_txt = _norm(_strip_leading_numbering(txt)).rstrip(" :")

    KNOWN_EQUALS = {
        "introduction","description",
        "contexte et usage des fonds",
        "facteurs de risque",
        "les bonnes raisons d investir",
        "presentation de l operation",
        "localisation",
        "administratif et timing","planning",
        "marche et references",
        "budget de l operation","budget",
        "l operateur",
        "track record et operations en cours","track record",
        "structure et management",
        "actionnariat et structure de l operation","actionnariat",
        "finances","finance"
    }
    # p tout en gras ? on n'exige pas forcément, mais on reste sur égalité stricte
    return norm_txt in KNOWN_EQUALS

def _html_escape(s: str) -> str:
    return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def strip_leading_title_block(html: str) -> str:
    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")

    # sauter les nœuds vides/espaces au tout début
    def _iter_first_blocks():
        for child in soup.div.children:
            if isinstance(child, NavigableString) and not str(child).strip():
                continue
            yield child
            break

    for child in list(_iter_first_blocks()):
        name = getattr(child, "name", None)
        if name in {"h1","h2","h3"}:
            child.decompose()
            break

        if name == "p":
            t = child.get_text(" ", strip=True) or ""
            # (a) Beaucoup de majuscules (= titre plaqué)
            letters = [c for c in t if c.isalpha()]
            upper_ratio = (sum(1 for c in letters if c.isupper()) / len(letters)) if letters else 0

            # (b) Ligne courte avec tiret long/court (pattern "Nom – Lieu"), pas de point final
            has_dash = (" - " in t) or (" – " in t)
            short_line = len(t) <= 120 and has_dash and not re.search(r"[\.!?]$", t)

            # (c) Ligne strictement en gras (un seul enfant <strong>/<b>) et courte
            is_strong_only = (len(child.contents) == 1 and getattr(child.contents[0], "name", None) in {"strong","b"} and len(t) <= 120)

            if upper_ratio >= 0.6 or short_line or is_strong_only:
                child.decompose()
            break

    return soup.div.decode_contents()

# ================= GESTION DES IMAGES =================

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

        # pixel transparent + marqueurs pour qu’on sache générer un bouton plus tard
        return {
            "src": "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==",
            "alt": "Image non affichable (EMF/WMF).",
            "data-unsupported": "1",
            "data-uid": uid,
        }

    import base64
    b64 = base64.b64encode(data).decode("ascii")
    return {"src": f"data:{ctype};base64,{b64}"}

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
    """
    Index = clef normalisée -> Titre Word (pas CRM).
    Ajoute des alias très limités (Description→Introduction, Contexte &/et, Finances/Finance).
    """
    idx: dict[str, str] = {}
    for wh in expected_headings:
        # variantes directes : forme telle quelle, version mappée (si tu en as), version sans numérotation
        for alias in {wh, word_to_pdf.get(wh, ''), _strip_leading_numbering(wh)}:
            if alias:
                idx[_norm(alias.rstrip(" :"))] = wh

    # alias minimaux, sûrs
    idx[_norm("description")] = "Introduction"
    idx[_norm("contexte & usage des fonds")] = "Contexte et usage des fonds"
    idx[_norm("finance")] = "Finances"
    idx[_norm("finances")] = "Finances"
    return idx

def split_sections_by_headings(html: str, heading_index: dict[str, str]) -> dict[str, str]:
    """
    Découpage conservateur et déterministe (style v22) :
    - on n'utilise que les titres connus (égalité stricte sur texte normalisé)
    - on range tout préambule avant 1er titre dans 'Introduction'
    - si 'Présentation de l'opération' est présent, on ignore 'Projet' comme titre top-niveau
    """
    def nrm(s: str) -> str:
        return _norm(_strip_leading_numbering((s or "")).rstrip(" :"))

    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    out = {v: "" for v in set(heading_index.values())}

    # Pré-scan : repérer les titres visibles
    present_titles = set()
    for el in soup.find_all(["h1","h2","h3","p"]):
        t = nrm(el.get_text(" ", strip=True))
        if t in heading_index:
            present_titles.add(heading_index[t])

    ignore_projet = "Présentation de l'opération" in present_titles

    current = None
    for el in soup.div.children:
        if not hasattr(el, "get_text"):
            continue

        name = getattr(el, "name", None)
        text = el.get_text(" ", strip=True) if name in {"h1","h2","h3","p"} else ""
        key = None

        if text:
            norm_text = nrm(text)
            if norm_text in heading_index:
                wh = heading_index[norm_text]   # Titre Word
                # ignorer 'Projet' si 'Présentation...' est présent
                if ignore_projet and wh == "Projet":
                    wh = None
                key = wh

        if key:
            current = key
            continue

        # Si rien encore détecté, tout va dans Introduction
        if current is None and text:
            current = "Introduction"

        if current in out:
            out[current] += str(el)

    return out

# ================= NETTOYAGE DES LISTES =================

def _fix_lists_in_soup(soup):
    changed = True
    while changed:
        changed = False

        # Retire <p> ou <li> "puces seules" / vides
        for p in list(soup.find_all("p")):
            if _is_bullet_only_text(p.get_text(" ", strip=True)):
                p.decompose()
                changed = True

        for li in list(soup.find_all("li")):
            txt = li.get_text(" ", strip=True)
            if not txt or _is_bullet_only_text(txt):
                li.decompose()
                changed = True

        # Nettoie <p> vides directement sous <li>
        for li in list(soup.find_all("li")):
            for p in list(li.find_all("p", recursive=False)):
                if not p.get_text(strip=True):
                    p.decompose()
                    changed = True

        # Aplatis quelques cas simples (li sans texte direct + une seule sous-liste)
        for li in list(soup.find_all("li")):
            direct_text = "".join(t for t in li.find_all(string=True, recursive=False)).strip()
            child_lists = [c for c in li.contents if getattr(c, "name", None) in ("ol", "ul")]
            if direct_text == "" and len(child_lists) == 1:
                inner = child_lists[0]
                parent = li.parent
                if getattr(parent, "name", None) not in ("ul", "ol"):
                    continue

                if parent.name == inner.name:
                    for sub_li in inner.find_all("li", recursive=False):
                        li.insert_before(sub_li)
                    li.decompose()
                    changed = True
                else:
                    li.replace_with(inner)
                    changed = True

        # Supprime <ul> sans <li> qui ne contiennent qu'une sous-liste
        for ul in list(soup.find_all("ul")):
            lis = ul.find_all("li", recursive=False)
            if len(lis) == 0:
                only_lists = [c for c in ul.contents if getattr(c, "name", None) in ("ol", "ul")]
                if len(only_lists) == 1:
                    ul.replace_with(only_lists[0])
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

    #_convert_numbered_paragraphs_to_ol(soup)
    #for cont in soup.find_all(["div", "section"]):
        #_convert_numbered_paragraphs_to_ol(cont)

    # récupérer les EMF/WMF stockées durant la conversion
    for img in list(soup.find_all("img")):
        if img.get("data-unsupported") == "1":
            uid = img.get("data-uid")
            if uid and "img_store" in st.session_state and uid in st.session_state["img_store"]:
                fname, data, ctype = st.session_state["img_store"][uid]
                downloads.append((uid, fname, data, ctype))
                # ⬇️ Remplacer l'image par un commentaire HTML "DL:<uid>"
                img.replace_with(Comment(f"DL:{uid}"))
            else:
                img.decompose()
            
    _fix_lists_in_soup(soup)

    cleaned = soup.div.decode_contents()
    cleaned = re.sub(r"(?is)le contenu\s+g[éè]n[éè]r[éè]\s+par l[''' ]?ia\s+peut\s+être\s+incorrect\.?", '', cleaned)

    return cleaned, downloads

# ================= CORRECTION DE LA NUMÉROTATION =================

def fix_section_numbering(html: str, section_key: str) -> str:
    if not html or not html.strip():
        return html

    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")

    EXPECTED = {
        'points_attention_fr': [
            "Risque lié au projet",
            "Risque lié au secteur",
            "Risque de défaut",
        ],
        'bonnes_raisons_fr': [
            "Une assurance sur 100% du capital investi",
            "Une fiducie-sûreté sur l'actif",
        ],
        'budget_fr': [
            "Prix de revient",
            "Financement et ratios",
            "Revenus et marges",
            "Couverture des intérêts",
            "Stress test",
        ],
    }
    if section_key not in EXPECTED:
        return html

    def nrm(s: str) -> str:
        s = re.sub(r'^\s*(?:[\(\[]?\d+(?:\.\d+)*[\)\.]?|[ivxlcdm]+[\)\.]|[A-Z]\)|•|–|—|-|\*)\s*', '', s or '', flags=re.I)
        s = _strip_accents((s or "").lower()).replace("\u00A0"," ").strip()
        s = s.replace(" d’"," d'").replace(" l’"," l'")
        return " ".join(s.split())

    expected = list(EXPECTED[section_key])

    # Carte libellé-normalisé -> libellé canonique
    expected_map = { nrm(lbl): lbl for lbl in expected }

    # Budget : variantes de "Couverture(s) des intérêts"
    if section_key == 'budget_fr':
        expected_map[nrm("Couvertures des intérêts")] = "Couverture des intérêts"
        expected_map[nrm("couverture des interets")]  = "Couverture des intérêts"
        expected_map[nrm("couvertures des interets")] = "Couverture des intérêts"

    # Bonnes raisons : variantes tolérantes sur "fiducie-sûreté"
    if section_key == 'bonnes_raisons_fr':
        expected_map[nrm("Une fiducie surete sur l actif")] = "Une fiducie-sûreté sur l'actif"
        expected_map[nrm("Une fiducie surete sur l'actif")] = "Une fiducie-sûreté sur l'actif"

    # Collecter la première occurrence de chaque libellé (match en début de ligne)
    first = {}
    for el in soup.find_all(['h1','h2','h3','h4','h5','h6','p','li','strong','b','em','i','u']):
        txt = el.get_text(" ", strip=True)
        if not txt or len(txt) > 180:
            continue
        nn = nrm(txt)
        for norm_lbl, canon in expected_map.items():
            if nn.startswith(norm_lbl) and canon not in first:
                first[canon] = el
                break

    # Bonnes raisons : Fiducie devient #1 si Assurance absente
    if section_key == 'bonnes_raisons_fr':
        if "Une assurance sur 100% du capital investi" not in first:
            expected = ["Une fiducie-sûreté sur l'actif"]
        else:
            expected = [
                "Une assurance sur 100% du capital investi",
                "Une fiducie-sûreté sur l'actif",
            ]

    # Ordre final (Stress test uniquement s'il existe)
    order = [x for x in expected if x in first]
    if section_key == 'budget_fr' and "Stress test" in first and "Stress test" not in order:
        order.append("Stress test")

    def _remove_leading_label_from_li(li: Tag, label_norm: str):
        """Retire la 'ligne-titre' au début de la <li>, qu'elle soit dans <p>/<em>/<u>/<strong> ou texte brut."""
        def norm(s: str) -> str:
            s = re.sub(r'^\s*(?:[\(\[]?\d+(?:\.\d+)*[\)\.]?|[ivxlcdm]+[\)\.]|[A-Z]\)|•|–|—|-|\*)\s*', '', s or '', flags=re.I)
            s = _strip_accents((s or "").lower()).replace("\u00A0"," ").strip()
            s = s.replace(" d’"," d'").replace(" l’"," l'")
            return " ".join(s.split())
    
        # 1) si premier enfant bloc ressemble au titre, on le supprime
        for child in list(li.children):
            if isinstance(child, NavigableString) and not str(child).strip():
                continue
            if getattr(child, "name", None) in {"p","strong","b","em","i","u","span"}:
                t = child.get_text(" ", strip=True) or ""
                if norm(t).startswith(label_norm):
                    child.decompose()
                break
            if isinstance(child, NavigableString):
                t = str(child)
                if norm(t).startswith(label_norm):
                    # on coupe juste le début correspondant
                    child.replace_with("")
                break
            # autre tag (ex: div) -> on arrête, pas de risque ici
            break

    # Appliquer la numérotation visible (et neutraliser la numérotation auto des <ol>)
    def set_title(el: Tag, label_text: str, italic_underline: bool):
        # repère si on est dans une <li>
        li = el if el.name == "li" else el.find_parent("li")
        if li:
            # liste la plus proche
            cur_list = li.find_parent(["ol","ul"])
            # remonter à la liste racine
            root_list = cur_list
            while True:
                pli = root_list.find_parent("li") if root_list else None
                plist = pli.find_parent(["ol","ul"]) if pli else None
                if plist: root_list = plist
                else: break
    
            # neutralise les marqueurs auto
            for lst in filter(None, [cur_list, root_list]):
                lst["data-noautonum"] = "1"
    
            # crée le <p> titre (soulignement inline)
            p = soup.new_tag("p")
            p["data-fixed-title"] = "1"
            if italic_underline:
                em = soup.new_tag("em")
                span = soup.new_tag("span")
                span["style"] = "text-decoration: underline;"
                span.string = label_text
                em.append(span)
                p.append(em)
            else:
                p.string = label_text
    
            # insère avant la liste racine (ou la plus proche à défaut)
            (root_list or cur_list or li).insert_before(p)
    
            # retire la ligne-titre d'origine dans la <li>
            lbl_norm = _strip_accents(label_text.lower()).replace("\u00A0"," ")
            lbl_norm = " ".join(lbl_norm.split())
            _remove_leading_label_from_li(li, lbl_norm)
    
            # si la <li> est vide après nettoyage, on l'enlève
            if not (li.get_text(strip=True) or li.find(True)):
                li.decompose()
            return
    
        # Pas dans une liste : réécriture in situ
        el.clear()
        if italic_underline:
            em = soup.new_tag("em")
            span = soup.new_tag("span")
            span["style"] = "text-decoration: underline;"
            span.string = label_text
            em.append(span)
            el.append(em)
        else:
            el.string = label_text

    for i, title in enumerate(order, 1):
        el = first[title]
        set_title(el, f"{i}. {title}", italic_underline=(section_key == 'budget_fr'))

    return soup.div.decode_contents()

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
      .sect p[data-fixed-title="1"] {
      margin-left: 0 !important;
      padding-left: 0 !important;
      text-indent: 0 !important;
      }
      .sect p { margin: .30rem 0; }
      .sect ol, .sect ul { margin: .40rem 0 .60rem 1.4rem; padding-left: 1.2rem; list-style-position: outside; }
      .sect ol { list-style-type: decimal; }
      .sect ul { list-style-type: disc; }
      .sect ul[data-noautonum="1"] { list-style: none; padding-left: 0; margin-left: 0; }
      .sect ul[data-noautonum="1"] > li { margin-left: 0; }
      
      .sect ol[data-noautonum="1"] { list-style: none; padding-left: 0; margin-left: 0; }
      .sect ol[data-noautonum="1"] > li { margin-left: 0; }
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
    
    # DÉBOGAGE : Afficher les sections détectées
    with st.expander("🔍 Débogage : Sections détectées dans le HTML"):
        st.write("**Sections avec contenu :**")
        for k, v in sections.items():
            if v.strip():
                st.write(f"- **{k}** : {len(v)} caractères")
        st.write("**Sections vides :**")
        for k, v in sections.items():
            if not v.strip():
                st.write(f"- {k}")

    st.header("Aperçu des sections (mise en forme préservée)")
    inject_css()
    
    for fdef in fields:
        key = fdef["key"]
        label = fdef["label"]
    
        raw_html = fr_payload.get(key, "")
        if key == "description_fr":
            raw_html = strip_leading_title_block(raw_html)
        clean_html, dls = prepare_section_html(raw_html)
        dlmap = {uid: (fname, data, ctype) for uid, fname, data, ctype in dls}
        
        parts = re.split(r'<!--DL:([0-9a-f]+)-->', clean_html, flags=re.I)
        
        st.subheader(label)
        for idx, part in enumerate(parts):
            if idx % 2 == 0:
                # morceau HTML normal
                if part.strip():
                    st.markdown(f"<div class='sect'>{part}</div>", unsafe_allow_html=True)
            else:
                # jeton DL -> bouton
                uid = part.lower()
                if uid in dlmap:
                    fname, data, ctype = dlmap[uid]
                    st.download_button(
                        f"Télécharger {fname}",
                        data=data,
                        file_name=fname,
                        mime=ctype,
                        key=f"dl_{uid}"
                    )
