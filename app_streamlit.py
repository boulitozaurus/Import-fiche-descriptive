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


HEADING_WORDS = (
    "introduction|contexte|les points d'attention|les bonnes raisons|"
    "projet|localisation|planning|budget|opérateur|opérateur|actionnariat|finances?"
)

_HEADING_RE = re.compile(rf"^\s*(?:{HEADING_WORDS})\b", re.I)

NUM_PREFIX_RE = re.compile(r"^\s*(\d+)[\.\)]\s+")

def _is_section_heading_p(p: Tag) -> bool:
    """Heuristique simple pour ne PAS aspirer des titres de section dans un <li>."""
    if p.name != "p":
        return False
    txt = p.get_text(" ", strip=True)
    if not txt:
        return False
    # court, fort, ou commence par un mot de section connu → probablement un titre
    if len(txt) <= 90 and (p.find(["strong", "b"]) or _HEADING_RE.match(txt)):
        return True
    return False

def _data_uri(image):
    data = base64.b64encode(image.read()).decode("ascii")
    return {"src": f"data:{image.content_type};base64,{data}"}

def split_sections_by_headings(html: str, heading_index: dict[str, str]) -> dict[str, str]:
    # normalise : retire numérotation + ":" final, supprime accents / casse / espaces
    def nrm(s: str) -> str:
        return _norm(_strip_leading_numbering((s or "").rstrip(" :")))

    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    out = {v: "" for v in set(heading_index.values())}
    current = None

    for el in soup.div.children:
        if not hasattr(el, "get_text"):   # ignorer les strings blanches
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

    # EMF/WMF -> téléchargement
    if ctype in ("image/x-emf", "image/emf", "image/x-wmf", "image/wmf", "application/octet-stream"):
        import uuid, base64
        uid = uuid.uuid4().hex
        fname = f"{uid}.{'emf' if 'emf' in ctype else ('wmf' if 'wmf' in ctype else 'bin')}"
        st.session_state.setdefault("img_store", {})[uid] = (fname, data, ctype)
        return {
            "src": "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==",
            "alt": "",                          # <= pas de texte alternatif visible
            "data-unsupported": "1",
            "data-uid": uid,
        }

    import base64
    b64 = base64.b64encode(data).decode("ascii")
    # IMPORTANT : forcer alt="" pour neutraliser tout alt auto de Mammoth/Word
    return {"src": f"data:{ctype};base64,{b64}", "alt": ""}

def docx_to_html(path: str) -> str:
    """Convertit le .docx en HTML avec Mammoth (heading robustes + handler d’images)."""
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

def prepare_section_html(html: str):
    soup = BeautifulSoup(f"<div>{html}</div>", "html.parser")
    downloads = []

    # (a) supprimer la phrase IA parasite
    IA_PAT = re.compile(r"le contenu\s+g[ée]n[ée]r[ée]\s+par l[’' ]?ia\s+peut\s+être\s+incorrect\.?", re.I)
    for p in list(soup.find_all("p")):
        if IA_PAT.search(p.get_text(" ", strip=True)):
            p.decompose()

    # (b) gérer les images EMF/WMF (déjà présent chez toi)
    for img in list(soup.find_all("img")):
        if img.get("data-unsupported") != "1" and img.has_attr("alt"):
            del img["alt"]

    # Unwrap des paragraphes entièrement en <strong>/<b> (Word "Caractères forts")
    for p in list(soup.find_all("p")):
        if len(p.contents) == 1 and getattr(p.contents[0], "name", None) in ("strong", "b"):
            p.contents[0].unwrap()
    #    (on le fait sur le root ET sur les conteneurs internes de type <div>/<section>)
    _convert_numbered_paragraphs_to_ol(soup)
    for cont in soup.find_all(["div", "section"]):
        _convert_numbered_paragraphs_to_ol(cont)

    # (c) corriger les listes fantômes
    _fix_lists_in_soup(soup)

    # (d) réparer les data URI orphelines -> vraies <img>
    cleaned = _fix_stray_data_uri(soup.div.decode_contents())

    # (e) 2e filet pour la phrase IA si encodage différent
    cleaned = re.sub(r'(?is)le contenu\s+g[ée]n[ée]r[ée]\s+par l[’\' ]?ia\s+peut\s+être\s+incorrect\.?', '', cleaned)

    return cleaned, downloads

def _strip_leading_numbering(s: str) -> str:
    """Supprime une numérotation de début de ligne: '1.2. ', 'I. ', 'A) ', '• ', '-', '– ' ..."""
    return re.sub(r'^\s*(?:[\(\[]?\d+(?:\.\d+)*[\)\.]?|[ivxlcdm]+[\)\.]|[A-Z]\)|•|–|-)\s*', '', s or '', flags=re.I)

def build_heading_index(expected_headings: list[str], word_to_pdf: dict[str, str]) -> dict[str, str]:
    """
    Construit un index de correspondance {titre_normalisé -> titre_attendu}
    en acceptant :
      - le titre Word,
      - le libellé PDF/CRM,
      - et ces mêmes titres sans numérotation initiale.
    + quelques alias manuels utiles.
    """
    idx: dict[str, str] = {}
    for wh in expected_headings:
        for alias in {wh, word_to_pdf.get(wh, ''), _strip_leading_numbering(wh)}:
            if alias:
                idx[_norm(alias)] = wh

    # Aliases manuels fréquents
    idx[_norm("description")] = "Introduction"
    idx[_norm("contexte & usage des fonds")] = "Contexte et usage des fonds"
    return idx

NBSP = "\u00A0"
_BULLETS = "•◦●▪▫·–—-o" + NBSP
_BULLET_CLASS = "".join(re.escape(ch) for ch in _BULLETS)
BULLET_ONLY_RE = re.compile(r'^[\s' + _BULLET_CLASS + r']+$', re.I)


def _is_bullet_only_text(text: str) -> bool:
    t = (text or "").replace(NBSP, " ").strip()
    return bool(BULLET_ONLY_RE.match(t))

def _merge_split_ol_blocks(soup: BeautifulSoup) -> bool:
    """
    Recoller des listes numérotées coupées par des <p> :
    <ol>…</ol><p>…</p><p>…</p><ol>…</ol>
    - rattache les <p> intermédiaires au dernier <li> du 1er <ol>
    - fusionne les <li> du 2e <ol> dans le 1er
    """
    changed = False
    for ol in list(soup.find_all("ol")):
        sib = ol.find_next_sibling()
        if sib is None:
            continue

        # collecter les <p> intercalés (en évitant les vrais titres de section)
        trail = []
        cur = sib
        while cur is not None and getattr(cur, "name", None) == "p" and not _is_section_heading_p(cur):
            trail.append(cur)
            cur = cur.find_next_sibling()

        # si derrière ces <p> on retombe sur un autre <ol> → fusion
        if cur is not None and getattr(cur, "name", None) == "ol":
            lis = ol.find_all("li", recursive=False)
            last_li = lis[-1] if lis else None

            # rattacher les <p> au dernier <li> (ou créer un <li> si besoin)
            if trail:
                if last_li is None:
                    last_li = soup.new_tag("li")
                    ol.append(last_li)
                for p in trail:
                    last_li.append(p.extract())

            # déplacer les <li> du second <ol> dans le premier
            for li in list(cur.find_all("li", recursive=False)):
                ol.append(li.extract())

            cur.decompose()
            changed = True

    return changed

def _looks_like_heading(p_tag: Tag) -> bool:
    """Heuristique simple pour NE PAS aspirer un titre dans un <li> précédent."""
    if p_tag.find(["h1", "h2", "h3", "h4", "h5", "h6"]):
        return True
    # court + en gras => probablement un intertitre
    txt = p_tag.get_text(" ", strip=True)
    if len(txt) <= 80 and p_tag.find(["strong", "b"]):
        return True
    return False


def _convert_numbered_paragraphs_to_ol(parent: Tag) -> bool:
    """
    Balaye les enfants directs de `parent` :
      - regroupe P commençant par '1. ', '2) '… dans une vraie <ol>
      - déplace les paragraphes 'explicatifs' qui suivent dans le <li> précédent
    Retourne True si des changements ont été faits.
    """
    changed = False
    children = list(parent.children)
    i = 0
    while i < len(children):
        node = children[i]
        if getattr(node, "name", None) == "p":
            text = node.get_text(" ", strip=True)
            m = NUM_PREFIX_RE.match(text or "")
            if m:
                # démarrage de groupe numéroté
                ol = parent.new_tag("ol")
                parent.insert_before(ol, node)
                last_li = None

                # Traitement des éléments consécutifs (P numérotés + P explicatifs)
                j = i
                while j < len(children):
                    cur = children[j]
                    name = getattr(cur, "name", None)

                    if name == "p":
                        t = cur.get_text(" ", strip=True)
                        m2 = NUM_PREFIX_RE.match(t or "")
                        if m2:
                            # créer un <li> en enlevant le préfixe "1. " du premier texte
                            li = parent.new_tag("li")
                            # retire le préfixe dans le premier NavigableString
                            first_txt = None
                            for c in cur.contents:
                                if isinstance(c, NavigableString):
                                    first_txt = c
                                    break
                            if first_txt:
                                m3 = NUM_PREFIX_RE.match(str(first_txt))
                                if m3:
                                    first_txt.replace_with(first_txt[m3.end():])
                            # déplacer le contenu du <p> dans le <li>
                            for c in list(cur.contents):
                                li.append(c.extract())
                            cur.decompose()
                            ol.append(li)
                            last_li = li
                            j += 1
                            changed = True
                            continue

                        # paragraphe explicatif directement après un <li> ?
                        if last_li and t and not _is_bullet_only_text(t) and not _looks_like_heading(cur):
                            last_li.append(cur.extract())
                            children.pop(j)
                            changed = True
                            continue

                        break  # fin du groupe

                    elif name in ("ul", "ol") and last_li:
                        # si une sous-liste arrive juste après, rattache-la au <li>
                        last_li.append(cur.extract())
                        children.pop(j)
                        changed = True
                        continue

                    else:
                        break

                # rafraîchir les enfants et se placer derrière le <ol> nouvellement inséré
                children = list(parent.children)
                i = list(parent.children).index(ol) + 1
                continue

        i += 1
    return changed

def _promote_nested_ol_to_siblings(soup: BeautifulSoup) -> bool:
    """
    Si un <li> contient une <ol> (sous-liste) et que cette sous-liste ressemble à une
    suite logique (2-6 items, phrases assez longues), on promeut les <li> internes
    au même niveau que le <li> parent. Ça corrige les '1. 1. 1.' vus dans
    'Les bonnes raisons d'investir'.
    """
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

        # Texte direct du li (hors sous-listes)
        direct_text = "".join(t for t in li.find_all(string=True, recursive=False)).strip()
        # Heuristique : on ne promeut pas si le li n'a *aucun* texte avant la sous-liste
        if len(direct_text) < 40:
            continue

        # Heuristique simple : items "longs" => vraies raisons/phrases
        long_items = sum(1 for s in sub_items if len(s.get_text(" ", strip=True)) >= 30)
        if long_items < len(sub_items) // 2:
            continue

        # Promotion : placer chaque sub <li> juste après le <li> courant
        anchor = li
        for sub in list(sub_items):
            anchor.insert_after(sub.extract())
            anchor = sub

        inner.decompose()
        changed = True

    return changed


def _fix_lists_in_soup(soup):
    """
    - Retire les paragraphes & <li> 'puces fantômes' (•, o, – tout seul).
    - Retire les <p> vides dans les <li>.
    - Aplati ul>li>ol / ul>li>ul mal emboîtés.
    - Fusionne les <ol> adjacents ou séparés par des <p> non titres.
    - *NOUVEAU* : promeut des <ol> internes au niveau du parent (corrige 1. 1. 1.).
    """
    changed = True
    while changed:
        changed = False

        # 1) <p> ne contenant qu'une puce
        for p in list(soup.find_all("p")):
            if _is_bullet_only_text(p.get_text(" ", strip=True)):
                p.decompose()
                changed = True

        # 2) <li> vides ou puces seules
        for li in list(soup.find_all("li")):
            txt = li.get_text(" ", strip=True)
            if not txt or _is_bullet_only_text(txt):
                li.decompose()
                changed = True

        # 3) <p> vides directement sous <li>
        for li in list(soup.find_all("li")):
            for p in list(li.find_all("p", recursive=False)):
                if not p.get_text(strip=True):
                    p.decompose()
                    changed = True

        # 4) Aplatir ul>li>ol / ul>li>ul
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

        # 4-bis) <ul> sans <li> -> remplacer par sa seule sous-liste
        for ul in list(soup.find_all("ul")):
            lis = ul.find_all("li", recursive=False)
            if len(lis) == 0:
                only_lists = [c for c in ul.contents if getattr(c, "name", None) in ("ol", "ul")]
                if len(only_lists) == 1:
                    ul.replace_with(only_lists[0])
                    changed = True

        # 4-ter) recoller les <ol> séparés par des <p> (non titres)
        if _merge_split_ol_blocks(soup):
            changed = True

        # 4-quater) *Promotion* des sous-<ol> raisonnables
        if _promote_nested_ol_to_siblings(soup):
            changed = True

        # 5) Fusionner <ol> consécutifs
        for ol in list(soup.find_all("ol")):
            nxt = ol.find_next_sibling()
            if getattr(nxt, "name", None) == "ol":
                for li in list(nxt.find_all("li", recursive=False)):
                    ol.append(li)
                nxt.decompose()
                changed = True

    return soup
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

def _fix_stray_data_uri(html: str) -> str:
    """
    EX-IMPORTANT : cette fonction a cassé des <img> en dupliquant l'attribut src.
    On la neutralise (Mammoth produit déjà des <img> valides).
    Si un jour tu dois re-traiter du texte brut 'src="data:..."' hors balise, fais-le
    avec un parseur HTML, pas un regex.
    """
    return html

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
    heading_index = build_heading_index(expected_word_headings, word_to_pdf)
    sections = split_sections_by_headings(html, heading_index)
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
    fr_payload = apply_fixed_numbering(fr_payload)

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
