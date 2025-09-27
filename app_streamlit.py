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

def _para_inner_html(p: Paragraph) -> str:
    """
    Construit l'HTML du paragraphe en respectant l'ordre réel des enfants XML :
    - texte/images/retours (runs ordinaires),
    - hyperliens <w:hyperlink r:id="...">...</w:hyperlink>,
    - hyperliens <w:fldSimple w:instr="HYPERLINK ...">...</w:fldSimple>.
    """
    frags: list[str] = []

    def _html_runs(children) -> str:
        chunks = []
        for r in children:
            if r.tag.endswith("}r"):
                chunks.append(_run_to_html(Run(r, p)))
        return "".join(chunks)

    for child in p._p.iterchildren():
        tag = child.tag

        # --- Cas 1 : <w:hyperlink> ---
        if tag.endswith("}hyperlink"):
            # URL via relation, sinon ancre (#bookmark)
            url = None
            rid = child.get(qn("r:id"))
            if rid:
                try:
                    rel = p.part.rels.get(rid)
                    if rel is not None:
                        url = getattr(rel, "target_ref", None)
                        if not url:
                            tp = getattr(rel, "target_part", None)
                            if tp is not None and hasattr(tp, "partname"):
                                url = str(tp.partname)
                except Exception:
                    url = None
            if not url:
                anchor = child.get(qn("w:anchor"))
                if anchor:
                    url = f"#{anchor}"

            inner_html = _html_runs(child.iterchildren())
            if url:
                frags.append(f'<a href="{_html_escape(str(url))}" target="_blank" rel="noopener noreferrer">{inner_html}</a>')
            else:
                frags.append(inner_html)
            continue

        # --- Cas 2 : <w:fldSimple w:instr="HYPERLINK ..."> ---
        if tag.endswith("}fldSimple"):
            instr = child.get(qn("w:instr")) or ""
            m = re.search(r'HYPERLINK\s+"([^"]+)"', instr, flags=re.I) or re.search(r'HYPERLINK\s+(\S+)', instr, flags=re.I)
            url = m.group(1) if m else None
            inner_html = _html_runs(child.iterchildren())
            if url:
                frags.append(f'<a href="{_html_escape(str(url))}" target="_blank" rel="noopener noreferrer">{inner_html}</a>')
            else:
                frags.append(inner_html)
            continue

        # --- Cas 3 : run ordinaire (texte/images/retours) ---
        if tag.endswith("}r"):
            frags.append(_run_to_html(Run(child, p)))
            continue

        # Autres balises (bookmarks, etc.) -> ignorées

    # Fallback si aucun enfant géré (cas rare)
    if not frags:
        return "".join(_run_to_html(run) for run in p.runs)

    return "".join(frags)

def _list_kind_from_numbering(p: Paragraph) -> str | None:
    """Retourne 'ol' (numérotée) ou 'ul' (puces) si le paragraphe appartient à une liste Word."""
    try:
        pPr  = getattr(p._p, "pPr", None)
        numPr = getattr(pPr, "numPr", None) if pPr is not None else None
        if numPr is None:
            return None

        numId_el = getattr(numPr, "numId", None)
        ilvl_el  = getattr(numPr, "ilvl", None)
        numId = str(getattr(numId_el, "val", None)) if numId_el is not None else None
        ilvl  = int(getattr(ilvl_el, "val", 0)) if ilvl_el is not None and getattr(ilvl_el, "val", None) is not None else 0
        if not numId:
            return None

        np = getattr(p.part, "numbering_part", None)
        if np is None:
            return None
        root = np.element

        # 1) numId -> abstractNumId
        abstract_id = None
        for num in root.iterchildren():
            if num.tag.endswith("}num") and num.get(qn("w:numId")) == numId:
                for ch in num.iterchildren():
                    if ch.tag.endswith("}abstractNumId"):
                        abstract_id = ch.get(qn("w:val"))
                        break
                break
        if abstract_id is None:
            return None

        # 2) abstractNumId -> numFmt (niveau ilvl si dispo)
        fmt = None
        for absn in root.iterchildren():
            if absn.tag.endswith("}abstractNum") and absn.get(qn("w:abstractNumId")) == abstract_id:
                # cherche le niveau exact
                for lvl in absn.iterchildren():
                    if lvl.tag.endswith("}lvl") and lvl.get(qn("w:ilvl")) == str(ilvl):
                        for comp in lvl.iterchildren():
                            if comp.tag.endswith("}numFmt"):
                                fmt = (comp.get(qn("w:val")) or "").lower()
                                break
                        if fmt: break
                # sinon prend le premier format trouvé
                if not fmt:
                    for lvl in absn.iterchildren():
                        if lvl.tag.endswith("}lvl"):
                            for comp in lvl.iterchildren():
                                if comp.tag.endswith("}numFmt"):
                                    fmt = (comp.get(qn("w:val")) or "").lower()
                                    break
                            if fmt: break
                break

        if not fmt:
            return None
        return "ul" if fmt == "bullet" else "ol"
    except Exception:
        return None

def _para_list_kind(p: Paragraph, text: str) -> str | None:
    """Renvoie 'ul', 'ol' ou None."""
    kind = _list_kind_from_numbering(p)
    if kind:
        return kind

    # Heuristiques de secours (si numbering non exploitable)
    sname = (p.style.name if getattr(p, "style", None) else "") or ""
    if "Number" in sname or re.match(r"^\s*\d+([.)]\s|$)", text or ""):
        return "ol"
    if any(k in sname for k in ["List", "Puces", "Bullet"]):
        return "ul"
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
    """
    Retourne ("p" | "li-ul" | "li-ol", html).
    - Conserve styles (gras/italique/souligné/couleur), <br/>, images et liens.
    - Détecte listes et retire le symbole de puce s'il est présent dans le texte.
    """
    # HTML interne du paragraphe (runs formatés + liens + images)
    inner = _para_inner_html(p) or _html_escape(p.text or "")
    # Auto-link des URLs brutes si Word n'a pas créé d'hyperlien
    inner = _autolink_html(inner)

    # Type de liste (ul/ol) + niveau (non utilisé ici, géré par la pile côté appelant)
    kind, _ = _para_list_info(p, p.text or "")

    if kind == "ol":
        return ("li-ol", f"<li>{inner}</li>")

    if kind == "ul":
        # Si le symbole de puce fait partie du texte, on l’enlève pour éviter le doublon
        for b in ("•", "◦", "▪", "-", "–", "—", "*"):
            if inner.startswith(b):
                inner = inner[len(b):].lstrip()
                break
        return ("li-ul", f"<li>{inner}</li>")

    # Paragraphe standard
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
    """
    - Conserve l'OL ouvert (continuité de numérotation) et insère les paragraphes explicatifs
      DANS le <li> courant, marqués <p class="cont"> pour les dé-décaler visuellement.
    - UL sous OL : force l'imbrication (sous-niveau) si Word ne renvoie pas le bon ilvl.
    - Redémarrage propre d'un OL après un titre / paragraphe d'intro finissant par ':'
      ou si Word change de numId entre deux listes.
    - Tableaux Word -> <table>.
    Prérequis : _norm, _looks_like_heading, _iter_blocks, _para_list_info, _para_to_html.
    """
    doc = Document(path)

    exp = {_norm(h): h for h in expected_headings}
    exp.update({_norm(h.rstrip(":")): h for h in expected_headings})

    sections: dict[str, str] = {}
    current: str | None = None
    buf: list[str] = []

    # pile de listes ouvertes avec numId (quand Word numérote nativement)
    list_stack: list[dict] = []   # ex: [{"tag":"ol","numId":"7"}, {"tag":"ul","numId":None}]
    # compteurs <ol> par niveau + valeur à appliquer au 1er <li> d’un <ol> rouvert (si on le fermait)
    ol_counters: dict[int, int] = {}
    ol_next_value: dict[int, int] = {}  # (reste utile si tu relies cette fonction avec d'autres)

    # indice contextuel : après un titre ou un paragraphe hors-liste qui finit par ':',
    # le prochain <ol> top-niveau doit redémarrer à 1
    restart_ol_hint: bool = False

    def _cleanup_counters(depth: int) -> None:
        for lvl in list(ol_counters.keys()):
            if lvl >= depth:
                del ol_counters[lvl]

    def _cleanup_nextvals(depth: int) -> None:
        for lvl in list(ol_next_value.keys()):
            if lvl >= depth:
                del ol_next_value[lvl]

    def _append_inside_last_li(html_fragment: str, as_cont: bool = False) -> bool:
        """
        Insère html_fragment juste avant </li> du dernier <li> émis.
        Si as_cont=True, on injecte en <p class="cont">…</p> (dé-décalage visuel via CSS).
        """
        frag = html_fragment
        if as_cont and frag.startswith("<p"):
            # ajoute class="cont" au premier <p ...>
            if frag.startswith("<p>"):
                frag = '<p class="cont">' + frag[3:]
            elif frag.startswith("<p "):
                frag = frag.replace("<p ", '<p class="cont" ', 1)
        for i in range(len(buf) - 1, -1, -1):
            s = buf[i]
            j = s.rfind("</li>")
            if j != -1:
                buf[i] = s[:j] + frag + s[j:]
                return True
        return False

    def _numref(p: Paragraph) -> tuple[str | None, int | None]:
        """Retourne (numId, ilvl) si liste Word native, sinon (None, None)."""
        try:
            pPr = getattr(p._p, "pPr", None)
            numPr = getattr(pPr, "numPr", None) if pPr is not None else None
            if numPr is None:
                return None, None
            numId_el = getattr(numPr, "numId", None)
            ilvl_el  = getattr(numPr, "ilvl", None)
            numId = str(getattr(numId_el, "val", None)) if numId_el is not None else None
            ilvl  = int(getattr(ilvl_el, "val", 0)) if ilvl_el is not None and getattr(ilvl_el, "val", None) is not None else 0
            return numId, ilvl
        except Exception:
            return None, None

    def flush() -> None:
        nonlocal buf, current, list_stack, ol_counters, ol_next_value, restart_ol_hint
        while list_stack:
            buf.append(f"</{list_stack.pop()['tag']}>")
        if current and buf:
            html = "".join(buf).strip()
            if html:
                sections[current] = (sections.get(current, "") + html)
        buf = []
        ol_counters.clear()
        ol_next_value.clear()
        restart_ol_hint = False

    for block in _iter_blocks(doc):

        # ---------- PARAGRAPHES ----------
        if isinstance(block, Paragraph):
            t = (block.text or "").strip()

            # Titre : nouvelle section
            if t and _looks_like_heading(t, block, exp):
                while list_stack:
                    buf.append(f"</{list_stack.pop()['tag']}>")
                flush()
                current = exp.get(_norm(t), exp.get(_norm(t.rstrip(":")), t))
                restart_ol_hint = True
                continue

            kind, level = _para_list_info(block, block.text or "")
            numId, _ilvl = _numref(block)
            manual_start1 = bool(re.match(r'^\s*1[.)]\s', block.text or ""))

            # UL juste après OL : force l’imbrication (Word renvoie parfois ilvl=0)
            if kind == "ul" and list_stack and list_stack[-1]["tag"] == "ol":
                if level is None or (level or 0) == 0:
                    level = len(list_stack)  # sous-niveau

            if kind is None:
                # Paragraphe normal : S'IL Y A UNE LISTE OUVERTE -> il appartient
                # au <li> courant (Word : texte explicatif d'un point), donc on N'EXPLLOSE PAS l'OL
                if list_stack:
                    _, p_html = _para_to_html(block)  # <p>…</p>
                    _append_inside_last_li(p_html, as_cont=True)  # p.cont = visuellement "plein gauche"
                else:
                    # pas de liste ouverte -> paragraphe standard
                    buf.append(_para_to_html(block)[1])
                    if t.endswith(":"):
                        restart_ol_hint = True
                continue

            # profondeur cible (0 => 1 liste ouverte)
            target_depth = (level or 0) + 1

            # Redémarrage après un titre / intro ':' (uniquement OL de niveau 0)
            section_restart = (kind == "ol" and restart_ol_hint and target_depth == 1)

            # Si une liste existe déjà à cette profondeur, vérifier si on redémarre l'OL :
            if target_depth <= len(list_stack):
                at_depth = list_stack[target_depth - 1]
                if kind == "ol":
                    different_num_id = (numId is not None and at_depth.get("numId") != numId)
                    # "1." manuel = restart seulement si on était déjà en OL
                    manual_restart   = (numId is None and manual_start1 and at_depth.get("tag") == "ol")
                    if section_restart or different_num_id or manual_restart:
                        # on ferme depuis ce niveau et on repart
                        while len(list_stack) >= target_depth:
                            buf.append(f"</{list_stack.pop()['tag']}>")
                        _cleanup_counters(target_depth - 1)
                        _cleanup_nextvals(target_depth - 1)

            # Réduction de profondeur
            while len(list_stack) > target_depth:
                buf.append(f"</{list_stack.pop()['tag']}>")
            _cleanup_counters(len(list_stack))
            _cleanup_nextvals(len(list_stack))

            # Ouverture jusqu'à la profondeur cible
            while len(list_stack) < target_depth:
                open_level = len(list_stack)
                to_open = kind if open_level + 1 == target_depth else "ul"
                if to_open == "ol":
                    # NOTE: on garde l’OL ouvert entre les paragraphes explicatifs,
                    # donc la continuité native fonctionne sans <li value="">
                    buf.append("<ol>")
                    list_stack.append({"tag": "ol", "numId": numId})
                else:
                    buf.append("<ul>")
                    list_stack.append({"tag": "ul", "numId": None})

            # Correction de type au niveau courant si besoin
            if list_stack and list_stack[-1]["tag"] != kind:
                buf.append(f"</{list_stack.pop()['tag']}>")
                open_level = len(list_stack)
                if kind == "ol":
                    buf.append("<ol>")
                    list_stack.append({"tag": "ol", "numId": numId})
                else:
                    buf.append("<ul>")
                    list_stack.append({"tag": "ul", "numId": None})

            # Ajoute l’item
            _, li_html = _para_to_html(block)
            buf.append(li_html)

            # Aligne numId si Word en fournit un
            if kind == "ol" and list_stack and list_stack[-1]["tag"] == "ol" and list_stack[-1].get("numId") is None and numId is not None:
                list_stack[-1]["numId"] = numId

            restart_ol_hint = False
            continue

        # ---------- TABLEAUX ----------
        if isinstance(block, Table):
            # On ferme les listes avant d’insérer le tableau
            while list_stack:
                buf.append(f"</{list_stack.pop()['tag']}>")

            rows_html: list[str] = []
            for row in block.rows:
                cells_html: list[str] = []
                for cell in row.cells:
                    cell_parts: list[str] = []
                    for pp in cell.paragraphs:
                        k, frag = _para_to_html(pp)
                        cell_parts.append(f"<ul>{frag}</ul>" if k.startswith("li") else frag)
                    cells_html.append(f"<td>{''.join(cell_parts) or '&nbsp;'}</td>")
                rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

            buf.append(
                "<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;width:100%'>"
                + "".join(rows_html)
                + "</table>"
            )
            restart_ol_hint = True  # souvent on repart sur une numé rotation distincte après un tableau
            continue

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
          .sect p { margin:.35rem 0; }
          .sect ol, .sect ul { margin: .4rem 0 .6rem 1.25rem; padding-left: 1.25rem; }
          .sect li { margin:.15rem 0; }
          /* quand on ferme la liste puis on revient à un paragraphe normal */
          .sect ol + p, .sect ul + p { margin-left: 0 !important; }
          /* table et images */
          .sect table { width:100%; border-collapse:collapse; }
          .sect table td, .sect table th { border:1px solid #ccc; padding:6px; }
        </style>
        """, unsafe_allow_html=True)


    for fdef in fields:
        key = fdef["key"]; label = fdef["label"]
        html_content = fr_payload.get(key, "")
        st.subheader(label)
        st.markdown(f"<div class='sect'>{html_content or '<p><em>(vide)</em></p>'}</div>", unsafe_allow_html=True)
        st.divider()


