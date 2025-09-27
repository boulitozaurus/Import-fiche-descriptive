# utils/docx_parser.py
from typing import Dict, List, Iterator, Union
from docx import Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
import base64

HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}

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

def _looks_like_heading(text: str, p: Paragraph, expected_map: Dict[str, str]) -> bool:
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

def _run_image_dataurl(run) -> str | None:
    try:
        ns = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        blips = run._r.xpath(".//a:blip/@r:embed", namespaces=ns)
        if not blips: return None
        rId = blips[0]
        part = run.part.related_parts[rId]
        content_type = getattr(part, "content_type", "image/png")
        b64 = base64.b64encode(part.blob).decode("ascii")
        return f"data:{content_type};base64,{b64}"
    except Exception:
        return None

def _run_to_html(run) -> str:
    # image
    dataurl = _run_image_dataurl(run)
    if dataurl:
        return f'<img src="{dataurl}" />'

    txt = _html_escape(run.text)
    if not txt: return ""

    open_tags, close_tags = "", ""
    color = getattr(getattr(run.font, "color", None), "rgb", None)
    if color:
        open_tags += f'<span style="color:#{str(color)}">'
        close_tags = "</span>" + close_tags
    if getattr(run, "underline", False):
        open_tags += "<u>"; close_tags = "</u>" + close_tags
    if getattr(run, "italic", False):
        open_tags += "<em>"; close_tags = "</em>" + close_tags
    if getattr(run, "bold", False):
        open_tags += "<strong>"; close_tags = "</strong>" + close_tags
    return f"{open_tags}{txt}{close_tags}"

def _para_list_kind(p: Paragraph, text: str) -> str | None:
    """Retourne 'ul', 'ol' ou None."""
    # numPr explicite
    pPr = getattr(p._p, "pPr", None)
    if getattr(pPr, "numPr", None) is not None:
        # approximation: 'ol' si le style contient 'Number' ou si le texte commence par un motif numéroté
        sname = (p.style.name if getattr(p, "style", None) else "") or ""
        if "Number" in sname:
            return "ol"
        # motif ex: "1) ", "2. ", "1.1 " présent dans run (parfois Word ne met pas le chiffre dans le texte)
        if text and (text[:3].strip().rstrip(".)").isdigit()):
            return "ol"
        return "ul"

    # styles usuels de listes
    sname = (p.style.name if getattr(p, "style", None) else "") or ""
    if any(k in sname for k in ["List", "Puces", "Bullet"]):
        return "ul"
    if "Number" in sname:
        return "ol"

    # heuristique : symbole en début
    start = (text or "").lstrip()
    if start.startswith(("•","◦","▪","-","–","—","*")):
        return "ul"
    return None

def _para_to_html(p: Paragraph) -> tuple[str, str]:
    inner_runs = "".join(_run_to_html(r) for r in p.runs)
    inner = inner_runs or _html_escape(p.text or "")
    # listes ?
    kind = _para_list_kind(p, p.text or "")
    if kind == "ol":
        return ("li-ol", f"<li>{inner}</li>")
    if kind == "ul":
        # enlève le bullet de texte si présent
        for b in ("•","◦","▪","-","–","—","*"):
            if inner.startswith(b):
                inner = inner[len(b):].lstrip()
                break
        return ("li-ul", f"<li>{inner}</li>")
    return ("p", f"<p>{inner}</p>")

def iter_block_items(parent) -> Iterator[Union[Paragraph, Table]]:
    """Parcourt Paragraph/Table dans l'ordre d'apparition."""
    if isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent._element
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def parse_docx_sections(path, expected_headings: List[str] = None) -> Dict[str, str]:
    """Retourne {heading: HTML} avec listes/tableaux/images/formatage."""
    doc = Document(path)
    expected_map = {_norm(h): h for h in (expected_headings or [])}
    expected_map.update({_norm(h.rstrip(":")): h for h in (expected_headings or [])})

    sections: Dict[str, str] = {}
    current = None
    html_chunks: List[str] = []
    in_list = False
    list_kind = None  # 'ul'/'ol'

    def flush():
        nonlocal html_chunks, in_list, list_kind, current
        if in_list:
            html_chunks.append(f"</{list_kind}>")
            in_list = False; list_kind = None
        if current and html_chunks:
            html = "".join(html_chunks).strip()
            if html:
                sections[current] = (sections.get(current,"") + html)
        html_chunks = []

    # Parcours ordonné
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = (block.text or "").strip()
            if t and _looks_like_heading(t, block, expected_map):
                flush()
                current = expected_map.get(_norm(t), expected_map.get(_norm(t.rstrip(":")), t))
                continue
            kind, frag = _para_to_html(block)
            if kind == "p":
                if in_list:
                    html_chunks.append(f"</{list_kind}>")
                    in_list = False; list_kind = None
                html_chunks.append(frag)
            else:
                target = "ol" if kind == "li-ol" else "ul"
                if not in_list or list_kind != target:
                    if in_list:
                        html_chunks.append(f"</{list_kind}>")
                    html_chunks.append(f"<{target}>")
                    in_list = True; list_kind = target
                html_chunks.append(frag)

        else:  # Table
            # ferme la liste si ouverte
            if in_list:
                html_chunks.append(f"</{list_kind}>")
                in_list = False; list_kind = None
            rows = []
            for row in block.rows:
                tds = []
                for cell in row.cells:
                    cell_parts = []
                    for pp in cell.paragraphs:
                        k, frag = _para_to_html(pp)
                        if k.startswith("li"):
                            cell_parts.append(f"<ul>{frag}</ul>")
                        else:
                            cell_parts.append(frag)
                    tds.append(f"<td>{''.join(cell_parts) or '&nbsp;'}</td>")
                rows.append(f"<tr>{''.join(tds)}</tr>")
            html_chunks.append(
                f"<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;width:100%'>"
                + "".join(rows) + "</table>"
            )

    flush()
    return sections
