from typing import Dict, List
from docx import Document

HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}

def _strip_accents(x: str) -> str:
    if x is None:
        return ""
    try:
        import unicodedata  # import local, pas de dépendance externe
        nfkd = unicodedata.normalize("NFKD", x)
        return "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    except Exception:
        return x

def _norm(s: str) -> str:
    return " ".join(_strip_accents((s or "")).lower().replace("’", "'").split())

def _is_heading_style(p) -> bool:
    s = p.style.name if getattr(p, "style", None) else ""
    return (s in HEADING_STYLES) or s.startswith("Heading")

def _looks_like_heading(text: str, paragraph, expected_map: Dict[str, str]) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    # 1) si le texte correspond à un titre attendu -> heading
    if _norm(t) in expected_map:
        return True
    # 2) sinon, si style "Heading" court et sans ponctuation -> heading
    if _is_heading_style(paragraph):
        if len(t) <= 80 and t.count(" ") <= 11 and all(p not in t for p in [".","!","?"]):
            return True
    return False

def _html_escape(s: str) -> str:
    return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def _run_to_html(run) -> str:
    txt = _html_escape(run.text)
    if not txt:
        return ""
    # couleur
    color = getattr(getattr(run.font, "color", None), "rgb", None)
    open_tags, close_tags = "", ""
    if color:
        open_tags += f'<span style="color:#{str(color)}">'
        close_tags = "</span>" + close_tags
    if getattr(run, "underline", False):
        open_tags += "<u>"
        close_tags = "</u>" + close_tags
    if getattr(run, "italic", False):
        open_tags += "<em>"
        close_tags = "</em>" + close_tags
    if getattr(run, "bold", False):
        open_tags += "<strong>"
        close_tags = "</strong>" + close_tags
    return f"{open_tags}{txt}{close_tags}"

def _para_to_html(p) -> str:
    inner = "".join(_run_to_html(r) for r in p.runs) or _html_escape(p.text)
    style = p.style.name if getattr(p, "style", None) else ""
    # Détection liste
    has_num = getattr(getattr(getattr(p._p, "pPr", None), "numPr", None), "numId", None) is not None
    is_bullet = ("Bullet" in style) or (has_num and "Number" not in style)
    is_number = ("Number" in style)
    if has_num and is_number:
        return ("li-ol", f"<li>{inner}</li>")
    elif has_num or is_bullet:
        return ("li-ul", f"<li>{inner}</li>")
    else:
        return ("p", f"<p>{inner}</p>")

def parse_docx_sections(path, expected_headings: List[str] = None) -> Dict[str, str]:
    """
    Retourne {heading: HTML}.
    - Titre détecté s'il est dans expected_headings (insensible casse/accents) OU si paragraphe Heading court sans ponctuation.
    - Listes préservées (<ul>/<ol>), emphases (gras/italique/souligné), couleurs.
    - Tableaux convertis en <table>.
    """
    doc = Document(path)
    expected_map = {_norm(h): h for h in (expected_headings or [])}

    sections: Dict[str, str] = {}
    current = None
    html_chunks: List[str] = []
    in_list = False
    list_kind = None  # "ul" / "ol"

    def flush():
        nonlocal current, html_chunks, in_list, list_kind
        # ferme liste si ouverte
        if in_list:
            html_chunks.append(f"</{list_kind}>")
            in_list = False
            list_kind = None
        if current and html_chunks:
            html = "".join(html_chunks).strip()
            if html:
                sections[current] = (sections.get(current,"") + ("" if not sections.get(current) else "") + html)
        html_chunks = []

    # Pass 1: paragraphs
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        if _looks_like_heading(t, p, expected_map):
            flush()
            current = expected_map.get(_norm(t), t)
            continue

        kind, frag = _para_to_html(p)
        if kind == "p":
            if in_list:
                html_chunks.append(f"</{list_kind}>")
                in_list = False
                list_kind = None
            html_chunks.append(frag)
        else:
            target_kind = "ol" if kind == "li-ol" else "ul"
            if not in_list or list_kind != target_kind:
                if in_list:
                    html_chunks.append(f"</{list_kind}>")
                html_chunks.append(f"<{target_kind}>")
                in_list = True
                list_kind = target_kind
            html_chunks.append(frag)

    flush()

    # Pass 2: tables -> append sous le dernier heading détecté
    last_heading = None
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t and _looks_like_heading(t, p, expected_map):
            last_heading = expected_map.get(_norm(t), t)

    def cell_to_html(cell) -> str:
        parts = []
        for pp in cell.paragraphs:
            k, frag = _para_to_html(pp)
            if k.startswith("li"):
                parts.append(f"<ul>{frag}</ul>")
            else:
                parts.append(frag)
        return "".join(parts) or "&nbsp;"

    if doc.tables:
        tbls = []
        for tb in doc.tables:
            rows_html = []
            for row in tb.rows:
                cells_html = "".join(f"<td>{cell_to_html(c)}</td>" for c in row.cells)
                rows_html.append(f"<tr>{cells_html}</tr>")
            tbls.append(f"<table border='1' cellspacing='0' cellpadding='4'>{''.join(rows_html)}</table>")
        if tbls:
            target = last_heading or "TABLES"
            sections[target] = (sections.get(target,"") + ("<br/>" if sections.get(target) else "") + "<br/>".join(tbls))

    return sections
