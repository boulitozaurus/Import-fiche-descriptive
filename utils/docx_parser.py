# utils/docx_parser.py
from typing import Dict, List
from docx import Document
import base64

HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}

def _strip_accents(x: str) -> str:
    if x is None:
        return ""
    try:
        import unicodedata  # import local
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
    # 1) titre exact attendu
    if _norm(t) in expected_map or _norm(t.rstrip(":")) in expected_map:
        return True
    # 2) petit paragraphe en style Heading, sans ponctuation de phrase
    if _is_heading_style(paragraph):
        if len(t) <= 80 and t.count(" ") <= 11 and all(p not in t for p in [".","!","?"]):
            return True
    return False

def _html_escape(s: str) -> str:
    return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def _run_image_dataurl(run) -> str | None:
    """Retourne une data-URI si le run contient une image, sinon None."""
    try:
        # run._r est un élément lxml ; on récupère l'id d'embed
        ns = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        }
        blips = run._r.xpath(".//a:blip/@r:embed", namespaces=ns)
        if not blips:
            return None
        rId = blips[0]
        image_part = run.part.related_parts[rId]  # python-docx
        content_type = getattr(image_part, "content_type", "image/png")
        blob = image_part.blob  # bytes
        b64 = base64.b64encode(blob).decode("ascii")
        return f"data:{content_type};base64,{b64}"
    except Exception:
        return None

def _run_to_html(run) -> str:
    # image d'abord
    dataurl = _run_image_dataurl(run)
    if dataurl:
        return f'<img src="{dataurl}" />'

    txt = _html_escape(run.text)
    if not txt:
        return ""

    # styles inline
    open_tags, close_tags = "", ""
    color = getattr(getattr(run.font, "color", None), "rgb", None)
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

def _para_to_html(p) -> tuple[str, str]:
    inner = "".join(_run_to_html(r) for r in p.runs) or _html_escape(p.text)
    # détection des listes
    pPr = getattr(p._p, "pPr", None)
    has_num = (getattr(pPr, "numPr", None) is not None)
    style = p.style.name if getattr(p, "style", None) else ""
    is_number = "Number" in style
    is_bullet = ("Bullet" in style) or (has_num and not is_number)

    if has_num and is_number:
        return ("li-ol", f"<li>{inner}</li>")
    elif has_num or is_bullet:
        return ("li-ul", f"<li>{inner}</li>")
    else:
        return ("p", f"<p>{inner}</p>")

def parse_docx_sections(path, expected_headings: List[str] = None) -> Dict[str, str]:
    """
    Retourne {heading: HTML} :
      - titres reconnus via expected_headings (insensible casse/accents/':') ou vrai style Heading court,
      - listes <ul>/<ol>,
      - tableaux <table>,
      - images <img src="data:...">,
      - gras/italique/souligné/couleurs conservés.
    """
    doc = Document(path)
    expected_map = {_norm(h): h for h in (expected_headings or [])}
    # ajoute aussi la variante avec ':' final
    expected_map.update({_norm(h.rstrip(":")): h for h in (expected_headings or [])})

    sections: Dict[str, str] = {}
    current = None
    html_chunks: List[str] = []
    in_list = False
    list_kind = None  # "ul" / "ol"

    def flush():
        nonlocal current, html_chunks, in_list, list_kind
        if in_list:
            html_chunks.append(f"</{list_kind}>")
            in_list = False
            list_kind = None
        if current and html_chunks:
            html = "".join(html_chunks).strip()
            if html:
                sections[current] = (sections.get(current,"") + html)
        html_chunks = []

    # Parcours des paragraphes
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t and _looks_like_heading(t, p, expected_map):
            flush()
            current = expected_map.get(_norm(t), expected_map.get(_norm(t.rstrip(":")), t))
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

    # Tables -> ajout sous le dernier heading détecté
    last_heading = None
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t and _looks_like_heading(t, p, expected_map):
            last_heading = expected_map.get(_norm(t), expected_map.get(_norm(t.rstrip(":")), t))

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
            tbls.append(f"<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse'>{''.join(rows_html)}</table>")
        if tbls:
            target = last_heading or "TABLES"
            sections[target] = (sections.get(target,"") + "".join(f"<div>{t}</div>" for t in tbls))

    return sections
