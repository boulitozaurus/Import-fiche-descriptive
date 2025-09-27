from typing import Dict, List
from docx import Document

def _strip_accents(x: str) -> str:
    if x is None:
        return ""
    try:
        import unicodedata  # import local, safe
        nfkd = unicodedata.normalize("NFKD", x)
        return "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    except Exception:
        # fallback: on renvoie tel quel (les titres restent comparables en exact match)
        return x

def _norm(s: str) -> str:
    return " ".join(_strip_accents((s or "")).lower().replace("’", "'").split())


HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}

def _strip_accents(x: str) -> str:
    if x is None:
        return ""
    nfkd = unicodedata.normalize("NFKD", x)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))

def _norm(s: str) -> str:
    return " ".join(_strip_accents((s or "")).lower().replace("’", "'").split())

def _is_heading_style(p) -> bool:
    s = p.style.name if getattr(p, "style", None) else ""
    return (s in HEADING_STYLES) or s.startswith("Heading")

def _looks_like_heading(text: str, paragraph, expected_map: Dict[str, str]) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    # titre exact attendu -> heading
    if _norm(t) in expected_map:
        return True
    # sinon, style Heading court et sans ponctuation de phrase -> heading
    if _is_heading_style(paragraph):
        if len(t) <= 80 and t.count(" ") <= 11 and all(p not in t for p in [".", "!", "?"]):
            return True
    return False

def parse_docx_sections(path, expected_headings: List[str] = None) -> Dict[str, str]:
    """
    Retourne {heading: texte}.
    Détection d'un titre si:
      - il appartient à expected_headings (insensible casse/accents), OU
      - c'est un style Heading court sans ponctuation.
    Les tableaux sont concaténés en Markdown sous le dernier heading vu.
    """
    doc = Document(path)
    expected_map = {_norm(h): h for h in (expected_headings or [])}

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

    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        if _looks_like_heading(t, p, expected_map):
            flush()
            current = expected_map.get(_norm(t), t)  # étiquette canonique si liste attendue
        else:
            buff.append(t)
    flush()

    # Tables -> ajout sous le dernier heading détecté
    last_heading = None
    for p in doc.paragraphs:
        if _looks_like_heading((p.text or "").strip(), p, expected_map):
            last_heading = expected_map.get(_norm(p.text.strip()), p.text.strip())

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
