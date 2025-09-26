
from typing import Dict, List, Tuple
from docx import Document

HEADING_STYLES = {"Heading 1","Heading 2","Heading 3","Titre 1","Titre 2","Titre 3","Title","Subtitle"}

def _is_heading(p) -> bool:
    s = p.style.name if p.style else ""
    return (s in HEADING_STYLES) or s.startswith("Heading")

def parse_docx_sections(path) -> Dict[str, str]:
    """
    Return {heading: text} merging paragraphs under each heading.
    Also appends any tables found under the last seen heading as a simple Markdown table.
    """
    doc = Document(path)
    sections: Dict[str, str] = {}
    current = None
    buff: List[str] = []

    def flush():
        nonlocal current, buff
        if current and buff:
            text = "\n\n".join([x for x in buff if x.strip()])
            if text.strip():
                if current in sections and sections[current].strip():
                    sections[current] = sections[current].strip() + "\n\n" + text.strip()
                else:
                    sections[current] = text.strip()
        buff = []

    # paragraphs
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if _is_heading(p) and t:
            flush()
            current = t
        else:
            if t:
                buff.append(t)
    flush()

    # tables -> append to the last heading
    last_heading = None
    for p in doc.paragraphs:
        if _is_heading(p) and p.text.strip():
            last_heading = p.text.strip()

    if doc.tables:
        if not last_heading:
            last_heading = "TABLES"
        # concatenate all tables as markdown under the last heading (or create if missing)
        md_chunks = []
        for table in doc.tables:
            rows = []
            for row in table.rows:
                rows.append([c.text.strip() for c in row.cells])
            md = "\n".join("| " + " | ".join(r) + " |" for r in rows if any(x for x in r))
            if md.strip():
                md_chunks.append(md)
        if md_chunks:
            merged = "\n\n".join(md_chunks)
            if last_heading in sections and sections[last_heading].strip():
                sections[last_heading] = sections[last_heading].strip() + "\n\n" + merged
            else:
                sections[last_heading] = merged

    return sections
