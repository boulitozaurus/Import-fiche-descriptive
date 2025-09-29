# -*- coding: utf-8 -*-
"""
Import fiche descriptive (.docx) -> HTML + nettoyage + forçages stricts.

Points clés / changements par rapport aux versions précédentes :
- Conversion unique via Mammoth -> HTML (plus de "python-docx").
- AUCUNE suppression des <ol>/<ul> au début des post-traitements (sinon
  on perdait la liste et le "forçage" ne trouvait plus rien).
- "Forçages stricts" :
  * Facteurs de risque : 3 items dans l'ordre : projet, secteur, défaut.
  * Bonnes raisons : #1 = assurance si présente, sinon fiducie ; #2 = fiducie si assurance présente.
  * Budget : titres fixes mis en <em><u>…</u></em>.
- Nettoyage des "img alt=" affichés en texte (on supprime les chaînes parasites
  qui ressemblent à un tag <img alt="…"> mal converti).
- Téléchargements EMF/WMF par section (non pris en charge par le navigateur).
- Code allégé : imports ou fonctions docx inutiles supprimés.
"""

import base64
import re
from typing import Dict, List, Tuple, Optional

import mammoth
import streamlit as st
from bs4 import BeautifulSoup, Tag, NavigableString


# -----------------------------------------------------------------------------
#                               CONFIG / CONSTANTES
# -----------------------------------------------------------------------------

st.set_page_config(page_title="Import fiche descriptive", layout="wide")

CSS = """
<style>
/* Rendre le contenu plus lisible */
section.main > div { max-width: 1400px; }

/* Titres de section */
h2,h3 { margin: 0.6rem 0 0.35rem; }
hr { margin: 1rem 0; }

/* Paragraphes/listes */
p { margin: 0.25rem 0 0.45rem; line-height: 1.45; }
li { margin: 0.15rem 0; line-height: 1.45; }

/* Boîte section */
.card {
  border: 1px solid #e8e8e8;
  border-radius: 8px;
  padding: 12px 14px;
  margin: 12px 0 20px;
  background: #fff;
}
.card h3 {
  margin: 0 0 8px;
}

/* Lien de retour (ancre) */
.anchor {
  margin-left: 6px;
  font-size: 0.9rem;
  opacity: 0.5;
}

/* Placeholders d'images EMF */
.img-placeholder {
  display: inline-block;
  padding: 6px 8px;
  margin: 4px 0;
  background: #fff7e6;
  border: 1px dashed #f0ad4e;
  color: #9a6c00;
  border-radius: 5px;
  font-size: 0.9rem;
}

/* Table mapping */
table thead tr th {
  background: #fafafa;
}

/* Numérotation (on laisse le navigateur faire) */
ol { margin-left: 20px; }
ul { margin-left: 20px; }

/* Texte trop grand provenant de Word (title/strong)  */
h1 { font-size: 1.15rem; font-weight: 600; }
h2 { font-size: 1.10rem; font-weight: 600; }
h3 { font-size: 1.05rem; font-weight: 600; }
strong { font-weight: 600; }
</style>
"""

# Headings attendus côté CRM (à adapter si besoin)
EXPECTED_HEADINGS = [
    "Description",
    "Contexte et usage des fonds",
    "Les points d'attention",
    "Les bonnes raisons d'investir",
    "Présentation de l'opération",
    "Localisation",
    "Planning",
    "Marché et références",
    "Budget",
    "Présentation de l'opérateur",
    "Actionnariat",
    "Finances",
]

# Mappage CRM (à ajuster à votre système)
CRM_KEY_BY_LABEL = {
    "Description": "description_fr",
    "Contexte et usage des fonds": "contexte_fonds_fr",
    "Les points d'attention": "points_attention_fr",
    "Les bonnes raisons d'investir": "bonnes_raisons_fr",
    "Présentation de l'opération": "operation_presentation_fr",
    "Localisation": "localisation_fr",
    "Planning": "planning_fr",
    "Marché et références": "marche_references_fr",
    "Budget": "budget_fr",
    "Présentation de l'opérateur": "operateur_presentation_fr",
    "Actionnariat": "actionnariat_fr",
    "Finances": "finances_fr",
}

# Forçage des labels (détection textuelle, insensibilité à la casse)
RISKS_LABELS = [
    "risque lié au projet",
    "risque lié au secteur",
    "risque de défaut",
]

BONNES_RAISONS_ASSURANCE = "une assurance sur 100% du capital investi"
BONNES_RAISONS_FIDUCIE = "une fiducie-sûreté sur l'actif"

BUDGET_TITLES = [
    "prix de revient",
    "financement et ratios",
    "revenus et marges",
    "couverture des intérêts",
    "stress test",
]

# -----------------------------------------------------------------------------
#                           OUTILS de parsing / nettoyage
# -----------------------------------------------------------------------------

def docx_to_html(file_bytes: bytes) -> str:
    """Convertit un DOCX en HTML via Mammoth (avec styles par défaut)."""
    # style_map minimal : on laisse Mammoth générer h1..hN, p, ul/ol/li...
    result = mammoth.convert_to_html(
        file_bytes,
        style_map="",
        include_default_style_map=True,
        convert_image=mammoth.images.inline(lambda image: {"src": _img_to_data_uri(image)}),
    )
    html = result.value or ""
    # Aucun traitement "magique" sur les images ici : on renvoie tel quel.
    return html


def _img_to_data_uri(image: mammoth.images.Image) -> str:
    """Image Mammoth -> data URI (png/jpeg/gif, sinon base64 générique)."""
    content_type = image.content_type or "application/octet-stream"
    data = image.read()
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{content_type};base64,{b64}"


def strip_leading_numbering(txt: str) -> str:
    """Supprime les préfixes de numérotation variés (1., 2), I., a), etc.)."""
    if not txt:
        return ""
    # Prefixes : chiffres/lettres/romains + ., ), - etc.
    return re.sub(r"^\s*([0-9ivxlcdmIVXLCDM]+[\.\)\-:])\s*", "", txt).strip()


def _drop_stray_img_alt_strings(soup: BeautifulSoup) -> None:
    """
    Certains flux cassés affichent une chaîne comme '<img alt="...">' en texte.
    On supprime ces NavigableString qui ressemblent à un tag img littéral.
    """
    pattern = re.compile(r"^\s*<img\s+alt=.*?>\s*/?>\s*$", re.I)
    for node in list(soup.find_all(text=True)):
        if isinstance(node, NavigableString) and pattern.match(str(node)):
            node.extract()


def _merge_adjacent_ol(soup: BeautifulSoup) -> None:
    """
    Fusionne des <ol> consécutifs au même niveau.
    Approche simple et robuste pour éviter les 1..16 quand Word segmente.
    """
    changed = True
    while changed:
        changed = False
        for ol in soup.find_all("ol"):
            nxt = ol.find_next_sibling()
            if nxt and nxt.name == "ol":
                # déplacer les <li> du suivant dans l'actuel
                for li in list(nxt.find_all("li", recursive=False)):
                    ol.append(li)
                nxt.decompose()
                changed = True
                break


def prepare_section_html(raw_html: str) -> Tuple[str, List[Tuple[str, bytes]]]:
    """
    Nettoyage générique d'un bloc HTML (une section).
    Retourne (html nettoyé, fichiers à télécharger dans la section).
    """
    soup = BeautifulSoup(raw_html, "html.parser")

    # Supprimer les chaînes "img alt=" laissées comme du texte
    _drop_stray_img_alt_strings(soup)

    # Fusion simple des <ol> contigus
    _merge_adjacent_ol(soup)

    # Gérer les images non affichables (EMF/WMF) : on remplace par placeholder
    # + on offre un download_button.
    downloads: List[Tuple[str, bytes]] = []
    for img in list(soup.find_all("img")):
        src = img.get("src", "")
        if src.startswith("data:image/"):
            # mimetype
            m = re.match(r"^data:(image/[^;]+);base64,(.*)$", src, re.I)
            if m:
                mime = m.group(1).lower()
                b64 = m.group(2)
                if mime in ("image/emf", "image/x-emf", "image/wmf", "image/x-wmf"):
                    try:
                        blob = base64.b64decode(b64)
                        fname = f"{mime.split('/')[-1]}_image.emf"
                        downloads.append((fname, blob))
                        # Placeholder dans le HTML
                        ph = soup.new_tag("span", **{"class": "img-placeholder"})
                        ph.string = f"[Image {mime} non affichable - disponible en téléchargement]"
                        img.replace_with(ph)
                    except Exception:
                        # si décodage échoue, on laisse l'image telle quelle
                        pass

    return str(soup), downloads


# -----------------------------------------------------------------------------
#                      Post-traitement “forçage strict” par section
# -----------------------------------------------------------------------------

def _normalize_text(text: str) -> str:
    """Pour matcher les libellés, neutraliser ponctuation/casse/espaces."""
    if not text:
        return ""
    t = re.sub(r"\s+", " ", text).strip().lower()
    return t


def _li_starts_with(li: Tag, label: str) -> bool:
    """True si un <li> commence (après numéro éventuel) par label."""
    txt = strip_leading_numbering(li.get_text(" ", strip=True))
    return _normalize_text(txt).startswith(_normalize_text(label))


def _find_items_by_labels_in_soup(soup: BeautifulSoup, labels: List[str]) -> Dict[str, Tag]:
    """
    Cherche dans TOUTES les <li> d'une section, renvoie un dict label-><li> (première occurrence).
    On ne touche jamais aux listes à ce stade (pas de decompose()).
    """
    mapping: Dict[str, Tag] = {}
    for li in soup.find_all("li"):
        for lab in labels:
            if lab not in mapping and _li_starts_with(li, lab):
                mapping[lab] = li
    return mapping


def _wrap_whole_tag_with_em_u(tag: Tag) -> None:
    """Transforme tout le contenu du tag en <em><u>…</u></em> (robuste)."""
    if not isinstance(tag, Tag) or not hasattr(tag, "new_tag"):
        return
    children = list(tag.contents)
    tag.clear()
    em = tag.new_tag("em")
    u = tag.new_tag("u")
    tag.append(em)
    em.append(u)
    for c in children:
        u.append(c)


def _postprocess_risks(html: str) -> str:
    """
    Forçage strict – Facteurs de risque :
    1. Risque lié au projet
    2. Risque lié au secteur
    3. Risque de défaut
    """
    soup = BeautifulSoup(html, "html.parser")

    # Récupère la première <ol> existante si possible, sinon on créera une neuve.
    all_ol = soup.find_all("ol")
    container = all_ol[0] if all_ol else None

    found = _find_items_by_labels_in_soup(soup, RISKS_LABELS)
    # Construit un <ol> propre avec les items dans l'ordre strict
    new_ol = soup.new_tag("ol")
    for lab in RISKS_LABELS:
        if lab in found:
            new_ol.append(found[lab].extract())

    # Insère new_ol à la place de la première <ol>, sinon en fin de bloc
    if container:
        container.replace_with(new_ol)
    else:
        soup.append(new_ol)

    return str(soup)


def _postprocess_bonnes_raisons(html: str) -> str:
    """
    Forçage strict – Bonnes raisons :
    - #1 = "Une assurance..." si présente, sinon "Une fiducie-sûreté..."
    - #2 = "Une fiducie-sûreté..." si "assurance" était présente
    """
    soup = BeautifulSoup(html, "html.parser")

    all_ol = soup.find_all("ol")
    container = all_ol[0] if all_ol else None

    wanted_order: List[str] = []
    # Priorité #1
    labels_pool = [BONNES_RAISONS_ASSURANCE, BONNES_RAISONS_FIDUCIE]
    found = _find_items_by_labels_in_soup(soup, labels_pool)

    if BONNES_RAISONS_ASSURANCE in found:
        wanted_order = [BONNES_RAISONS_ASSURANCE, BONNES_RAISONS_FIDUCIE]
    else:
        # Assurance absente => Fiducie devient #1
        if BONNES_RAISONS_FIDUCIE in found:
            wanted_order = [BONNES_RAISONS_FIDUCIE]
        else:
            # Aucun des deux → ne change rien
            return str(soup)

    new_ol = soup.new_tag("ol")
    for lab in wanted_order:
        if lab in found:
            new_ol.append(found[lab].extract())

    if container:
        container.replace_with(new_ol)
    else:
        soup.append(new_ol)

    return str(soup)


def _postprocess_budget(html: str) -> str:
    """
    Met en italique + souligné tous les titres fixes du budget, s’ils existent :
      1. Prix de revient
      2. Financement et ratios
      3. Revenus et marges
      4. Couverture des intérêts
      5. Stress test
    """
    soup = BeautifulSoup(html, "html.parser")
    # Chercher p/li/strong/em … on stylise le bloc qui commence par le titre.
    # Ici on applique un wrap <em><u>…</u></em> sur TOUT le tag parent.
    targets = soup.find_all(["p", "li", "h1", "h2", "h3", "strong"])
    for t in targets:
        txt = strip_leading_numbering(t.get_text(" ", strip=True))
        nrm = _normalize_text(txt)
        for title in BUDGET_TITLES:
            if nrm.startswith(_normalize_text(title)):
                _wrap_whole_tag_with_em_u(t)
                break
    return str(soup)


def postprocess_domain_section(section_label: str, clean_html: str) -> str:
    """
    Applique, si nécessaire, un forçage strict propre à une section.
    """
    label = (section_label or "").strip().lower()

    if label == _normalize_text("Les points d'attention"):
        return _postprocess_risks(clean_html)

    if label == _normalize_text("Les bonnes raisons d'investir"):
        return _postprocess_bonnes_raisons(clean_html)

    if label == _normalize_text("Budget"):
        return _postprocess_budget(clean_html)

    # Par défaut : inchangé
    return clean_html


# -----------------------------------------------------------------------------
#                         SEGMENTATION DU DOC EN SECTIONS
# -----------------------------------------------------------------------------

def split_into_sections(full_html: str) -> Dict[str, str]:
    """
    Découpe le HTML global en sections, en s'appuyant sur les h1..h3.
    S'il manque un titre attendu, la section restera vide.
    """
    soup = BeautifulSoup(full_html, "html.parser")
    sections: Dict[str, str] = {}

    # On récupère tous les blocs de niveau "titre" (h1..h4), simple et robuste.
    headers = soup.find_all(re.compile(r"^h[1-4]$", re.I))
    if not headers:
        # Aucun titre → tout dans "Description" (fallback)
        sections["Description"] = str(soup)
        return sections

    def header_text(h: Tag) -> str:
        return h.get_text(" ", strip=True)

    # On scanne, chaque h* ouvre une nouvelle section.
    for i, h in enumerate(headers):
        title = header_text(h)
        # Trouver borne fin
        content_nodes: List[Tag] = []
        cur = h.next_sibling
        while cur and cur not in headers:
            if isinstance(cur, Tag):
                content_nodes.append(cur)
            cur = cur.next_sibling

        # Ajout au dict (si le titre exact fait partie de l'attendu, sinon on garde le texte tel quel)
        sections[title] = "".join(str(n) for n in content_nodes)

    return sections


# -----------------------------------------------------------------------------
#                                      UI
# -----------------------------------------------------------------------------

def main():
    st.markdown(CSS, unsafe_allow_html=True)
    st.title("Import fiche descriptive (.docx)")

    file = st.file_uploader("Choisir un fichier Word (.docx)", type=["docx"])
    if not file:
        st.info("Charge un .docx pour commencer.")
        return

    # Conversion
    with st.spinner("Conversion Word → HTML…"):
        file_bytes = file.read()
        full_html = docx_to_html(file_bytes)

    # Découpe par sections
    raw_sections = split_into_sections(full_html)

    # Mapping auto (détection brute -> heading attendu)
    detected = list(raw_sections.keys())
    rows = []
    for expected in EXPECTED_HEADINGS:
        present = ""
        matched_title = ""
        for d in detected:
            if _normalize_text(d) == _normalize_text(expected):
                present = "Oui"
                matched_title = d
                break
        rows.append((expected, present or "Non", matched_title, CRM_KEY_BY_LABEL.get(expected, "")))

    st.subheader("Résultat du mapping automatique")
    st.table({"Word heading attendu": [r[0] for r in rows],
              "Dans le .docx ?": [r[1] for r in rows],
              "PDF/CRM heading": [r[2] for r in rows],
              "CRM key": [r[3] for r in rows]})

    st.subheader("Aperçu des sections (mise en forme préservée)")

    # Affichage section par section (dans l'ordre attendu)
    for expected in EXPECTED_HEADINGS:
        # On cherche dans le splitted (exact match) sinon fallback label approx
        the_key = None
        for k in raw_sections.keys():
            if _normalize_text(k) == _normalize_text(expected):
                the_key = k
                break

        raw = raw_sections.get(the_key or "", "")
        with st.container():
            st.markdown(f"### {expected} <span class='anchor'>↪</span>", unsafe_allow_html=True)

            if not raw.strip():
                st.caption("(vide)")
                st.markdown("<hr/>", unsafe_allow_html=True)
                continue

            # Nettoyage de base (pas de suppression des listes ici !)
            clean_html, downloads = prepare_section_html(raw)

            # Forçage strict (risques, bonnes raisons, budget)
            clean_html = postprocess_domain_section(expected, clean_html)

            # Affichage HTML
            st.markdown(f"<div class='card'>{clean_html}</div>", unsafe_allow_html=True)

            # Downloads EMF/WMF de la section
            if downloads:
                st.info("Certaines images (EMF/WMF) ne peuvent pas s’afficher dans le navigateur. Télécharge-les :")
                cols = st.columns(max(1, min(3, len(downloads))))
                for i, (fname, blob) in enumerate(downloads):
                    with cols[i % len(cols)]:
                        st.download_button(
                            label=f"Télécharger {fname}",
                            data=blob,
                            file_name=fname,
                            mime="application/octet-stream",
                        )

            st.markdown("<hr/>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
