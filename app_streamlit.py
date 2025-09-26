
import os
import io
import json
import yaml
import streamlit as st
from pathlib import Path
from typing import Dict, List
from utils.docx_parser import parse_docx_sections
import requests
try:
    from Unidecode import Unidecode  # si installé, on l'utilise
except Exception:
    import unicodedata
    def unidecode(x: str) -> str:
        if x is None:
            return ""
        # suppression des accents via NFKD (approximation d'Unidecode)
        nfkd = unicodedata.normalize("NFKD", x)
        return "".join(ch for ch in nfkd if not unicodedata.combining(ch))

# -------------- Config --------------
st.set_page_config(page_title="Prototype Mapping Word -> CRM + NL", layout="wide")

DEFAULT_SCHEMA_PATH = Path("crm_schema.yaml")
DEFAULT_MAPPING_PATH = Path("sample_mapping.yaml")

# ------------------ Sidebar: Traduction ---------------
st.sidebar.header("Traduction")
provider = st.sidebar.selectbox(
    "Moteur",
    ["Local (Argos – offline)", "LibreTranslate (HTTP)", "MyMemory (HTTP)"],
    index=0
)
lt_endpoint = st.sidebar.text_input("LibreTranslate endpoint", "http://localhost:5000/translate")

def split_blocks(text: str):
    # on traduit par paragraphes pour éviter de casser les listes
    return [b for b in text.split("\n\n")]

def join_blocks(blocks):
    return "\n\n".join(blocks)

@st.cache_resource(show_spinner=False)
def _load_argos():
    import argostranslate.package, argostranslate.translate
    # Télécharge/installe FR->NL si absent
    argostranslate.package.update_package_index()
    pkgs = argostranslate.package.get_available_packages()
    fr_nl = next((p for p in pkgs if p.from_code=="fr" and p.to_code=="nl"), None)
    if fr_nl:
        argostranslate.package.install_from_path(fr_nl.download())
    return argostranslate.translate

def translate_fr_to_nl(text: str) -> str:
    if provider.startswith("Local"):
        # Option A — 100% gratuit/offline
        tr = _load_argos()                     # charge une fois
        blocks = [tr.translate(b, "fr", "nl") if b.strip() else "" for b in split_blocks(text)]
        return join_blocks(blocks)

    elif provider.startswith("LibreTranslate"):
        # Option B — API self-hostée gratuite (docker)
        # POST {q, source, target, format}
        blocks = []
        for b in split_blocks(text):
            if not b.strip():
                blocks.append("")
                continue
            r = requests.post(lt_endpoint, json={"q": b, "source":"fr", "target":"nl", "format":"text"}, timeout=60)
            r.raise_for_status()
            blocks.append(r.json()["translatedText"])
        return join_blocks(blocks)

    else:
        # Option C — API publique gratuite (limites/qualité variables)
        blocks = []
        for b in split_blocks(text):
            if not b.strip():
                blocks.append("")
                continue
            r = requests.get("https://api.mymemory.translated.net/get",
                             params={"q": b, "langpair":"fr|nl"}, timeout=30)
            r.raise_for_status()
            blocks.append(r.json()["responseData"]["translatedText"])
        return join_blocks(blocks)

# -------------- Step 1: Load CRM schema --------------
st.header("Étape 1 — Schéma des champs du CRM")
if DEFAULT_SCHEMA_PATH.exists():
    with open(DEFAULT_SCHEMA_PATH, "r", encoding="utf-8") as f:
        schema = yaml.safe_load(f)
else:
    schema = {"fields":[]}
schema_text = st.text_area("Éditez si besoin le schéma (YAML) :", value=yaml.safe_dump(schema, sort_keys=False, allow_unicode=True), height=250)
try:
    schema = yaml.safe_load(schema_text) or {"fields":[]}
    field_keys = [f["key"] for f in schema.get("fields",[])]
    nl_lookup = {f["key"]: f.get("nl_key") for f in schema.get("fields",[])}
    st.success(f"{len(field_keys)} champs chargés.")
except Exception as e:
    st.error(f"YAML invalide: {e}")
    st.stop()

# -------------- Step 2: Upload Word --------------
st.header("Étape 2 — Import du document Word")
uploaded = st.file_uploader("Glissez votre fiche .docx", type=["docx"])

sections: Dict[str, str] = {}
if uploaded is not None:
    tmp_path = Path("uploaded.docx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())
    sections = parse_docx_sections(tmp_path)
    st.write(f"Titres détectés: {list(sections.keys())}")
    st.dataframe(
        [{"Heading": h, "Snippet": s[:180] + ("…" if len(s)>180 else "")} for h,s in sections.items()],
        use_container_width=True
    )

# -------------- Step 3: Mapping --------------
st.header("Étape 3 — Mapping sections Word -> champs CRM")
mapping = {}
if DEFAULT_MAPPING_PATH.exists():
    with open(DEFAULT_MAPPING_PATH, "r", encoding="utf-8") as f:
        mapping = yaml.safe_load(f).get("mapping", {})
else:
    mapping = {}

# UI mapping
st.caption("Associez chaque champ CRM à une section du Word (ou laissez vide). Vous pouvez aussi éditer le texte final FR.")
tab_map, tab_preview = st.tabs(["Construire le mapping", "Aperçu FR/NL"])

with tab_map:
    final_fr: Dict[str,str] = {}
    for field in schema.get("fields", []):
        key = field["key"]
        label = field.get("label", key)
        choices = ["-- (vide) --"] + list(sections.keys())
        default_choice = mapping.get(label) or mapping.get(key) or "-- (vide) --"
        sel = st.selectbox(f"{label} → {key}", options=choices, index=choices.index(default_choice) if default_choice in choices else 0, key=f"sel_{key}")
        base_text = sections.get(sel) if sel and sel != "-- (vide) --" else ""
        edited = st.text_area(f"Texte FR pour {label} ({key})", value=base_text, height=180, key=f"txt_{key}")
        final_fr[key] = edited

    # Save mapping file
    if st.button("Exporter le mapping YAML"):
        out = {"mapping": {f.get("label", f["key"]): st.session_state.get(f"sel_{f['key']}") for f in schema.get("fields", [])}}
        st.download_button("Télécharger sample_mapping.yaml", data=yaml.safe_dump(out, allow_unicode=True), file_name="sample_mapping.yaml")

with tab_preview:
    st.write("Cliquez sur **Générer NL** pour traduire chaque champ FR.")
    if st.button("Générer NL"):

        results = []
        for fdef in fields:
            key = fdef["key"]
            nl_key = nl_key_by_key.get(key)
            text_fr = edited_fr.get(key, "")
            text_nl = translate_fr_to_nl(text_fr) if text_fr.strip() else ""
            results.append({"key": key, "fr": text_fr, "nl_key": nl_key, "nl": text_nl})
        st.session_state["results_auto"] = results

    # Show results
    results = st.session_state.get("results")
    if results:
        st.subheader("Aperçu / Export")
        st.dataframe(results, use_container_width=True)
        payload = {"fields": results}
        st.download_button("Télécharger payload.json", data=json.dumps(payload, ensure_ascii=False, indent=2), file_name="payload.json")
        # CSV simple
        import csv
        out_csv = io.StringIO()
        w = csv.writer(out_csv)
        w.writerow(["key","fr","nl_key","nl"])
        for r in results:
            w.writerow([r["key"], r["fr"], r["nl_key"], r["nl"]])
        st.download_button("Télécharger payload.csv", data=out_csv.getvalue(), file_name="payload.csv", mime="text/csv")

st.info("Ce prototype ne se connecte PAS au CRM. Il prépare un payload (FR+NL) que votre équipe IT pourra intégrer.")
