
import os
import io
import json
import yaml
import streamlit as st
from pathlib import Path
from typing import Dict, List
from utils.docx_parser import parse_docx_sections

# -------------- Config --------------
st.set_page_config(page_title="Prototype Mapping Word -> CRM + NL", layout="wide")

DEFAULT_SCHEMA_PATH = Path("crm_schema.yaml")
DEFAULT_MAPPING_PATH = Path("sample_mapping.yaml")

# Sidebar: OpenAI key
st.sidebar.header("Configuration")
openai_key = st.sidebar.text_input("OPENAI_API_KEY (optionnel pour tester la traduction)", type="password", value=os.getenv("OPENAI_API_KEY",""))
model_name = st.sidebar.text_input("Modèle OpenAI", value="gpt-4o-mini")
system_prompt = st.sidebar.text_area("System prompt traduction", value=(
    "You are a professional translator. Translate from French to Dutch (Belgium). "
    "Preserve structure, bullet lists, tables and numbers. Keep a neutral finance tone."
))

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
        # lazy import
        try:
            import openai
            client = openai.OpenAI(api_key=openai_key) if openai_key else None
        except Exception as e:
            client = None
            st.warning(f"SDK OpenAI non disponible ou clé absente. Erreur: {e}. On simule la traduction (préfixe [NL]).")

        results = []
        for f in schema.get("fields", []):
            key = f["key"]
            nl_key = nl_lookup.get(key)
            text_fr = st.session_state.get(f"txt_{key}", "")
            if not text_fr.strip():
                results.append({"key": key, "fr": "", "nl_key": nl_key, "nl": ""})
                continue

            if client:
                try:
                    resp = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role":"system","content": system_prompt},
                            {"role":"user","content": text_fr}
                        ],
                        temperature=0.2,
                    )
                    text_nl = resp.choices[0].message.content.strip()
                except Exception as e:
                    text_nl = f"[NL ERREUR: {e}]"
            else:
                text_nl = "[NL] " + text_fr  # fallback demo

            results.append({"key": key, "fr": text_fr, "nl_key": nl_key, "nl": text_nl})

        st.session_state["results"] = results

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
