
# Prototype Word → CRM (FR + NL) — Streamlit

Ce projet permet de :
1. Importer une fiche `.docx` (structure stable),
2. Extraire les sections (titres + contenus, tables incluses en Markdown),
3. Mapper chaque section vers vos **champs CRM** (schéma éditable),
4. Traduire chaque champ **FR → NL (Belgique)** via l'API OpenAI,
5. Exporter un **payload JSON/CSV** prêt pour l'intégration côté CRM.

> **Aucun accès au CRM** ici : c'est volontaire. Vous testez, ajustez le mapping, et remettez le payload + le code à l'IT qui branchera l'API interne.

## Lancer
```bash
pip install -r requirements.txt
# Optionnel : exportez votre clé si vous voulez la vraie traduction
export OPENAI_API_KEY="sk-..."
streamlit run app_streamlit.py
```
Ouvrez l'URL locale affichée par Streamlit.

## Fichiers clés
- `crm_schema.yaml` : liste des champs de votre CRM (clé FR + clé NL + label). Éditable directement dans l'app.
- `sample_mapping.yaml` : exemple de correspondance *titre Word* → *clé CRM FR*.
- `utils/docx_parser.py` : parseur `.docx` basé sur `python-docx`.

## Flux d’usage
1. Chargez votre `.docx`.
2. Ajustez le schéma YAML si besoin (noms techniques FR/NL).
3. Allez dans **Étape 3** : associez chaque *champ CRM* à un *titre Word*.
4. Modifiez/affinez le texte FR si besoin.
5. Cliquez **Générer NL** (utilise `OPENAI_API_KEY` si présent, sinon un *mock* `[NL] ...`).
6. Téléchargez le `payload.json` et/ou `payload.csv` pour IT.

## Intégration côté IT
Le `payload.json` a la forme :
```json
{
  "fields": [
    {"key": "intro_fr", "fr": "…", "nl_key": "intro_nl", "nl": "…"},
    ...
  ]
}
```
L’IT peut alors pousser ces valeurs dans le CRM (via API interne ou script navigateur).

## Remarques
- Les tableaux dans le Word sont ajoutés **en Markdown** à la fin de la dernière section détectée.
- Si la structure du Word évolue : mettez à jour le `mapping` ou les titres.
- Si l’éditeur NL impose des limites de longueur, il faudra éventuellement tronquer côté IT.
```
