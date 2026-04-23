# 🌱 Gaea21 Alumni Hub

**Data governance workflow for alumni mapping — Python + Streamlit + Plotly**

Outil de cartographie et d'enrichissement des alumni d'une association, développé pour Gaea21 (Suisse) en préparation de son AG annuelle.

**🎯 Objectif** : transformer une liste brute de 500+ alumni en une base de données qualifiée, visualisable, avec export de livrables prêts pour des présentations stratégiques.

---

## ✨ Fonctionnalités

### 🔍 Enrichissement semi-automatique
- Interface Streamlit pour qualifier un alumni à la fois
- Génération automatique d'URLs de recherche LinkedIn (respect des CGU)
- Extraction automatique depuis un profil LinkedIn collé (regex + NLP basique)
- Classification automatique du secteur et du pays (dictionnaires enrichis)
- Sauvegarde incrémentale dans un fichier Excel maître

### 📊 Dashboard & Visualisations
- Vue d'ensemble avec KPIs live
- Carte du monde interactive (Plotly Scattergeo)
- Donut des secteurs d'activité
- Timeline des sorties par année
- Top entreprises employeuses
- Export PNG haute qualité pour slides Canva

### 📧 Brief AG auto-généré
- Email au président avec stats injectées
- Discours AG (versions courte 2 min et longue 5 min)
- Liste des "vedettes alumni" (marquées ⭐ dans les Notes)
- Fiche synthèse imprimable
- Messages LinkedIn pré-rédigés pour solliciter des témoignages

### 🔐 Authentification
- Mot de passe optionnel (via `st.secrets`)
- Déploiement sécurisé sur Streamlit Cloud

---

## 🏗️ Architecture technique

```
alumni_workflow/
├── app.py                    # Application Streamlit principale (4 pages)
├── utils.py                  # Fonctions utilitaires (lecture Excel, classification)
├── visualisations.py         # Graphiques Plotly
├── brief.py                  # Génération email / discours / fiche synthèse
├── marquer_vedettes.py       # Script one-shot pour marquer les alumni vedettes
├── requirements.txt          # Dépendances Python
└── .gitignore                # Exclut les données sensibles
```

### Technologies

- **Python 3.12**
- **Streamlit** — interface web
- **Pandas + Openpyxl** — manipulation du fichier Excel
- **Plotly + Kaleido** — graphiques interactifs + export PNG
- **Regex** — extraction d'entités depuis texte libre

---

## 🚀 Installation locale

```bash
# 1. Cloner le repo
git clone https://github.com/RAFCERAY/gaea21-alumni-workflow.git
cd gaea21-alumni-workflow

# 2. Créer un environnement virtuel
python3 -m venv venv
source venv/bin/activate

# 3. Installer les dépendances
pip install -r requirements.txt

# 4. Placer le fichier Excel source (non versionné) dans le dossier
# Format attendu : feuille "Alumni AG" avec colonnes Nom, Prénom, Année sortie, etc.

# 5. Lancer l'app
streamlit run app.py
```

L'app s'ouvre automatiquement sur `http://localhost:8501`.

---

## 🔐 Configuration du mot de passe (optionnel)

Pour activer la protection par mot de passe :

1. Créer un fichier `.streamlit/secrets.toml` :

```toml
password = "votre_mot_de_passe_ici"
```

2. Ce fichier est **ignoré par Git** (voir `.gitignore`).
3. Sur Streamlit Cloud : configurer le secret via **Settings → Secrets** dans le dashboard.

---

## 🎯 Conformité nLPD (Suisse)

L'outil respecte la **nouvelle Loi suisse sur la Protection des Données** (nLPD, Sept 2023) :

- ✅ **Pas de scraping** — aucune donnée n'est extraite automatiquement de LinkedIn (respect des CGU)
- ✅ **Séparation stricte** — les données sensibles (emails, motifs de sortie) sont dans un fichier `_INTERNE` séparé, jamais dans la version publique
- ✅ **Transparence** — chaque alumni dans la base doit être informé de sa présence
- ✅ **Accès restreint** — mot de passe sur l'app en ligne

---

## 🎓 Apprentissages & valeur pédagogique

Ce projet illustre des concepts clés du DAMA-DMBOK (Data Management Body of Knowledge) :

- **Entity Resolution** — rapprocher des identités entre sources (fichier Gaea21 ↔ LinkedIn)
- **Master Data Management** — règle du "Principal Attribute" pour les profils multi-rôles
- **Data Quality > Data Completeness** — préférer une donnée manquante à un faux match
- **Social Proof Validation** — utiliser la proximité réseau LinkedIn (2e degré) comme preuve
- **Agile Data Management** — livraison MVP itérative (Phase 1 → 2 → 3)
- **Pipeline data** — de la lecture à la visualisation, bout en bout

---

## 👤 Autrice

**Rafika Cervera**
Coordinatrice Data Gouvernance — Gaea21 (stage 2025-2026)

Projet développé dans le cadre d'un stage en data governance et préparation au CDMP Fundamentals (DAMA).

🐙 GitHub : [RAFCERAY](https://github.com/RAFCERAY)

---

## 📄 Licence

MIT — code open source. Les **données alumni** ne sont PAS incluses dans ce repo (fichier `.xlsx` exclu du versioning).
