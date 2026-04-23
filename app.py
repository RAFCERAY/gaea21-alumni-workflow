"""Gaea21 Alumni Hub - Interface d'enrichissement rapide.

Usage : streamlit run app.py
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import webbrowser

from utils import (
    load_alumni_from_excel,
    get_next_alumni_to_enrich,
    save_enrichment_to_excel,
    build_linkedin_search_url,
    extract_info_from_linkedin_text,
    get_stats,
    get_top_entreprises,
    get_by_country,
    get_by_sector,
)
from visualisations import (
    build_world_map,
    build_sector_donut,
    build_timeline,
    build_top_companies,
    fig_to_png_bytes,
)
from brief import (
    get_vedettes,
    build_email_president,
    build_speech_short,
    build_speech_long,
    build_message_temoignage,
    build_fiche_synthese,
)

# ============================================================================
# CONFIG
# ============================================================================

st.set_page_config(
    page_title="Gaea21 Alumni Hub",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded",
)

EXCEL_FILE_DEFAULT = "alumni-gaea21-AG-READY.xlsx"
SHEET_NAME = "Alumni AG"

COUNTRIES = [
    "Suisse", "France", "Allemagne", "Belgique", "Italie",
    "Royaume-Uni", "Espagne", "Portugal", "États-Unis", "Canada",
    "Brésil", "Kenya", "Sénégal", "Maroc", "Tunisie", "Côte d'Ivoire",
    "Burkina Faso", "Algérie", "Togo", "Bénin", "Luxembourg", "Mexique",
    "Autre",
]

SECTORS = [
    "Environnement / Climat", "Finance durable / ESG",
    "Conseil / Audit", "Éducation / Recherche",
    "ONG / Associatif", "Secteur public", "Entreprise privée",
    "Tech / Data", "Santé", "Culture / Média", "Autre",
]

STATUTS = [
    "✅ Enrichi",
    "❌ Introuvable sur LinkedIn",
    "🟡 Profil privé",
    "⭐ Vedette",
]


# ============================================================================
# AUTHENTIFICATION (mot de passe)
# ============================================================================

def check_password() -> bool:
    """Affiche une page de connexion avec mot de passe.

    Returns True si authentifié, False sinon.
    Le mot de passe est stocké dans st.secrets (non versionné).
    """
    # Si déjà authentifié dans cette session, on passe
    if st.session_state.get("password_correct", False):
        return True

    # Récupère le mot de passe attendu depuis les secrets
    try:
        expected_password = st.secrets["password"]
    except (KeyError, FileNotFoundError, Exception):
        # Mode développement local : pas de secrets configurés
        # On laisse passer (pour tester en local sans configurer secrets)
        return True

    # Affiche la page de connexion
    st.title("🌱 Gaea21 Alumni Hub")
    st.markdown("### 🔒 Accès protégé")
    st.caption("Cet outil contient des données alumni — accès réservé à l'équipe Gaea21.")

    with st.form("login_form"):
        password_input = st.text_input(
            "Mot de passe",
            type="password",
            placeholder="Entrez le mot de passe communiqué par Rafika...",
        )
        submitted = st.form_submit_button("🔓 Se connecter", type="primary")

        if submitted:
            if password_input == expected_password:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ Mot de passe incorrect")

    st.markdown("---")
    st.caption(
        "💡 Outil développé par Rafika Cervera — Coordinatrice Data Gouvernance Gaea21. "
        "Données conformes nLPD, restreint par mot de passe."
    )

    return False


# ============================================================================
# SIDEBAR - CONFIG & DASHBOARD
# ============================================================================

def restore_excel_from_secrets():
    """Restaure le fichier Excel depuis les secrets Streamlit Cloud (base64).

    Cette fonction est appelée au démarrage pour reconstituer le fichier Excel
    sur le serveur Streamlit Cloud (car le fichier n'est pas sur GitHub).

    En local (pas de secret excel_base64), ne fait rien → utilise le fichier sur disque.
    """
    import base64

    target_file = EXCEL_FILE_DEFAULT

    # Si fichier déjà présent sur disque, ne rien faire
    if Path(target_file).exists():
        return

    # Essaie de restaurer depuis les secrets
    try:
        excel_b64 = st.secrets["excel_base64"]
    except (KeyError, FileNotFoundError, Exception):
        # Pas de secret configuré : on laisse tomber, l'erreur de fichier sera gérée ailleurs
        return

    try:
        excel_bytes = base64.b64decode(excel_b64)
        with open(target_file, "wb") as f:
            f.write(excel_bytes)
    except Exception as e:
        st.error(f"Erreur restauration fichier Excel : {e}")


def render_sidebar():
    st.sidebar.title("🌱 Gaea21 Alumni Hub")
    st.sidebar.caption("Enrichissement data gouvernance")

    # Fichier
    file_path = st.sidebar.text_input(
        "📁 Fichier Excel",
        value=EXCEL_FILE_DEFAULT,
        help="Chemin vers ton fichier alumni-gaea21-INTERNE.xlsx"
    )

    if not Path(file_path).exists():
        st.sidebar.error(f"❌ Fichier introuvable : {file_path}")
        st.sidebar.info(
            "Mets le fichier `alumni-gaea21-INTERNE.xlsx` "
            "dans le même dossier que ce script."
        )
        return None

    # Stats live
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Progression")
    try:
        df = load_alumni_from_excel(file_path, SHEET_NAME)
        stats = get_stats(df)

        # Barre de progression
        pct = stats["pct_enrichis"]
        st.sidebar.progress(pct / 100)
        st.sidebar.markdown(
            f"**{stats['enrichis']} enrichis** / {stats['total']} total"
        )

        col_a, col_b = st.sidebar.columns(2)
        col_a.metric("✅ Enrichis", stats["enrichis"])
        col_b.metric("⏳ Restants", stats["restants"])

        col_c, col_d = st.sidebar.columns(2)
        col_c.metric("❌ Introuvables", stats["introuvables"])
        col_d.metric("🟡 Privés", stats["prives"])

        st.sidebar.markdown("---")
        st.sidebar.caption(f"Fichier : `{file_path}`")

    except Exception as e:
        st.sidebar.error(f"Erreur lecture fichier : {e}")
        return None

    return file_path


# ============================================================================
# PAGE 1 - DASHBOARD
# ============================================================================

def page_dashboard(file_path: str):
    st.title("📊 Dashboard Alumni Gaea21")
    st.caption("Vue d'ensemble de ta base alumni")

    df = load_alumni_from_excel(file_path, SHEET_NAME)
    stats = get_stats(df)

    # Row 1 : KPIs
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🎯 Total alumni", stats["total"])
    col2.metric("✅ Enrichis", stats["enrichis"], f"+{stats['pct_enrichis']}%")
    col3.metric("❌ Introuvables", stats["introuvables"])
    col4.metric("⏳ Restants", stats["restants"])

    st.markdown("---")

    # Row 2 : répartitions
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("### 🌍 Répartition par pays")
        by_country = get_by_country(df)
        if by_country:
            df_country = pd.DataFrame([
                {"Pays": k, "Nombre": v}
                for k, v in by_country.items()
            ])
            st.bar_chart(df_country.set_index("Pays"))
            st.dataframe(df_country, hide_index=True, use_container_width=True)
        else:
            st.info("Aucun alumni enrichi avec pays pour le moment.")

    with col_b:
        st.markdown("### 🎯 Répartition par secteur")
        by_sector = get_by_sector(df)
        if by_sector:
            df_sector = pd.DataFrame([
                {"Secteur": k, "Nombre": v}
                for k, v in by_sector.items()
            ])
            st.bar_chart(df_sector.set_index("Secteur"))
            st.dataframe(df_sector, hide_index=True, use_container_width=True)
        else:
            st.info("Aucun alumni enrichi avec secteur pour le moment.")

    st.markdown("---")

    # Row 3 : top entreprises
    st.markdown("### 🏢 Top 15 entreprises (mur de logos AG)")
    top = get_top_entreprises(df, n=15)
    if top:
        df_top = pd.DataFrame([
            {"Entreprise": k, "Alumni": v}
            for k, v in top.items()
        ])
        st.dataframe(df_top, hide_index=True, use_container_width=True)
    else:
        st.info("Pas encore d'entreprises enrichies.")


# ============================================================================
# PAGE 2 - ENRICHISSEMENT
# ============================================================================

def page_enrichissement(file_path: str):
    st.title("🔍 Enrichir un alumni")
    st.caption("1 clic LinkedIn → colle → ✨ magique → ✅ sauvé")

    df = load_alumni_from_excel(file_path, SHEET_NAME)

    # Récupère le prochain alumni
    alumni = get_next_alumni_to_enrich(df)

    if alumni is None:
        st.success("🎉 **TOUS les alumni sont enrichis !**")
        st.balloons()
        return

    # Affiche l'alumni courant
    st.markdown(f"### Prochain alumni à enrichir")

    col1, col2, col3 = st.columns(3)
    col1.metric("👤 Nom", f"{alumni['prenom']} {alumni['nom']}")
    col2.metric("🗓️ Sortie", str(alumni["annee_sortie"]))
    col3.metric("💼 Fonction Gaea21", str(alumni["fonction_gaea"])[:40])

    # Ligne de séparation
    st.markdown("---")

    # Étape 1 : ouvrir LinkedIn
    st.markdown("### Étape 1 · 🔍 Recherche LinkedIn")

    col_a, col_b, col_c = st.columns(3)

    url_gaea21 = build_linkedin_search_url(alumni["prenom"], alumni["nom"], "Gaea21")
    url_classic = build_linkedin_search_url(alumni["prenom"], alumni["nom"], "")
    url_fonction = build_linkedin_search_url(
        alumni["prenom"], alumni["nom"], str(alumni["fonction_gaea"]).split()[0]
    )

    col_a.link_button("🎯 + Gaea21", url_gaea21, use_container_width=True)
    col_b.link_button("🔍 Nom seul", url_classic, use_container_width=True)
    col_c.link_button("💼 + Fonction", url_fonction, use_container_width=True)

    # Étape 2 : collage texte
    st.markdown("### Étape 2 · 📋 Colle le texte du profil LinkedIn")
    st.caption("Copie-colle **tout le haut du profil** (nom, headline, localisation, expérience)")

    linkedin_text = st.text_area(
        "Contenu LinkedIn",
        height=200,
        placeholder="Colle ici le texte du profil LinkedIn...\n\n"
                    "Exemple :\n"
                    "Marie Dupont\n"
                    "Relation de 2e niveau\n"
                    "Sustainability Manager chez Patagonia\n"
                    "Genève, Suisse"
    )

    # Étape 3 : analyse auto
    if linkedin_text.strip():
        st.markdown("### Étape 3 · ✨ Extraction automatique")

        extracted = extract_info_from_linkedin_text(linkedin_text)

        with st.expander("🔎 Voir les données extraites automatiquement", expanded=True):
            col_e1, col_e2 = st.columns(2)
            col_e1.text(f"Nom détecté : {extracted.get('nom_complet', '')}")
            col_e1.text(f"Headline : {extracted.get('headline', '')}")
            col_e2.text(f"Localisation : {extracted.get('localisation', '')}")
            col_e2.text(f"Pays deviné : {extracted.get('pays_devine', '?')}")
            col_e2.text(f"Secteur deviné : {extracted.get('secteur_devine', '?')}")

        # Étape 4 : validation humaine
        st.markdown("### Étape 4 · ✏️ Valide ou corrige")

        col_f1, col_f2 = st.columns(2)

        with col_f1:
            poste = st.text_input(
                "Poste actuel",
                value=extracted.get("poste", ""),
                help="Tu peux garder ou modifier"
            )
            entreprise = st.text_input(
                "Entreprise actuelle",
                value=extracted.get("entreprise", ""),
            )
            ville = st.text_input(
                "Ville",
                value="",
                placeholder="Ex: Genève, Paris..."
            )

        with col_f2:
            # Pays avec valeur pré-remplie
            pays_default = extracted.get("pays_devine", "")
            if pays_default and pays_default in COUNTRIES:
                pays_idx = COUNTRIES.index(pays_default)
            else:
                pays_idx = 0
            pays = st.selectbox("Pays", COUNTRIES, index=pays_idx)

            # Secteur avec valeur pré-remplie
            secteur_default = extracted.get("secteur_devine", "")
            if secteur_default and secteur_default in SECTORS:
                sec_idx = SECTORS.index(secteur_default)
            else:
                sec_idx = 0
            secteur = st.selectbox("Secteur", SECTORS, index=sec_idx)

            statut = st.selectbox("Statut", STATUTS, index=0)

        url_linkedin = st.text_input(
            "URL LinkedIn",
            placeholder="https://www.linkedin.com/in/..."
        )

        notes = st.text_area(
            "Notes (optionnel)",
            placeholder="Ex: ⭐ Vedette AG - candidat témoignage...",
            height=80,
        )

        # Étape 5 : sauvegarder
        st.markdown("---")

        col_save1, col_save2, col_save3 = st.columns([3, 1, 1])

        with col_save1:
            if st.button("✅ Valider et passer au suivant", type="primary", use_container_width=True):
                data = {
                    "Poste actuel": poste,
                    "Entreprise": entreprise,
                    "Ville": ville,
                    "Pays": pays,
                    "Secteur": secteur,
                    "URL LinkedIn": url_linkedin,
                    "Statut": statut,
                    "Notes": notes,
                }
                try:
                    save_enrichment_to_excel(
                        file_path,
                        alumni["row_index_excel"],
                        data,
                    )
                    st.success(f"✅ {alumni['prenom']} {alumni['nom']} sauvé !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur sauvegarde : {e}")

        with col_save2:
            if st.button("❌ Introuvable", use_container_width=True):
                data = {
                    "Statut": "❌ Introuvable sur LinkedIn",
                }
                try:
                    save_enrichment_to_excel(
                        file_path,
                        alumni["row_index_excel"],
                        data,
                    )
                    st.success("Marqué introuvable, au suivant !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")

        with col_save3:
            if st.button("🟡 Privé", use_container_width=True, help="Profil LinkedIn existe mais peu/pas d'infos accessibles"):
                data = {
                    "Statut": "🟡 Profil privé",
                    "Notes": notes if notes else "Profil LinkedIn trouvé mais infos insuffisantes",
                }
                try:
                    save_enrichment_to_excel(
                        file_path,
                        alumni["row_index_excel"],
                        data,
                    )
                    st.success("Marqué profil privé, au suivant !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")

    else:
        # Cas : pas encore de texte collé
        st.info(
            "👆 **Étape suivante** : clique sur un bouton LinkedIn ci-dessus, "
            "trouve le profil, copie-colle son texte dans le champ ci-dessus."
        )

        # 2 options bypass : Introuvable ou Profil privé
        col_bp1, col_bp2 = st.columns(2)

        with col_bp1:
            if st.button("❌ Passer — Introuvable", use_container_width=True):
                try:
                    save_enrichment_to_excel(
                        file_path,
                        alumni["row_index_excel"],
                        {"Statut": "❌ Introuvable sur LinkedIn"},
                    )
                    st.success("Marqué introuvable, au suivant !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")

        with col_bp2:
            if st.button("🟡 Passer — Profil privé", use_container_width=True, help="Profil trouvé mais infos inaccessibles"):
                try:
                    save_enrichment_to_excel(
                        file_path,
                        alumni["row_index_excel"],
                        {
                            "Statut": "🟡 Profil privé",
                            "Notes": "Profil LinkedIn trouvé mais infos insuffisantes",
                        },
                    )
                    st.success("Marqué profil privé, au suivant !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")


# ============================================================================
# PAGE 3 - VISUALISATIONS (Phase 2 - pour l'AG)
# ============================================================================

def page_visualisations(file_path: str):
    st.title("🎯 Visualisations pour l'AG")
    st.caption("Graphiques prêts à intégrer dans ta slide Canva")

    df = load_alumni_from_excel(file_path, SHEET_NAME)
    stats = get_stats(df)

    # KPIs en haut
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🎯 Total", stats["total"])
    col2.metric("✅ Enrichis", stats["enrichis"])
    col3.metric("🌍 Pays", len(get_by_country(df)))
    col4.metric("🏢 Secteurs", len(get_by_sector(df)))

    if stats["enrichis"] == 0:
        st.warning(
            "⚠️ Aucun alumni enrichi pour le moment. "
            "Va sur la page 'Enrichir' pour commencer."
        )
        return

    st.markdown("---")

    # === CARTE DU MONDE ===
    st.markdown("## 🌍 Carte du monde")
    st.caption("Répartition géographique des alumni Gaea21")

    country_counts = get_by_country(df)
    fig_map = build_world_map(country_counts)
    st.plotly_chart(fig_map, use_container_width=True)

    # Bouton export
    col_dl1, col_info1 = st.columns([1, 3])
    with col_dl1:
        png = fig_to_png_bytes(fig_map, width=1400, height=800)
        if png:
            st.download_button(
                "📥 Télécharger PNG",
                data=png,
                file_name="carte_monde_alumni.png",
                mime="image/png",
                key="dl_map",
            )
        else:
            st.caption("Install: `pip3 install kaleido`")
    with col_info1:
        st.caption(f"🌍 **{len(country_counts)} pays** représentés parmi tes alumni enrichis")

    st.markdown("---")

    # === DONUT SECTEURS ===
    st.markdown("## 🎨 Répartition par secteur")
    st.caption("Dans quels secteurs travaillent nos alumni ?")

    sector_counts = get_by_sector(df)
    fig_donut = build_sector_donut(sector_counts)
    st.plotly_chart(fig_donut, use_container_width=True)

    col_dl2, col_info2 = st.columns([1, 3])
    with col_dl2:
        png = fig_to_png_bytes(fig_donut, width=1200, height=800)
        if png:
            st.download_button(
                "📥 Télécharger PNG",
                data=png,
                file_name="donut_secteurs.png",
                mime="image/png",
                key="dl_donut",
            )
    with col_info2:
        if sector_counts:
            top_sector = max(sector_counts.items(), key=lambda x: x[1])
            st.caption(f"🏆 **Secteur #1** : {top_sector[0]} ({top_sector[1]} alumni)")

    st.markdown("---")

    # === TIMELINE ===
    st.markdown("## 📅 Timeline des sorties")
    st.caption("Évolution du nombre d'alumni par année")

    fig_timeline = build_timeline(df)
    st.plotly_chart(fig_timeline, use_container_width=True)

    col_dl3, col_info3 = st.columns([1, 3])
    with col_dl3:
        png = fig_to_png_bytes(fig_timeline, width=1200, height=600)
        if png:
            st.download_button(
                "📥 Télécharger PNG",
                data=png,
                file_name="timeline_annees.png",
                mime="image/png",
                key="dl_timeline",
            )
    with col_info3:
        st.caption(f"📊 Calculé sur **{stats['total']} alumni** (toute la base, pas seulement les enrichis)")

    st.markdown("---")

    # === TOP ENTREPRISES ===
    st.markdown("## 🏢 Top entreprises (mur de logos)")
    st.caption("Où travaillent nos alumni aujourd'hui ?")

    top_companies = get_top_entreprises(df, n=15)
    fig_top = build_top_companies(top_companies)
    st.plotly_chart(fig_top, use_container_width=True)

    col_dl4, col_info4 = st.columns([1, 3])
    with col_dl4:
        png = fig_to_png_bytes(fig_top, width=1400, height=800)
        if png:
            st.download_button(
                "📥 Télécharger PNG",
                data=png,
                file_name="top_entreprises.png",
                mime="image/png",
                key="dl_top",
            )
    with col_info4:
        st.caption(
            f"🏢 **{len(top_companies)} entreprises** distinctes dans ta base. "
            "Garde les plus prestigieuses pour ton mur de logos Canva."
        )

    st.markdown("---")

    # === RÉSUMÉ POUR L'AG ===
    st.markdown("## 📋 Résumé à inclure dans ton pitch AG")

    nb_pays = len(country_counts)
    nb_secteurs = len(sector_counts)

    summary = f"""
**Chiffres clés à mettre en avant vendredi** :

- 🎯 **{stats['total']} alumni** dans notre communauté Gaea21
- 🌍 **{nb_pays} pays** différents représentés
- 🏢 **{nb_secteurs} secteurs** d'activité variés
- ✅ **{stats['enrichis']} profils** qualifiés à ce jour

**Top 3 secteurs** :
"""
    top3 = sorted(sector_counts.items(), key=lambda x: x[1], reverse=True)[:3]
    for i, (sec, count) in enumerate(top3, 1):
        summary += f"{i}. **{sec}** : {count} alumni\n"

    st.markdown(summary)

    # Bouton copier-coller
    st.text_area("📋 Copie-colle dans ton script AG :", value=summary, height=200)


# ============================================================================
# PAGE 4 - BRIEF PRÉSIDENT (Phase 3)
# ============================================================================

def page_brief(file_path: str):
    st.title("📧 Brief président")
    st.caption("Tous les livrables pour préparer l'AG en 1 page")

    df = load_alumni_from_excel(file_path, SHEET_NAME)
    stats = get_stats(df)
    country_counts = get_by_country(df)
    sector_counts = get_by_sector(df)
    top_companies = get_top_entreprises(df, n=15)
    vedettes = get_vedettes(df)

    if stats["enrichis"] == 0:
        st.warning(
            "⚠️ Aucun alumni enrichi. "
            "Va sur la page 'Enrichir' pour commencer."
        )
        return

    # KPIs rapides
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🎯 Total", stats["total"])
    col2.metric("✅ Enrichis", stats["enrichis"])
    col3.metric("🌍 Pays", len(country_counts))
    col4.metric("⭐ Vedettes", len(vedettes))

    st.markdown("---")

    # Onglets pour organiser les livrables
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📧 Email président",
        "🎤 Discours AG",
        "⭐ Vedettes",
        "📋 Fiche synthèse",
        "✉️ Messages témoignages",
    ])

    # --- Onglet 1 : Email président ---
    with tab1:
        st.markdown("### 📧 Email pré-rempli au président")
        st.caption("Copie-colle dans Gmail, personnalise si besoin, envoie.")

        email = build_email_president(
            stats, country_counts, sector_counts, top_companies, vedettes,
            president_name="Monsieur le Président",
        )

        st.text_input("📬 Objet de l'email", value=email["subject"])
        st.text_area("📝 Corps de l'email (cliquer puis Cmd+A + Cmd+C pour copier)",
                     value=email["body"], height=500)

        st.download_button(
            "📥 Télécharger en .txt",
            data=email["body"],
            file_name="email_president_ag.txt",
            mime="text/plain",
        )

    # --- Onglet 2 : Discours AG ---
    with tab2:
        st.markdown("### 🎤 Scripts de discours pour l'AG")

        discours_choix = st.radio(
            "Quel format ?",
            ["Court (2 min)", "Long (5 min)"],
            horizontal=True,
        )

        if "Court" in discours_choix:
            discours = build_speech_short(stats, country_counts, sector_counts, vedettes)
        else:
            discours = build_speech_long(stats, country_counts, sector_counts, top_companies, vedettes)

        st.markdown(discours)

        st.download_button(
            "📥 Télécharger le discours en .txt",
            data=discours,
            file_name=f"discours_ag_{discours_choix.split()[0].lower()}.txt",
            mime="text/plain",
            key=f"dl_discours_{discours_choix}",
        )

    # --- Onglet 3 : Vedettes ---
    with tab3:
        st.markdown(f"### ⭐ {len(vedettes)} vedettes identifiées")
        st.caption("Ces alumni sont marqués ⭐ dans les Notes — candidats idéaux pour témoignages")

        if not vedettes:
            st.info(
                "Aucune vedette identifiée pour le moment.\n\n"
                "**Astuce** : ajoute ⭐ dans la colonne Notes d'un alumni enrichi pour le marquer vedette."
            )
        else:
            for v in vedettes:
                stars = "⭐" * v["nb_stars"]
                with st.expander(f"{stars} **{v['prenom']} {v['nom']}** — {v['entreprise']}"):
                    col_v1, col_v2 = st.columns(2)
                    col_v1.markdown(f"**Poste** : {v['poste']}")
                    col_v1.markdown(f"**Pays** : {v['pays']}")
                    col_v2.markdown(f"**Secteur** : {v['secteur']}")
                    if v["url_linkedin"]:
                        col_v2.markdown(f"**LinkedIn** : [Profil]({v['url_linkedin']})")
                    st.caption(f"Notes : {v['notes'][:200]}")

    # --- Onglet 4 : Fiche synthèse ---
    with tab4:
        st.markdown("### 📋 Fiche 1 page à imprimer pour le président")

        fiche = build_fiche_synthese(stats, country_counts, sector_counts, top_companies, vedettes)

        st.markdown(fiche)

        st.download_button(
            "📥 Télécharger la fiche en .md",
            data=fiche,
            file_name="fiche_synthese_ag.md",
            mime="text/markdown",
        )

    # --- Onglet 5 : Messages de témoignages ---
    with tab5:
        st.markdown("### ✉️ Messages LinkedIn à envoyer aux vedettes")
        st.caption("Pour demander un témoignage à intégrer dans la slide AG")

        if not vedettes:
            st.info("Marque d'abord tes vedettes avec ⭐ dans les Notes.")
        else:
            for v in vedettes[:5]:
                with st.expander(f"✉️ Message pour **{v['prenom']} {v['nom']}**"):
                    message = build_message_temoignage(v)
                    st.text_area(
                        "Message à envoyer",
                        value=message,
                        height=250,
                        key=f"msg_{v['prenom']}_{v['nom']}",
                    )
                    if v["url_linkedin"]:
                        st.link_button(
                            f"🔗 Ouvrir le profil de {v['prenom']}",
                            v["url_linkedin"],
                            use_container_width=False,
                        )


# ============================================================================
# ROUTING
# ============================================================================

def main():
    # Restaure le fichier Excel depuis les secrets (si déployé sur Streamlit Cloud)
    restore_excel_from_secrets()

    # Check authentification avant tout affichage
    if not check_password():
        return

    file_path = render_sidebar()

    if not file_path:
        st.title("🌱 Gaea21 Alumni Hub")
        st.warning("Configure le fichier Excel dans la sidebar pour commencer.")
        return

    # Navigation
    page = st.sidebar.radio(
        "📑 Navigation",
        ["🔍 Enrichir", "📊 Dashboard", "🎯 Visualisations", "📧 Brief président"],
    )

    if page == "🔍 Enrichir":
        page_enrichissement(file_path)
    elif page == "📊 Dashboard":
        page_dashboard(file_path)
    elif page == "🎯 Visualisations":
        page_visualisations(file_path)
    elif page == "📧 Brief président":
        page_brief(file_path)


if __name__ == "__main__":
    main()
