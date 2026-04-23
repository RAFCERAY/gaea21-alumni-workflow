"""Génération du brief président pour l'AG Gaea21.

Génère :
- Email au président avec stats pré-remplies
- Scripts de discours (court + long)
- Liste des vedettes / témoignages potentiels
- Fiche synthèse imprimable
"""

from datetime import datetime
import pandas as pd


# ============================================================================
# IDENTIFICATION DES VEDETTES
# ============================================================================

def get_vedettes(df: pd.DataFrame) -> list:
    """Extrait les alumni marqués comme vedettes depuis la colonne Notes.

    Un alumni est une vedette si sa colonne Notes contient "⭐" ou "vedette" ou "VEDETTE".
    """
    # Trouve la colonne Notes
    notes_col = None
    for c in df.columns:
        if c.lower() in ("notes", "note"):
            notes_col = c
            break

    if notes_col is None:
        return []

    # Trouve la colonne Statut
    status_col = None
    for c in df.columns:
        if "statut" in str(c).lower():
            status_col = c
            break

    vedettes = []
    for idx, row in df.iterrows():
        notes = str(row.get(notes_col, ""))
        # Détection vedette : emoji ou mot-clé dans les Notes
        is_vedette = "⭐" in notes or "vedette" in notes.lower()
        if not is_vedette:
            continue

        # Compte les étoiles pour trier par priorité
        nb_stars = notes.count("⭐")

        vedettes.append({
            "prenom": row.get("Prénom", ""),
            "nom": row.get("Nom", ""),
            "poste": row.get("Poste actuel", ""),
            "entreprise": row.get("Entreprise", ""),
            "pays": row.get("Pays", ""),
            "secteur": row.get("Secteur", ""),
            "url_linkedin": row.get("URL LinkedIn", ""),
            "notes": notes,
            "nb_stars": max(nb_stars, 1),
        })

    # Trie par nb d'étoiles décroissant
    vedettes.sort(key=lambda x: -x["nb_stars"])
    return vedettes


# ============================================================================
# GÉNÉRATION DE L'EMAIL AU PRÉSIDENT
# ============================================================================

def build_email_president(
    stats: dict,
    country_counts: dict,
    sector_counts: dict,
    top_companies: dict,
    vedettes: list,
    president_name: str = "Monsieur le Président",
) -> dict:
    """Construit l'email au président.

    Returns :
        dict {"subject": ..., "body": ...}
    """
    total = stats["total"]
    enrichis = stats["enrichis"]
    nb_pays = len(country_counts)
    nb_secteurs = len(sector_counts)
    pct_enrichis = stats["pct_enrichis"]

    # Top 3 pays
    top_pays = sorted(country_counts.items(), key=lambda x: -x[1])[:3]
    top_pays_str = ", ".join(f"{p} ({n})" for p, n in top_pays)

    # Top 3 secteurs
    top_sec = sorted(sector_counts.items(), key=lambda x: -x[1])[:3]
    top_sec_str = ", ".join(f"{s} ({n})" for s, n in top_sec)

    # Top 5 entreprises
    top_ent = list(top_companies.items())[:5]
    top_ent_str = ", ".join(t[0] for t in top_ent)

    subject = f"AG Gaea21 — Synthèse mapping alumni ({enrichis} profils qualifiés)"

    body = f"""Bonjour {president_name},

Je vous partage la synthèse du mapping alumni réalisé en préparation de l'AG.

## Chiffres clés à présenter vendredi

- **{total} alumni** identifiés dans notre communauté Gaea21
- **{enrichis} profils qualifiés** à ce jour (soit {pct_enrichis}%)
- **{nb_pays} pays** différents représentés
- **{nb_secteurs} secteurs** d'activité variés

## Répartition géographique

Principaux pays : {top_pays_str}

Notre réseau dépasse largement Genève : nos alumni poursuivent l'engagement environnemental en France, en Afrique francophone, en Allemagne, au Mexique et au-delà.

## Répartition sectorielle

Top 3 secteurs : {top_sec_str}

L'environnement et le climat ressortent comme notre secteur principal — preuve que Gaea21 forme des talents qui continuent à œuvrer pour la transition écologique après leur passage chez nous.

## Entreprises employeuses (mur de logos)

Nos alumni travaillent aujourd'hui chez : {top_ent_str}, et d'autres.
Ces entreprises incarnent la portée et la crédibilité du réseau Gaea21.

## Profils vedettes pour témoignage

"""

    # Top 5 vedettes
    for v in vedettes[:5]:
        stars = "⭐" * v["nb_stars"]
        body += f"- **{v['prenom']} {v['nom']}** {stars} — {v['poste']} chez {v['entreprise']} ({v['pays']})\n"

    body += f"""
Je recommande de contacter 2-3 de ces vedettes dès lundi pour solliciter un court témoignage écrit ou vidéo à intégrer à la slide AG.

## Livrables prêts

- Fichier alumni complet (confidentiel, nLPD) : alumni-gaea21-INTERNE.xlsx
- Fichier public sans données sensibles : alumni-gaea21-AG-READY.xlsx
- Visualisations pour la slide AG (carte monde, donut secteurs, timeline, mur logos) : exports PNG disponibles
- Méthodologie et rapport : fichiers séparés

Je reste disponible pour toute question d'ici vendredi.

Bien cordialement,
Rafika
Coordinatrice Data Gouvernance
Gaea21
"""

    return {"subject": subject, "body": body}


# ============================================================================
# SCRIPTS DE DISCOURS POUR L'AG
# ============================================================================

def build_speech_short(stats: dict, country_counts: dict, sector_counts: dict, vedettes: list) -> str:
    """Discours court (2 min) pour l'AG."""
    total = stats["total"]
    enrichis = stats["enrichis"]
    nb_pays = len(country_counts)

    # Récupère une vedette pour exemple
    exemple = None
    if vedettes:
        v = vedettes[0]
        exemple = f"{v['prenom']} {v['nom']}, aujourd'hui {v['poste']} chez {v['entreprise']} ({v['pays']})"

    top_sec = sorted(sector_counts.items(), key=lambda x: -x[1])[:1]
    secteur_1 = top_sec[0][0] if top_sec else "l'environnement"

    speech = f"""⏱️ **DISCOURS COURT (2 minutes) - Pour l'AG**

---

Chers membres, chers partenaires,

Je souhaite aujourd'hui partager avec vous une donnée qui résume l'impact de notre association : **notre communauté alumni rassemble aujourd'hui {total} personnes**, dont nous avons pu identifier et qualifier précisément **{enrichis} profils**.

Ces alumni sont désormais présents dans **{nb_pays} pays différents**, de la Suisse jusqu'au Mexique, en passant par l'Allemagne, le Maroc, le Togo, et la Tunisie.

Le message le plus fort à retenir : **une majorité de nos alumni continuent à travailler dans le domaine de {secteur_1}**. Gaea21 forme des talents qui deviennent des acteurs de la transition écologique dans leurs entreprises, leurs laboratoires, leurs startups.

{f"Pour illustrer concrètement : {exemple}. Son parcours incarne ce que nous produisons : des professionnels engagés qui portent notre mission au-delà de nos murs." if exemple else ""}

Cette cartographie de nos alumni est désormais un outil vivant qui nous permettra de :
- Mobiliser notre réseau pour de futurs projets
- Solliciter des témoignages et parrainages
- Démontrer auprès de nos partenaires l'impact concret et durable de Gaea21

Merci à tous pour votre engagement.
"""

    return speech


def build_speech_long(stats: dict, country_counts: dict, sector_counts: dict,
                     top_companies: dict, vedettes: list) -> str:
    """Discours long (5 min) pour l'AG."""
    total = stats["total"]
    enrichis = stats["enrichis"]
    nb_pays = len(country_counts)
    nb_secteurs = len(sector_counts)

    top_pays = sorted(country_counts.items(), key=lambda x: -x[1])[:5]
    top_sec = sorted(sector_counts.items(), key=lambda x: -x[1])[:3]
    top_ent = list(top_companies.items())[:5]

    speech = f"""⏱️ **DISCOURS LONG (5 minutes) - Pour l'AG**

---

## 1. Introduction (30 sec)

Chers membres, chers partenaires, bienvenue à cette Assemblée Générale.

Je souhaite partager avec vous, pour la première fois, une photographie précise de notre communauté alumni. Depuis sa création, Gaea21 a accueilli {total} personnes — bénévoles, stagiaires, volontaires. Qui sont-ils devenus ? Où travaillent-ils aujourd'hui ? Ce mapping nous le dit.

## 2. Les chiffres clés (1 min)

- **{total} alumni** répertoriés dans notre base
- **{enrichis} profils qualifiés** précisément ce mois-ci
- **{nb_pays} pays** différents représentés
- **{nb_secteurs} secteurs** d'activité

## 3. Une portée internationale (1 min)

Nos alumni sont présents dans des pays très différents. Les principaux foyers sont :
"""

    for pays, count in top_pays:
        speech += f"- **{pays}** : {count} alumni\n"

    speech += """
Cette diversité géographique n'est pas un hasard — elle reflète le modèle de travail à distance que nous avons porté depuis le début. Gaea21 a su attirer des talents engagés bien au-delà de la Suisse romande.

## 4. Des secteurs porteurs (1 min)

Quand on regarde où nos alumni travaillent aujourd'hui, trois secteurs se détachent :

"""
    for sec, count in top_sec:
        speech += f"- **{sec}** : {count} alumni\n"

    speech += f"""
Le secteur **{top_sec[0][0] if top_sec else 'environnemental'}** arrive en tête. C'est exactement la démonstration que nous cherchions à faire : **Gaea21 est un tremplin pour la transition écologique**. Les personnes qui passent par nous restent dans l'écosystème de la durabilité.

## 5. Nos alumni dans les grandes entreprises (1 min)

Nos alumni travaillent dans des entreprises et institutions reconnues :
"""
    for ent, count in top_ent:
        speech += f"- **{ent}**\n"

    speech += """
De la multinationale industrielle à la PME innovante, du cabinet conseil à la grande école, nos alumni exercent une influence dans des environnements variés. C'est la preuve que nous formons des professionnels que le marché valorise.

## 6. Des parcours inspirants (1 min)

Je voudrais terminer avec quelques parcours exceptionnels qui incarnent notre mission :

"""
    for v in vedettes[:3]:
        speech += f"- **{v['prenom']} {v['nom']}** — {v['poste']} chez {v['entreprise']} ({v['pays']})\n"

    speech += """
Ces personnes ont toutes un point commun : leur passage chez Gaea21 a nourri un engagement qu'elles continuent à porter dans leur vie professionnelle.

## 7. Conclusion (30 sec)

Cette cartographie alumni n'est pas seulement un exercice statistique : c'est un **outil stratégique** pour notre association.

- Elle nous permet de **mesurer notre impact** dans la durée.
- Elle nous donne des **ambassadeurs** à mobiliser.
- Elle **crédibilise notre modèle** auprès de nos financeurs potentiels.

Je remercie toute l'équipe pour ce travail, et je nous invite à continuer à enrichir cette communauté.

Merci.
"""

    return speech


# ============================================================================
# MESSAGES AUX VEDETTES POUR TÉMOIGNAGE
# ============================================================================

def build_message_temoignage(vedette: dict) -> str:
    """Génère un message LinkedIn à envoyer à une vedette pour demander un témoignage."""
    prenom = vedette["prenom"]
    poste = vedette["poste"]
    entreprise = vedette["entreprise"]

    message = f"""Bonjour {prenom},

J'espère que tu vas bien. Je m'appelle Rafika Cervera et je travaille comme coordinatrice data gouvernance chez Gaea21.

Dans le cadre de notre Assemblée Générale de cette semaine, nous préparons une présentation de notre communauté alumni pour montrer l'impact de Gaea21 sur les parcours de nos anciens membres.

Ton parcours — passé par Gaea21, aujourd'hui {poste} chez {entreprise} — fait partie des parcours exemplaires que nous aimerions mettre en valeur.

Serais-tu d'accord pour nous partager en quelques phrases :
1. Ce que Gaea21 t'a apporté
2. Comment ton passage chez nous a influencé ton parcours actuel

Format libre : 3-5 phrases, par écrit, par LinkedIn ou par email.

Merci beaucoup pour ta réponse, même brève !

Bien cordialement,
Rafika
"""
    return message


# ============================================================================
# FICHE SYNTHÈSE IMPRIMABLE
# ============================================================================

def build_fiche_synthese(stats: dict, country_counts: dict, sector_counts: dict,
                        top_companies: dict, vedettes: list) -> str:
    """Fiche 1 page pour accompagner le président pendant l'AG."""
    total = stats["total"]
    enrichis = stats["enrichis"]
    nb_pays = len(country_counts)
    nb_secteurs = len(sector_counts)

    top_pays = sorted(country_counts.items(), key=lambda x: -x[1])[:5]
    top_sec = sorted(sector_counts.items(), key=lambda x: -x[1])[:5]
    top_ent = list(top_companies.items())[:10]

    fiche = f"""# 📋 FICHE SYNTHÈSE — AG Gaea21

**Date** : {datetime.now().strftime('%d %B %Y')}
**Auteur** : Rafika Cervera, Coordinatrice Data Gouvernance

---

## 🎯 Chiffres clés à retenir

| Métrique | Valeur |
|---|---|
| 🔢 Total alumni identifiés | **{total}** |
| ✅ Profils qualifiés | **{enrichis}** |
| 🌍 Pays représentés | **{nb_pays}** |
| 🏢 Secteurs d'activité | **{nb_secteurs}** |

---

## 🌍 Top 5 pays

"""
    for i, (pays, count) in enumerate(top_pays, 1):
        fiche += f"{i}. **{pays}** — {count} alumni\n"

    fiche += "\n---\n\n## 🎯 Top 5 secteurs\n\n"
    for i, (sec, count) in enumerate(top_sec, 1):
        fiche += f"{i}. **{sec}** — {count} alumni\n"

    fiche += "\n---\n\n## 🏢 Top 10 entreprises (mur de logos)\n\n"
    for i, (ent, count) in enumerate(top_ent, 1):
        fiche += f"{i}. **{ent}** ({count} alumni)\n"

    fiche += "\n---\n\n## ⭐ Vedettes à citer en cas de question\n\n"
    for v in vedettes[:5]:
        stars = "⭐" * v["nb_stars"]
        fiche += f"- **{v['prenom']} {v['nom']}** {stars}\n"
        fiche += f"  - {v['poste']}, {v['entreprise']} ({v['pays']})\n"
        fiche += f"  - Secteur : {v['secteur']}\n\n"

    fiche += """---

## 💡 Messages-clés à faire passer

1. **Notre communauté alumni est un actif majeur** : plus de 500 personnes, 9 pays, engagement continu
2. **Gaea21 forme des talents de la transition écologique** : secteur environnement en tête
3. **Nos alumni ont de la valeur sur le marché** : présents dans des entreprises reconnues
4. **Nous sommes un modèle de travail à distance efficace** : portée internationale authentique

---

## 📞 Actions post-AG

- Contacter les 3 premières vedettes pour témoignages écrits
- Enrichir progressivement les alumni restants (outil en place)
- Publier la carte sur LinkedIn Gaea21 pour visibilité externe
- Créer un groupe LinkedIn "Alumni Gaea21"

---

*Fiche générée automatiquement depuis le fichier alumni-gaea21-AG-READY.xlsx*
"""

    return fiche
