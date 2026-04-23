"""Visualisations pour l'AG Gaea21.

Génère avec Plotly :
- Carte du monde par pays
- Donut secteurs
- Timeline par année de sortie
- Top entreprises (mur de logos)
"""

from typing import Optional
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px


# ============================================================================
# COORDONNÉES GPS PAR PAYS (pour la carte du monde)
# ============================================================================

COUNTRY_COORDS = {
    "Suisse": (46.8182, 8.2275),
    "France": (46.2276, 2.2137),
    "Allemagne": (51.1657, 10.4515),
    "Belgique": (50.5039, 4.4699),
    "Italie": (41.8719, 12.5674),
    "Royaume-Uni": (55.3781, -3.4360),
    "Espagne": (40.4637, -3.7492),
    "Portugal": (39.3999, -8.2245),
    "États-Unis": (37.0902, -95.7129),
    "Canada": (56.1304, -106.3468),
    "Brésil": (-14.2350, -51.9253),
    "Kenya": (-0.0236, 37.9062),
    "Sénégal": (14.4974, -14.4524),
    "Maroc": (31.7917, -7.0926),
    "Tunisie": (33.8869, 9.5375),
    "Côte d'Ivoire": (7.5399, -5.5471),
    "Burkina Faso": (12.2383, -1.5616),
    "Algérie": (28.0339, 1.6596),
    "Togo": (8.6195, 0.8248),
    "Bénin": (9.3077, 2.3158),
    "Luxembourg": (49.8153, 6.1296),
    "Mexique": (23.6345, -102.5528),
    "Autre": (0, 0),
}

# Couleurs thématiques Gaea21 (palette verte environnementale)
GAEA_COLORS = {
    "primary": "#2E7D32",      # vert foncé
    "secondary": "#66BB6A",    # vert clair
    "accent": "#1B5E20",       # vert très foncé
    "neutral": "#78909C",      # gris-bleu
    "highlight": "#FFA726",    # orange (accent)
}

# Palette pour camemberts/donuts (couleurs cohérentes par secteur)
SECTOR_COLORS = {
    "Environnement / Climat": "#2E7D32",       # vert foncé
    "Finance durable / ESG": "#1B5E20",        # vert très foncé
    "Tech / Data": "#1976D2",                  # bleu
    "Conseil / Audit": "#5E35B1",              # violet
    "Éducation / Recherche": "#FFA726",        # orange
    "ONG / Associatif": "#66BB6A",             # vert clair
    "Secteur public": "#546E7A",               # gris-bleu
    "Entreprise privée": "#EC407A",            # rose
    "Santé": "#EF5350",                        # rouge
    "Culture / Média": "#AB47BC",              # violet clair
    "Autre": "#90A4AE",                        # gris
}


# ============================================================================
# CARTE DU MONDE
# ============================================================================

def build_world_map(country_counts: dict, title: str = "🌍 Alumni Gaea21 dans le monde") -> go.Figure:
    """Crée une carte du monde avec bulles par pays.

    Args:
        country_counts: dict {pays: nombre} venant de get_by_country()
    """
    # Prépare les données avec coordonnées
    rows = []
    for country, count in country_counts.items():
        if country in COUNTRY_COORDS:
            lat, lon = COUNTRY_COORDS[country]
            rows.append({
                "Pays": country,
                "Alumni": count,
                "lat": lat,
                "lon": lon,
            })

    if not rows:
        # Carte vide avec message
        fig = go.Figure()
        fig.add_annotation(
            text="Aucune donnée géographique à afficher",
            showarrow=False, xref="paper", yref="paper", x=0.5, y=0.5,
            font=dict(size=16)
        )
        return fig

    df_map = pd.DataFrame(rows)

    fig = go.Figure()

    fig.add_trace(go.Scattergeo(
        lon=df_map["lon"],
        lat=df_map["lat"],
        text=df_map.apply(
            lambda r: f"<b>{r['Pays']}</b><br>{r['Alumni']} alumni", axis=1
        ),
        marker=dict(
            size=df_map["Alumni"] * 8 + 10,  # taille proportionnelle
            color=GAEA_COLORS["primary"],
            line=dict(width=2, color=GAEA_COLORS["accent"]),
            sizemode="diameter",
            opacity=0.75,
        ),
        hoverinfo="text",
        name="Alumni",
    ))

    fig.update_geos(
        showland=True, landcolor="#F5F5F5",
        showocean=True, oceancolor="#E8F5E9",
        showcountries=True, countrycolor="#BDBDBD",
        showcoastlines=False,
        projection_type="natural earth",
        showframe=False,
    )

    fig.update_layout(
        title=dict(text=title, font=dict(size=20), x=0.5),
        geo=dict(
            scope="world",
            projection_scale=1.2,
            center=dict(lat=30, lon=10),  # centré sur l'Europe/Afrique
        ),
        margin=dict(l=0, r=0, t=50, b=0),
        height=500,
        paper_bgcolor="white",
    )

    return fig


# ============================================================================
# DONUT SECTEURS
# ============================================================================

def build_sector_donut(sector_counts: dict, title: str = "🎯 Répartition par secteur d'activité") -> go.Figure:
    """Crée un donut chart des secteurs."""
    if not sector_counts:
        fig = go.Figure()
        fig.add_annotation(
            text="Aucune donnée de secteur à afficher",
            showarrow=False, xref="paper", yref="paper", x=0.5, y=0.5,
            font=dict(size=16)
        )
        return fig

    labels = list(sector_counts.keys())
    values = list(sector_counts.values())
    colors = [SECTOR_COLORS.get(l, "#90A4AE") for l in labels]

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.5,
        marker=dict(colors=colors, line=dict(color="white", width=2)),
        textposition="outside",
        textinfo="label+percent",
        textfont=dict(size=13),
        hovertemplate="<b>%{label}</b><br>%{value} alumni (%{percent})<extra></extra>",
    )])

    total = sum(values)
    fig.update_layout(
        title=dict(text=title, font=dict(size=20), x=0.5),
        annotations=[dict(
            text=f"<b>{total}</b><br>alumni",
            x=0.5, y=0.5,
            font=dict(size=22, color=GAEA_COLORS["primary"]),
            showarrow=False,
        )],
        showlegend=True,
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05),
        margin=dict(l=20, r=150, t=60, b=20),
        height=450,
        paper_bgcolor="white",
    )

    return fig


# ============================================================================
# TIMELINE PAR ANNÉE
# ============================================================================

def build_timeline(df: pd.DataFrame, title: str = "📅 Sorties Gaea21 par année") -> go.Figure:
    """Crée un bar chart des sorties par année (basé sur toute la base)."""
    annee_col = "Année sortie" if "Année sortie" in df.columns else None
    if annee_col is None:
        # Essaie d'autres noms
        for c in df.columns:
            if "nn" in str(c).lower() and "sort" in str(c).lower():
                annee_col = c
                break

    if annee_col is None:
        fig = go.Figure()
        fig.add_annotation(text="Colonne année introuvable", showarrow=False,
                           xref="paper", yref="paper", x=0.5, y=0.5)
        return fig

    # Compte par année
    df_filter = df[df[annee_col].notna()].copy()
    df_filter[annee_col] = df_filter[annee_col].astype(float).astype(int)
    counts = df_filter[annee_col].value_counts().sort_index()

    fig = go.Figure(data=[go.Bar(
        x=[str(y) for y in counts.index],
        y=counts.values,
        marker=dict(
            color=counts.values,
            colorscale=[[0, GAEA_COLORS["secondary"]], [1, GAEA_COLORS["accent"]]],
            line=dict(color=GAEA_COLORS["accent"], width=1),
        ),
        text=counts.values,
        textposition="outside",
        textfont=dict(size=14, color=GAEA_COLORS["accent"]),
        hovertemplate="<b>%{x}</b><br>%{y} sorties<extra></extra>",
    )])

    fig.update_layout(
        title=dict(text=title, font=dict(size=20), x=0.5),
        xaxis=dict(title="Année de sortie", tickfont=dict(size=13)),
        yaxis=dict(title="Nombre d'alumni", showgrid=True, gridcolor="#E0E0E0"),
        margin=dict(l=50, r=30, t=60, b=50),
        height=400,
        paper_bgcolor="white",
        plot_bgcolor="white",
        bargap=0.3,
    )

    return fig


# ============================================================================
# TOP ENTREPRISES (MUR DE LOGOS)
# ============================================================================

def build_top_companies(top_companies: dict, title: str = "🏢 Top entreprises où nos alumni travaillent") -> go.Figure:
    """Crée un bar chart horizontal des top entreprises."""
    if not top_companies:
        fig = go.Figure()
        fig.add_annotation(text="Aucune entreprise à afficher", showarrow=False,
                           xref="paper", yref="paper", x=0.5, y=0.5)
        return fig

    # Trie par nombre décroissant
    items = sorted(top_companies.items(), key=lambda x: x[1], reverse=True)
    labels = [str(k)[:40] for k, _ in items]  # tronque si trop long
    values = [v for _, v in items]

    fig = go.Figure(data=[go.Bar(
        x=values,
        y=labels,
        orientation="h",
        marker=dict(
            color=values,
            colorscale=[[0, GAEA_COLORS["secondary"]], [1, GAEA_COLORS["primary"]]],
            line=dict(color=GAEA_COLORS["accent"], width=1),
        ),
        text=values,
        textposition="outside",
        textfont=dict(size=12, color=GAEA_COLORS["accent"]),
        hovertemplate="<b>%{y}</b><br>%{x} alumni<extra></extra>",
    )])

    fig.update_layout(
        title=dict(text=title, font=dict(size=20), x=0.5),
        xaxis=dict(title="Nombre d'alumni", showgrid=True, gridcolor="#E0E0E0"),
        yaxis=dict(title="", autorange="reversed", tickfont=dict(size=12)),
        margin=dict(l=200, r=50, t=60, b=50),
        height=max(300, len(items) * 35 + 100),
        paper_bgcolor="white",
        plot_bgcolor="white",
    )

    return fig


# ============================================================================
# EXPORT PNG
# ============================================================================

def fig_to_png_bytes(fig: go.Figure, width: int = 1200, height: int = 800) -> Optional[bytes]:
    """Convertit une figure Plotly en bytes PNG pour téléchargement.

    Nécessite kaleido (sera installé automatiquement).
    """
    try:
        return fig.to_image(format="png", width=width, height=height, scale=2)
    except Exception as e:
        return None
