"""
PSG Social Media Analytics Dashboard
Analyse comparative des clubs de football sur TikTok, Instagram et X (Twitter)
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
import warnings
warnings.filterwarnings("ignore")

# ─── Config page ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PSG Social Media Analytics",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Couleurs clubs ───────────────────────────────────────────────────────────
CLUB_COLORS = {
    "Paris Saint-Germain": "#004170",   # PSG bleu
    "Real Madrid C.F.":    "#FEBE10",   # RM or
    "Borussia Dortmund":   "#FDE100",   # BVB jaune
    "Tottenham Hotspur":   "#132257",   # Spurs marine
}
CLUB_COLORS_LIGHT = {
    "Paris Saint-Germain": "#004170",
    "Real Madrid C.F.":    "#C8A415",
    "Borussia Dortmund":   "#C8B400",
    "Tottenham Hotspur":   "#132257",
}
CLUB_SHORT = {
    "Paris Saint-Germain": "PSG",
    "Real Madrid C.F.":    "Real Madrid",
    "Borussia Dortmund":   "Dortmund",
    "Tottenham Hotspur":   "Tottenham",
}
PSG_COLOR   = "#004170"
PSG_ACCENT  = "#DA291C"
DARK_BG     = "#0A0E1A"
CARD_BG     = "#111827"

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
    background-color: #0A0E1A;
    color: #E8EAF0;
}

.stApp { background-color: #0A0E1A; }
.block-container { padding: 1.5rem 2rem 2rem; max-width: 1400px; }

/* Hero banner */
.hero {
    background: linear-gradient(135deg, #004170 0%, #001A35 50%, #DA291C22 100%);
    border: 1px solid #004170;
    border-radius: 16px;
    padding: 2.5rem 3rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: "PSG";
    position: absolute;
    right: 3rem;
    top: 50%;
    transform: translateY(-50%);
    font-family: 'Bebas Neue', sans-serif;
    font-size: 8rem;
    color: rgba(255,255,255,0.04);
    pointer-events: none;
}
.hero h1 {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.8rem;
    letter-spacing: 3px;
    margin: 0;
    color: white;
}
.hero p { color: #94A3B8; margin: 0.4rem 0 0; font-size: 0.95rem; }

/* KPI cards */
.kpi-card {
    background: #111827;
    border: 1px solid #1E2A3A;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    text-align: center;
    transition: border-color 0.2s;
}
.kpi-card:hover { border-color: #004170; }
.kpi-value {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.2rem;
    color: white;
    line-height: 1;
}
.kpi-label { font-size: 0.78rem; color: #64748B; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
.kpi-delta { font-size: 0.8rem; margin-top: 6px; }
.kpi-delta.up   { color: #22C55E; }
.kpi-delta.down { color: #EF4444; }

/* Section headers */
.section-header {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.6rem;
    letter-spacing: 2px;
    color: white;
    border-left: 4px solid #DA291C;
    padding-left: 12px;
    margin: 2rem 0 1rem;
}

/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    background: #111827;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
    border: 1px solid #1E2A3A;
}
.stTabs [data-baseweb="tab"] {
    background: transparent;
    color: #64748B;
    border-radius: 8px;
    font-family: 'DM Sans', sans-serif;
    font-weight: 500;
    font-size: 0.85rem;
    padding: 8px 20px;
    border: none;
}
.stTabs [aria-selected="true"] {
    background: #004170 !important;
    color: white !important;
}

/* Plotly charts */
.js-plotly-plot { border-radius: 12px; }

/* Sidebar */
.css-1d391kg { background: #0D1526; }

div[data-testid="metric-container"] {
    background: #111827;
    border: 1px solid #1E2A3A;
    border-radius: 10px;
    padding: 0.8rem 1rem;
}

hr { border-color: #1E2A3A; }
</style>
""", unsafe_allow_html=True)

# ─── Chargement données ────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    df_tt = pd.read_excel("Group E_IUT_tiktok.xlsx")
    df_ig = pd.read_excel("Groupe E_IUT_insta_X.xlsx")

    # TikTok : nettoyage
    df_tt["Published_Date"] = pd.to_datetime(df_tt["Published_Date"])
    df_tt["day_of_week"]    = df_tt["Published_Date"].dt.day_name()
    df_tt["month"]          = df_tt["Published_Date"].dt.month
    df_tt["week"]           = df_tt["Published_Date"].dt.isocalendar().week.astype(int)
    df_tt["Club"] = df_tt["Creator"].map({
        "psg": "Paris Saint-Germain",
        "spursofficial": "Tottenham Hotspur",
        "Real Madrid C.F.": "Real Madrid C.F.",
        "Borussia Dortmund": "Borussia Dortmund",
    }).fillna(df_tt["Creator"])

    # Insta/X : nettoyage
    df_ig["Date"]        = pd.to_datetime(df_ig["Date"])
    df_ig["day_of_week"] = df_ig["Date"].dt.day_name()
    df_ig["month"]       = df_ig["Date"].dt.month
    df_ig["week"]        = df_ig["Date"].dt.isocalendar().week.astype(int)
    df_ig["Club"]        = df_ig["Profile name"]
    df_ig["Virality Rate"] = pd.to_numeric(df_ig["Virality Rate"], errors="coerce")
    df_ig["Total interactions"] = pd.to_numeric(df_ig["Total interactions"], errors="coerce")
    df_ig["Total shares"]       = pd.to_numeric(df_ig["Total shares"], errors="coerce")
    df_ig["Total reactions"]    = pd.to_numeric(df_ig["Total reactions"], errors="coerce")
    df_ig["Interactions per 1000 followers"] = pd.to_numeric(
        df_ig["Interactions per 1000 followers"], errors="coerce")

    return df_tt, df_ig

df_tt, df_ig = load_data()

# ─── Helpers graphiques ────────────────────────────────────────────────────────
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="DM Sans", color="#94A3B8", size=12),
    margin=dict(l=40, r=20, t=50, b=40),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#CBD5E1")),
    xaxis=dict(gridcolor="#1E2A3A", zerolinecolor="#1E2A3A"),
    yaxis=dict(gridcolor="#1E2A3A", zerolinecolor="#1E2A3A"),
)

def apply_layout(fig, title=""):
    fig.update_layout(**PLOTLY_LAYOUT, title=dict(text=title, font=dict(
        family="Bebas Neue", size=18, color="white"), x=0))
    return fig

def club_color_list(clubs):
    return [CLUB_COLORS.get(c, "#64748B") for c in clubs]

# ─── HERO ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>⚽ PSG — Social Media Analytics</h1>
    <p>Analyse comparative des performances digitales · Janvier–Mars 2026 · TikTok · Instagram · X</p>
</div>
""", unsafe_allow_html=True)

# ─── KPI PSG ──────────────────────────────────────────────────────────────────
psg_tt = df_tt[df_tt["Club"] == "Paris Saint-Germain"]
psg_ig = df_ig[(df_ig["Club"] == "Paris Saint-Germain") & (df_ig["Platform"] == "instagram")]
psg_x  = df_ig[(df_ig["Club"] == "Paris Saint-Germain") & (df_ig["Platform"] == "twitter")]

avg_views_tt   = psg_tt["Views"].mean()
avg_eng_ig     = psg_ig["Total interactions"].mean()
avg_inter_1k   = psg_ig["Interactions per 1000 followers"].mean()
vir_rate       = psg_ig["Virality Rate"].mean()

# Comparaison avec moyenne concurrents
competitors = ["Real Madrid C.F.", "Borussia Dortmund", "Tottenham Hotspur"]
comp_tt   = df_tt[df_tt["Club"].isin(competitors)]
comp_ig   = df_ig[(df_ig["Club"].isin(competitors)) & (df_ig["Platform"] == "instagram")]

avg_views_comp = comp_tt["Views"].mean()
avg_eng_comp   = comp_ig["Total interactions"].mean()

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    delta_v = ((avg_views_tt - avg_views_comp) / avg_views_comp * 100)
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-value">{avg_views_tt/1e6:.1f}M</div>
        <div class="kpi-label">Vues moy. TikTok</div>
        <div class="kpi-delta {'up' if delta_v>0 else 'down'}">{'▲' if delta_v>0 else '▼'} {abs(delta_v):.0f}% vs concurrents</div>
    </div>""", unsafe_allow_html=True)
with col2:
    delta_e = ((avg_eng_ig - avg_eng_comp) / avg_eng_comp * 100)
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-value">{avg_eng_ig/1e3:.0f}K</div>
        <div class="kpi-label">Engagement moy. Instagram</div>
        <div class="kpi-delta {'up' if delta_e>0 else 'down'}">{'▲' if delta_e>0 else '▼'} {abs(delta_e):.0f}% vs concurrents</div>
    </div>""", unsafe_allow_html=True)
with col3:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-value">{avg_inter_1k:.2f}</div>
        <div class="kpi-label">Inter. / 1000 followers</div>
        <div class="kpi-delta">Instagram</div>
    </div>""", unsafe_allow_html=True)
with col4:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-value">{vir_rate:.3f}</div>
        <div class="kpi-label">Virality Rate</div>
        <div class="kpi-delta">Partages / Total interactions</div>
    </div>""", unsafe_allow_html=True)
with col5:
    total_posts = len(psg_tt) + len(psg_ig) + len(psg_x)
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-value">{total_posts}</div>
        <div class="kpi-label">Posts totaux</div>
        <div class="kpi-delta">Jan–Mar 2026</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── ONGLETS ──────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "Performance globale",
    "TikTok",
    "Instagram & X",
    "Facteurs d'engagement",
    "Recommandations",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — PERFORMANCE GLOBALE
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<div class="section-header">Performance Globale — Comparaison des 4 clubs</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    # Engagement moyen par club (Instagram)
    with col_a:
        eng_ig = df_ig[df_ig["Platform"] == "instagram"].groupby("Club")["Total interactions"].mean().reset_index()
        eng_ig = eng_ig.sort_values("Total interactions", ascending=True)
        fig = go.Figure(go.Bar(
            x=eng_ig["Total interactions"],
            y=[CLUB_SHORT.get(c, c) for c in eng_ig["Club"]],
            orientation="h",
            marker_color=[CLUB_COLORS.get(c, "#64748B") for c in eng_ig["Club"]],
            text=[f"{v/1e3:.0f}K" for v in eng_ig["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Engagement moyen Instagram par club")
        fig.update_xaxes(title="Total interactions moyen")
        st.plotly_chart(fig, use_container_width=True)

    # Interactions / 1000 followers Instagram
    with col_b:
        i1k = df_ig[df_ig["Platform"] == "instagram"].groupby("Club")["Interactions per 1000 followers"].mean().reset_index()
        i1k = i1k.sort_values("Interactions per 1000 followers", ascending=True)
        fig2 = go.Figure(go.Bar(
            x=i1k["Interactions per 1000 followers"],
            y=[CLUB_SHORT.get(c, c) for c in i1k["Club"]],
            orientation="h",
            marker_color=[CLUB_COLORS.get(c, "#64748B") for c in i1k["Club"]],
            text=[f"{v:.2f}" for v in i1k["Interactions per 1000 followers"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig2, "Interactions / 1000 followers — Instagram")
        fig2.update_xaxes(title="Interactions pour 1000 abonnés")
        st.plotly_chart(fig2, use_container_width=True)

    col_c, col_d = st.columns(2)

    # Vues moyennes TikTok
    with col_c:
        views_tt = df_tt.groupby("Club")["Views"].mean().reset_index().sort_values("Views", ascending=True)
        fig3 = go.Figure(go.Bar(
            x=views_tt["Views"],
            y=[CLUB_SHORT.get(c, c) for c in views_tt["Club"]],
            orientation="h",
            marker_color=[CLUB_COLORS.get(c, "#64748B") for c in views_tt["Club"]],
            text=[f"{v/1e6:.1f}M" for v in views_tt["Views"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig3, "Vues moyennes TikTok par club")
        fig3.update_xaxes(title="Vues moyennes")
        st.plotly_chart(fig3, use_container_width=True)

    # Virality Rate Instagram
    with col_d:
        vr = df_ig[df_ig["Platform"] == "instagram"].groupby("Club")["Virality Rate"].mean().reset_index()
        vr = vr.dropna().sort_values("Virality Rate", ascending=True)
        fig4 = go.Figure(go.Bar(
            x=vr["Virality Rate"],
            y=[CLUB_SHORT.get(c, c) for c in vr["Club"]],
            orientation="h",
            marker_color=[CLUB_COLORS.get(c, "#64748B") for c in vr["Club"]],
            text=[f"{v:.4f}" for v in vr["Virality Rate"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig4, "Virality Rate moyen — Instagram")
        fig4.update_xaxes(title="Virality Rate (partages/interactions)")
        st.plotly_chart(fig4, use_container_width=True)

    # Fréquence de publication par club et plateforme
    st.markdown('<div class="section-header">Fréquence de publication</div>', unsafe_allow_html=True)

    col_e, col_f = st.columns(2)
    with col_e:
        freq_ig = df_ig.groupby(["Club", "Platform"]).size().reset_index(name="nb_posts")
        freq_ig["Club_short"] = freq_ig["Club"].map(CLUB_SHORT)
        fig5 = px.bar(freq_ig, x="Club_short", y="nb_posts", color="Platform",
                      color_discrete_map={"instagram": "#E1306C", "twitter": "#1DA1F2"},
                      barmode="group")
        apply_layout(fig5, "Nombre de posts par club (Instagram & X)")
        fig5.update_xaxes(title="")
        fig5.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig5, use_container_width=True)

    with col_f:
        freq_tt = df_tt.groupby("Club").size().reset_index(name="nb_posts")
        freq_tt["Club_short"] = freq_tt["Club"].map(CLUB_SHORT)
        fig6 = go.Figure(go.Bar(
            x=freq_tt["Club_short"],
            y=freq_tt["nb_posts"],
            marker_color=[CLUB_COLORS.get(c, "#64748B") for c in freq_tt["Club"]],
            text=freq_tt["nb_posts"], textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig6, "Nombre de posts TikTok par club")
        fig6.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig6, use_container_width=True)

    # Répartition des posts par club
    col_g, col_h = st.columns(2)
    with col_g:
        rep_ig = df_ig.groupby("Club").size().reset_index(name="nb_posts")
        fig7 = go.Figure(go.Pie(
            labels=[CLUB_SHORT.get(c, c) for c in rep_ig["Club"]],
            values=rep_ig["nb_posts"],
            marker_colors=[CLUB_COLORS.get(c, "#64748B") for c in rep_ig["Club"]],
            hole=0.45,
            textinfo="label+percent",
            textfont=dict(color="white"),
        ))
        apply_layout(fig7, "Répartition des posts Instagram+X par club")
        st.plotly_chart(fig7, use_container_width=True)

    with col_h:
        rep_tt = df_tt.groupby("Club").size().reset_index(name="nb_posts")
        fig8 = go.Figure(go.Pie(
            labels=[CLUB_SHORT.get(c, c) for c in rep_tt["Club"]],
            values=rep_tt["nb_posts"],
            marker_colors=[CLUB_COLORS.get(c, "#64748B") for c in rep_tt["Club"]],
            hole=0.45,
            textinfo="label+percent",
            textfont=dict(color="white"),
        ))
        apply_layout(fig8, "Répartition des posts TikTok par club")
        st.plotly_chart(fig8, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — TIKTOK
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="section-header">TikTok — Analyse détaillée</div>', unsafe_allow_html=True)

    # Evolution vues et engagement par jour
    col_a, col_b = st.columns([1, 3])
    with col_a:
        club_filter = st.selectbox("Club", ["Tous les clubs"] + list(df_tt["Club"].unique()),
                                   key="tt_club_filter")

    df_tt_f = df_tt if club_filter == "Tous les clubs" else df_tt[df_tt["Club"] == club_filter]

    # Courbe des vues par jour
    views_day = df_tt_f.groupby(df_tt_f["Published_Date"].dt.date).agg(
        Views=("Views", "sum"),
        Likes=("Likes", "sum"),
        Comments=("Comments", "sum"),
        Total_Engagements=("Total_Engagements", "sum"),
    ).reset_index()
    views_day.columns = ["Date", "Views", "Likes", "Comments", "Total_Engagements"]

    fig_tt1 = go.Figure()
    fig_tt1.add_trace(go.Scatter(
        x=views_day["Date"], y=views_day["Views"],
        mode="lines", name="Vues",
        line=dict(color="#004170", width=2),
        fill="tozeroy", fillcolor="rgba(0,65,112,0.15)",
    ))
    apply_layout(fig_tt1, "Evolution des vues TikTok par jour")
    fig_tt1.update_yaxes(title="Vues totales")
    st.plotly_chart(fig_tt1, use_container_width=True)

    # Engagement : Likes vs Comments
    col_c, col_d = st.columns(2)
    with col_c:
        fig_tt2 = go.Figure()
        fig_tt2.add_trace(go.Scatter(
            x=views_day["Date"], y=views_day["Likes"],
            mode="lines", name="Likes",
            line=dict(color="#DA291C", width=2),
            fill="tozeroy", fillcolor="rgba(218,41,28,0.12)",
        ))
        fig_tt2.add_trace(go.Scatter(
            x=views_day["Date"], y=views_day["Comments"],
            mode="lines", name="Comments",
            line=dict(color="#F59E0B", width=2),
        ))
        apply_layout(fig_tt2, "Engagement TikTok : Likes vs Comments")
        st.plotly_chart(fig_tt2, use_container_width=True)

    with col_d:
        fig_tt3 = go.Figure()
        fig_tt3.add_trace(go.Scatter(
            x=views_day["Date"], y=views_day["Total_Engagements"],
            mode="lines", name="Total Engagements",
            line=dict(color="#22C55E", width=2),
            fill="tozeroy", fillcolor="rgba(34,197,94,0.12)",
        ))
        apply_layout(fig_tt3, "Total Engagement TikTok par jour")
        st.plotly_chart(fig_tt3, use_container_width=True)

    # Durée vidéo vs vues
    st.markdown('<div class="section-header">Durée de vidéo vs Vues</div>', unsafe_allow_html=True)

    col_e, col_f = st.columns(2)
    with col_e:
        # Scatter durée/vues par club
        fig_dur = go.Figure()
        for club in df_tt["Club"].unique():
            sub = df_tt[df_tt["Club"] == club]
            fig_dur.add_trace(go.Scatter(
                x=sub["Duration (seconds)"],
                y=sub["Views"],
                mode="markers",
                name=CLUB_SHORT.get(club, club),
                marker=dict(color=CLUB_COLORS.get(club, "#64748B"), size=5, opacity=0.7),
            ))
        apply_layout(fig_dur, "Durée vidéo (sec) vs Vues TikTok")
        fig_dur.update_xaxes(title="Durée (secondes)")
        fig_dur.update_yaxes(title="Vues")
        st.plotly_chart(fig_dur, use_container_width=True)

    with col_f:
        # Vues moyennes par tranche de durée
        df_tt["duration_bucket"] = pd.cut(df_tt["Duration (seconds)"],
            bins=[0, 15, 30, 60, 120, 999],
            labels=["< 15s", "15–30s", "30–60s", "60–120s", "> 120s"])
        dur_views = df_tt.groupby("duration_bucket", observed=True)["Views"].mean().reset_index()
        fig_dur2 = go.Figure(go.Bar(
            x=dur_views["duration_bucket"].astype(str),
            y=dur_views["Views"],
            marker_color=["#004170", "#1A6A9A", "#2E8BC0", "#DA291C", "#7B2020"],
            text=[f"{v/1e6:.1f}M" for v in dur_views["Views"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_dur2, "Vues moyennes par durée de vidéo (tous clubs)")
        fig_dur2.update_yaxes(title="Vues moyennes")
        st.plotly_chart(fig_dur2, use_container_width=True)

    # Distribution posts par jour de la semaine
    st.markdown('<div class="section-header">Distribution par jour de la semaine</div>', unsafe_allow_html=True)

    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    day_labels = {"Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi",
                  "Thursday": "Jeudi", "Friday": "Vendredi", "Saturday": "Samedi", "Sunday": "Dimanche"}

    col_g, col_h = st.columns(2)
    with col_g:
        posts_dow = df_tt.groupby(["Club", "day_of_week"]).size().reset_index(name="nb_posts")
        posts_dow["day_fr"] = posts_dow["day_of_week"].map(day_labels)
        posts_dow["day_order"] = posts_dow["day_of_week"].map({d: i for i, d in enumerate(day_order)})
        posts_dow = posts_dow.sort_values("day_order")

        fig_dow = px.bar(posts_dow, x="day_fr", y="nb_posts",
                         color="Club",
                         color_discrete_map={c: CLUB_COLORS[c] for c in CLUB_COLORS},
                         barmode="group")
        apply_layout(fig_dow, "Distribution des posts TikTok par jour")
        fig_dow.update_xaxes(title="")
        fig_dow.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig_dow, use_container_width=True)

    with col_h:
        views_dow = df_tt.groupby("day_of_week")["Views"].mean().reset_index()
        views_dow["day_fr"] = views_dow["day_of_week"].map(day_labels)
        views_dow["day_order"] = views_dow["day_of_week"].map({d: i for i, d in enumerate(day_order)})
        views_dow = views_dow.sort_values("day_order")

        fig_vdow = go.Figure(go.Bar(
            x=views_dow["day_fr"], y=views_dow["Views"],
            marker_color=["#DA291C" if v == views_dow["Views"].max() else "#004170"
                          for v in views_dow["Views"]],
            text=[f"{v/1e6:.1f}M" for v in views_dow["Views"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_vdow, "Vues moyennes TikTok par jour de la semaine")
        fig_vdow.update_yaxes(title="Vues moyennes")
        st.plotly_chart(fig_vdow, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — INSTAGRAM & X
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="section-header">Instagram & X — Analyse détaillée</div>', unsafe_allow_html=True)

    platform_choice = st.radio("Plateforme", ["Instagram", "X (Twitter)", "Comparaison"],
                                horizontal=True, key="ig_platform")

    if platform_choice == "Instagram":
        df_plat = df_ig[df_ig["Platform"] == "instagram"]
        plat_color = "#E1306C"
    elif platform_choice == "X (Twitter)":
        df_plat = df_ig[df_ig["Platform"] == "twitter"]
        plat_color = "#1DA1F2"
    else:
        df_plat = df_ig
        plat_color = PSG_COLOR

    col_a, col_b = st.columns(2)

    # Evolution interactions par semaine
    with col_a:
        evo_week = df_plat.groupby(["Club", "week"])["Total interactions"].mean().reset_index()
        fig_ev = go.Figure()
        for club in evo_week["Club"].unique():
            sub = evo_week[evo_week["Club"] == club]
            fig_ev.add_trace(go.Scatter(
                x=sub["week"], y=sub["Total interactions"],
                mode="lines+markers",
                name=CLUB_SHORT.get(club, club),
                line=dict(color=CLUB_COLORS.get(club, "#64748B"), width=2),
                marker=dict(size=5),
            ))
        apply_layout(fig_ev, "Evolution interactions moyennes par semaine")
        fig_ev.update_xaxes(title="Semaine")
        fig_ev.update_yaxes(title="Interactions moyennes")
        st.plotly_chart(fig_ev, use_container_width=True)

    # Total interactions par club
    with col_b:
        total_int = df_plat.groupby("Club")["Total interactions"].sum().reset_index()
        total_int = total_int.sort_values("Total interactions", ascending=True)
        fig_ti = go.Figure(go.Bar(
            x=total_int["Total interactions"],
            y=[CLUB_SHORT.get(c, c) for c in total_int["Club"]],
            orientation="h",
            marker_color=[CLUB_COLORS.get(c, "#64748B") for c in total_int["Club"]],
            text=[f"{v/1e6:.1f}M" for v in total_int["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_ti, "Total interactions sur la période")
        st.plotly_chart(fig_ti, use_container_width=True)

    col_c, col_d = st.columns(2)

    # Virality Rate
    with col_c:
        vr_club = df_plat.groupby("Club")["Virality Rate"].mean().reset_index().dropna()
        vr_club = vr_club.sort_values("Virality Rate", ascending=True)
        fig_vr = go.Figure(go.Bar(
            x=vr_club["Virality Rate"],
            y=[CLUB_SHORT.get(c, c) for c in vr_club["Club"]],
            orientation="h",
            marker_color=[CLUB_COLORS.get(c, "#64748B") for c in vr_club["Club"]],
            text=[f"{v:.4f}" for v in vr_club["Virality Rate"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_vr, "Virality Rate moyen")
        st.plotly_chart(fig_vr, use_container_width=True)

    # Distribution media type
    with col_d:
        media_dist = df_plat.groupby(["Club", "Media type"]).size().reset_index(name="n")
        fig_mt = px.bar(media_dist, x=[CLUB_SHORT.get(c, c) for c in media_dist["Club"]],
                        y="n", color="Media type", barmode="stack",
                        color_discrete_sequence=px.colors.qualitative.Set2)
        apply_layout(fig_mt, "Répartition des types de média")
        fig_mt.update_xaxes(title="")
        fig_mt.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig_mt, use_container_width=True)

    # Evolution nb posts
    st.markdown('<div class="section-header">Evolution du nombre de posts</div>', unsafe_allow_html=True)

    evo_posts = df_plat.groupby(["Club", "week"]).size().reset_index(name="nb_posts")
    fig_ep = go.Figure()
    for club in evo_posts["Club"].unique():
        sub = evo_posts[evo_posts["Club"] == club]
        fig_ep.add_trace(go.Scatter(
            x=sub["week"], y=sub["nb_posts"],
            mode="lines+markers",
            name=CLUB_SHORT.get(club, club),
            line=dict(color=CLUB_COLORS.get(club, "#64748B"), width=2),
        ))
    apply_layout(fig_ep, "Evolution du nombre de posts par semaine")
    fig_ep.update_xaxes(title="Semaine")
    fig_ep.update_yaxes(title="Nombre de posts")
    st.plotly_chart(fig_ep, use_container_width=True)

    # Distribution posts par jour de la semaine
    st.markdown('<div class="section-header">Distribution par jour de la semaine</div>', unsafe_allow_html=True)

    col_e, col_f = st.columns(2)
    with col_e:
        pdow = df_plat.groupby("day_of_week").size().reset_index(name="n")
        pdow["day_fr"]    = pdow["day_of_week"].map(day_labels)
        pdow["day_order"] = pdow["day_of_week"].map({d: i for i, d in enumerate(day_order)})
        pdow = pdow.sort_values("day_order")
        fig_pdow = go.Figure(go.Bar(
            x=pdow["day_fr"], y=pdow["n"],
            marker_color=["#DA291C" if v == pdow["n"].max() else "#004170" for v in pdow["n"]],
            text=pdow["n"], textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_pdow, "Distribution des posts par jour de la semaine")
        st.plotly_chart(fig_pdow, use_container_width=True)

    with col_f:
        # Interactions par jour de la semaine
        idow = df_plat.groupby("day_of_week")["Total interactions"].mean().reset_index()
        idow["day_fr"]    = idow["day_of_week"].map(day_labels)
        idow["day_order"] = idow["day_of_week"].map({d: i for i, d in enumerate(day_order)})
        idow = idow.sort_values("day_order")
        fig_idow = go.Figure(go.Bar(
            x=idow["day_fr"], y=idow["Total interactions"],
            marker_color=["#DA291C" if v == idow["Total interactions"].max() else "#004170"
                          for v in idow["Total interactions"]],
            text=[f"{v/1e3:.0f}K" for v in idow["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_idow, "Interactions moyennes par jour de la semaine")
        st.plotly_chart(fig_idow, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — FACTEURS D'ENGAGEMENT
# ══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="section-header">Facteurs d\'engagement — Focus PSG</div>', unsafe_allow_html=True)

    psg_ig_full = df_ig[(df_ig["Club"] == "Paris Saint-Germain") & (df_ig["Platform"] == "instagram")]
    psg_x_full  = df_ig[(df_ig["Club"] == "Paris Saint-Germain") & (df_ig["Platform"] == "twitter")]

    col_a, col_b = st.columns(2)

    # Engagement par type de média (Instagram PSG)
    with col_a:
        eng_media = psg_ig_full.groupby("Media type")["Total interactions"].mean().reset_index()
        eng_media = eng_media.sort_values("Total interactions", ascending=True)
        fig_em = go.Figure(go.Bar(
            x=eng_media["Total interactions"],
            y=eng_media["Media type"],
            orientation="h",
            marker_color=["#DA291C" if v == eng_media["Total interactions"].max() else "#004170"
                          for v in eng_media["Total interactions"]],
            text=[f"{v/1e3:.0f}K" for v in eng_media["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_em, "PSG Instagram — Engagement par type de média")
        st.plotly_chart(fig_em, use_container_width=True)

    # Grade distribution PSG
    with col_b:
        grade_dist = psg_ig_full["Grade"].dropna().value_counts().reset_index()
        grade_dist.columns = ["Grade", "count"]
        grade_order = ["A", "B", "C", "D", "F"]
        grade_colors = {"A": "#22C55E", "B": "#86EFAC", "C": "#F59E0B", "D": "#EF4444", "F": "#7F1D1D"}
        grade_dist["color"] = grade_dist["Grade"].map(grade_colors).fillna("#64748B")
        grade_dist["order"] = grade_dist["Grade"].map({g: i for i, g in enumerate(grade_order)}).fillna(99)
        grade_dist = grade_dist.sort_values("order")
        fig_gr = go.Figure(go.Bar(
            x=grade_dist["Grade"], y=grade_dist["count"],
            marker_color=grade_dist["color"],
            text=grade_dist["count"], textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig_gr, "PSG Instagram — Distribution des Grades")
        st.plotly_chart(fig_gr, use_container_width=True)

    col_c, col_d = st.columns(2)

    # Sentiment PSG
    with col_c:
        sent = psg_ig_full["Sentiment"].dropna().value_counts().reset_index()
        sent.columns = ["Sentiment", "count"]
        sent_colors = {"positive": "#22C55E", "negative": "#EF4444", "neutral": "#64748B",
                       "Positive": "#22C55E", "Negative": "#EF4444", "Neutral": "#64748B"}
        fig_sent = go.Figure(go.Pie(
            labels=sent["Sentiment"],
            values=sent["count"],
            marker_colors=[sent_colors.get(s, "#004170") for s in sent["Sentiment"]],
            hole=0.45,
            textinfo="label+percent",
            textfont=dict(color="white"),
        ))
        apply_layout(fig_sent, "PSG Instagram — Sentiment des posts")
        st.plotly_chart(fig_sent, use_container_width=True)

    # Engagement par heure de publication
    with col_d:
        psg_ig_full = psg_ig_full.copy()
        psg_ig_full["hour"] = psg_ig_full["Date"].dt.hour
        hour_eng = psg_ig_full.groupby("hour")["Total interactions"].mean().reset_index()
        fig_hr = go.Figure(go.Scatter(
            x=hour_eng["hour"], y=hour_eng["Total interactions"],
            mode="lines+markers",
            line=dict(color="#DA291C", width=2),
            marker=dict(color="#004170", size=8),
            fill="tozeroy", fillcolor="rgba(218,41,28,0.1)",
        ))
        apply_layout(fig_hr, "PSG Instagram — Engagement moyen par heure de publication")
        fig_hr.update_xaxes(title="Heure de publication", tickvals=list(range(0, 24, 2)))
        fig_hr.update_yaxes(title="Interactions moyennes")
        st.plotly_chart(fig_hr, use_container_width=True)

    # Comparaison PSG vs concurrents par media type
    st.markdown('<div class="section-header">PSG vs Concurrents — Type de contenu</div>', unsafe_allow_html=True)

    ig_media_comp = df_ig[df_ig["Platform"] == "instagram"].groupby(
        ["Club", "Media type"])["Total interactions"].mean().reset_index()
    ig_media_comp["Club_short"] = ig_media_comp["Club"].map(CLUB_SHORT)

    fig_comp = px.bar(ig_media_comp, x="Media type", y="Total interactions",
                      color="Club_short",
                      color_discrete_map={CLUB_SHORT[c]: CLUB_COLORS[c] for c in CLUB_COLORS},
                      barmode="group")
    apply_layout(fig_comp, "Engagement moyen par type de média et par club (Instagram)")
    fig_comp.update_yaxes(title="Interactions moyennes")
    st.plotly_chart(fig_comp, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — RECOMMANDATIONS
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="section-header">Recommandations Business — PSG</div>', unsafe_allow_html=True)

    st.markdown("""
    <div style="background:#111827;border:1px solid #DA291C;border-radius:12px;padding:1.5rem;margin-bottom:1rem;">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.1rem;color:#DA291C;letter-spacing:1px;margin-bottom:0.8rem;">✅ À FAIRE</div>
        <ul style="color:#CBD5E1;font-size:0.88rem;line-height:1.8;padding-left:1.2rem;">
            <li>Privilégier le format <b style="color:white">reel</b> sur Instagram (+engagement)</li>
            <li>Publier davantage le <b style="color:white">Lundi</b> (meilleur engagement)</li>
            <li>Vidéos TikTok de <b style="color:white">60–120s</b> pour maximiser les vues</li>
            <li>Miser sur les posts en soirée (18h–22h)</li>
            <li>Augmenter la fréquence sur X lors des matchs</li>
            <li>Accentuer les contenus "coulisses" et joueurs stars</li>
        </ul>
    </div>

    <div style="background:#111827;border:1px solid #F59E0B;border-radius:12px;padding:1.5rem;margin-bottom:1rem;">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.1rem;color:#F59E0B;letter-spacing:1px;margin-bottom:0.8rem;">⚠️ À OPTIMISER</div>
        <ul style="color:#CBD5E1;font-size:0.88rem;line-height:1.8;padding-left:1.2rem;">
            <li>Le Virality Rate du PSG est faible — travailler le contenu "partageable"</li>
            <li>L'engagement Instagram/1000 followers est inférieur à Tottenham</li>
            <li>La fréquence de publication sur TikTok peut être augmentée</li>
            <li>Harmoniser les thématiques entre plateformes</li>
            <li>Tester des formats interactifs (sondages, Q&A)</li>
        </ul>
    </div>

    <div style="background:#111827;border:1px solid #64748B;border-radius:12px;padding:1.5rem;">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.1rem;color:#94A3B8;letter-spacing:1px;margin-bottom:0.8rem;">❌ À ÉVITER</div>
        <ul style="color:#CBD5E1;font-size:0.88rem;line-height:1.8;padding-left:1.2rem;">
            <li>Vidéos TikTok de plus de 2 minutes (sous-performance)</li>
            <li>Posts statiques isolés sur Instagram (faible engagement)</li>
            <li>Publications en milieu de nuit (0h–6h)</li>
            <li>Contenu générique sans personnalisation joueur</li>
            <li>Surpublier sur X les jours sans actualité</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    # Benchmark final
    st.markdown('<div class="section-header">Benchmark PSG vs Concurrents</div>', unsafe_allow_html=True)

    bench_data = []
    for club in df_ig["Club"].unique():
        ig_sub = df_ig[(df_ig["Club"]==club) & (df_ig["Platform"]=="instagram")]
        tt_sub = df_tt[df_tt["Club"]==club]
        bench_data.append({
            "Club": CLUB_SHORT.get(club, club),
            "Posts Instagram": len(ig_sub),
            "Eng. moy IG": round(ig_sub["Total interactions"].mean(), 0),
            "Inter/1k IG": round(ig_sub["Interactions per 1000 followers"].mean(), 3),
            "Virality Rate": round(ig_sub["Virality Rate"].mean(), 4),
            "Vues moy TikTok": round(tt_sub["Views"].mean(), 0),
            "Posts TikTok": len(tt_sub),
        })

    bench_df = pd.DataFrame(bench_data)
    st.dataframe(
        bench_df.style
        .highlight_max(subset=["Eng. moy IG", "Inter/1k IG", "Virality Rate", "Vues moy TikTok"],
                       color="#1A3A5C")
        .format({
            "Eng. moy IG": "{:,.0f}",
            "Inter/1k IG": "{:.3f}",
            "Virality Rate": "{:.4f}",
            "Vues moy TikTok": "{:,.0f}",
        }),
        use_container_width=True,
        hide_index=True,
    )

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center;color:#374151;font-size:0.8rem;padding:1rem 0;">
    PSG Social Media Analytics · Données Jan–Mar 2026 · Sources : TikTok, Instagram, X (Twitter)
</div>
""", unsafe_allow_html=True)
