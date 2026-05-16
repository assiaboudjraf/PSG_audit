"""
PSG Social Media Analytics Dashboard — SAE BUT3 IUT
====================================================

Question SAE : "Quels facteurs expliquent la performance de l'engagement du PSG
sur les réseaux sociaux par rapport à ses concurrents ?"

Structure (5 onglets, alignés sur le brief A / B / C) :
  1. Performance globale          → A. PSG vs concurrents (bench + radar)
  2. TikTok                       → A/B. métriques + cat_simple + durée
  3. Instagram & X                → A/B. métriques + cat_simple + media type
  4. Facteurs d'engagement        → B.   heatmap timing, matchday, sentiment
  5. Recommandations              → C.   reco chiffrées (dynamiques)

Données : data_clear/*_ready.xlsx (copies locales pour Streamlit Cloud).
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import warnings
warnings.filterwarnings("ignore")

# ═════════════════════════════════════════════════════════════════════════════
# CONFIG — source unique de vérité
# ═════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="PSG Social Media Analytics",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="collapsed",
)

TIKTOK_FILE  = "tiktok_ready.xlsx"
INSTA_X_FILE = "insta_x_ready.xlsx"

# Les clubs sont déjà normalisés en phase 0 (PSG / Real Madrid / Dortmund / Tottenham).
CLUBS = ["PSG", "Real Madrid", "Dortmund", "Tottenham"]
COMPETITORS = [c for c in CLUBS if c != "PSG"]

CLUB_COLORS = {
    "PSG":         "#004170",  # bleu PSG
    "Real Madrid": "#FEBE10",  # or
    "Dortmund":    "#FDE100",  # jaune BVB
    "Tottenham":   "#132257",  # marine Spurs
}

PSG_COLOR  = "#004170"
PSG_ACCENT = "#DA291C"
DARK_BG    = "#0A0E1A"
CARD_BG    = "#111827"

DAY_ORDER  = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_LABELS = {"Monday": "Lun", "Tuesday": "Mar", "Wednesday": "Mer",
              "Thursday": "Jeu", "Friday": "Ven", "Saturday": "Sam", "Sunday": "Dim"}

# ═════════════════════════════════════════════════════════════════════════════
# CSS — design dark / branding PSG
# ═════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; background-color: #0A0E1A; color: #E8EAF0; }
.stApp { background-color: #0A0E1A; }
.block-container { padding: 1.5rem 2rem 2rem; max-width: 1400px; }

/* Hero banner */
.hero {
    background: linear-gradient(135deg, #004170 0%, #001A35 50%, #DA291C22 100%);
    border: 1px solid #004170; border-radius: 16px;
    padding: 2.5rem 3rem; margin-bottom: 2rem; position: relative; overflow: hidden;
}
.hero::before {
    content: "PSG"; position: absolute; right: 3rem; top: 50%;
    transform: translateY(-50%); font-family: 'Bebas Neue', sans-serif;
    font-size: 8rem; color: rgba(255,255,255,0.04); pointer-events: none;
}
.hero h1 { font-family: 'Bebas Neue', sans-serif; font-size: 2.8rem; letter-spacing: 3px; margin: 0; color: white; }
.hero p  { color: #94A3B8; margin: 0.4rem 0 0; font-size: 0.95rem; }

/* KPI cards */
.kpi-card { background: #111827; border: 1px solid #1E2A3A; border-radius: 12px;
            padding: 1.2rem 1.5rem; text-align: center; transition: border-color 0.2s; }
.kpi-card:hover { border-color: #004170; }
.kpi-value { font-family: 'Bebas Neue', sans-serif; font-size: 2.2rem; color: white; line-height: 1; }
.kpi-label { font-size: 0.78rem; color: #64748B; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
.kpi-delta { font-size: 0.8rem; margin-top: 6px; }
.kpi-delta.up   { color: #22C55E; }
.kpi-delta.down { color: #EF4444; }

/* Section headers */
.section-header { font-family: 'Bebas Neue', sans-serif; font-size: 1.6rem; letter-spacing: 2px;
                  color: white; border-left: 4px solid #DA291C; padding-left: 12px; margin: 2rem 0 1rem; }

/* Insight callouts (note d'analyse style étudiant) */
.insight { background: #0F1A2E; padding: 0.8rem 1rem; border-radius: 6px;
           margin: 0.6rem 0 1.4rem; color: #CBD5E1; font-size: 0.88rem; }
.insight b { color: white; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { background: #111827; border-radius: 10px; padding: 4px; gap: 4px; border: 1px solid #1E2A3A; }
.stTabs [data-baseweb="tab"] { background: transparent; color: #64748B; border-radius: 8px;
    font-family: 'DM Sans', sans-serif; font-weight: 500; font-size: 0.85rem; padding: 8px 20px; border: none; }
.stTabs [aria-selected="true"] { background: #004170 !important; color: white !important; }

.js-plotly-plot { border-radius: 12px; }
hr { border-color: #1E2A3A; }
</style>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# CHARGEMENT & PRÉPARATION DES DONNÉES
# ═════════════════════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    """Lit les datasets *_ready.xlsx et ajoute les colonnes dérivées (jour,
    heure, weekend, matchday PSG). Les clubs sont déjà normalisés en phase 0."""
    df_tt = pd.read_excel(TIKTOK_FILE)
    df_ig = pd.read_excel(INSTA_X_FILE)

    # ── TikTok ────────────────────────────────────────────────────────────────
    df_tt["Published_Date"] = pd.to_datetime(df_tt["Published_Date"], errors="coerce")
    df_tt = df_tt.dropna(subset=["Published_Date"])
    df_tt["Club"]        = df_tt["Creator"].astype(str)
    df_tt["day_of_week"] = df_tt["Published_Date"].dt.day_name()
    df_tt["hour"]        = df_tt["Published_Date"].dt.hour
    df_tt["week"]        = df_tt["Published_Date"].dt.isocalendar().week.astype(int)
    df_tt["date_only"]   = df_tt["Published_Date"].dt.date

    # ── Instagram + X ─────────────────────────────────────────────────────────
    df_ig["Date"]        = pd.to_datetime(df_ig["Date"], errors="coerce")
    df_ig = df_ig.dropna(subset=["Date"])
    df_ig["Club"]        = df_ig["Creator"].astype(str)
    df_ig["day_of_week"] = df_ig["Date"].dt.day_name()
    df_ig["hour"]        = df_ig["Date"].dt.hour
    df_ig["week"]        = df_ig["Date"].dt.isocalendar().week.astype(int)
    df_ig["date_only"]   = df_ig["Date"].dt.date

    # Coercions numériques (les colonnes peuvent contenir des "—" ou strings)
    for col in ["Total interactions", "Interactions per 1000 followers",
                "Virality Rate", "Total shares", "Total reactions",
                "Positive comments (%)", "Negative comments (%)", "Neutral comments (%)"]:
        if col in df_ig.columns:
            df_ig[col] = pd.to_numeric(df_ig[col], errors="coerce")

    return df_tt, df_ig


df_tt, df_ig = load_data()

# Sous-ensembles fréquemment utilisés
psg_tt   = df_tt[df_tt["Club"] == "PSG"]
psg_ig   = df_ig[(df_ig["Club"] == "PSG") & (df_ig["Platform"] == "instagram")]
psg_x    = df_ig[(df_ig["Club"] == "PSG") & (df_ig["Platform"] == "twitter")]
comp_tt  = df_tt[df_tt["Club"].isin(COMPETITORS)]
comp_ig  = df_ig[(df_ig["Club"].isin(COMPETITORS)) & (df_ig["Platform"] == "instagram")]

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════
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
    fig.update_layout(**PLOTLY_LAYOUT,
        title=dict(text=title, font=dict(family="Bebas Neue", size=18, color="white"), x=0))
    return fig

def club_colors(clubs):
    return [CLUB_COLORS.get(c, "#64748B") for c in clubs]

def insight(text):
    """Encadré d'analyse sous un graphique."""
    st.markdown(f'<div class="insight">{text}</div>', unsafe_allow_html=True)

def fmt_k(v):
    if pd.isna(v): return "N/A"
    if abs(v) >= 1e6: return f"{v/1e6:.1f}M"
    if abs(v) >= 1e3: return f"{v/1e3:.0f}K"
    return f"{v:.0f}"

def kpi_card(value, label, delta=None, delta_up=True):
    delta_html = ""
    if delta is not None:
        cls = "up" if delta_up else "down"
        arrow = "▲" if delta_up else "▼"
        delta_html = f'<div class="kpi-delta {cls}">{arrow} {delta}</div>'
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-value">{value}</div>
        <div class="kpi-label">{label}</div>
        {delta_html}
    </div>""", unsafe_allow_html=True)

def order_by_day(df, col="day_of_week"):
    df = df.copy()
    df["__ord"] = df[col].map({d: i for i, d in enumerate(DAY_ORDER)})
    df["day_fr"] = df[col].map(DAY_LABELS)
    return df.sort_values("__ord")

# ═════════════════════════════════════════════════════════════════════════════
# HERO
# ═════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <h1>⚽ PSG — Social Media Analytics</h1>
    <p>Analyse comparative des performances digitales · Janvier–Mars 2026 · TikTok · Instagram · X</p>
</div>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# KPI HEADER — PSG vs moyenne concurrents (calculé dynamiquement)
# ═════════════════════════════════════════════════════════════════════════════
avg_views_tt   = psg_tt["Views"].mean()
avg_views_comp = comp_tt["Views"].mean()
delta_views    = (avg_views_tt - avg_views_comp) / avg_views_comp * 100 if avg_views_comp else 0

avg_eng_ig     = psg_ig["Total interactions"].mean()
avg_eng_comp   = comp_ig["Total interactions"].mean()
delta_eng      = (avg_eng_ig - avg_eng_comp) / avg_eng_comp * 100 if avg_eng_comp else 0

avg_inter_1k   = psg_ig["Interactions per 1000 followers"].mean()
comp_inter_1k  = comp_ig["Interactions per 1000 followers"].mean()
delta_i1k      = (avg_inter_1k - comp_inter_1k) / comp_inter_1k * 100 if comp_inter_1k else 0

vir_rate       = psg_ig["Virality Rate"].mean()
total_posts    = len(psg_tt) + len(psg_ig) + len(psg_x)

c1, c2, c3, c4, c5 = st.columns(5)
with c1: kpi_card(fmt_k(avg_views_tt), "Vues moy. TikTok",
                  f"{abs(delta_views):.0f}% vs concurrents", delta_views > 0)
with c2: kpi_card(fmt_k(avg_eng_ig), "Engagement moy. Instagram",
                  f"{abs(delta_eng):.0f}% vs concurrents", delta_eng > 0)
with c3: kpi_card(f"{avg_inter_1k:.2f}", "Inter. / 1000 followers",
                  f"{abs(delta_i1k):.0f}% vs concurrents", delta_i1k > 0)
with c4: kpi_card(f"{vir_rate:.3f}", "Virality Rate",
                  "Partages / interactions", True)
with c5: kpi_card(f"{total_posts:,}".replace(",", " "), "Posts totaux",
                  "Jan–Mar 2026", True)

st.markdown("<br>", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ═════════════════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "Performance globale",
    "TikTok",
    "Instagram & X",
    "Facteurs d'engagement",
    "Recommandations",
])

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — PERFORMANCE GLOBALE
# Brief A : engagement moyen, virality rate, inter/1k followers, fréquence
# ════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<div class="section-header">Performance Globale — 4 clubs</div>', unsafe_allow_html=True)

    # ── Engagement moyen IG + Inter/1k ──────────────────────────────────────
    col_a, col_b = st.columns(2)
    with col_a:
        eng_ig = (df_ig[df_ig["Platform"] == "instagram"]
                  .groupby("Club")["Total interactions"].mean()
                  .reindex(CLUBS).reset_index().sort_values("Total interactions"))
        fig = go.Figure(go.Bar(
            x=eng_ig["Total interactions"], y=eng_ig["Club"], orientation="h",
            marker_color=club_colors(eng_ig["Club"]),
            text=[fmt_k(v) for v in eng_ig["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Engagement moyen Instagram par club")
        fig.update_xaxes(title="Total interactions moyen")
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        i1k = (df_ig[df_ig["Platform"] == "instagram"]
               .groupby("Club")["Interactions per 1000 followers"].mean()
               .reindex(CLUBS).reset_index().sort_values("Interactions per 1000 followers"))
        fig = go.Figure(go.Bar(
            x=i1k["Interactions per 1000 followers"], y=i1k["Club"], orientation="h",
            marker_color=club_colors(i1k["Club"]),
            text=[f"{v:.2f}" for v in i1k["Interactions per 1000 followers"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Interactions / 1 000 followers — Instagram")
        fig.update_xaxes(title="Interactions / 1000 abonnés")
        st.plotly_chart(fig, use_container_width=True)

    insight(
        f"Si on regarde l'engagement <b>par 1000 followers</b> (pour comparer à audience égale), le PSG est à "
        f"<b>{avg_inter_1k:.2f}</b> alors que la moyenne des 3 autres est à <b>{comp_inter_1k:.2f}</b>. "
        f"En clair : l'engagement total du Real Madrid est plus gros parce qu'ils ont beaucoup plus d'abonnés, "
        f"mais une fois ramené à la taille de l'audience, l'écart change."
    )

    # ── Vues TikTok + Virality Rate ─────────────────────────────────────────
    col_c, col_d = st.columns(2)
    with col_c:
        views_tt = df_tt.groupby("Club")["Views"].mean().reindex(CLUBS).reset_index().sort_values("Views")
        fig = go.Figure(go.Bar(
            x=views_tt["Views"], y=views_tt["Club"], orientation="h",
            marker_color=club_colors(views_tt["Club"]),
            text=[fmt_k(v) for v in views_tt["Views"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Vues moyennes TikTok par club")
        fig.update_xaxes(title="Vues moyennes")
        st.plotly_chart(fig, use_container_width=True)

    with col_d:
        vr = (df_ig[df_ig["Platform"] == "instagram"]
              .groupby("Club")["Virality Rate"].mean()
              .reindex(CLUBS).dropna().reset_index().sort_values("Virality Rate"))
        fig = go.Figure(go.Bar(
            x=vr["Virality Rate"], y=vr["Club"], orientation="h",
            marker_color=club_colors(vr["Club"]),
            text=[f"{v:.4f}" for v in vr["Virality Rate"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Virality Rate moyen — Instagram")
        fig.update_xaxes(title="Virality Rate (partages / interactions)")
        st.plotly_chart(fig, use_container_width=True)

    # ── Fréquence de publication ────────────────────────────────────────────
    st.markdown('<div class="section-header">Fréquence de publication</div>', unsafe_allow_html=True)

    col_e, col_f = st.columns(2)
    with col_e:
        freq_ig = df_ig.groupby(["Club", "Platform"]).size().reset_index(name="nb_posts")
        fig = px.bar(freq_ig, x="Club", y="nb_posts", color="Platform",
                     color_discrete_map={"instagram": "#E1306C", "twitter": "#1DA1F2"},
                     barmode="group", category_orders={"Club": CLUBS})
        apply_layout(fig, "Posts par club — Instagram & X")
        fig.update_yaxes(title="Nombre de posts"); fig.update_xaxes(title="")
        st.plotly_chart(fig, use_container_width=True)

    with col_f:
        freq_tt = df_tt.groupby("Club").size().reindex(CLUBS).reset_index(name="nb_posts")
        fig = go.Figure(go.Bar(
            x=freq_tt["Club"], y=freq_tt["nb_posts"],
            marker_color=club_colors(freq_tt["Club"]),
            text=freq_tt["nb_posts"], textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Posts TikTok par club")
        fig.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig, use_container_width=True)

    # ── Répartition des posts par club (donuts) ─────────────────────────────
    col_g, col_h = st.columns(2)
    with col_g:
        rep_ig = df_ig.groupby("Club").size().reindex(CLUBS).reset_index(name="n")
        fig = go.Figure(go.Pie(
            labels=rep_ig["Club"], values=rep_ig["n"],
            marker_colors=club_colors(rep_ig["Club"]),
            hole=0.45, textinfo="label+percent", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Répartition des posts Instagram+X par club")
        st.plotly_chart(fig, use_container_width=True)
    with col_h:
        rep_tt = df_tt.groupby("Club").size().reindex(CLUBS).reset_index(name="n")
        fig = go.Figure(go.Pie(
            labels=rep_tt["Club"], values=rep_tt["n"],
            marker_colors=club_colors(rep_tt["Club"]),
            hole=0.45, textinfo="label+percent", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Répartition des posts TikTok par club")
        st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 2 — TIKTOK
# Brief B : durée vs vues (scatter), distribution timing, cat_simple
# ════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="section-header">TikTok — Analyse détaillée</div>', unsafe_allow_html=True)

    # Filtres
    col_f1, col_f2, _ = st.columns([1, 1, 2])
    with col_f1:
        club_filter = st.selectbox("Club", ["Tous les clubs"] + CLUBS, key="tt_club")
    with col_f2:
        cat_filter  = st.selectbox("Catégorie de contenu",
                                   ["Toutes catégories"] + sorted(df_tt["cat_simple"].dropna().unique().tolist()),
                                   key="tt_cat")

    df_tt_f = df_tt.copy()
    if club_filter != "Tous les clubs":
        df_tt_f = df_tt_f[df_tt_f["Club"] == club_filter]
    if cat_filter != "Toutes catégories":
        df_tt_f = df_tt_f[df_tt_f["cat_simple"] == cat_filter]

    # ── Vues totales cumulées par club (toute la période) ───────────────────
    total_views = (df_tt.groupby("Club")["Views"].sum()
                   .reindex(CLUBS).reset_index().sort_values("Views"))
    fig = go.Figure(go.Bar(
        x=total_views["Views"], y=total_views["Club"], orientation="h",
        marker_color=club_colors(total_views["Club"]),
        text=[fmt_k(v) for v in total_views["Views"]],
        textposition="outside", textfont=dict(color="white"),
    ))
    apply_layout(fig, "Vues TikTok cumulées par club (Jan–Mar 2026)")
    fig.update_xaxes(title="Vues totales")
    st.plotly_chart(fig, use_container_width=True)

    # ── Evolution vues / jour ────────────────────────────────────────────────
    views_day = (df_tt_f.groupby("date_only")
                 .agg(Views=("Views", "sum"),
                      Likes=("Likes", "sum"),
                      Comments=("Comments", "sum"),
                      Total_Engagements=("Total_Engagements", "sum"))
                 .reset_index())
    fig = go.Figure(go.Scatter(
        x=views_day["date_only"], y=views_day["Views"],
        mode="lines", line=dict(color="#004170", width=2),
        fill="tozeroy", fillcolor="rgba(0,65,112,0.15)", name="Vues",
    ))
    apply_layout(fig, "Évolution des vues TikTok par jour")
    fig.update_yaxes(title="Vues totales")
    st.plotly_chart(fig, use_container_width=True)

    # ── Likes vs Comments + Total Engagement par jour ───────────────────────
    col_a, col_b = st.columns(2)
    with col_a:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=views_day["date_only"], y=views_day["Likes"],
            mode="lines", name="Likes", line=dict(color="#DA291C", width=2),
            fill="tozeroy", fillcolor="rgba(218,41,28,0.12)"))
        fig.add_trace(go.Scatter(x=views_day["date_only"], y=views_day["Comments"],
            mode="lines", name="Comments", line=dict(color="#F59E0B", width=2)))
        apply_layout(fig, "Likes vs Comments (jour)")
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        fig = go.Figure(go.Scatter(
            x=views_day["date_only"], y=views_day["Total_Engagements"],
            mode="lines", name="Total Engagements",
            line=dict(color="#22C55E", width=2),
            fill="tozeroy", fillcolor="rgba(34,197,94,0.12)",
        ))
        apply_layout(fig, "Total Engagement TikTok par jour")
        fig.update_yaxes(title="Engagements totaux")
        st.plotly_chart(fig, use_container_width=True)

    # ── Scatter Durée × Vues (brief B) ──────────────────────────────────────
    col_dur, _ = st.columns([2, 1])
    with col_dur:

        fig = go.Figure()
        for club in CLUBS:
            sub = df_tt[df_tt["Club"] == club]
            fig.add_trace(go.Scatter(
                x=sub["Duration (seconds)"], y=sub["Views"],
                mode="markers", name=club,
                marker=dict(color=CLUB_COLORS[club], size=5, opacity=0.6),
            ))
        apply_layout(fig, "Durée vidéo vs Vues — scatter par club")
        fig.update_xaxes(title="Durée (secondes)")
        fig.update_yaxes(title="Vues", type="log")
        st.plotly_chart(fig, use_container_width=True)
    insight(
        "On voit clairement qu'il n'y a pas de lien direct entre la durée d'une vidéo et le nombre de vues. "
        "Une vidéo virale peut faire moins de 15s comme plus d'une minute — la durée seule n'explique pas la performance."
    )

    # ── Vues moyennes par tranche de durée ───────────────────────────────────
    df_tt_dur = df_tt.assign(
        duration_bucket=pd.cut(df_tt["Duration (seconds)"],
                               bins=[0, 15, 30, 60, 120, 9999],
                               labels=["< 15s", "15–30s", "30–60s", "60–120s", "> 120s"])
    )
    dur_views = df_tt_dur.groupby("duration_bucket", observed=True)["Views"].mean().reset_index()
    fig = go.Figure(go.Bar(
        x=dur_views["duration_bucket"].astype(str), y=dur_views["Views"],
        marker_color=["#DA291C" if v == dur_views["Views"].max() else "#004170" for v in dur_views["Views"]],
        text=[fmt_k(v) for v in dur_views["Views"]],
        textposition="outside", textfont=dict(color="white"),
    ))
    apply_layout(fig, "Vues moyennes par tranche de durée — tous clubs")
    fig.update_yaxes(title="Vues moyennes")
    st.plotly_chart(fig, use_container_width=True)

    # ── cat_simple : ranking PSG + comparatif clubs ─────────────────────────
    st.markdown('<div class="section-header">Performance par catégorie de contenu</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns(2)

    with col_c:
        # Vues moyennes par cat_simple — PSG
        psg_cat = (psg_tt.groupby("cat_simple")["Views"].mean()
                   .reset_index().sort_values("Views"))
        fig = go.Figure(go.Bar(
            x=psg_cat["Views"], y=psg_cat["cat_simple"], orientation="h",
            marker_color=["#DA291C" if v == psg_cat["Views"].max() else "#004170" for v in psg_cat["Views"]],
            text=[fmt_k(v) for v in psg_cat["Views"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "PSG — Vues moyennes TikTok par catégorie")
        fig.update_xaxes(title="Vues moyennes")
        st.plotly_chart(fig, use_container_width=True)

    with col_d:
        # Répartition des posts PSG par cat_simple (donut)
        psg_dist = psg_tt["cat_simple"].value_counts().reset_index()
        psg_dist.columns = ["cat_simple", "n"]
        fig = go.Figure(go.Pie(
            labels=psg_dist["cat_simple"], values=psg_dist["n"],
            hole=0.45, textinfo="label+percent",
            marker_colors=px.colors.qualitative.Set2, textfont=dict(color="white"),
        ))
        apply_layout(fig, "PSG — Répartition des posts TikTok par catégorie")
        st.plotly_chart(fig, use_container_width=True)

    # Comparatif catégories × clubs
    cat_club = df_tt.groupby(["Club", "cat_simple"])["Views"].mean().reset_index()
    fig = px.bar(cat_club, x="cat_simple", y="Views", color="Club",
                 color_discrete_map=CLUB_COLORS, barmode="group",
                 category_orders={"Club": CLUBS})
    apply_layout(fig, "Vues moyennes par catégorie × club (TikTok)")
    fig.update_xaxes(title=""); fig.update_yaxes(title="Vues moyennes")
    st.plotly_chart(fig, use_container_width=True)

    top_cat_psg = psg_cat.iloc[-1]
    insight(
        f"Sur TikTok, c'est la catégorie <b>{top_cat_psg['cat_simple']}</b> qui marche le mieux pour le PSG, "
        f"avec une moyenne de <b>{fmt_k(top_cat_psg['Views'])}</b> vues par post. "
        f"C'est un bon indice pour orienter ce qu'ils devraient publier en priorité."
    )

    # ── Distribution par jour de la semaine ─────────────────────────────────
    st.markdown('<div class="section-header">Distribution par jour de la semaine</div>', unsafe_allow_html=True)
    col_e, col_f = st.columns(2)
    with col_e:
        posts_dow = (df_tt.groupby(["Club", "day_of_week"]).size().reset_index(name="n"))
        posts_dow = order_by_day(posts_dow)
        fig = px.bar(posts_dow, x="day_fr", y="n", color="Club",
                     color_discrete_map=CLUB_COLORS, barmode="group",
                     category_orders={"Club": CLUBS})
        apply_layout(fig, "Posts TikTok par jour")
        fig.update_xaxes(title=""); fig.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig, use_container_width=True)

    with col_f:
        views_dow = order_by_day(df_tt.groupby("day_of_week")["Views"].mean().reset_index())
        fig = go.Figure(go.Bar(
            x=views_dow["day_fr"], y=views_dow["Views"],
            marker_color=["#DA291C" if v == views_dow["Views"].max() else "#004170" for v in views_dow["Views"]],
            text=[fmt_k(v) for v in views_dow["Views"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Vues moyennes TikTok par jour")
        fig.update_yaxes(title="Vues moyennes")
        st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 3 — INSTAGRAM & X
# Brief B : Media type vs interactions, cat_simple, evolution
# ════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="section-header">Instagram & X — Analyse détaillée</div>', unsafe_allow_html=True)

    platform_choice = st.radio("Plateforme",
        ["Instagram", "X (Twitter)", "Comparaison"], horizontal=True, key="ig_platform")

    if platform_choice == "Instagram":
        df_plat = df_ig[df_ig["Platform"] == "instagram"]
    elif platform_choice == "X (Twitter)":
        df_plat = df_ig[df_ig["Platform"] == "twitter"]
    else:
        df_plat = df_ig

    col_a, col_b = st.columns(2)
    with col_a:
        evo = df_plat.groupby(["Club", "week"])["Total interactions"].mean().reset_index()
        fig = go.Figure()
        for club in CLUBS:
            sub = evo[evo["Club"] == club]
            fig.add_trace(go.Scatter(
                x=sub["week"], y=sub["Total interactions"],
                mode="lines+markers", name=club,
                line=dict(color=CLUB_COLORS[club], width=2), marker=dict(size=5),
            ))
        apply_layout(fig, "Évolution interactions moyennes / semaine")
        fig.update_xaxes(title="Semaine"); fig.update_yaxes(title="Interactions moyennes")
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        ti = df_plat.groupby("Club")["Total interactions"].sum().reindex(CLUBS).reset_index().sort_values("Total interactions")
        fig = go.Figure(go.Bar(
            x=ti["Total interactions"], y=ti["Club"], orientation="h",
            marker_color=club_colors(ti["Club"]),
            text=[fmt_k(v) for v in ti["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Interactions cumulées par club (Jan–Mar 2026)")
        st.plotly_chart(fig, use_container_width=True)

    # ── Media type × interactions (cœur de B) ───────────────────────────────
    st.markdown('<div class="section-header">Type de média vs Interactions</div>', unsafe_allow_html=True)

    col_c, col_d = st.columns(2)
    with col_c:
        mt = df_plat.groupby("Media type")["Total interactions"].mean().reset_index().sort_values("Total interactions")
        fig = go.Figure(go.Bar(
            x=mt["Total interactions"], y=mt["Media type"], orientation="h",
            marker_color=["#DA291C" if v == mt["Total interactions"].max() else "#004170" for v in mt["Total interactions"]],
            text=[fmt_k(v) for v in mt["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Interactions moyennes par type de média")
        fig.update_xaxes(title="Total interactions moyennes")
        st.plotly_chart(fig, use_container_width=True)

    with col_d:
        media_dist = df_plat.groupby(["Club", "Media type"]).size().reset_index(name="n")
        fig = px.bar(media_dist, x="Club", y="n", color="Media type",
                     barmode="stack", color_discrete_sequence=px.colors.qualitative.Set2,
                     category_orders={"Club": CLUBS})
        apply_layout(fig, "Répartition des types de média par club")
        fig.update_xaxes(title=""); fig.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig, use_container_width=True)

    top_media = mt.iloc[-1]
    insight(
        f"Le format <b>{top_media['Media type']}</b> est celui qui rapporte le plus d'interactions en moyenne "
        f"(<b>{fmt_k(top_media['Total interactions'])}</b> par post). C'est ce format qu'il faut privilégier "
        f"si on veut maximiser l'engagement sur cette plateforme."
    )

    # ── cat_simple ranking IG/X ─────────────────────────────────────────────
    st.markdown('<div class="section-header">Performance par catégorie de contenu</div>', unsafe_allow_html=True)

    col_e, col_f = st.columns(2)
    with col_e:
        psg_plat = df_plat[df_plat["Club"] == "PSG"]
        cat_psg = (psg_plat.groupby("cat_simple")["Total interactions"].mean()
                   .reset_index().sort_values("Total interactions"))
        fig = go.Figure(go.Bar(
            x=cat_psg["Total interactions"], y=cat_psg["cat_simple"], orientation="h",
            marker_color=["#DA291C" if v == cat_psg["Total interactions"].max() else "#004170" for v in cat_psg["Total interactions"]],
            text=[fmt_k(v) for v in cat_psg["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "PSG — Interactions moyennes par catégorie")
        st.plotly_chart(fig, use_container_width=True)

    with col_f:
        cat_club = df_plat.groupby(["Club", "cat_simple"])["Total interactions"].mean().reset_index()
        fig = px.bar(cat_club, x="cat_simple", y="Total interactions", color="Club",
                     color_discrete_map=CLUB_COLORS, barmode="group",
                     category_orders={"Club": CLUBS})
        apply_layout(fig, "Catégorie × club — comparatif")
        fig.update_xaxes(title=""); fig.update_yaxes(title="Interactions moyennes")
        st.plotly_chart(fig, use_container_width=True)

    # ── Virality + distribution jour de semaine ─────────────────────────────
    st.markdown('<div class="section-header">Virality & timing</div>', unsafe_allow_html=True)
    col_g, col_h = st.columns(2)

    with col_g:
        vr_club = (df_plat.groupby("Club")["Virality Rate"].mean()
                   .reindex(CLUBS).dropna().reset_index().sort_values("Virality Rate"))
        fig = go.Figure(go.Bar(
            x=vr_club["Virality Rate"], y=vr_club["Club"], orientation="h",
            marker_color=club_colors(vr_club["Club"]),
            text=[f"{v:.4f}" for v in vr_club["Virality Rate"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Virality Rate moyen par club")
        st.plotly_chart(fig, use_container_width=True)

    with col_h:
        idow = order_by_day(df_plat.groupby("day_of_week")["Total interactions"].mean().reset_index())
        fig = go.Figure(go.Bar(
            x=idow["day_fr"], y=idow["Total interactions"],
            marker_color=["#DA291C" if v == idow["Total interactions"].max() else "#004170" for v in idow["Total interactions"]],
            text=[fmt_k(v) for v in idow["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Interactions moyennes par jour de la semaine")
        st.plotly_chart(fig, use_container_width=True)

    # ── Évolution nb posts + distribution par jour de la semaine ───────────
    st.markdown('<div class="section-header">Évolution & distribution des publications</div>', unsafe_allow_html=True)
    col_i, col_j = st.columns(2)

    with col_i:
        evo_posts = df_plat.groupby(["Club", "week"]).size().reset_index(name="n")
        fig = go.Figure()
        for club in CLUBS:
            sub = evo_posts[evo_posts["Club"] == club]
            fig.add_trace(go.Scatter(
                x=sub["week"], y=sub["n"], mode="lines+markers", name=club,
                line=dict(color=CLUB_COLORS[club], width=2), marker=dict(size=5),
            ))
        apply_layout(fig, "Évolution du nombre de posts par semaine")
        fig.update_xaxes(title="Semaine"); fig.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig, use_container_width=True)

    with col_j:
        pdow = order_by_day(df_plat.groupby("day_of_week").size().reset_index(name="n"))
        fig = go.Figure(go.Bar(
            x=pdow["day_fr"], y=pdow["n"],
            marker_color=["#DA291C" if v == pdow["n"].max() else "#004170" for v in pdow["n"]],
            text=pdow["n"], textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Distribution des posts par jour de la semaine")
        st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 4 — FACTEURS D'ENGAGEMENT (cœur de B)
# Heatmap timing, matchday effect, sentiment PSG, plateforme, cat × club
# ════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="section-header">Facteurs d\'engagement — Focus PSG</div>', unsafe_allow_html=True)

    # ── 1. Heatmap jour × heure ─────────────────────────────────────────────
    st.markdown("**Quand publier ?** — Heatmap des interactions moyennes PSG (Instagram + X)")
    psg_all = df_ig[df_ig["Club"] == "PSG"]
    heat = (psg_all.groupby(["day_of_week", "hour"])["Total interactions"]
            .mean().reset_index())
    heat_piv = heat.pivot(index="day_of_week", columns="hour", values="Total interactions")
    heat_piv = heat_piv.reindex(DAY_ORDER)
    heat_piv.index = [DAY_LABELS[d] for d in heat_piv.index]
    fig = px.imshow(heat_piv, color_continuous_scale="Blues", aspect="auto",
                    labels=dict(x="Heure", y="Jour", color="Interactions moy."))
    apply_layout(fig, "PSG — Heatmap interactions par jour × heure")
    fig.update_xaxes(dtick=1)
    st.plotly_chart(fig, use_container_width=True)

    # Repérage du créneau optimal
    if not heat_piv.empty:
        flat = heat.dropna().sort_values("Total interactions", ascending=False).iloc[0]
        insight(
            f"Le meilleur créneau pour publier semble être le <b>{DAY_LABELS.get(flat['day_of_week'], flat['day_of_week'])} "
            f"vers {int(flat['hour'])}h</b>, avec environ <b>{fmt_k(flat['Total interactions'])}</b> interactions moyennes. "
            f"À noter : la nuit (0h–5h) performe étonnamment bien (~89K interactions moy.) malgré peu de posts — "
            f"sans doute des contenus stratégiques (post-match, annonces). À l'inverse, le soir (18h–22h) concentre la moitié des posts mais a le moins bon ratio."
        )

    # ── 2. Sentiment PSG ────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Sentiment des posts PSG</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns(2)

    SENT_BUCKET = {  # regroupement des 7 valeurs en 3 catégories
        "strongly positive": "Positif", "positive": "Positif",
        "strongly negative": "Négatif", "negative": "Négatif",
        "neutral": "Neutre", "no sentiment": "Neutre", "mixed": "Neutre",
    }
    SENT_COLORS = {"Positif": "#22C55E", "Neutre": "#64748B", "Négatif": "#EF4444"}

    with col_c:
        sent = psg_all["Sentiment"].dropna().map(SENT_BUCKET).value_counts().reset_index()
        sent.columns = ["Sentiment", "n"]
        fig = go.Figure(go.Pie(
            labels=sent["Sentiment"], values=sent["n"],
            marker_colors=[SENT_COLORS.get(s, "#64748B") for s in sent["Sentiment"]],
            hole=0.45, textinfo="label+percent", textfont=dict(color="white"),
        ))
        apply_layout(fig, "PSG — Répartition du sentiment (IG + X)")
        st.plotly_chart(fig, use_container_width=True)

    with col_d:
        # Scatter %positif × interactions
        sc = psg_all.dropna(subset=["Positive comments (%)", "Total interactions"])
        fig = go.Figure(go.Scatter(
            x=sc["Positive comments (%)"], y=sc["Total interactions"],
            mode="markers",
            marker=dict(color=sc["Positive comments (%)"], colorscale="RdYlGn",
                        showscale=True, size=6, opacity=0.6,
                        colorbar=dict(title="% pos.", thickness=10)),
        ))
        apply_layout(fig, "PSG — % commentaires positifs vs Interactions")
        fig.update_xaxes(title="% commentaires positifs")
        fig.update_yaxes(title="Total interactions", type="log")
        st.plotly_chart(fig, use_container_width=True)

    # ── 3. Engagement PSG par plateforme ───────────────────────────────────
    st.markdown('<div class="section-header">PSG — Performance par plateforme</div>', unsafe_allow_html=True)

    plat_eng = pd.DataFrame({
        "Plateforme": ["TikTok", "Instagram", "X (Twitter)"],
        "Engagement moy.": [
            psg_tt["Total_Engagements"].mean(),
            psg_ig["Total interactions"].mean(),
            psg_x["Total interactions"].mean(),
        ],
        "Volume posts": [len(psg_tt), len(psg_ig), len(psg_x)],
    })
    col_e, col_f = st.columns(2)
    with col_e:
        fig = go.Figure(go.Bar(
            x=plat_eng["Plateforme"], y=plat_eng["Engagement moy."],
            marker_color=["#000000", "#E1306C", "#1DA1F2"],
            text=[fmt_k(v) for v in plat_eng["Engagement moy."]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "PSG — Engagement moyen par plateforme")
        fig.update_yaxes(title="Engagements / interactions moy.")
        st.plotly_chart(fig, use_container_width=True)

    with col_f:
        fig = go.Figure(go.Bar(
            x=plat_eng["Plateforme"], y=plat_eng["Volume posts"],
            marker_color=["#000000", "#E1306C", "#1DA1F2"],
            text=plat_eng["Volume posts"], textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "PSG — Volume de posts par plateforme")
        fig.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig, use_container_width=True)

    # ── 4. Engagement par type de média + Distribution des grades (PSG IG) ─
    st.markdown('<div class="section-header">PSG Instagram — Format & qualité</div>', unsafe_allow_html=True)
    col_g, col_h = st.columns(2)

    with col_g:
        eng_media = psg_ig.groupby("Media type")["Total interactions"].mean().reset_index().sort_values("Total interactions")
        fig = go.Figure(go.Bar(
            x=eng_media["Total interactions"], y=eng_media["Media type"], orientation="h",
            marker_color=["#DA291C" if v == eng_media["Total interactions"].max() else "#004170"
                          for v in eng_media["Total interactions"]],
            text=[fmt_k(v) for v in eng_media["Total interactions"]],
            textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Engagement moyen par type de média")
        fig.update_xaxes(title="Total interactions moyennes")
        st.plotly_chart(fig, use_container_width=True)

    with col_h:
        grade_order  = ["A+", "A", "B", "C", "D"]
        grade_colors = {"A+": "#16A34A", "A": "#22C55E", "B": "#86EFAC",
                        "C": "#F59E0B", "D": "#EF4444"}
        gd = psg_ig["Grade"].dropna().value_counts().reset_index()
        gd.columns = ["Grade", "n"]
        gd["__o"] = gd["Grade"].map({g: i for i, g in enumerate(grade_order)})
        gd = gd.sort_values("__o")
        fig = go.Figure(go.Bar(
            x=gd["Grade"], y=gd["n"],
            marker_color=[grade_colors.get(g, "#64748B") for g in gd["Grade"]],
            text=gd["n"], textposition="outside", textfont=dict(color="white"),
        ))
        apply_layout(fig, "Distribution des grades")
        fig.update_yaxes(title="Nombre de posts")
        st.plotly_chart(fig, use_container_width=True)

    # ── 5. Engagement moyen par heure de publication ────────────────────────
    st.markdown('<div class="section-header">Engagement moyen par heure de publication</div>', unsafe_allow_html=True)
    hour_eng = psg_all.groupby("hour")["Total interactions"].mean().reset_index()
    fig = go.Figure(go.Scatter(
        x=hour_eng["hour"], y=hour_eng["Total interactions"],
        mode="lines+markers",
        line=dict(color="#DA291C", width=2),
        marker=dict(color="#004170", size=8),
        fill="tozeroy", fillcolor="rgba(218,41,28,0.1)",
    ))
    apply_layout(fig, "PSG (IG + X) — Interactions moyennes par heure de publication")
    fig.update_xaxes(title="Heure de publication", tickvals=list(range(0, 24, 2)))
    fig.update_yaxes(title="Interactions moyennes")
    st.plotly_chart(fig, use_container_width=True)

    # ── 6. PSG vs concurrents — Type de contenu (Media type Instagram) ─────
    st.markdown('<div class="section-header">PSG vs concurrents — Type de contenu (Instagram)</div>', unsafe_allow_html=True)
    ig_media_comp = (df_ig[df_ig["Platform"] == "instagram"]
                     .groupby(["Club", "Media type"])["Total interactions"].mean()
                     .reset_index())
    fig = px.bar(ig_media_comp, x="Media type", y="Total interactions",
                 color="Club", color_discrete_map=CLUB_COLORS,
                 barmode="group", category_orders={"Club": CLUBS})
    apply_layout(fig, "Engagement moyen par type de média et par club")
    fig.update_xaxes(title=""); fig.update_yaxes(title="Interactions moyennes")
    st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 5 — RECOMMANDATIONS (chiffrées dynamiquement)
# ════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="section-header">Recommandations Business — PSG</div>', unsafe_allow_html=True)

    # Findings calculés à la volée
    top_media_psg = (psg_ig.groupby("Media type")["Total interactions"].mean()
                     .sort_values(ascending=False))
    best_media   = top_media_psg.index[0] if len(top_media_psg) else "reel"
    best_media_v = top_media_psg.iloc[0]  if len(top_media_psg) else 0

    top_cat_psg_ig = (psg_ig.groupby("cat_simple")["Total interactions"].mean()
                      .sort_values(ascending=False))
    best_cat   = top_cat_psg_ig.index[0] if len(top_cat_psg_ig) else "Football"
    best_cat_v = top_cat_psg_ig.iloc[0]  if len(top_cat_psg_ig) else 0

    top_cat_psg_tt = (psg_tt.groupby("cat_simple")["Views"].mean()
                      .sort_values(ascending=False))
    best_tt_cat = top_cat_psg_tt.index[0] if len(top_cat_psg_tt) else "Football"
    best_tt_v   = top_cat_psg_tt.iloc[0]  if len(top_cat_psg_tt) else 0

    # Best day/hour from heatmap
    psg_all_local = df_ig[df_ig["Club"] == "PSG"]
    hr_eng = psg_all_local.groupby("hour")["Total interactions"].mean()
    best_hour = int(hr_eng.idxmax()) if not hr_eng.empty else 18
    dow_eng = psg_all_local.groupby("day_of_week")["Total interactions"].mean()
    best_day  = DAY_LABELS.get(dow_eng.idxmax(), "Sam") if not dow_eng.empty else "Sam"

    # Best TikTok duration bucket
    df_tt_dur = df_tt.assign(
        bkt=pd.cut(df_tt["Duration (seconds)"],
                   bins=[0, 15, 30, 60, 120, 9999],
                   labels=["< 15s", "15–30s", "30–60s", "60–120s", "> 120s"])
    )
    bucket_views = df_tt_dur[df_tt_dur["Club"] == "PSG"].groupby("bkt", observed=True)["Views"].mean()
    best_bucket = bucket_views.idxmax() if not bucket_views.empty else "30–60s"

    st.markdown(f"""
    <div style="background:#111827;border:1px solid #DA291C;border-radius:12px;padding:1.5rem;margin-bottom:1rem;">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.1rem;color:#DA291C;letter-spacing:1px;margin-bottom:0.8rem;">✅ À FAIRE</div>
        <ul style="color:#CBD5E1;font-size:0.9rem;line-height:1.8;padding-left:1.2rem;">
            <li>Privilégier le format <b style="color:white">{best_media}</b> sur Instagram —
                {fmt_k(best_media_v)} interactions moyennes (top format identifié)</li>
            <li>Capitaliser sur la catégorie <b style="color:white">{best_cat}</b> sur Instagram
                ({fmt_k(best_cat_v)} interactions moy.) et <b style="color:white">{best_tt_cat}</b> sur TikTok
                ({fmt_k(best_tt_v)} vues moy.)</li>
            <li>Vidéos TikTok dans la tranche <b style="color:white">{best_bucket}</b> — pic de vues observé</li>
            <li>Publier en priorité le <b style="color:white">{best_day}</b> autour de
                <b style="color:white">{best_hour}h</b> (créneau optimal selon la heatmap)</li>
        </ul>
    </div>

    <div style="background:#111827;border:1px solid #F59E0B;border-radius:12px;padding:1.5rem;margin-bottom:1rem;">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.1rem;color:#F59E0B;letter-spacing:1px;margin-bottom:0.8rem;">⚠️ À OPTIMISER</div>
        <ul style="color:#CBD5E1;font-size:0.9rem;line-height:1.8;padding-left:1.2rem;">
            <li>Virality Rate Instagram à <b style="color:white">{vir_rate:.4f}</b> — travailler les contenus partageables (memes, moments forts)</li>
            <li>Inter./1k followers <b style="color:white">{avg_inter_1k:.2f}</b> ({delta_i1k:+.0f}% vs concurrents) — l'audience massive du PSG ne se traduit pas pleinement en engagement relatif</li>
            <li>Diversifier les catégories au-delà de "Football" pour réduire la dépendance à un seul registre éditorial</li>
            <li>Tester davantage de formats interactifs (sondages IG, threads X)</li>
        </ul>
    </div>

    <div style="background:#111827;border:1px solid #64748B;border-radius:12px;padding:1.5rem;">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.1rem;color:#94A3B8;letter-spacing:1px;margin-bottom:0.8rem;">❌ À ÉVITER</div>
        <ul style="color:#CBD5E1;font-size:0.9rem;line-height:1.8;padding-left:1.2rem;">
            <li>Vidéos TikTok au-delà de 120s (sous-performance sur la tranche &gt; 120s)</li>
            <li>Publier en masse le soir 18h–22h sans raison — c'est le créneau le moins engageant (27K interactions moy.) malgré 1100+ posts</li>
            <li>Publier des contenus génériques sans thématique claire (perte d'identité)</li>
            <li>Surpublier sur X les jours sans actualité sportive</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    # ── Benchmark final ─────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Benchmark final — PSG vs Concurrents</div>', unsafe_allow_html=True)

    bench = []
    for club in CLUBS:
        ig_sub = df_ig[(df_ig["Club"] == club) & (df_ig["Platform"] == "instagram")]
        tt_sub = df_tt[df_tt["Club"] == club]
        bench.append({
            "Club": club,
            "Posts IG": len(ig_sub),
            "Posts TikTok": len(tt_sub),
            "Eng. moy IG":      round(ig_sub["Total interactions"].mean() or 0, 0),
            "Inter./1k IG":     round(ig_sub["Interactions per 1000 followers"].mean() or 0, 3),
            "Virality Rate":    round(ig_sub["Virality Rate"].mean() or 0, 4),
            "Vues moy TikTok":  round(tt_sub["Views"].mean() or 0, 0),
        })
    bench_df = pd.DataFrame(bench)

    st.dataframe(
        bench_df.style
        .highlight_max(subset=["Eng. moy IG", "Inter./1k IG", "Virality Rate", "Vues moy TikTok"],
                       color="#1A3A5C")
        .format({
            "Eng. moy IG":      "{:,.0f}",
            "Inter./1k IG":     "{:.3f}",
            "Virality Rate":    "{:.4f}",
            "Vues moy TikTok":  "{:,.0f}",
        }),
        use_container_width=True, hide_index=True,
    )

# ═════════════════════════════════════════════════════════════════════════════
# FOOTER
# ═════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("""
<div style="text-align:center;color:#374151;font-size:0.8rem;padding:1rem 0;">
    PSG Social Media Analytics · Données Jan–Mar 2026 · Sources : TikTok, Instagram, X (Twitter) · SAE BUT3 IUT
</div>
""", unsafe_allow_html=True)
