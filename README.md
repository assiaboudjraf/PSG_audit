# PSG Social Media Analytics

Dashboard d'analyse comparative des performances digitales du PSG vs Real Madrid, Borussia Dortmund et Tottenham sur TikTok, Instagram et X (Twitter). Période : janvier–mars 2026.

Projet SAE BUT3 — IUT Paris Rives de Seine.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Données

- `tiktok_ready.xlsx` — 856 posts TikTok nettoyés, avec colonne `cat_simple` (7 catégories de contenu)
- `insta_x_ready.xlsx` — 8 481 posts Instagram + X, avec `Media type`, `Grade`, `Sentiment`, `cat_simple`

## Structure (5 onglets)

1. **Performance globale** — comparaison des 4 clubs (engagement IG, vues TikTok, inter./1k, virality)
2. **TikTok** — vues, durée, catégories, distribution temporelle
3. **Instagram & X** — interactions, formats, catégories, jours
4. **Facteurs d'engagement** — heatmap timing, sentiment, plateforme, type de média, grades
5. **Recommandations** — actions chiffrées dynamiquement
