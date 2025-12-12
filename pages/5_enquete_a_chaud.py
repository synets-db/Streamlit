import streamlit as st
import pandas as pd
import altair as alt

# ==============================================================
# Configuration de la page
# ==============================================================
st.set_page_config(page_title="Enquête à chaud — Items", layout="wide")
st.title("📝 Enquête à chaud — Analyse de tous les items")

st.markdown("""
Cette page analyse les items de l’enquête (colonnes **C à AZ**) :

- On ignore les colonnes **A et B** (identité / technique).  
- Pour chaque item, on prend la **dernière ligne** comme ligne de totaux ("Les Oui").  
- On affiche un **camembert** avec les pourcentages par modalité,  
  puis les **commentaires** liés à l’item, répartis sur **deux colonnes** pour limiter le scroll.  
- En plus, on affiche les **questions ouvertes finales** (colonnes BO, BQ, BS).
""")

# ==============================================================
# Upload du fichier
# ==============================================================
uploaded_file = st.file_uploader(
    "Déposez le fichier CSV (export Evento, séparateur `;`)",
    type=["csv"]
)

if not uploaded_file:
    st.info("En attente d'un fichier CSV…")
    st.stop()

# Lecture du CSV avec ; et gestion des accents
try:
    df_raw = pd.read_csv(uploaded_file, sep=";", encoding="utf-8")
except UnicodeDecodeError:
    df_raw = pd.read_csv(uploaded_file, sep=";", encoding="latin-1")

if df_raw.shape[1] <= 2:
    st.error("Le fichier ne semble pas contenir plus de 2 colonnes (A et B).")
    st.stop()

# ==============================================================
# Sélection des colonnes C → AZ (indices 2 à 51)
# ==============================================================
max_col = min(df_raw.shape[1], 52)  # au cas où il y aurait moins de colonnes que prévu
df = df_raw.iloc[:, 2:max_col].copy()  # colonnes C..AZ

st.write(f"Nombre de lignes dans le fichier : {len(df)}")
st.write(f"Nombre de colonnes analysées (C → AZ) : {df.shape[1]}")

last_row = df.iloc[-1]           # ligne "Les Oui"
df_without_total = df.iloc[:-1]  # lignes des répondants (pour les commentaires)

# ==============================================================
# Préparation des items : base = texte avant le dernier '.'
# ==============================================================
# items = {
#   item_label: {
#       "mod_cols": [(modalité, nom_col), ...],
#       "comment_cols": [nom_col, ...]
#   }
# }
items = {}

for col in df.columns:
    full = str(col)
    if "." in full:
        base, modality = full.rsplit(".", 1)
    else:
        base, modality = full, ""

    base = base.strip()
    modality = modality.strip()

    if base not in items:
        items[base] = {"mod_cols": [], "comment_cols": []}

    if "commentaire" in modality.lower():
        items[base]["comment_cols"].append(col)
    else:
        items[base]["mod_cols"].append((modality, col))

# ==============================================================
# Palette satisfaction + ordre dédié
# ==============================================================
ordre_satisfaction = [
    "Très satisfait",
    "Satisfait",
    "Peu satisfait",
    "Pas du tout satisfait",
]

couleurs_satisfaction = {
    "Très satisfait": "#1b7837",      # vert foncé
    "Satisfait": "#5aae61",           # vert clair
    "Peu satisfait": "#80cdc1",       # bleu doux
    "Pas du tout satisfait": "#f46d43"  # orange
}

# ==============================================================
# Affichage des items (camembert + commentaires en 2 colonnes)
# ==============================================================
for item_label, info in items.items():
    mod_cols = info["mod_cols"]
    comment_cols = info["comment_cols"]

    # Si vraiment rien à afficher, on saute
    if not mod_cols and not comment_cols:
        continue

    # ====== 1) Calcul des données pour les modalités ======
    chart_df = None
    pie = None
    total = 0

    if mod_cols:
        labels = []
        counts = []

        for modality, col in mod_cols:
            labels.append(modality)
            try:
                # la dernière ligne contient les totaux pour cette modalité
                val = int(last_row[col])
            except (ValueError, TypeError):
                val = 0
            counts.append(val)

        total = sum(counts)

        if total > 0:
            percentages = [round(c / total * 100, 1) for c in counts]

            chart_df = pd.DataFrame({
                "Modalité": labels,
                "Nombre": counts,
                "Pourcentage": percentages,
            })

            # Déterminer si c'est une question de satisfaction (cas spécial)
            if set(labels).issubset(set(ordre_satisfaction)):
                # Ordre satisfaction
                chart_df["Modalité"] = pd.Categorical(
                    chart_df["Modalité"],
                    categories=ordre_satisfaction,
                    ordered=True
                )
                chart_df = chart_df.sort_values("Modalité")

                # Camembert avec palette satisfaction
                pie = (
                    alt.Chart(chart_df)
                    .mark_arc(innerRadius=40)
                    .encode(
                        theta="Nombre:Q",
                        color=alt.Color(
                            "Modalité:N",
                            scale=alt.Scale(
                                domain=ordre_satisfaction,
                                range=[couleurs_satisfaction[m] for m in ordre_satisfaction]
                            ),
                            legend=alt.Legend(title="Modalité")
                        ),
                        tooltip=["Modalité", "Nombre", "Pourcentage"]
                    )
                    .properties(width=350, height=350)
                )
            else:
                # Autres types de modalités (ex : Oui / Non / Partiellement, délais, etc.)
                pie = (
                    alt.Chart(chart_df)
                    .mark_arc(innerRadius=40)
                    .encode(
                        theta="Nombre:Q",
                        color=alt.Color("Modalité:N", legend=alt.Legend(title="Modalité")),
                        tooltip=["Modalité", "Nombre", "Pourcentage"]
                    )
                    .properties(width=350, height=350)
                )

    # ====== 2) Récupération des commentaires pour cet item ======
    commentaires = []
    if comment_cols:
        for col in comment_cols:
            s = df_without_total[col].dropna().astype(str).str.strip()
            s = s[s != ""]
            commentaires.extend(list(s))

    # ====== 3) Affichage en 2 colonnes ======
    st.markdown(f"## ❓ {item_label}")
    col1, col2 = st.columns([1, 1])

    # Colonne de gauche : camembert
    # with col1:
    #     if mod_cols:
    #         if total > 0 and pie is not None:
    #             st.markdown(f"**{total} réponses**")
    #             st.altair_chart(pie, use_container_width=True)
    #         else:
    #             st.info("Aucun total disponible pour calculer les pourcentages sur cet item.")
    #     else:
    #         st.info("Aucune modalité fermée pour cet item (uniquement des commentaires).")
# ====== 3) Affichage en 2 colonnes ======
    # Colonne de gauche : camembert
    with col1:
        if mod_cols:
            if total > 0 and pie is not None:
                st.markdown(f"**{total} réponses**")
                # Camembert avec taille "compacte" (ancienne version)
                st.altair_chart(
                    pie.properties(width=400, height=200), 
                    use_container_width=False
                )
            else:
                st.info("Aucun total disponible pour calculer les pourcentages sur cet item.")
        else:
            st.info("Aucune modalité fermée pour cet item (uniquement des commentaires).")


    # Colonne de droite : détail + commentaires
    with col2:
        if mod_cols and chart_df is not None:
            st.markdown("### 📊 Détail des réponses")
            for lab, c, p in zip(chart_df["Modalité"], chart_df["Nombre"], chart_df["Pourcentage"]):
                st.markdown(f"- **{lab}** : {c} réponses ({p}%)")

        if comment_cols:
            st.markdown("### 💬 Commentaires")
            if not commentaires:
                st.info("Aucun commentaire renseigné pour cet item.")
            else:
                for i, txt in enumerate(commentaires, start=1):
                    st.markdown(f"- **Commentaire {i}** : {txt}")

    st.markdown("---")

# ==============================================================
# Questions ouvertes finales (colonnes BO, BQ, BS)
# ==============================================================
st.header("📝 Questions ouvertes finales")

# Indices Excel -> index pandas (0-based) :
# BO = 66, BQ = 68, BS = 70
open_indices = [66, 68, 70]
ncols = df_raw.shape[1]

for idx in open_indices:
    if idx >= ncols:
        continue  # au cas où le fichier aurait moins de colonnes

    col_name = df_raw.columns[idx]

    # On supprime systématiquement le suffixe ".Commentaire"
    if ".Commentaire" in col_name:
        question_label = col_name.replace(".Commentaire", "").strip()
    else:
        # fallback si le nom est inhabituel
        question_label = col_name.split(".")[0].strip()

    st.subheader(f"❓ {question_label}")

    # On prend toutes les lignes sauf la dernière (si c'est la ligne de totaux)
    serie = df_raw.iloc[:-1, idx].dropna().astype(str).str.strip()
    serie = serie[serie != ""]

    if serie.empty:
        st.info("Aucune réponse renseignée pour cette question.")
    else:
        for i, txt in enumerate(serie, start=1):
            st.markdown(f"- **Réponse {i}** : {txt}")

