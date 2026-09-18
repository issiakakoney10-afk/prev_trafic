"""
Outil local de gestion des prévisions de trafics — interface Streamlit.

Organisée en onglets : Saisie mensuelle / Coefficients saisonniers /
Prévisions (génération) / Visualisation / Aide.

Cette appli fait le lien entre :
  - le fichier maître d'origine (schema/master_reference.xlsx), jamais modifié,
  - les modèles remplis par l'utilisateur (détail mensuel, paramètres annuels),
  - les informations complémentaires saisies ici (chiffres annoncés,
    frontières réalisé/prévision, méthode de projection),
et appelle engine.inject.generate_output() pour produire un classeur au même
format que l'original, entièrement recalculé. L'onglet Visualisation relit
ensuite ce classeur (ou le fichier de référence) pour l'afficher en tableaux
et graphiques.
"""
import datetime as dt
import shutil
import sys
import tempfile
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
ENGINE_DIR = APP_DIR / "engine"
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))

import common  # noqa: E402
import viz  # noqa: E402
from inject import generate_output  # noqa: E402

st.set_page_config(
    page_title="Prévisions de trafics — outil local",
    page_icon="🚢",
    layout="wide",
)

# --------------------------------------------------------------------------
# Palette (méthode data-viz : couleurs catégorielles à ordre fixe, jamais
# recyclées au hasard — chaque KPI garde la même couleur partout dans l'appli)
# --------------------------------------------------------------------------
PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
INK_PRIMARY, INK_SECONDARY, INK_MUTED = "#0b0b0b", "#52514e", "#898781"
GRID, SURFACE = "#e1e0d9", "#fcfcfb"
STATUS_GOOD, STATUS_CRITICAL = "#0ca30c", "#d03b3b"

KPI_COLOR = {kpi: PALETTE[i % len(PALETTE)] for i, kpi in enumerate(viz.KPI_ORDER)}
YEAR_COLOR_SEQ = PALETTE  # les années d'un historique reprennent l'ordre fixe


def style_fig(fig, ytitle="", legend=True):
    fig.update_layout(
        plot_bgcolor=SURFACE, paper_bgcolor=SURFACE,
        font=dict(color=INK_SECONDARY, family="system-ui, -apple-system, 'Segoe UI', sans-serif"),
        margin=dict(l=10, r=10, t=30, b=10),
        hovermode="x unified",
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        height=340,
    )
    fig.update_xaxes(showgrid=False, linecolor=GRID, tickfont=dict(color=INK_MUTED))
    fig.update_yaxes(showgrid=True, gridcolor=GRID, zeroline=False, title=ytitle,
                      tickfont=dict(color=INK_MUTED))
    return fig


def line_chart(df, kpi_key, ytitle=""):
    color = KPI_COLOR.get(kpi_key, PALETTE[0])
    fig = go.Figure(go.Scatter(
        x=df["Mois"], y=df["Valeur"], mode="lines+markers",
        line=dict(color=color, width=2), marker=dict(size=8, color=color),
        hovertemplate="%{x} : %{y:,.0f}<extra></extra>",
    ))
    return style_fig(fig, ytitle, legend=False)


def realise_projete_chart(df_kpi, ytitle=""):
    """Ligne continue + marqueurs pleins (réalisé) / creux (projeté) —
    la distinction ne repose jamais sur la couleur seule."""
    color = KPI_COLOR.get(df_kpi["kpi"].iloc[0], PALETTE[0])
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_kpi["Mois"], y=df_kpi["Valeur"], mode="lines",
                              line=dict(color=color, width=2), showlegend=False,
                              hoverinfo="skip"))
    for statut, symbol in (("Réalisé", "circle"), ("Projeté", "circle-open")):
        sub = df_kpi[df_kpi["Statut"] == statut]
        if sub.empty:
            continue
        fig.add_trace(go.Scatter(
            x=sub["Mois"], y=sub["Valeur"], mode="markers", name=statut,
            marker=dict(size=9, color=color, symbol=symbol,
                        line=dict(width=2, color=color)),
            hovertemplate="%{x} : %{y:,.0f}<extra>" + statut + "</extra>",
        ))
    return style_fig(fig, ytitle)


def yearly_bar_chart(df, x_col, y_col, color_col, ytitle=""):
    fig = go.Figure()
    cats = list(dict.fromkeys(df[color_col]))
    for i, cat in enumerate(cats):
        sub = df[df[color_col] == cat]
        fig.add_trace(go.Bar(x=sub[x_col], y=sub[y_col], name=str(cat),
                              marker_color=PALETTE[i % len(PALETTE)],
                              hovertemplate="%{x} : %{y:,.0f}<extra>" + str(cat) + "</extra>"))
    fig.update_layout(barmode="group")
    return style_fig(fig, ytitle)


def history_line_chart(hist_df, ytitle=""):
    fig = go.Figure()
    for i, year in enumerate(sorted(hist_df["Année"].unique())):
        sub = hist_df[hist_df["Année"] == year]
        fig.add_trace(go.Scatter(x=sub["Mois"], y=sub["Valeur"], mode="lines+markers",
                                  name=str(year), line=dict(color=YEAR_COLOR_SEQ[i % 8], width=2),
                                  marker=dict(size=6)))
    return style_fig(fig, ytitle)


def coef_bar_chart(coef_df):
    colors = [STATUS_GOOD if v >= 1 else STATUS_CRITICAL for v in coef_df["Coefficient"]]
    fig = go.Figure(go.Bar(x=coef_df["Mois"], y=coef_df["Coefficient"], marker_color=colors,
                            hovertemplate="%{x} : %{y:.3f}<extra></extra>"))
    fig.add_hline(y=1.0, line_dash="dash", line_color=INK_MUTED)
    return style_fig(fig, "Coefficient (1.0 = moyenne)", legend=False)


MONTH_OPTIONS = ["— Ne pas modifier —"] + [
    f"{i:02d} — {name}" for i, name in enumerate(common.MONTHS_FR, start=1)
]


def select_month_or_none(label: str, help_text: str, key: str):
    choice = st.selectbox(label, MONTH_OPTIONS, index=0, help=help_text, key=key)
    if choice == MONTH_OPTIONS[0]:
        return None
    return int(choice.split(" — ")[0])


def active_workbook_path() -> Path:
    last = st.session_state.get("last_generated_path")
    if last and Path(last).exists():
        return Path(last)
    return common.MASTER


# --------------------------------------------------------------------------
# Vérifications de démarrage
# --------------------------------------------------------------------------
st.title("🚢 Prévisions de trafics")

if not common.MASTER.exists():
    st.error(f"Fichier maître introuvable : {common.MASTER}\n\nLe dossier 'schema/' doit rester à côté de ce script.")
    st.stop()

if shutil.which("soffice") is None:
    st.warning("LibreOffice ('soffice') introuvable sur ce serveur : le recalcul échouera. "
               "Voir l'onglet Aide.", icon="🛑")

tab_saisie, tab_coef, tab_prev, tab_viz, tab_aide = st.tabs([
    "📅 Saisie mensuelle", "📊 Coefficients saisonniers", "🔮 Prévisions",
    "📈 Visualisation", "❓ Aide",
])

# --------------------------------------------------------------------------
# Onglet 1 — Saisie mensuelle
# --------------------------------------------------------------------------
with tab_saisie:
    st.header("Données du mois")
    st.caption("Modèle vierge, détail mensuel rempli, et chiffre annoncé du mois en cours.")

    tpl_trafic = common.TEMPLATES_DIR / "modele_trafic_mens_2026.xlsx"
    if tpl_trafic.exists():
        st.download_button("📥 Télécharger le modèle vierge (détail mensuel)",
                            data=tpl_trafic.read_bytes(), file_name=tpl_trafic.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    trafic_mens_file = st.file_uploader(
        "Détail mensuel rempli (modèle 'Trafic mens 2026')", type=["xlsx"])

    st.divider()
    st.subheader("Chiffre annoncé du mois (optionnel)")
    st.caption("Pour un mois dont le détail complet n'est pas encore prêt : ces 2 chiffres "
               "'flash' suffisent, le fichier estime automatiquement les autres statistiques "
               "annoncées à partir d'eux, comme il le fait déjà aujourd'hui.")
    anchor_enabled = st.checkbox("Renseigner un chiffre annoncé ce mois-ci")
    default_month_idx = min(max(dt.datetime.now().month, 1), 12) - 1
    anchor_month_label = st.selectbox("Mois concerné", common.MONTHS_FR,
                                       index=default_month_idx, disabled=not anchor_enabled)
    anchor_month = common.MONTHS_FR.index(anchor_month_label) + 1
    c1, c2 = st.columns(2)
    with c1:
        anchor_escales = st.number_input("Nombre d'escales de navires", min_value=0, step=1,
                                          value=0, disabled=not anchor_enabled)
    with c2:
        anchor_tonnage = st.number_input("Tonnage global (tonnes)", min_value=0.0, step=100.0,
                                          value=0.0, format="%.0f", disabled=not anchor_enabled)

# --------------------------------------------------------------------------
# Onglet 2 — Coefficients saisonniers (séparé des prévisions : c'est la
# donnée de fond qui alimente la méthode de projection, pas la génération
# elle-même)
# --------------------------------------------------------------------------
with tab_coef:
    st.header("Coefficients saisonniers & paramètres annuels")
    st.caption("Objectifs annuels, historique 5 ans et coefficients saisonniers par catégorie, "
               "engagement contractuel TC2, statistiques de clôture. C'est cette saisonnalité "
               "qui sert de base à la méthode « coefficients saisonniers » dans l'onglet Prévisions.")

    tpl_param = common.TEMPLATES_DIR / "modele_parametres_annuels.xlsx"
    if tpl_param.exists():
        st.download_button("📥 Télécharger le modèle vierge (paramètres annuels)",
                            data=tpl_param.read_bytes(), file_name=tpl_param.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    parametres_file = st.file_uploader(
        "Paramètres annuels remplis (objectifs, historique, TC2, stats de clôture)", type=["xlsx"])

    st.divider()
    st.subheader("Aperçu de la saisonnalité actuelle")
    try:
        wb_coef = viz.load_wb(active_workbook_path())
        obj_label = st.selectbox("Catégorie", [lbl for _, lbl in common.OBJECTIFS_ANNUELS],
                                  key="coef_preview_kpi")
        cell = dict((lbl, c) for c, lbl in common.OBJECTIFS_ANNUELS)[obj_label]
        hist_df, coef_df, objectif = viz.seasonal_block(wb_coef, cell)
        m1, m2 = st.columns([1, 3])
        with m1:
            st.metric("Objectif annuel", f"{objectif:,.0f}".replace(",", " "))
        with m2:
            st.plotly_chart(coef_bar_chart(coef_df), use_container_width=True,
                             config={"displaylogo": False})
        st.caption("Historique mensuel 5 ans (sert au calcul du coefficient ci-dessus)")
        st.plotly_chart(history_line_chart(hist_df), use_container_width=True,
                         config={"displaylogo": False})
    except Exception as e:
        st.info(f"Aperçu indisponible pour le moment ({e}).")

# --------------------------------------------------------------------------
# Onglet 3 — Prévisions (frontières réalisé/prévision + génération)
# --------------------------------------------------------------------------
with tab_prev:
    st.header("Générer les prévisions")
    st.info("Chaque génération repart du fichier d'origine et applique uniquement ce qui est "
             "renseigné dans les 3 onglets précédents et ci-dessous.", icon="⚠️")

    last_full_detail_month = select_month_or_none(
        "Dernier mois en détail complet",
        "Jusqu'à ce mois inclus, CALC. PREV utilise le détail complet saisi. Au-delà, elle "
        "utilise le chiffre annoncé du mois (onglet Saisie mensuelle), s'il a été renseigné.",
        key="last_full_detail_month")
    st.caption("ℹ️ Pilote la feuille **CALC. PREV**.")

    last_known_month = select_month_or_none(
        "Dernier mois connu (détail ou annoncé)",
        "Jusqu'à ce mois inclus, ANNUEL reprend directement le réalisé de CALC. PREV. Au-delà, "
        "elle bascule sur la méthode de projection choisie ci-dessous.",
        key="last_known_month")
    st.caption("ℹ️ Pilote la feuille **ANNUEL**, à partir des coefficients saisonniers si cette "
               "méthode est choisie.")

    if (last_known_month is not None and last_full_detail_month is not None
            and last_known_month < last_full_detail_month):
        st.warning("Le dernier mois connu (ANNUEL) est antérieur au dernier mois en détail "
                   "complet (CALC. PREV) — vérifiez que c'est bien voulu.")

    method_keys = list(common.METHOD_LABELS.keys())
    method_labels = [common.METHOD_LABELS[k] for k in method_keys]
    default_method_idx = method_keys.index("correction_des_previsions") if "correction_des_previsions" in method_keys else 0
    chosen_method_label = st.selectbox("Méthode de projection au-delà du dernier mois connu",
                                        method_labels, index=default_method_idx,
                                        disabled=last_known_month is None)
    annuel_method_key = method_keys[method_labels.index(chosen_method_label)]

    st.divider()
    output_name = st.text_input("Nom du fichier généré (optionnel)",
                                 placeholder="Laisser vide pour un nom automatique horodaté")

    if st.button("🚀 Générer le fichier Excel", type="primary", use_container_width=True):
        tmp_dir = None
        try:
            tmp_dir = Path(tempfile.mkdtemp(prefix="trafic_tool_"))
            kwargs = {}

            if trafic_mens_file is not None:
                p = tmp_dir / "trafic_mens_upload.xlsx"
                p.write_bytes(trafic_mens_file.getvalue())
                kwargs["trafic_mens_template_path"] = p

            if parametres_file is not None:
                p = tmp_dir / "parametres_upload.xlsx"
                p.write_bytes(parametres_file.getvalue())
                kwargs["parametres_template_path"] = p

            if anchor_enabled:
                kwargs["monthly_anchor"] = {"month": anchor_month, "escales": anchor_escales,
                                             "tonnage_global": anchor_tonnage}
            if last_full_detail_month is not None:
                kwargs["last_full_detail_month"] = last_full_detail_month
            if last_known_month is not None:
                kwargs["last_known_month"] = last_known_month
                kwargs["annuel_method"] = annuel_method_key
            if output_name.strip():
                name = output_name.strip()
                kwargs["output_name"] = name if name.lower().endswith(".xlsx") else name + ".xlsx"

            with st.spinner("Génération en cours… le recalcul complet peut prendre jusqu'à une minute."):
                final_path, report = generate_output(**kwargs)

            st.session_state["last_generated_path"] = str(final_path)
            st.success("Fichier généré avec succès. Voir l'onglet **Visualisation** pour l'explorer.")
            with st.expander("Détails du traitement", expanded=False):
                for msg in report.messages:
                    st.write(f"- {msg}")
            st.download_button("⬇️ Télécharger le fichier généré", data=final_path.read_bytes(),
                                file_name=final_path.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary", use_container_width=True)
        except Exception as e:
            st.error(f"Une erreur est survenue pendant la génération : {e}")
        finally:
            if tmp_dir is not None:
                shutil.rmtree(tmp_dir, ignore_errors=True)

# --------------------------------------------------------------------------
# Onglet 4 — Visualisation (tableaux + graphiques de toutes les feuilles)
# --------------------------------------------------------------------------
with tab_viz:
    is_generated = "last_generated_path" in st.session_state and Path(st.session_state["last_generated_path"]).exists()
    st.caption(("🟢 Vous visualisez le **dernier fichier généré**." if is_generated
                else "🔵 Vous visualisez le **fichier de référence d'origine** (aucune génération faite dans cette session).")
               + " " + f"`{active_workbook_path().name}`")

    wb_v = viz.load_wb(active_workbook_path())
    wb_f = viz.load_wb(active_workbook_path(), data_only=False)

    sub = st.tabs(["Vue d'ensemble (ANNUEL)", "CALC. PREV", "Détail mensuel", "Répartition", "TC2", "Stats de clôture"])

    with sub[0]:
        st.caption("Marqueur plein = réalisé, marqueur creux = projeté (méthode de prévision).")
        df = viz.annuel_df(wb_v, wb_f)
        cols = st.columns(2)
        for i, kpi in enumerate(viz.KPI_ORDER):
            with cols[i % 2]:
                st.markdown(f"**{viz.KPI_LABELS[kpi]}**")
                st.plotly_chart(realise_projete_chart(df[df["kpi"] == kpi]),
                                 use_container_width=True, config={"displaylogo": False})
        with st.expander("Tableau complet"):
            st.dataframe(df.pivot(index="KPI", columns="Mois", values="Valeur")[common.MONTHS_FR],
                         use_container_width=True)

    with sub[1]:
        df = viz.calc_prev_df(wb_v)
        kpi_lbl = st.selectbox("Catégorie", [viz.KPI_LABELS[k] for k in viz.KPI_ORDER], key="calcprev_kpi")
        kpi_key = viz.KPI_ORDER[[viz.KPI_LABELS[k] for k in viz.KPI_ORDER].index(kpi_lbl)]
        st.plotly_chart(line_chart(df[df["kpi"] == kpi_key], kpi_key), use_container_width=True,
                         config={"displaylogo": False})
        with st.expander("Tableau complet"):
            st.dataframe(df.pivot(index="KPI", columns="Mois", values="Valeur")[common.MONTHS_FR],
                         use_container_width=True)

    with sub[2]:
        year_sheet = st.radio("Année", [common.YEAR_SHEET, "Trafic mens2025"], horizontal=True)
        df = viz.trafic_mens_kpi_df(wb_v, year_sheet)
        kpi_lbl = st.selectbox("Catégorie", [viz.KPI_LABELS[k] for k in viz.KPI_ORDER], key="mens_kpi")
        kpi_key = viz.KPI_ORDER[[viz.KPI_LABELS[k] for k in viz.KPI_ORDER].index(kpi_lbl)]
        sub_df = df[df["kpi"] == kpi_key]
        if not sub_df.empty:
            st.plotly_chart(line_chart(sub_df, kpi_key), use_container_width=True,
                             config={"displaylogo": False})
        with st.expander("Tableau complet"):
            st.dataframe(df.pivot(index="KPI", columns="Mois", values="Valeur")[common.MONTHS_FR],
                         use_container_width=True)

    with sub[3]:
        df = viz.repartition_df(wb_v)
        serie = st.selectbox("Série", df["Série"].unique(), key="repartition_serie")
        sub_df = df[df["Série"] == serie]
        fig = go.Figure(go.Scatter(x=sub_df["Année"], y=sub_df["Valeur"], mode="lines+markers",
                                    line=dict(color=PALETTE[0], width=2), marker=dict(size=8)))
        st.plotly_chart(style_fig(fig, legend=False), use_container_width=True, config={"displaylogo": False})
        with st.expander("Tableau complet"):
            st.dataframe(df.pivot(index="Série", columns="Année", values="Valeur"), use_container_width=True)

    with sub[4]:
        df = viz.tc2_df(wb_v)
        st.caption("Engagement contractuel de trafic de conteneurs transbordés (business plan TC2) "
                   "vs projection du plan d'affaires, par année.")
        melted = df.melt(id_vars="Année", var_name="Série", value_name="Valeur")
        st.plotly_chart(yearly_bar_chart(melted, "Année", "Valeur", "Série", "EVP"),
                         use_container_width=True, config={"displaylogo": False})
        st.dataframe(df, use_container_width=True, hide_index=True)

    with sub[5]:
        section = st.selectbox("Vue", list(viz.STATS_EXO_SECTIONS.keys()), key="stats_exo_section")
        df = viz.stats_exo_df(wb_v, section)
        st.plotly_chart(yearly_bar_chart(df, "Année", "Valeur", "Catégorie"),
                         use_container_width=True, config={"displaylogo": False})
        with st.expander("Tableau complet"):
            st.dataframe(df.pivot(index="Catégorie", columns="Année", values="Valeur"),
                         use_container_width=True)

# --------------------------------------------------------------------------
# Onglet 5 — Aide
# --------------------------------------------------------------------------
with tab_aide:
    st.header("Comment utiliser l'outil")
    st.markdown("""
**1. Saisie mensuelle** — téléchargez le modèle vierge, remplissez uniquement les cellules
jaunes dans Excel, redéposez-le. Le chiffre annoncé (2 valeurs flash) sert quand le détail
complet du mois n'est pas encore prêt.

**2. Coefficients saisonniers** — objectifs annuels, historique 5 ans et engagement TC2.
C'est cette saisonnalité (le graphique en barres, vert = mois fort, rouge = mois faible par
rapport à la moyenne) qui alimente la méthode « coefficients saisonniers » disponible dans
l'onglet Prévisions — d'où la séparation : on renseigne l'hypothèse de fond ici, on choisit
comment l'utiliser dans Prévisions.

**3. Prévisions** — deux réglages à ne pas confondre :
- *Dernier mois en détail complet* : pilote CALC. PREV (détail vs chiffre annoncé).
- *Dernier mois connu* : pilote ANNUEL (réalisé vs méthode de projection), en général égal ou
  postérieur au précédent.

Cliquez sur **Générer** : le fichier est recalculé (LibreOffice) puis téléchargeable.

**4. Visualisation** — explore le dernier fichier généré (ou le fichier de référence si vous
n'avez encore rien généré dans cette session) : toutes les feuilles du classeur d'origine,
en tableaux et graphiques interactifs (survolez un point pour voir sa valeur).

### Limites connues
- La feuille `ANNUEL_DCAQ_AVEC_TC2_Option2` (ancienne version, erreurs `#REF!` héritées) est
  recopiée telle quelle, sans recalcul, comme dans le fichier d'origine.
- Le roulement annuel (nouvel onglet, décalage de l'historique 5 ans) reste manuel.
- Chaque génération repart du fichier de référence fixe : les réglages de l'onglet Prévisions
  ne sont pas mémorisés d'une génération à l'autre.

### En cas de problème
- **Erreur "LibreOffice introuvable"** : contactez la personne qui a déployé l'outil —
  `packages.txt` doit contenir `libreoffice-calc`.
- **Erreur pendant la génération** : le message affiché indique la cause (souvent, un modèle
  déposé ne correspond pas à celui téléchargé depuis l'outil — redéposez un modèle frais).
""")

st.divider()
st.caption("Outil local — tout le traitement se fait sur ce serveur, aucune donnée n'est partagée ailleurs.")
