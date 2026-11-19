import io
import re
import textwrap
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib.backends.backend_pdf import PdfPages

st.set_page_config(
    page_title="Archétype Asso — Audit ANC 2018-06, Pilotage & Prospective",
    page_icon="🤝",
    layout="wide"
)

# ==========================================
# 1. RÉFÉRENTIEL & NORMES PAR SECTEUR ASSOCIATIF (ANC 2018-06)
# ==========================================
def get_normes_asso(secteur):
    if "Social" in secteur:
        return {
            "type_court": "Action sociale & Insertion",
            "strate": "Associations d'action sociale, solidarité et insertion (IAE)",
            "financeurs": "État (DDETS), Département, FSE",
            "dep_pub_max": 75.0,
            "conc_fin_max": 50.0,
            "taux_ebe": 5.5,
            "taux_caf": 4.5,
            "mois_treso": 4.0,
            "poids_ms": 68.0,
            "poids_support": 14.0,
            "taux_benevolat": 8.0,
            "desend_moy": 4.0,
            "desend_max": 7.0
        }
    elif "Culture" in secteur:
        return {
            "type_court": "Culture & Spectacle vivant",
            "strate": "Associations culturelles, festivals et lieux de création",
            "financeurs": "DRAC, Région, Commune, Billetterie",
            "dep_pub_max": 60.0,
            "conc_fin_max": 40.0,
            "taux_ebe": 6.5,
            "taux_caf": 5.0,
            "mois_treso": 3.5,
            "poids_ms": 54.0,
            "poids_support": 16.0,
            "taux_benevolat": 15.0,
            "desend_moy": 3.5,
            "desend_max": 6.5
        }
    elif "Sport" in secteur:
        return {
            "type_court": "Sport & Éducation populaire",
            "strate": "Clubs sportifs employeurs, MJC et centres sociaux",
            "financeurs": "Commune, ANS, CAF, Cotisations",
            "dep_pub_max": 45.0,
            "conc_fin_max": 35.0,
            "taux_ebe": 7.0,
            "taux_caf": 5.5,
            "mois_treso": 3.5,
            "poids_ms": 58.0,
            "poids_support": 13.0,
            "taux_benevolat": 25.0,
            "desend_moy": 3.5,
            "desend_max": 6.0
        }
    elif "Médico" in secteur:
        return {
            "type_court": "Santé & Médico-social",
            "strate": "Associations gestionnaires d'établissements et services",
            "financeurs": "ARS, Département, Assurance Maladie",
            "dep_pub_max": 85.0,
            "conc_fin_max": 60.0,
            "taux_ebe": 6.5,
            "taux_caf": 5.5,
            "mois_treso": 3.0,
            "poids_ms": 72.0,
            "poids_support": 12.0,
            "taux_benevolat": 5.0,
            "desend_moy": 5.0,
            "desend_max": 8.0
        }
    elif "Fédération" in secteur:
        return {
            "type_court": "Fédération & Tête de réseau",
            "strate": "Fédérations régionales, têtes de réseau et syndicats",
            "financeurs": "Cotisations adhérents, Prestations, Subv.",
            "dep_pub_max": 35.0,
            "conc_fin_max": 30.0,
            "taux_ebe": 8.0,
            "taux_caf": 6.5,
            "mois_treso": 5.0,
            "poids_ms": 55.0,
            "poids_support": 18.0,
            "taux_benevolat": 10.0,
            "desend_moy": 3.0,
            "desend_max": 6.0
        }
    else:
        return {
            "type_court": "Environnement & Tiers-lieux",
            "strate": "Associations de transition écologique et développement local",
            "financeurs": "ADEME, Région, Mécénat, Prestations",
            "dep_pub_max": 55.0,
            "conc_fin_max": 40.0,
            "taux_ebe": 7.0,
            "taux_caf": 5.5,
            "mois_treso": 4.0,
            "poids_ms": 60.0,
            "poids_support": 15.0,
            "taux_benevolat": 18.0,
            "desend_moy": 3.5,
            "desend_max": 6.5
        }

def afficher_synthese_unique(titre, constat, points_appui, leviers_budget):
    st.markdown("---")
    with st.container(border=True):
        st.markdown(f"### Synthèse & Clés de lecture — {titre}")
        st.markdown(f"**1. Diagnostic & positionnement normatif :** {constat}")
        st.markdown(f"**2. Points d'appui & marges de manœuvre :** {points_appui}")
        st.markdown(f"**3. Arbitrages pour le Budget Prévisionnel (BP) :** {leviers_budget}")

def afficher_aide_pedagogique_onglet(titre_aide, points_cles):
    with st.expander(f"💡 Comprendre les indicateurs — {titre_aide}", expanded=False):
        for notion, explication in points_cles:
            st.markdown(f"* **{notion} :** {explication}")

# ==========================================
# 2. INGESTION BALANCES ASSOCIATIVES (ANC 2018-06)
# ==========================================
def _read_tabular_file(file_obj):
    name = getattr(file_obj, "name", "balance.csv")
    raw = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
    if name.lower().endswith((".xlsx", ".xls")):
        try:
            return pd.read_excel(io.BytesIO(raw))
        except Exception:
            return None
    for enc in ("utf-8", "latin-1", "cp1252"):
        for sep in (";", ",", "\t"):
            try:
                tmp = pd.read_csv(io.BytesIO(raw), sep=sep, encoding=enc)
                if tmp.shape[1] >= 2:
                    return tmp
            except Exception:
                continue
    return None

def _to_num(series):
    return pd.to_numeric(
        series.astype(str).str.replace(r"[\xa0\s]", "", regex=True).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0.0)

def parse_balance_asso(file_obj):
    name = getattr(file_obj, "name", "balance_asso_2025.csv")
    ym = re.search(r"(20[1-3]\d)", name)
    annee = ym.group(1) if ym else None
    df = _read_tabular_file(file_obj)
    if df is None or df.empty:
        return None

    cols_l = {c: str(c).strip().lower() for c in df.columns}
    compte_col = next((c for c, cl in cols_l.items() if any(k in cl for k in ("compte", "article", "nature"))), df.columns[0])
    lib_col = next((c for c, cl in cols_l.items() if "libell" in cl or "intitul" in cl), None)
    deb_col = next((c for c, cl in cols_l.items() if "débit" in cl or "debit" in cl), None)
    cred_col = next((c for c, cl in cols_l.items() if "crédit" in cl or "credit" in cl), None)
    solde_col = next((c for c, cl in cols_l.items() if "solde" in cl or "montant" in cl), None)

    raw_c = df[compte_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    debits = _to_num(df[deb_col]) if deb_col else (_to_num(df[solde_col]).abs() if solde_col else pd.Series(0.0, index=df.index))
    credits = _to_num(df[cred_col]) if cred_col else (_to_num(df[solde_col]).abs() if solde_col else pd.Series(0.0, index=df.index))

    valid = raw_c.str.match(r"^\d{2,}$")
    if lib_col:
        valid = valid & ~df[lib_col].astype(str).str.lower().str.contains(r"\b(total|totaux)\b", regex=True)

    c_val, d_val, cr_val = raw_c[valid], debits[valid], credits[valid]
    norm_c = c_val.apply(lambda c: c.rstrip("0") if len(c.rstrip("0")) >= 2 else c[:2]) if (c_val.str.len().nunique() == 1 and c_val.str.len().iloc[0] >= 5) else c_val
    uniques = set(norm_c.unique())
    parents = {c for c in uniques if any(o != c and o.startswith(c) for o in uniques)}
    is_leaf = ~norm_c.isin(parents)

    comptes, debits, credits = c_val[is_leaf], d_val[is_leaf], cr_val[is_leaf]

    def sum_p(pref, side="debit"):
        m = comptes.apply(lambda x: any(x.startswith(p) for p in pref))
        return float(debits[m].sum() if side == "debit" else credits[m].sum())

    subv_74 = sum_p(["74"], "credit")
    prest_70 = sum_p(["70"], "credit")
    cotis_756 = sum_p(["756"], "credit")
    mecenat_754 = sum_p(["754"], "credit")
    rep_fd_789 = sum_p(["789"], "credit")
    aut_p = max(0.0, sum_p(["75", "76", "77", "78", "79"], "credit") - cotis_756 - mecenat_754 - rep_fd_789)

    ach_60 = sum_p(["60"], "debit")
    ext_61_62 = sum_p(["61", "62"], "debit")
    imp_63 = sum_p(["63"], "debit")
    ms_64 = sum_p(["64"], "debit")
    aut_65_67 = sum_p(["65", "67"], "debit")
    fin_66 = sum_p(["66"], "debit")
    amo_681 = sum_p(["681", "686", "687"], "debit")
    dot_fd_689 = sum_p(["689"], "debit")

    fonds_prop_10 = sum_p(["10", "11"], "credit")
    fonds_ded_194 = sum_p(["194"], "credit")
    dette_16 = sum_p(["164", "168"], "credit")
    remb_16 = sum_p(["164", "168"], "debit")
    tres_51 = sum_p(["50", "512", "514", "53"], "debit")
    vol_87 = sum_p(["87"], "credit")

    tp_hors_fd = subv_74 + prest_70 + cotis_756 + mecenat_754 + aut_p
    tc_hors_fd = ach_60 + ext_61_62 + imp_63 + ms_64 + aut_65_67 + fin_66 + amo_681
    if tp_hors_fd <= 0 and tc_hors_fd <= 0:
        return None

    rn = (tp_hors_fd + rep_fd_789) - (tc_hors_fd + dot_fd_689)
    return {
        "fichier": name, "annee": annee,
        "subv_74": subv_74, "prest_70": prest_70, "cotis_756": cotis_756, "mecenat_754": mecenat_754, "aut_p": aut_p,
        "rep_fd_789": rep_fd_789, "dot_fd_689": dot_fd_689,
        "ach_60": ach_60, "ext_61_62": ext_61_62 + imp_63 + aut_65_67,
        "ms_64": ms_64, "fin_66": fin_66, "amo_681": amo_681,
        "tp_hors_fd": tp_hors_fd, "tc_hors_fd": tc_hors_fd,
        "rn": rn, "caf": rn + amo_681 + (dot_fd_689 - rep_fd_789),
        "fonds_prop": fonds_prop_10, "fonds_ded": fonds_ded_194,
        "dette": dette_16, "remb": remb_16, "tres": tres_51, "vol_87": vol_87
    }

# ==========================================
# 3. GÉNÉRATEURS DE RAPPORTS PDF EXÉCUTIFS
# ==========================================
def _box(ax, x, y, w, h, fc="#ffffff", ec="#cbd5e1", lw=1.0, r=0.012):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}", facecolor=fc, edgecolor=ec, linewidth=lw, transform=ax.transAxes))

def _draw_lines(ax, x, y_start, items, wrap_w=115, step_y=0.016, gap_item=0.005, fontsize=7.0, color="#1e293b", bold_prefix=False):
    y = y_start
    for it in items:
        wrapped = textwrap.wrap(it, width=wrap_w)
        for idx, wl in enumerate(wrapped):
            ax.text(x, y, wl, fontsize=fontsize, color=color, fontweight=("bold" if (bold_prefix and idx == 0) else "normal"), transform=ax.transAxes)
            y -= step_y
        y -= gap_item
    return y

def _header_pdf(ax, titre_doc, sous_titre):
    ax.add_patch(Rectangle((0, 0.915), 1, 0.085, facecolor="#0f172a", transform=ax.transAxes))
    ax.add_patch(Rectangle((0, 0.911), 1, 0.004, facecolor="#0d9488", transform=ax.transAxes))
    ax.text(0.04, 0.962, titre_doc, color="white", fontsize=10.2, fontweight="bold", transform=ax.transAxes)
    ax.text(0.04, 0.930, sous_titre, color="#cbd5e1", fontsize=7.4, transform=ax.transAxes)

def _footer_pdf(ax, gauche, page_txt):
    ax.plot([0.04, 0.96], [0.038, 0.038], color="#cbd5e1", lw=0.8, transform=ax.transAxes)
    ax.text(0.04, 0.020, gauche, fontsize=7.0, color="#64748b", transform=ax.transAxes)
    ax.text(0.96, 0.020, page_txt, ha="right", fontsize=7.0, fontweight="bold", color="#475569", transform=ax.transAxes)

def build_pdf_asso_retro(asso, normes, ans_r, p_r, c_r, ms_r, caf_r, dep_r, fd_r, tcma_p, tcma_c, tcma_ms, gvt_pt, conv_pt, vol_pt, atouts_txt, vig_txt):
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        fig = plt.figure(figsize=(8.27, 11.69), dpi=150)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis("off")
        _header_pdf(ax, "ARCHÉTYPE ASSO  |  DOSSIER D'AUDIT RÉTROSPECTIF ANC 2018-06 (2021-2025)", f"ASSOCIATION : {asso.upper()}  •  {normes['type_court'].upper()}  •  TRAJECTOIRE & DÉPENDANCE")

        ciseau = tcma_c - tcma_p
        var_caf = caf_r[-1] - caf_r[0]
        kpis = [
            ("TCMA PRODUITS", f"+{tcma_p:.2f} % / an", "Rythme annuel ressources", "#059669"),
            ("TCMA CHARGES", f"+{tcma_c:.2f} % / an", f"Effet ciseau : {ciseau:+.2f} pts", "#dc2626" if ciseau > 0.3 else "#d97706"),
            ("DÉPENDANCE PUBLIQUE", f"{dep_r[-1]:.1f} %", f"Seuil vigilance : {normes['dep_pub_max']} %", "#0d9488" if dep_r[-1] <= normes["dep_pub_max"] else "#d97706"),
            ("FONDS DÉDIÉS 31/12", f"{fd_r[-1]:.2f} M€", f"CAF : {caf_r[-1]:.2f} M€ ({var_caf:+.2f} M€)", "#1d4ed8"),
        ]
        for i, (kt, kv, ks, kc) in enumerate(kpis):
            kx = 0.04 + i * 0.233
            _box(ax, kx, 0.815, 0.218, 0.084, fc="#f8fafc")
            ax.text(kx + 0.012, 0.880, kt, fontsize=6.9, fontweight="bold", color="#475569", transform=ax.transAxes)
            ax.text(kx + 0.012, 0.845, kv, fontsize=11.8, fontweight="bold", color=kc, transform=ax.transAxes)
            ax.text(kx + 0.012, 0.823, ks, fontsize=6.3, color="#64748b", transform=ax.transAxes)

        ax.text(0.04, 0.785, "1. EFFET DE CISEAU HISTORIQUE (PRODUITS VS CHARGES HORS FD)", fontsize=8.0, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        g1 = fig.add_axes([0.06, 0.595, 0.41, 0.175])
        g1.plot(ans_r, p_r, marker="o", color="#2563eb", lw=2, label="Produits")
        g1.plot(ans_r, c_r, marker="s", color="#dc2626", lw=2, label="Charges")
        g1.grid(True, linestyle="--", alpha=0.5)
        g1.tick_params(labelsize=7)
        g1.legend(fontsize=6.8, loc="upper left")

        ax.text(0.52, 0.785, "2. TRAJECTOIRE DE LA CAF & STOCK DE FONDS DÉDIÉS (M€)", fontsize=8.0, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        g2 = fig.add_axes([0.54, 0.595, 0.41, 0.175])
        g2.plot(ans_r, caf_r, marker="^", color="#059669", lw=2, label="CAF annuelle")
        g2.plot(ans_r, fd_r, marker="s", color="#0d9488", lw=2, label="Fonds dédiés (194)")
        g2.grid(True, linestyle="--", alpha=0.5)
        g2.tick_params(labelsize=7)
        g2.legend(fontsize=6.8, loc="upper left")

        ax.text(0.04, 0.545, "3. ÉVOLUTION DU TAUX DE DÉPENDANCE AUX SUBVENTIONS PUBLIQUES (%)", fontsize=8.0, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        g3 = fig.add_axes([0.06, 0.375, 0.41, 0.155])
        g3.bar(ans_r, dep_r, color="#0d9488", width=0.5)
        g3.axhline(normes["dep_pub_max"], color="#dc2626", linestyle="--", lw=1.2, label=f"Seuil ({normes['dep_pub_max']}%)")
        g3.grid(True, axis="y", linestyle="--", alpha=0.5)
        g3.tick_params(labelsize=7)
        g3.legend(fontsize=6.5, loc="upper left")

        _box(ax, 0.51, 0.375, 0.45, 0.155, fc="#f8fafc")
        ax.text(0.525, 0.508, "LECTURE ANC 2018-06 : FONDS DÉDIÉS & DÉPENDANCE", fontsize=7.6, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        diag_fd = [
            f"• Stock de fonds dédiés (194) : {fd_r[-1]:.2f} M€ à la clôture (subventions affectées à réaliser sur les prochains exercices).",
            f"• Dépendance publique : {dep_r[-1]:.1f} % des produits (seuil sectoriel : {normes['dep_pub_max']} %).",
            "• Enjeu : Sécuriser les conventions pluriannuelles (CPO) et diversifier les ressources propres."
        ]
        _draw_lines(ax, 0.525, 0.480, diag_fd, wrap_w=56, step_y=0.017, gap_item=0.005, fontsize=6.8)

        ax.text(0.04, 0.335, "TABLEAU RÉTROSPECTIF DES SOLDES ASSOCIATIFS ANC 2018-06 (2021-2025)", fontsize=8.0, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        tax = fig.add_axes([0.04, 0.145, 0.92, 0.175])
        tax.axis("off")
        t_rows = [
            ["Produits d'exploitation hors FD (M€)"] + [f"{v:.2f}" for v in p_r],
            ["Charges d'exploitation hors FD (M€)"] + [f"{v:.2f}" for v in c_r],
            ["Masse salariale Cpte 64 (M€)"] + [f"{v:.2f}" for v in ms_r],
            ["Capacité d'Autofinancement — CAF (M€)"] + [f"{v:.2f}" for v in caf_r],
            ["Fonds dédiés au bilan — Cpte 194 (M€)"] + [f"{v:.2f}" for v in fd_r],
            [f"Taux dépendance publique (Seuil {normes['dep_pub_max']}%)"] + [f"{v:.1f} %" for v in dep_r],
        ]
        tbl = tax.table(cellText=t_rows, colLabels=["Indicateurs ANC 2018-06"] + ans_r, loc="center", cellLoc="center", colWidths=[0.44] + [0.112]*len(ans_r))
        tbl.set_fontsize(6.9)
        tbl.scale(1, 1.18)
        for (r, c), cell in tbl.get_celld().items():
            cell.set_edgecolor("#cbd5e1")
            if r == 0:
                cell.set_facecolor("#1e293b")
                cell.get_text().set_color("white")
                cell.get_text().set_weight("bold")
            elif c == 0:
                cell.get_text().set_ha("left")

        _box(ax, 0.04, 0.050, 0.92, 0.080, fc="#fffbeb", ec="#fcd34d")
        ax.text(0.055, 0.110, "SYNTHÈSE DE LA TRAJECTOIRE PASSÉE", fontsize=7.6, fontweight="bold", color="#92400e", transform=ax.transAxes)
        cis_lines = [
            f"• Effet ciseau ({ciseau:+.2f} pts/an) : les charges (+{tcma_c:.2f} %/an) évoluent {'plus vite' if ciseau > 0 else 'moins vite'} que les produits (+{tcma_p:.2f} %/an).",
            f"• La CAF s'établit à {caf_r[-1]:.2f} M€ ({(caf_r[-1]/p_r[-1]*100):.1f} % vs cible sectorielle de {normes['taux_caf']} %)."
        ]
        _draw_lines(ax, 0.055, 0.089, cis_lines, wrap_w=124, step_y=0.016, gap_item=0.004, fontsize=6.9)
        _footer_pdf(ax, "Dossier Rétrospectif ANC 2018-06  •  Cabinet QUINZE", "Page 1/2")
        pdf.savefig(fig)
        plt.close(fig)

        fig2 = plt.figure(figsize=(8.27, 11.69), dpi=150)
        ax2 = fig2.add_axes([0, 0, 1, 1])
        ax2.axis("off")
        _header_pdf(ax2, "ARCHÉTYPE ASSO  |  DÉCOMPOSITION DE LA DÉRIVE RH & SYNTHÈSE", f"ASSOCIATION : {asso.upper()}  •  LEVIERS MANAGÉRIAUX & PISTES D'AUDIT")

        ax2.text(0.04, 0.875, "1. DÉCOMPOSITION DE LA PROGRESSION DE LA MASSE SALARIALE (COMPTE 64)", fontsize=8.2, fontweight="bold", color="#0f172a", transform=ax2.transAxes)
        grh = fig2.add_axes([0.24, 0.675, 0.70, 0.180])
        cats = ["Effet Volume\n(Postes / Rempl.)", "Revalorisations\n(Conv. / SMIC)", "GVT Mécanique\n(Ancienneté)"]
        vals = [vol_pt, conv_pt, gvt_pt]
        cols_rh = ["#059669" if vol_pt <= 0.3 else "#dc2626", "#d97706", "#64748b"]
        grh.barh(cats, vals, color=cols_rh, height=0.52)
        grh.set_xlabel("Contribution annuelle (% / an)", fontsize=7.5)
        grh.grid(True, axis="x", linestyle="--", alpha=0.5)
        grh.tick_params(labelsize=7.5)

        _box(ax2, 0.04, 0.515, 0.92, 0.125, fc="#f8fafc", ec="#cbd5e1")
        ax2.text(0.055, 0.618, "LECTURE MANAGÉRIALE : FACTEURS SUBIS VS LEVIER INTERNE (NORME VOLUME <= 0,3 PT/AN)", fontsize=7.7, fontweight="bold", color="#0f172a", transform=ax2.transAxes)
        rh_lines = [
            f"• Sur +{tcma_ms:.2f} %/an de hausse salariale, {(gvt_pt+conv_pt):.2f} points sont automatiques (ancienneté GVT + accords de branche).",
            f"• {vol_pt:+.2f} points relèvent des choix d'organisation (créations de postes non financées, remplacements, heures supplémentaires).",
            f"• Le poids global de la masse salariale atteint {(ms_r[-1]/p_r[-1]*100):.1f} % des produits (repère sectoriel : {normes['poids_ms']} %)."
        ]
        _draw_lines(ax2, 0.055, 0.592, rh_lines, wrap_w=122, step_y=0.017, gap_item=0.005, fontsize=7.1)

        ax2.text(0.04, 0.480, "2. BILAN STRATÉGIQUE : ATOUTS & LEVIERS D'AJUSTEMENT BUDGÉTAIRE", fontsize=8.2, fontweight="bold", color="#0f172a", transform=ax2.transAxes)
        _box(ax2, 0.04, 0.115, 0.445, 0.350, fc="#ecfdf5", ec="#6ee7b7")
        ax2.text(0.055, 0.442, "ATOUTS & MARGES DE MANŒUVRE", fontsize=7.8, fontweight="bold", color="#065f46", transform=ax2.transAxes)
        _draw_lines(ax2, 0.055, 0.412, [f"• {x}" for x in atouts_txt], wrap_w=54, step_y=0.019, gap_item=0.010, fontsize=7.1)

        _box(ax2, 0.515, 0.115, 0.445, 0.350, fc="#fffbeb", ec="#fcd34d")
        ax2.text(0.530, 0.442, "LEVIERS BUDGÉTAIRES & POINTS DE VIGILANCE", fontsize=7.8, fontweight="bold", color="#92400e", transform=ax2.transAxes)
        _draw_lines(ax2, 0.530, 0.412, [f"• {x}" for x in vig_txt], wrap_w=54, step_y=0.019, gap_item=0.010, fontsize=7.1)

        _footer_pdf(ax2, "Dossier Rétrospectif ANC 2018-06  •  Cabinet QUINZE", "Page 2/2")
        pdf.savefig(fig2)
        plt.close(fig2)
    return buf.getvalue()

def build_pdf_asso_tbi(asso, normes, etp_sal, etp_ben, score, tp, ach, ext_f, va, ms, ebe, fin_val, amo_val, var_fd, rn, caf, remb, aut, fonds_prop, fonds_ded, dette, cap_d, treso, mois_tr, dep_pub, conc_fin, poids_sup, levier_vol, synthese_tbi):
    buf = io.BytesIO()
    fig = plt.figure(figsize=(8.27, 11.69), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    _header_pdf(ax, "ARCHÉTYPE ASSO  |  TABLEAU DE BORD EXÉCUTIF ANC 2018-06 & NORMES", f"ASSOCIATION : {asso.upper()} ({etp_sal:.1f} ETP SALARIÉS • {etp_ben:.1f} ETP BÉNÉVOLES • {normes['type_court'].upper()})")

    sc_col = "#059669" if score >= 70 else ("#d97706" if score >= 45 else "#dc2626")
    sc_bg = "#ecfdf5" if score >= 70 else ("#fffbeb" if score >= 45 else "#fef2f2")
    _box(ax, 0.04, 0.815, 0.92, 0.084, fc=sc_bg, ec=sc_col, lw=1.4)
    ax.text(0.055, 0.876, "HEALTH SCORE ANC 2018-06", fontsize=7.0, fontweight="bold", color="#475569", transform=ax.transAxes)
    ax.text(0.055, 0.832, f"{score}/100", fontsize=17.0, fontweight="bold", color=sc_col, transform=ax.transAxes)
    ax.plot([0.255, 0.255], [0.823, 0.891], color="#cbd5e1", lw=1.0, transform=ax.transAxes)

    statut_msg = "MODÈLE ASSOCIATIF SOLIDE ET CONFORME AUX NORMES" if score >= 70 else ("TENSION SUR LA TRÉSORERIE OU DÉPENDANCE ÉLEVÉE" if score >= 45 else "ALERTE STRUCTURELLE : RISQUE DE LIQUIDITÉ OU DE MODÈLE")
    ax.text(0.275, 0.874, statut_msg, fontsize=8.0, fontweight="bold", color=sc_col, transform=ax.transAxes)
    tcaf = (caf / tp * 100) if tp > 0 else 0
    tebe = (ebe / tp * 100) if tp > 0 else 0
    ax.text(0.275, 0.850, f"EBE : {tebe:.1f}% (Moy. {normes['taux_ebe']}%) • CAF : {tcaf:.1f}% (Cible {normes['taux_caf']}%) • Réserve Tréso : {mois_tr:.1f} mois (Cible {normes['mois_treso']} m.)", fontsize=6.8, color="#1e293b", transform=ax.transAxes)
    ax.text(0.275, 0.827, f"Dépendance Publique : {dep_pub:.1f}% (Max {normes['dep_pub_max']}%) • Fonds Propres : {fonds_prop:.2f} M€ • Levier Bénévolat : +{levier_vol:.1f} %", fontsize=6.8, color="#1e293b", transform=ax.transAxes)

    _box(ax, 0.04, 0.450, 0.50, 0.350, fc="#ffffff")
    ax.add_patch(Rectangle((0.04, 0.762), 0.50, 0.038, facecolor="#1e293b", transform=ax.transAxes))
    ax.text(0.055, 0.775, "1. CASCADE DES SIG ASSOCIATIFS (ANC 2018-06)", color="white", fontsize=7.9, fontweight="bold", transform=ax.transAxes)
    sig_rows = [
        ("Produits d'Exploitation (hors FD)", f"{tp:.2f} M€", "100%", True, "#1d4ed8"),
        ("- Achats & Fournitures (60)", f"-{ach:.2f} M€", "-", False, "#475569"),
        ("- Services extérieurs & Locaux (61-62)", f"-{ext_f:.2f} M€", "-", False, "#475569"),
        ("= Valeur Ajoutée (VA)", f"{va:.2f} M€", "-", True, "#0f172a"),
        ("- Masse Salariale chargée (64)", f"-{ms:.2f} M€", f"Moy. {normes['poids_ms']}%", False, "#475569"),
        ("= EBE (Excédent Brut)", f"{ebe:.2f} M€ ({tebe:.1f}%)", f"Moy. {normes['taux_ebe']}%", True, "#059669" if tebe >= normes['taux_ebe'] else "#dc2626"),
        ("- Intérêts emprunts (66)", f"-{fin_val:.2f} M€", "< 2,0%", False, "#475569"),
        ("= CAPACITÉ D'AUTOFINANCEMENT", f"{caf:.2f} M€ ({tcaf:.1f}%)", f"Cible {normes['taux_caf']}%", True, "#0d9488"),
        ("- Dotations Amortissements (681)", f"-{amo_val:.2f} M€", "-", False, "#475569"),
        ("+/- Variation Fonds Dédiés (789-689)", f"{var_fd:+.2f} M€", "Neutre", False, "#475569"),
        ("= Résultat Net Comptable", f"{rn:+.2f} M€", "Norme >= 0", True, "#059669" if rn >= 0 else "#dc2626"),
    ]
    ys = 0.735
    for lbl, val, nrm, bld, col in sig_rows:
        if bld:
            ax.add_patch(Rectangle((0.046, ys - 0.007), 0.488, 0.024, facecolor="#f1f5f9", transform=ax.transAxes))
        ax.text(0.052, ys, lbl, fontsize=6.8, fontweight="bold" if bld else "normal", color="#1e293b", transform=ax.transAxes)
        ax.text(0.420, ys, val, ha="right", fontsize=6.9, fontweight="bold" if bld else "normal", color=col, transform=ax.transAxes)
        ax.text(0.532, ys, nrm, ha="right", fontsize=6.4, color="#64748b", transform=ax.transAxes)
        ys -= 0.027

    _box(ax, 0.56, 0.635, 0.40, 0.165, fc="#ffffff")
    ax.add_patch(Rectangle((0.56, 0.762), 0.40, 0.038, facecolor="#1e293b", transform=ax.transAxes))
    ax.text(0.575, 0.775, "2. FINANCEMENTS & CLASSE 8", color="white", fontsize=7.9, fontweight="bold", transform=ax.transAxes)
    sc_items = [
        ("Dépendance publique :", f"{dep_pub:.1f} % (Max {normes['dep_pub_max']} %)"),
        ("1er financeur :", f"{conc_fin:.1f} % (Max {normes['conc_fin_max']} %)"),
        ("Poids frais support :", f"{poids_sup:.1f} % (Max {normes['poids_support']} %)"),
        ("Levier Bénévolat (Cl. 8) :", f"+{levier_vol:.1f} % du budget"),
    ]
    ysc = 0.732
    for k, v in sc_items:
        ax.text(0.572, ysc, k, fontsize=6.7, color="#475569", transform=ax.transAxes)
        ax.text(0.948, ysc, v, ha="right", fontsize=6.8, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        ysc -= 0.028

    _box(ax, 0.56, 0.450, 0.40, 0.170, fc="#ffffff")
    ax.add_patch(Rectangle((0.56, 0.582), 0.40, 0.038, facecolor="#1e293b", transform=ax.transAxes))
    ax.text(0.575, 0.595, "3. BILAN, FONDS DÉDIÉS & TRÉSO", color="white", fontsize=7.9, fontweight="bold", transform=ax.transAxes)
    det_items = [
        ("Fonds propres (10/11) :", f"{fonds_prop:.2f} M€"),
        ("Fonds dédiés (194) :", f"{fonds_ded:.2f} M€"),
        ("Encours Dette (164) :", f"{dette:.2f} M€ ({cap_d:.1f} ans)"),
        ("Trésorerie disponible :", f"{treso:.2f} M€ ({mois_tr:.1f} mois)"),
    ]
    yd = 0.552
    for k, v in det_items:
        ax.text(0.572, yd, k, fontsize=6.7, color="#475569", transform=ax.transAxes)
        ax.text(0.948, yd, v, ha="right", fontsize=6.8, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        yd -= 0.028

    _box(ax, 0.04, 0.275, 0.92, 0.160, fc="#f8fafc", ec="#94a3b8")
    ax.text(0.055, 0.415, "4. SYNTHÈSE EXÉCUTIVE & CLÉS D'AJUSTEMENT DU BUDGET PRÉVISIONNEL", fontsize=8.0, fontweight="bold", color="#0f172a", transform=ax.transAxes)
    _draw_lines(ax, 0.055, 0.390, synthese_tbi, wrap_w=124, step_y=0.017, gap_item=0.005, fontsize=7.0)

    ax.text(0.04, 0.248, f"REPÈRES PÉDAGOGIQUES ASSOCIATIFS ({normes['type_court'].upper()})", fontsize=7.5, fontweight="bold", color="#0f172a", transform=ax.transAxes)
    refs = [
        ("FONDS DÉDIÉS (ANC 2018-06)", "Subventions à consommer", ["689 = report sur N+1 (-)", "789 = reprise sur N (+)", "Ne pas confondre avec bénéfice"]),
        ("RÉSERVE DE TRÉSORERIE", "Absorber délais subventions", [f"Cible : >= {normes['mois_treso']} mois", "Alerte : < 2 mois de charges", "Sécurise le BFR public"]),
        ("DÉPENDANCE PUBLIQUE", "Risque de concentration", [f"Subv. 74 <= {normes['dep_pub_max']} % prod.", f"1er financeur <= {normes['conc_fin_max']} %", "Développer prestations / mécénat"]),
        ("CLASSE 8 (BÉNÉVOLAT)", "Contributions en nature", ["Valorise l'engagement réel", "Prouve l'ancrage aux bailleurs", "Équilibre Pied de Bilan"]),
    ]
    for i, (rt, rf, rbullets) in enumerate(refs):
        rx = 0.04 + i * 0.233
        _box(ax, rx, 0.055, 0.218, 0.180, fc="#ffffff")
        ax.text(rx + 0.012, 0.215, rt, fontsize=6.9, fontweight="bold", color="#0f172a", transform=ax.transAxes)
        ax.text(rx + 0.012, 0.193, rf, fontsize=6.2, color="#64748b", transform=ax.transAxes)
        _draw_lines(ax, rx + 0.012, 0.160, [f"• {b}" for b in rbullets], wrap_w=32, step_y=0.024, gap_item=0.004, fontsize=6.6)

    _footer_pdf(ax, "Rapport certifié conforme aux règles ANC 2018-06  •  Cabinet QUINZE", "Page 1/1")
    fig.savefig(buf, format="pdf")
    plt.close(fig)
    return buf.getvalue()

def build_pdf_asso_prospective(asso, normes, ppi_annuel, df_proj, feuille_route):
    buf = io.BytesIO()
    fig = plt.figure(figsize=(8.27, 11.69), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    _header_pdf(ax, "ARCHÉTYPE ASSO  |  TRAJECTOIRE PROSPECTIVE & PROJET ASSOCIATIF À 5 ANS (2026-2030)", f"ASSOCIATION : {asso.upper()}  •  {normes['type_court'].upper()}  •  INVESTISSEMENTS : {ppi_annuel:.2f} M€/AN")

    caf_30 = df_proj["CAF (M€)"].iloc[-1]
    tcaf_30 = (caf_30 / df_proj["Produits (M€)"].iloc[-1] * 100)
    fp_30 = df_proj["Fonds Propres (M€)"].iloc[-1]
    dep_30 = df_proj["Dépendance Pub. (%)"].iloc[-1]
    m_30 = df_proj["Trésorerie (mois)"].iloc[-1]
    tr_30 = df_proj["Trésorerie (M€)"].iloc[-1]

    kpis = [
        ("CAF 2030", f"+{caf_30:.2f} M€", f"Taux : {tcaf_30:.1f}% (Cible {normes['taux_caf']}%)", "#059669" if tcaf_30 >= normes["taux_caf"] else "#d97706"),
        ("FONDS PROPRES 2030", f"{fp_30:.2f} M€", f"Variation : {(fp_30 - df_proj['Fonds Propres (M€)'].iloc[0]):+.2f} M€", "#0f172a"),
        ("DÉPENDANCE 2030", f"{dep_30:.1f} %", f"Seuil max : {normes['dep_pub_max']} %", "#0d9488" if dep_30 <= normes["dep_pub_max"] else "#d97706"),
        ("TRÉSORERIE 2030", f"{m_30:.1f} mois", f"{tr_30:.2f} M€ (Cible {normes['mois_treso']} mois)", "#1d4ed8" if m_30 >= 3.0 else "#dc2626"),
    ]
    for i, (kt, kv, ks, kc) in enumerate(kpis):
        kx = 0.04 + i * 0.233
        _box(ax, kx, 0.815, 0.218, 0.084, fc="#f8fafc")
        ax.text(kx + 0.012, 0.880, kt, fontsize=6.9, fontweight="bold", color="#475569", transform=ax.transAxes)
        ax.text(kx + 0.012, 0.845, kv, fontsize=11.8, fontweight="bold", color=kc, transform=ax.transAxes)
        ax.text(kx + 0.012, 0.823, ks, fontsize=6.5, color="#64748b", transform=ax.transAxes)

    ax.text(0.04, 0.785, "1. PROJECTION PRODUITS VS CHARGES (INCLUANT DÉVELOPPEMENT & OPEX)", fontsize=7.9, fontweight="bold", color="#0f172a", transform=ax.transAxes)
    gp1 = fig.add_axes([0.06, 0.595, 0.41, 0.175])
    gp1.plot(df_proj["Année"], df_proj["Produits (M€)"], marker="o", color="#2563eb", lw=2, label="Produits")
    gp1.plot(df_proj["Année"], df_proj["Charges (M€)"], marker="s", color="#dc2626", lw=2, label="Charges")
    gp1.grid(True, linestyle="--", alpha=0.5)
    gp1.tick_params(labelsize=7)
    gp1.legend(fontsize=6.8, loc="upper left")

    ax.text(0.52, 0.785, "2. TRAJECTOIRE FONDS PROPRES & TRÉSORERIE (M€)", fontsize=7.9, fontweight="bold", color="#0f172a", transform=ax.transAxes)
    gp2 = fig.add_axes([0.54, 0.595, 0.41, 0.175])
    gp2.plot(df_proj["Année"], df_proj["Fonds Propres (M€)"], marker="^", color="#059669", lw=2, label="Fonds Propres")
    gp2.plot(df_proj["Année"], df_proj["Trésorerie (M€)"], marker="s", linestyle=":", color="#2563eb", lw=2, label="Trésorerie")
    gp2.grid(True, linestyle="--", alpha=0.5)
    gp2.tick_params(labelsize=7)
    gp2.legend(fontsize=6.8, loc="upper left")

    ax.text(0.04, 0.548, "TABLEAU PLURIANNUEL DES TRAJECTOIRES ASSOCIATIVES (2026-2030)", fontsize=8.0, fontweight="bold", color="#0f172a", transform=ax.transAxes)
    tp_ax = fig.add_axes([0.04, 0.340, 0.92, 0.195])
    tp_ax.axis("off")
    t_data = [
        ["Produits d'exploitation (M€)"] + [f"{v:.2f}" for v in df_proj["Produits (M€)"]],
        ["Charges d'exploitation (M€)"] + [f"{v:.2f}" for v in df_proj["Charges (M€)"]],
        [f"CAF (Norme cible : >= {normes['taux_caf']} %)"] + [f"{v:.2f}" for v in df_proj["CAF (M€)"]],
        ["Fonds propres cumulés (M€)"] + [f"{v:.2f}" for v in df_proj["Fonds Propres (M€)"]],
        [f"Dépendance publique (Max {normes['dep_pub_max']} %)"] + [f"{v:.1f} %" for v in df_proj["Dépendance Pub. (%)"]],
        ["Encours dette (M€)"] + [f"{v:.2f}" for v in df_proj["Dette (M€)"]],
        [f"Trésorerie de sécurité (Cible {normes['mois_treso']} mois)"] + [f"{v:.1f} m." for v in df_proj["Trésorerie (mois)"]],
    ]
    tbl = tp_ax.table(cellText=t_data, colLabels=["Indicateurs & Normes ANC 2018-06"] + list(df_proj["Année"]), loc="center", cellLoc="center", colWidths=[0.46] + [0.108]*5)
    tbl.set_fontsize(6.9)
    tbl.scale(1, 1.18)
    for (r, c), cell in tbl.get_celld().items():
        cell.set_edgecolor("#cbd5e1")
        if r == 0:
            cell.set_facecolor("#1e293b")
            cell.get_text().set_color("white")
            cell.get_text().set_weight("bold")
        elif c == 0:
            cell.get_text().set_ha("left")

    _box(ax, 0.04, 0.055, 0.92, 0.270, fc="#ecfdf5", ec="#6ee7b7")
    ax.text(0.055, 0.305, "SYNTHÈSE PROSPECTIVE & FEUILLE DE ROUTE POUR LE CONSEIL D'ADMINISTRATION", fontsize=7.8, fontweight="bold", color="#065f46", transform=ax.transAxes)
    _draw_lines(ax, 0.055, 0.280, feuille_route, wrap_w=124, step_y=0.016, gap_item=0.005, fontsize=7.0)

    _footer_pdf(ax, "Document de prospective ANC 2018-06  •  Cabinet QUINZE — Ingénierie & Conseil Associatif", "Page 1/1")
    fig.savefig(buf, format="pdf")
    plt.close(fig)
    return buf.getvalue()

# ==========================================
# 4. BARRE LATÉRALE : INGESTION & PARAMÉTRAGE
# ==========================================
st.sidebar.markdown("## 🤝 Archétype Asso")
secteur_asso = st.sidebar.selectbox(
    "Secteur & Convention de l'association",
    [
        "Action sociale, solidarité & insertion (IAE / CHRS)",
        "Culture, patrimoine & spectacle vivant",
        "Sport, jeunesse & éducation populaire",
        "Santé, médico-social & handicap",
        "Environnement, tiers-lieux & développement local",
        "Fédération, tête de réseau & syndicat professionnel"
    ]
)
normes_base = get_normes_asso(secteur_asso)
asso_nom = st.sidebar.text_input("Nom de l'association", "Association Passerelles & Territoires")

with st.sidebar.expander("🎯 Barème & Normes de référence (Personnalisable)", expanded=False):
    st.caption("Ajustez les seuils selon les exigences de vos financeurs ou de votre fédération.")
    cn1, cn2 = st.columns(2)
    n_dep_max = cn1.number_input("Dépendance pub. max (%)", value=float(normes_base["dep_pub_max"]), step=5.0)
    n_conc_max = cn2.number_input("1er financeur max (%)", value=float(normes_base["conc_fin_max"]), step=5.0)
    cn3, cn4 = st.columns(2)
    n_ebe = cn3.number_input("Taux EBE moy. (%)", value=float(normes_base["taux_ebe"]), step=0.5)
    n_caf = cn4.number_input("Taux CAF cible (%)", value=float(normes_base["taux_caf"]), step=0.5)
    cn5, cn6 = st.columns(2)
    n_ms = cn5.number_input("Poids MS moy. (%)", value=float(normes_base["poids_ms"]), step=1.0)
    n_sup = cn6.number_input("Frais support max (%)", value=float(normes_base["poids_support"]), step=1.0)
    cn7, cn8 = st.columns(2)
    n_mois_tr = cn7.number_input("Tréso cible (mois)", value=float(normes_base["mois_treso"]), step=0.5)
    n_des_max = cn8.number_input("Désendettement max (ans)", value=float(normes_base["desend_max"]), step=0.5)

normes = dict(normes_base)
normes.update({
    "dep_pub_max": n_dep_max,
    "conc_fin_max": n_conc_max,
    "taux_ebe": n_ebe,
    "taux_caf": n_caf,
    "poids_ms": n_ms,
    "poids_support": n_sup,
    "mois_treso": n_mois_tr,
    "desend_max": n_des_max,
})

c_sb1, c_sb2 = st.sidebar.columns(2)
etp_salaries = c_sb1.number_input("Effectif salarié (ETP)", min_value=0.0, value=24.0, step=1.0)
etp_benevoles = c_sb2.number_input("Bénévolat (ETP estimé)", min_value=0.0, value=3.5, step=0.5)

st.sidebar.markdown("### 1. Ingestion Balances (2021-2025)")
fichiers = st.sidebar.file_uploader("Importer balances comptables (.csv, .xlsx)", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
balances_parsees = []
if fichiers:
    for f in fichiers:
        rb = parse_balance_asso(f)
        if rb:
            balances_parsees.append(rb)
balances_parsees.sort(key=lambda x: x["annee"] or x["fichier"])
source_reelle = len(balances_parsees) > 0

if source_reelle:
    dern = balances_parsees[-1]
    d_subv = dern["subv_74"]
    d_prest = dern["prest_70"]
    d_cotis = dern["cotis_756"]
    d_mec = dern["mecenat_754"]
    d_autp = dern["aut_p"]
    d_rep_fd = dern["rep_fd_789"]
    d_dot_fd = dern["dot_fd_689"]
    d_ach = dern["ach_60"]
    d_ext = dern["ext_61_62"]
    d_ms = dern["ms_64"]
    d_fin = dern["fin_66"]
    d_amo = dern["amo_681"]
    d_fp = dern["fonds_prop"] if dern["fonds_prop"] > 0 else 380000.0
    d_fd_bil = dern["fonds_ded"] if dern["fonds_ded"] > 0 else 95000.0
    d_det = dern["dette"] if dern["dette"] > 0 else 210000.0
    d_remb = dern["remb"] if dern["remb"] > 0 else 35000.0
    d_tres = dern["tres"] if dern["tres"] > 0 else 480000.0
    d_vol = dern["vol_87"] if dern["vol_87"] > 0 else round(etp_benevoles * 34000)
else:
    budget_base = max(150000.0, etp_salaries * 62000.0)
    d_subv = round(budget_base * (min(80.0, normes["dep_pub_max"] - 5.0) / 100.0))
    d_prest = round(budget_base * 0.22)
    d_cotis = round(budget_base * 0.04)
    d_mec = round(budget_base * 0.04)
    d_autp = max(0.0, budget_base - d_subv - d_prest - d_cotis - d_mec)
    d_rep_fd = 45000.0
    d_dot_fd = 52000.0
    d_ach = round(budget_base * 0.09)
    d_ext = round(budget_base * 0.18)
    d_ms = round(budget_base * (normes["poids_ms"] / 100.0))
    d_fin = 9000.0
    d_amo = 48000.0
    d_fp = 380000.0
    d_fd_bil = 95000.0
    d_det = 210000.0
    d_remb = 35000.0
    d_tres = round((budget_base / 12.0) * (normes["mois_treso"] + 0.3))
    d_vol = round(etp_benevoles * 34000 + 25000)

st.sidebar.markdown("### 2. Ressources & Structure des Financements")
subv_74 = st.sidebar.number_input("Subventions d'exploitation (74) (€)", value=float(d_subv), step=10000.0)
part_1er_fin = st.sidebar.slider("Part du 1er financeur dans le total produits (%)", 5, 90, int(min(85, round(d_subv / max(1.0, d_subv + d_prest + d_cotis + d_mec + d_autp) * 100 * 0.65))), 5)
prest_70 = st.sidebar.number_input("Prestations, ventes & billetterie (70) (€)", value=float(d_prest), step=10000.0)
cotis_756 = st.sidebar.number_input("Cotisations des membres (756) (€)", value=float(d_cotis), step=2000.0)
mecenat_754 = st.sidebar.number_input("Dons, mécénat & fonds privés (754) (€)", value=float(d_mec), step=5000.0)
aut_p = st.sidebar.number_input("Autres produits (75/76/79) (€)", value=float(d_autp), step=5000.0)

with st.sidebar.expander("📦 3. Fonds Dédiés (ANC 2018-06) & Classe 8", expanded=True):
    rep_fd_789 = st.number_input("Reprises sur fonds dédiés (789) (+) (€)", value=float(d_rep_fd), step=5000.0)
    dot_fd_689 = st.number_input("Reports en fonds dédiés (689) (-) (€)", value=float(d_dot_fd), step=5000.0)
    vol_87 = st.number_input("Contributions volontaires en nature (Classe 8) (€)", value=float(d_vol), step=5000.0)

with st.sidebar.expander("👥 4. Granularité Charges (Missions vs Support)", expanded=True):
    ms_missions = st.number_input("Salaires chargés — Missions sociales terrain (64)", value=float(round(d_ms * 0.84)), step=10000.0)
    ms_support = st.number_input("Salaires chargés — Direction, admin & support (64)", value=float(round(d_ms * 0.16)), step=5000.0)
    ms_totale = ms_missions + ms_support
    ach_60 = st.number_input("Achats & fournitures directes (60) (€)", value=float(d_ach), step=5000.0)
    ext_missions = st.number_input("Charges externes — Actions & projets (61-62)", value=float(round(d_ext * 0.72)), step=5000.0)
    ext_support = st.number_input("Charges externes — Siège, assur., honoraires (61-62)", value=float(round(d_ext * 0.28)), step=5000.0)
    ext_totale = ext_missions + ext_support

st.sidebar.markdown("### 5. Bilan, Fonds Propres & Trésorerie")
amo_681 = st.sidebar.number_input("Dotations aux amortissements (681) (€)", value=float(d_amo), step=5000.0)
fin_66 = st.sidebar.number_input("Intérêts des emprunts (66) (€)", value=float(d_fin), step=1000.0)
fonds_prop_10 = st.sidebar.number_input("Fonds propres / associatifs (10/11) (€)", value=float(d_fp), step=10000.0)
fonds_ded_194 = st.sidebar.number_input("Stock fonds dédiés au bilan (194) (€)", value=float(d_fd_bil), step=5000.0)
dette_16 = st.sidebar.number_input("Encours dette financière (164) (€)", value=float(d_det), step=10000.0)
remb_16 = st.sidebar.number_input("Remboursement annuel capital emprunts (€)", value=float(d_remb), step=5000.0)
treso_51 = st.sidebar.number_input("Trésorerie & placements liquides (50/51) (€)", value=float(d_tres), step=10000.0)

# ==========================================
# 5. CALCULS DU MOTEUR FINANCIER ASSOCIATIF (ANC 2018-06)
# ==========================================
tp_hors_fd = subv_74 + prest_70 + cotis_756 + mecenat_754 + aut_p
tc_hors_fd = ach_60 + ext_totale + ms_totale + fin_66 + amo_681

marge_brute = tp_hors_fd - ach_60
valeur_ajoutee = marge_brute - ext_totale
ebe = valeur_ajoutee - ms_totale
tebe = (ebe / tp_hors_fd * 100) if tp_hors_fd > 0 else 0.0

caf = ebe - fin_66
tcaf = (caf / tp_hors_fd * 100) if tp_hors_fd > 0 else 0.0
var_fd = rep_fd_789 - dot_fd_689
rn = caf - amo_681 + var_fd
autofin_net = caf - remb_16

dep_pub_pct = (subv_74 / tp_hors_fd * 100) if tp_hors_fd > 0 else 0.0
ress_propres_pct = 100.0 - dep_pub_pct
poids_ms_tp = (ms_totale / tp_hors_fd * 100) if tp_hors_fd > 0 else 0.0
cout_support_tot = ms_support + ext_support
poids_support_pct = (cout_support_tot / max(1.0, tc_hors_fd) * 100)
poids_missions_pct = 100.0 - poids_support_pct
levier_benevolat_pct = (vol_87 / tp_hors_fd * 100) if tp_hors_fd > 0 else 0.0

ch_dec_mensuelle = max(1.0, (tc_hors_fd - amo_681) / 12.0)
mois_treso = treso_51 / ch_dec_mensuelle
cap_desend = (dette_16 / caf) if caf > 0 else 99.9
couv_fd_treso = (treso_51 / max(1.0, fonds_ded_194) * 100) if fonds_ded_194 > 0 else 200.0

s_treso = min(35, max(0, int(round((mois_treso / max(0.5, normes["mois_treso"])) * 30))))
s_caf = min(25, max(0, int(round((tcaf / max(0.5, normes["taux_caf"])) * 22))))
s_dep = 20 if (dep_pub_pct <= normes["dep_pub_max"] and part_1er_fin <= normes["conc_fin_max"]) else (12 if dep_pub_pct <= normes["dep_pub_max"] + 12 else 5)
s_fp = 20 if (fonds_prop_10 > 0 and treso_51 >= fonds_ded_194 and cap_desend <= normes["desend_max"]) else (10 if fonds_prop_10 > 0 else 0)
health_score = min(100, s_treso + s_caf + s_dep + s_fp)

# ==========================================
# 6. INTERFACE PRINCIPALE + GUIDE PÉDAGOGIQUE
# ==========================================
st.title("Archétype Asso — Audit ANC 2018-06, Pilotage & Prospective")
badge_src = "Balances ANC 2018-06 certifiées" if source_reelle else "Modèle calibré Secteur Associatif"
st.caption(f"Association : **{asso_nom}** • {normes['type_court']} • {etp_salaries:.1f} ETP salariés • {etp_benevoles:.1f} ETP bénévoles • {badge_src}")

with st.expander("📖 Guide pédagogique : comprendre les comptes associatifs ANC 2018-06 (Spécial Président, Trésorier & CA)", expanded=False):
    st.markdown("Ce guide traduit les spécificités comptables des associations (**règlement ANC n° 2018-06**) en repères concrets pour le Conseil d'Administration :")
    gp1, gp2, gp3 = st.columns(3)
    with gp1:
        st.markdown("**1. Fonds dédiés (Comptes 194, 689, 789)**")
        st.markdown("""* **Pourquoi ils faussent la lecture du résultat :** lorsqu'une association encaisse une subvention affectée à un projet précis (ex. 50 000 €) et n'en dépense que 30 000 € avant le 31 décembre, les 20 000 € restants ne sont pas un bénéfice libre : ils sont mis en réserve obligatoire dans les **fonds dédiés (compte 689 au résultat / 194 au bilan)** pour être dépensés l'année suivante (**reprise 789**).
* **Réflexe clé :** toujours vérifier que la trésorerie disponible couvre au moins **100 % du stock de fonds dédiés (194)**.""")
    with gp2:
        st.markdown("**2. Réserve de trésorerie (en mois) & CAF**")
        st.markdown("""* **Pourquoi compter en mois :** les subventions publiques sont souvent versées en deux fois, avec plusieurs mois de décalage après le début de l'action. Une réserve de **3 à 5 mois de charges** évite de recourir à un découvert bancaire coûteux.
* **EBE & CAF associative :** une association n'a pas vocation à distribuer des dividendes, mais elle doit dégager une **CAF de 4,5 à 6,5 %** pour renouveler son matériel, absorber les imprévus et renforcer ses **fonds propres (10/11)**.""")
    with gp3:
        st.markdown("**3. Dépendance publique & Classe 8 (Bénévolat)**")
        st.markdown("""* **Dépendance & concentration :** dépasser **70–75 % de subventions** ou **50 % sur un seul financeur** expose l'association à un risque vital en cas de changement de politique publique.
* **Classe 8 (Contributions volontaires en nature) :** valorise en pied de compte de résultat le temps des bénévoles et les locaux prêtés gratuitement. Cela ne change pas la trésorerie, mais démontre aux financeurs l'effet de levier citoyen du projet.""")

onglets = st.tabs([
    "📊 1. Rétrospective (Trajectoire & Dépendance)",
    "⚖️ 2. Situation à date (TBI, SIG & Fonds Dédiés)",
    "🔪 3. Scalpel (Missions, Support & Bénévolat)",
    "🔮 4. Prospective (Projet Associatif à 5 ans)"
])

# ------------------------------------------
# ONGLET 1 : RÉTROSPECTIVE
# ------------------------------------------
with onglets[0]:
    st.subheader("D'où venons-nous ? Trajectoire financière, dépendance et fonds dédiés (2021-2025)")
    afficher_aide_pedagogique_onglet("Lecture de la trajectoire passée", [
        ("Effet Ciseau (TCMA Charges vs Produits)", "Rythme annuel moyen d'évolution hors fonds dédiés. Si les salaires et frais généraux progressent plus vite que les subventions et prestations, l'association consomme ses réserves."),
        ("Évolution de la dépendance publique", "Permet de voir si l'association réussit à diversifier ses ressources propres (prestations, cotisations, mécénat) ou si elle devient plus vulnérable aux arbitrages publics."),
        ("Effet Volume RH", "Part de la hausse de la masse salariale liée aux créations de postes et remplacements, par opposition aux hausses subies (ancienneté GVT et revalorisations conventionnelles).")
    ])

    if len(balances_parsees) >= 2:
        ans_r = [b["annee"] or f"Ex.{i+1}" for i, b in enumerate(balances_parsees)]
        p_r = [b["tp_hors_fd"]/1e6 for b in balances_parsees]
        c_r = [b["tc_hors_fd"]/1e6 for b in balances_parsees]
        ms_r = [b["ms_64"]/1e6 for b in balances_parsees]
        caf_r = [b["caf"]/1e6 for b in balances_parsees]
        dep_r = [(b["subv_74"]/max(1.0, b["tp_hors_fd"])*100) for b in balances_parsees]
        fd_r = [(b["fonds_ded"] if b["fonds_ded"] > 0 else fonds_ded_194)/1e6 for b in balances_parsees]
    else:
        ans_r = ["2021", "2022", "2023", "2024", "2025"]
        tp_m, tc_m, ms_m, caf_m, fd_m = tp_hors_fd/1e6, tc_hors_fd/1e6, ms_totale/1e6, caf/1e6, fonds_ded_194/1e6
        p_r = [round(tp_m*0.86, 2), round(tp_m*0.90, 2), round(tp_m*0.93, 2), round(tp_m*0.97, 2), round(tp_m, 2)]
        c_r = [round(tc_m*0.83, 2), round(tc_m*0.88, 2), round(tc_m*0.92, 2), round(tc_m*0.96, 2), round(tc_m, 2)]
        ms_r = [round(ms_m*0.83, 2), round(ms_m*0.88, 2), round(ms_m*0.92, 2), round(ms_m*0.96, 2), round(ms_m, 2)]
        caf_r = [round(p - c + (amo_681/1e6)*0.95, 2) for p, c in zip(p_r[:-1], c_r[:-1])] + [round(caf_m, 2)]
        dep_r = [round(min(95.0, dep_pub_pct + 4.0), 1), round(min(95.0, dep_pub_pct + 2.8), 1), round(min(95.0, dep_pub_pct + 1.5), 1), round(min(95.0, dep_pub_pct + 0.6), 1), round(dep_pub_pct, 1)]
        fd_r = [round(fd_m*0.75, 2), round(fd_m*0.85, 2), round(fd_m*1.10, 2), round(fd_m*0.92, 2), round(fd_m, 2)]

    n_int = max(1, len(ans_r) - 1)
    tcma_p = ((p_r[-1]/p_r[0])**(1/n_int) - 1)*100 if p_r[0] > 0 else 0.0
    tcma_c = ((c_r[-1]/c_r[0])**(1/n_int) - 1)*100 if c_r[0] > 0 else 0.0
    tcma_ms = ((ms_r[-1]/ms_r[0])**(1/n_int) - 1)*100 if ms_r[0] > 0 else 0.0
    ciseau = tcma_c - tcma_p

    kr1, kr2, kr3, kr4 = st.columns(4)
    kr1.metric("TCMA Produits", f"+{tcma_p:.2f} % / an", "Rythme des ressources")
    kr2.metric("TCMA Charges", f"+{tcma_c:.2f} % / an", f"Effet ciseau : {ciseau:+.2f} pts (Cible <= 0)", delta_color="inverse" if ciseau > 0 else "normal")
    kr3.metric("Dépendance Publique", f"{dep_r[-1]:.1f} %", f"Évolution 5 ans : {(dep_r[-1]-dep_r[0]):+.1f} pts (Max {normes['dep_pub_max']} %)", delta_color="inverse" if dep_r[-1] > normes["dep_pub_max"] else "normal")
    kr4.metric("Fonds Dédiés (194)", f"{fd_r[-1]:.2f} M€", f"CAF actuelle : {caf_r[-1]:.2f} M€ ({tcaf:.1f} %)")

    cg1, cg2, cg3 = st.columns(3)
    with cg1:
        st.markdown("**1. Effet de ciseau (Produits vs Charges M€)**")
        st.line_chart(pd.DataFrame({"Exercice": ans_r, "Produits (M€)": p_r, "Charges (M€)": c_r}).set_index("Exercice"))
    with cg2:
        st.markdown("**2. CAF & Stock de Fonds Dédiés (M€)**")
        st.line_chart(pd.DataFrame({"Exercice": ans_r, "CAF (M€)": caf_r, "Fonds Dédiés (M€)": fd_r}).set_index("Exercice"))
    with cg3:
        st.markdown("**3. Taux de dépendance publique (%)**")
        st.bar_chart(pd.DataFrame({"Exercice": ans_r, "Subventions / Produits (%)": dep_r}).set_index("Exercice"), color="#0d9488")

    st.markdown("**4. Décomposition de la dérive salariale (Compte 64)**")
    cd1, cd2, cd3 = st.columns(3)
    gvt_pt = cd1.slider("GVT Mécanique — Ancienneté (pts/an)", 0.5, 2.5, 1.2, 0.1)
    conv_pt = cd2.slider("Accords de branche & SMIC (pts/an)", 0.5, 3.5, 1.6, 0.1)
    vol_pt = round(tcma_ms - gvt_pt - conv_pt, 2)
    cd3.metric("Effet Volume (Levier interne)", f"{vol_pt:+.2f} pts / an", "Norme cible : <= +0,30 pt/an", delta_color="inverse" if vol_pt > 0.3 else "normal")

    df_retro_tbl = pd.DataFrame({
        "Indicateurs ANC 2018-06": [
            "Produits d'exploitation hors FD (M€)", "Charges d'exploitation hors FD (M€)",
            "Masse salariale Cpte 64 (M€)", "Capacité d'Autofinancement — CAF (M€)",
            "Fonds dédiés au bilan — Cpte 194 (M€)", f"Dépendance publique (Seuil {normes['dep_pub_max']} %)"
        ],
        **{ans_r[i]: [f"{p_r[i]:.2f}", f"{c_r[i]:.2f}", f"{ms_r[i]:.2f}", f"{caf_r[i]:.2f}", f"{fd_r[i]:.2f}", f"{dep_r[i]:.1f} %"] for i in range(len(ans_r))}
    })
    st.dataframe(df_retro_tbl, hide_index=True, use_container_width=True)

    atouts_list = [
        f"Socle d'activité de {tp_hors_fd/1e6:.2f} M€ avec une dépendance publique de {dep_r[-1]:.1f} % ({'maîtrisée' if dep_r[-1] <= normes['dep_pub_max'] else 'élevée'}).",
        f"CAF de {caf_r[-1]:.2f} M€ ({tcaf:.1f} % des produits vs cible sectorielle de {normes['taux_caf']} %).",
        f"Engagement bénévole valorisé à {vol_87:,.0f} € ({etp_benevoles:.1f} ETP) renforçant la légitimité auprès des partenaires.".replace(",", " ")
    ]
    vig_list = [
        f"Résorber l'effet ciseau ({ciseau:+.2f} pts/an) en indexant les tarifs de prestations et demandes de subventions.",
        f"Contenir l'effet volume RH ({vol_pt:+.2f} pts/an) sous +0,30 pt/an en conditionnant toute création de poste à un financement pérenne.",
        f"Suivre l'exécution des projets liés aux {fonds_ded_194:,.0f} € de fonds dédiés inscrits au bilan.".replace(",", " ")
    ]
    pdf_retro_bytes = build_pdf_asso_retro(asso_nom, normes, ans_r, p_r, c_r, ms_r, caf_r, dep_r, fd_r, tcma_p, tcma_c, tcma_ms, gvt_pt, conv_pt, vol_pt, atouts_list, vig_list)
    st.download_button("📑 Télécharger le Dossier Rétrospectif Associatif (PDF 2 pages)", data=pdf_retro_bytes, file_name=f"Dossier_Retro_Asso_{asso_nom.replace(' ','_')}.pdf", mime="application/pdf", type="primary")

    afficher_synthese_unique(
        "Rétrospective Associative (2021-2025)",
        f"Les charges progressent de +{tcma_c:.2f} %/an contre +{tcma_p:.2f} %/an pour les produits (effet ciseau de {ciseau:+.2f} pts/an). Le taux de dépendance aux subventions publiques s'établit à {dep_r[-1]:.1f} % (seuil de vigilance : {normes['dep_pub_max']} %) et le stock de fonds dédiés (compte 194) atteint {fonds_ded_194:,.0f} €.".replace(",", " "),
        f"Sur +{tcma_ms:.2f} %/an de dérive salariale, {(gvt_pt+conv_pt):.2f} pts sont structurels (ancienneté GVT et revalorisations conventionnelles), tandis que {vol_pt:+.2f} pts relèvent des choix internes d'effectifs.",
        f"Plafonner l'effet volume RH sous +0,30 pt/an au prochain Budget Prévisionnel, sécuriser le solde des actions liées aux {fonds_ded_194:,.0f} € de fonds dédiés et diversifier les ressources propres ({ress_propres_pct:.1f} % actuellement).".replace(",", " ")
    )

# ------------------------------------------
# ONGLET 2 : SITUATION À DATE
# ------------------------------------------
with onglets[1]:
    st.subheader("Où en est-on ? Situation à date, Cascade des SIG ANC 2018-06 & Trésorerie")
    afficher_aide_pedagogique_onglet("Lecture des SIG Associatifs & Couverture des Fonds Dédiés", [
        ("Distinction CAF vs Résultat Net", "La CAF mesure l'épargne dégagée par l'activité courante avant amortissements et mouvements de fonds dédiés. Le Résultat Net intègre ensuite les dotations aux amortissements (681) et le solde des fonds dédiés (789 - 689)."),
        ("Couverture des Fonds Dédiés par la trésorerie", "Les fonds dédiés au bilan (194) correspondent à de l'argent déjà encaissé mais fléché vers un projet précis. La trésorerie disponible (50/51) doit impérativement être supérieure à ce montant."),
        ("Réserve de trésorerie en mois", "Mesure combien de mois l'association peut payer ses salaires et fournisseurs sans aucun encaissement nouveau. Cible saine : 3 à 5 mois.")
    ])

    hs1, hs2 = st.columns([1, 2.5])
    with hs1:
        st.metric("HEALTH SCORE ANC 2018-06", f"{health_score} / 100", "Sain >= 70 | Vigilance 45-69")
    with hs2:
        if health_score >= 70:
            st.success(f"**Structure associative saine ({health_score}/100) :** Réserve de trésorerie de {mois_treso:.1f} mois (cible {normes['mois_treso']} mois), CAF à {tcaf:.1f} % (cible {normes['taux_caf']} %) et fonds dédiés couverts à {couv_fd_treso:.0f} %.")
        elif health_score >= 45:
            st.warning(f"**Zone de vigilance ({health_score}/100) :** Tensions détectées sur la réserve de trésorerie ({mois_treso:.1f} mois vs {normes['mois_treso']} mois attendus) ou la dépendance publique ({dep_pub_pct:.1f} %).")
        else:
            st.error(f"**Alerte structurelle ({health_score}/100) :** Risque de tension de liquidité ({mois_treso:.1f} mois) ou de sous-couverture des fonds dédiés. Plan d'ajustement immédiat requis.")

    if treso_51 < fonds_ded_194:
        st.error(f"**Alerte Fonds Dédiés :** La trésorerie disponible ({treso_51:,.0f} €) est inférieure au stock de fonds dédiés à réaliser ({fonds_ded_194:,.0f} €). L'association a déjà consommé une partie de sa trésorerie fléchée.".replace(",", " "))

    st.markdown("**1. Cascade des Soldes Intermédiaires de Gestion Associatifs (ANC 2018-06) vs Normes**")
    df_sig = pd.DataFrame([
        {"Solde Intermédiaire de Gestion": "Total Produits d'Exploitation hors Fonds Dédiés (70/74/75)", "Montant (€)": f"{tp_hors_fd:,.0f} €".replace(",", " "), "% Produits": "100,0 %", "Moyenne / Norme": "Base 100 %"},
        {"Solde Intermédiaire de Gestion": "- Achats & Fournitures directes (Cpte 60)", "Montant (€)": f"-{ach_60:,.0f} €".replace(",", " "), "% Produits": f"{ach_60/tp_hors_fd*100:.1f} %", "Moyenne / Norme": "6,0 à 12,0 %"},
        {"Solde Intermédiaire de Gestion": "- Services extérieurs, Locaux & Frais de mission (61-62)", "Montant (€)": f"-{ext_totale:,.0f} €".replace(",", " "), "% Produits": f"{ext_totale/tp_hors_fd*100:.1f} %", "Moyenne / Norme": "15,0 à 22,0 %"},
        {"Solde Intermédiaire de Gestion": "= VALEUR AJOUTÉE ASSOCIATIVE (VA)", "Montant (€)": f"{valeur_ajoutee:,.0f} €".replace(",", " "), "% Produits": f"{valeur_ajoutee/tp_hors_fd*100:.1f} %", "Moyenne / Norme": "65,0 à 78,0 %"},
        {"Solde Intermédiaire de Gestion": "- Masse Salariale chargée (Cpte 64)", "Montant (€)": f"-{ms_totale:,.0f} €".replace(",", " "), "% Produits": f"{poids_ms_tp:.1f} %", "Moyenne / Norme": f"Moyenne secteur : {normes['poids_ms']} %"},
        {"Solde Intermédiaire de Gestion": "= EXCÉDENT BRUT D'EXPLOITATION (EBE)", "Montant (€)": f"{ebe:,.0f} €".replace(",", " "), "% Produits": f"{tebe:.1f} %", "Moyenne / Norme": f"Moyenne secteur : {normes['taux_ebe']} %"},
        {"Solde Intermédiaire de Gestion": "- Frais financiers / Intérêts emprunts (66)", "Montant (€)": f"-{fin_66:,.0f} €".replace(",", " "), "% Produits": f"{fin_66/tp_hors_fd*100:.1f} %", "Moyenne / Norme": "Norme : < 2,0 %"},
        {"Solde Intermédiaire de Gestion": "= CAPACITÉ D'AUTOFINANCEMENT (CAF)", "Montant (€)": f"{caf:,.0f} €".replace(",", " "), "% Produits": f"{tcaf:.1f} %", "Moyenne / Norme": f"Cible sectorielle : >= {normes['taux_caf']} %"},
        {"Solde Intermédiaire de Gestion": "- Dotations aux amortissements (681)", "Montant (€)": f"-{amo_681:,.0f} €".replace(",", " "), "% Produits": f"{amo_681/tp_hors_fd*100:.1f} %", "Moyenne / Norme": "2,0 à 5,0 %"},
        {"Solde Intermédiaire de Gestion": "+/- Mouvements nets des Fonds Dédiés (Reprises 789 - Reports 689)", "Montant (€)": f"{var_fd:+,.0f} €".replace(",", " "), "% Produits": f"{var_fd/tp_hors_fd*100:+.1f} %", "Moyenne / Norme": "Neutralisation projets"},
        {"Solde Intermédiaire de Gestion": "= RÉSULTAT NET COMPTABLE (Affecté aux Fonds Propres)", "Montant (€)": f"{rn:+,.0f} €".replace(",", " "), "% Produits": f"{rn/tp_hors_fd*100:+.1f} %", "Moyenne / Norme": "Norme : >= 0,0 %"},
        {"Solde Intermédiaire de Gestion": "Pour mémoire : Contributions volontaires en nature (Classe 8)", "Montant (€)": f"{vol_87:,.0f} €".replace(",", " "), "% Produits": f"+{levier_benevolat_pct:.1f} %", "Moyenne / Norme": f"Repère secteur : ~{normes['taux_benevolat']} %"},
    ])
    st.dataframe(
        df_sig,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Solde Intermédiaire de Gestion": st.column_config.TextColumn("Solde Intermédiaire de Gestion", width="large"),
            "Montant (€)": st.column_config.TextColumn("Montant (€)", width="small"),
            "% Produits": st.column_config.TextColumn("% Produits", width="small"),
            "Moyenne / Norme": st.column_config.TextColumn("Moyenne / Norme", width="medium"),
        }
    )

    st.markdown("---")
    c_b1, c_b2 = st.columns(2)
    with c_b1:
        st.markdown("**2. Liquidité, Fonds Propres & Fonds Dédiés**")
        b1, b2 = st.columns(2)
        b1.metric("Réserve de Trésorerie", f"{mois_treso:.1f} mois", f"Cible : {normes['mois_treso']} mois ({treso_51/1e6:.2f} M€)", delta_color="normal" if mois_treso >= 2.5 else "inverse")
        b2.metric("Couverture Fonds Dédiés", f"{couv_fd_treso:.0f} %", "Tréso / FD 194 (Norme >= 100 %)", delta_color="normal" if couv_fd_treso >= 100 else "inverse")
        b3, b4 = st.columns(2)
        b3.metric("Fonds Propres (10/11)", f"{fonds_prop_10:,.0f} €".replace(",", " "), f"Autonomie : {(fonds_prop_10/max(1,tp_hors_fd)*100):.0f} % du budget")
        b4.metric("Désendettement", f"{cap_desend:.1f} ans", f"Dette : {dette_16:,.0f} € (Max {normes['desend_max']} a.)".replace(",", " "), delta_color="normal" if cap_desend <= normes["desend_max"] else "inverse")

    with c_b2:
        st.markdown("**3. Tableau Comparatif : Association vs Barème Sectoriel**")
        df_bench = pd.DataFrame([
            {"Indicateur": "Dépendance subventions publiques", "Association": f"{dep_pub_pct:.1f} %", "Norme / Cible": f"<= {normes['dep_pub_max']} %", "Statut": "OK" if dep_pub_pct <= normes["dep_pub_max"] else "Élevée"},
            {"Indicateur": "Concentration 1er financeur", "Association": f"{part_1er_fin:.1f} %", "Norme / Cible": f"<= {normes['conc_fin_max']} %", "Statut": "OK" if part_1er_fin <= normes["conc_fin_max"] else "Risque"},
            {"Indicateur": "Taux d'EBE / Taux de CAF", "Association": f"{tebe:.1f}% / {tcaf:.1f}%", "Norme / Cible": f"{normes['taux_ebe']}% / {normes['taux_caf']}%", "Statut": "OK" if tcaf >= normes["taux_caf"] else "Vigilance"},
            {"Indicateur": "Poids frais de support / charges", "Association": f"{poids_support_pct:.1f} %", "Norme / Cible": f"<= {normes['poids_support']} %", "Statut": "OK" if poids_support_pct <= normes["poids_support"] else "Élevé"},
            {"Indicateur": "Levier bénévolat (Classe 8)", "Association": f"+{levier_benevolat_pct:.1f} %", "Norme / Cible": f"~{normes['taux_benevolat']} %", "Statut": "Valorisé"},
        ])
        st.dataframe(df_bench, hide_index=True, use_container_width=True)

    synthese_tbi_lines = [
        f"• Santé financière : Health Score de {health_score}/100. La trésorerie couvre {mois_treso:.1f} mois de charges (cible : {normes['mois_treso']} mois) et {couv_fd_treso:.0f} % des fonds dédiés.",
        f"• Mix de financement : {dep_pub_pct:.1f} % de subventions publiques (seuil : {normes['dep_pub_max']} %) et {ress_propres_pct:.1f} % de ressources propres (prestations, cotisations, mécénat).",
        f"• Rentabilité & Bénévolat : EBE à {tebe:.1f} %, CAF à {tcaf:.1f} % et contributions volontaires en nature valorisées à {vol_87:,.0f} € (+{levier_benevolat_pct:.1f} % du budget).".replace(",", " ")
    ]
    pdf_tbi_bytes = build_pdf_asso_tbi(
        asso_nom, normes, etp_salaries, etp_benevoles, health_score,
        tp_hors_fd/1e6, ach_60/1e6, ext_totale/1e6, valeur_ajoutee/1e6, ms_totale/1e6,
        ebe/1e6, fin_66/1e6, amo_681/1e6, var_fd/1e6, rn/1e6, caf/1e6, remb_16/1e6, autofin_net/1e6,
        fonds_prop_10/1e6, fonds_ded_194/1e6, dette_16/1e6, cap_desend, treso_51/1e6, mois_treso,
        dep_pub_pct, float(part_1er_fin), poids_support_pct, levier_benevolat_pct, synthese_tbi_lines
    )
    st.download_button("📑 Télécharger le TBI Exécutif Associatif (PDF 1 page)", data=pdf_tbi_bytes, file_name=f"TBI_Asso_{asso_nom.replace(' ','_')}.pdf", mime="application/pdf", type="primary")

    lev_tbi = (
        f"Sanctuariser l'autofinancement net ({autofin_net:+,.0f} €) en réserve associative pour porter la trésorerie au-dessus de {normes['mois_treso']} mois."
        if autofin_net >= 0 else
        f"Corriger le déficit d'autofinancement net ({autofin_net:+,.0f} €) au prochain Budget Prévisionnel."
    ).replace(",", " ")
    afficher_synthese_unique(
        "Situation à date (ANC 2018-06)",
        f"L'association dégage un EBE de {ebe/1e6:.2f} M€ ({tebe:.1f} %) et une CAF de {caf/1e6:.2f} M€ ({tcaf:.1f} % vs {normes['taux_caf']} % attendus). Après dotations aux amortissements ({amo_681:,.0f} €) et mouvements nets de fonds dédiés ({var_fd:+,.0f} €), le résultat net s'établit à {rn:+,.0f} €.".replace(",", " "),
        f"La réserve de trésorerie représente {mois_treso:.1f} mois de charges décaissables (cible : {normes['mois_treso']} mois) et les contributions volontaires en nature ({vol_87:,.0f} €) apportent un effet de levier de +{levier_benevolat_pct:.1f} % sur le budget monétaire.".replace(",", " "),
        lev_tbi
    )

# ------------------------------------------
# ONGLET 3 : SCALPEL
# ------------------------------------------
with onglets[2]:
    st.subheader("Où faire ses choix ? Pilotage des missions sociales, frais support et ressources")
    afficher_aide_pedagogique_onglet("Arbitrer les moyens dans le Budget Prévisionnel", [
        ("Ratio Missions Sociales vs Frais de Support", "Les bailleurs publics et fondations vérifient que l'essentiel des fonds va au terrain (missions sociales >= 80–85 %) et que les frais de siège/administration restent contenus (<= 15–20 %)."),
        ("Risque de dépendance à un financeur unique", "Si un financeur pèse plus de 40 à 50 % du budget, le non-renouvellement d'une convention menace la survie de la structure (plan social impossible sans réserves)."),
        ("Fiscalisation / Secteur lucratif", "Si les prestations marchandes (compte 70) deviennent prépondérantes et concurrencent le secteur privé lucratif, vérifier les règles fiscales (franchise des impôts commerciaux ou sectorisation).")
    ])

    st.markdown("#### 1. Répartition Missions Sociales (Terrain) vs Frais de Support (Siège)")
    cs1, cs2 = st.columns([1.4, 1])
    with cs1:
        cout_missions_tot = ms_missions + ext_missions + ach_60
        df_ms_sup = pd.DataFrame([
            {"Affectation Analytique": "Missions sociales — Personnel terrain (64)", "Montant (€)": f"{ms_missions:,.0f} €".replace(",", " "), "% Charges": f"{ms_missions/max(1,tc_hors_fd)*100:.1f} %", "Repère Bailleurs": "60 - 75 %"},
            {"Affectation Analytique": "Missions sociales — Achats & actions projets (60-62)", "Montant (€)": f"{(ach_60+ext_missions):,.0f} €".replace(",", " "), "% Charges": f"{(ach_60+ext_missions)/max(1,tc_hors_fd)*100:.1f} %", "Repère Bailleurs": "12 - 20 %"},
            {"Affectation Analytique": "Frais de support — Direction, admin & RH siège (64)", "Montant (€)": f"{ms_support:,.0f} €".replace(",", " "), "% Charges": f"{ms_support/max(1,tc_hors_fd)*100:.1f} %", "Repère Bailleurs": "8 - 12 %"},
            {"Affectation Analytique": "Frais de support — Locaux siège, assur., honoraires", "Montant (€)": f"{ext_support:,.0f} €".replace(",", " "), "% Charges": f"{ext_support/max(1,tc_hors_fd)*100:.1f} %", "Repère Bailleurs": "4 - 7 %"},
            {"Affectation Analytique": "TOTAL FRAIS DE SUPPORT (Siège & Gouvernance)", "Montant (€)": f"{cout_support_tot:,.0f} €".replace(",", " "), "% Charges": f"{poids_support_pct:.1f} %", "Repère Bailleurs": f"Max {normes['poids_support']} %"},
        ])
        st.dataframe(df_ms_sup, hide_index=True, use_container_width=True)
        if poids_support_pct > normes["poids_support"]:
            st.warning(f"**Point de vigilance Bailleurs :** Vos frais de support représentent **{poids_support_pct:.1f} %** des charges contre un repère maximum de **{normes['poids_support']} %** dans votre secteur.")
        else:
            st.success(f"**Ratio conforme ({poids_support_pct:.1f} %) :** **{poids_missions_pct:.1f} %** des moyens sont directement affectés aux missions sociales de terrain.")
    with cs2:
        st.bar_chart(pd.DataFrame({
            "Pôle": ["Missions Terrain", "Frais Support"],
            "Montant (€)": [cout_missions_tot, cout_support_tot]
        }).set_index("Pôle"), color="#0d9488")

    st.divider()

    st.markdown("#### 2. Cartographie des Financements & Risque de Concentration")
    cf1, cf2, cf3, cf4 = st.columns(4)
    cf1.metric("Subventions Publiques (74)", f"{dep_pub_pct:.1f} %", f"Max secteur : {normes['dep_pub_max']} %", delta_color="inverse" if dep_pub_pct > normes["dep_pub_max"] else "normal")
    cf2.metric("1er Financeur Unique", f"{part_1er_fin:.1f} %", f"Max prudentiel : {normes['conc_fin_max']} %", delta_color="inverse" if part_1er_fin > normes["conc_fin_max"] else "normal")
    cf3.metric("Prestations & Ventes (70)", f"{(prest_70/tp_hors_fd*100):.1f} %", f"{prest_70:,.0f} €".replace(",", " "))
    cf4.metric("Mécénat & Cotisations", f"{((mecenat_754+cotis_756)/tp_hors_fd*100):.1f} %", f"{(mecenat_754+cotis_756):,.0f} €".replace(",", " "))

    if part_1er_fin > normes["conc_fin_max"]:
        st.error(f"**Risque de dépendance critique :** Un seul financeur représente **{part_1er_fin:.1f} %** des ressources (seuil d'alerte : {normes['conc_fin_max']} %). Prévoir un plan de diversification sur 3 ans.")

    st.divider()

    st.markdown("#### 3. Valorisation de l'Engagement Bénévole (Classe 8) & Productivité")
    cb1, cb2, cb3 = st.columns(3)
    cb1.metric("Contributions Volontaires (Cl. 8)", f"{vol_87:,.0f} €".replace(",", " "), f"+{levier_benevolat_pct:.1f} % du budget monétaire")
    cb2.metric("Force de travail totale", f"{(etp_salaries + etp_benevoles):.1f} ETP", f"{etp_salaries:.1f} salariés + {etp_benevoles:.1f} bénévoles")
    cb3.metric("Coût moyen par ETP salarié", f"{(ms_totale/max(0.5,etp_salaries)):,.0f} € / an".replace(",", " "), "Salaires + charges patronales")

    afficher_synthese_unique(
        "Indicateurs de pilotage budgétaire (Le Scalpel)",
        f"Les missions sociales mobilisent {poids_missions_pct:.1f} % des charges et les frais de support représentent {poids_support_pct:.1f} % (repère maximum : {normes['poids_support']} %). La dépendance aux financements publics atteint {dep_pub_pct:.1f} % avec un financeur principal à {part_1er_fin:.1f} %.",
        f"L'engagement bénévole ({etp_benevoles:.1f} ETP valorisés à {vol_87:,.0f} € en Classe 8) complète la force salariale ({etp_salaries:.1f} ETP) et constitue un argument clé dans les dossiers de subvention et de mécénat.".replace(",", " "),
        "Pour le prochain Budget Prévisionnel : flécher toute nouvelle embauche uniquement sur des recettes affectées ou pluriannuelles, maintenir les frais de siège sous le plafond sectoriel et développer les prestations ou le mécénat."
    )

# ------------------------------------------
# ONGLET 4 : PROSPECTIVE
# ------------------------------------------
with onglets[3]:
    st.subheader("Où allons-nous ? Poser financièrement le projet associatif à 5 ans (2026-2030)")
    afficher_aide_pedagogique_onglet("Tester la soutenabilité du projet associatif à 5 ans", [
        ("Effet de ciseau prospectif", "Dans le monde associatif, les subventions publiques sont souvent gelées (0 à +1 %/an) alors que les salaires augmentent mécaniquement (+2 à +3 %/an). Le simulateur permet de mesurer combien de prestations ou de mécénat il faut développer pour compenser."),
        ("Consolidation des Fonds Propres & Trésorerie", "Un projet associatif est soutenable si la trésorerie reste au-dessus de 3 mois de fonctionnement sur toute la période et si les fonds propres progressent régulièrement.")
    ])

    sp1, sp2, sp3, sp4 = st.columns(4)
    with sp1:
        st.markdown("**1. Charges & Salaires**")
        tx_ms_proj = st.slider("Évolution masse salariale (%/an)", 0.5, 5.0, 2.4, 0.1)
        tx_ext_proj = st.slider("Inflation achats & locaux (%/an)", 0.5, 5.0, 1.8, 0.1)
    with sp2:
        st.markdown("**2. Évolution des Ressources**")
        tx_subv_proj = st.slider("Évolution subventions 74 (%/an)", -2.0, 4.0, 1.0, 0.1)
        tx_prop_proj = st.slider("Croissance ressources propres (%/an)", 0.0, 10.0, 3.5, 0.5)
    with sp3:
        st.markdown("**3. Investissements & Développement**")
        ppi_annuel = st.slider("Investissements matériels (M€/an)", 0.0, 1.0, 0.06, 0.02)
        taux_subv_inv = st.slider("Subventions d'investissement (%)", 0, 80, 40, 5)
    with sp4:
        st.markdown("**4. Financement**")
        opex_ppi_pct = st.slider("Charges nouvelles induites (%)", -2.0, 5.0, 0.5, 0.5)
        duree_emprunt = st.selectbox("Durée emprunts", [7, 10, 15], index=1)

    ans_p = ["2026", "2027", "2028", "2029", "2030"]
    subv_inv_an = ppi_annuel * (taux_subv_inv / 100.0)
    emprunt_an = max(0.0, (ppi_annuel - subv_inv_an) * 0.50)
    autofin_inv_an = max(0.0, (ppi_annuel - subv_inv_an) * 0.50)

    p_list, c_list, caf_list, fp_list, dep_list, det_list, tr_list, mtr_list = [], [], [], [], [], [], [], []
    curr_subv = subv_74 / 1e6
    curr_prop = (tp_hors_fd - subv_74) / 1e6
    curr_ms = ms_totale / 1e6
    curr_autres_ch = (ach_60 + ext_totale) / 1e6
    curr_amo = amo_681 / 1e6
    curr_fin = fin_66 / 1e6
    curr_fp = fonds_prop_10 / 1e6
    curr_det = dette_16 / 1e6
    curr_remb = remb_16 / 1e6
    curr_treso = treso_51 / 1e6
    opex_cumul = 0.0

    for _ in ans_p:
        curr_subv *= (1 + tx_subv_proj / 100.0)
        curr_prop *= (1 + tx_prop_proj / 100.0)
        curr_p = curr_subv + curr_prop
        curr_ms *= (1 + tx_ms_proj / 100.0)
        curr_autres_ch *= (1 + tx_ext_proj / 100.0)
        opex_cumul += ppi_annuel * (opex_ppi_pct / 100.0)
        curr_amo += ppi_annuel / 10.0
        curr_fin += emprunt_an * 0.035

        tot_c = curr_ms + curr_autres_ch + opex_cumul + curr_amo + curr_fin
        curr_caf = (curr_p - tot_c) + curr_amo
        curr_rn = curr_caf - curr_amo
        curr_fp += curr_rn
        curr_det = max(0.0, curr_det - curr_remb + emprunt_an)
        curr_remb_tot = curr_remb + (emprunt_an / duree_emprunt)
        curr_treso = curr_treso + curr_caf - curr_remb_tot - autofin_inv_an
        ch_dec_m = max(0.005, (tot_c - curr_amo) / 12.0)
        curr_mois = curr_treso / ch_dec_m
        curr_dep = (curr_subv / max(0.01, curr_p)) * 100

        p_list.append(round(curr_p, 2))
        c_list.append(round(tot_c, 2))
        caf_list.append(round(curr_caf, 2))
        fp_list.append(round(curr_fp, 2))
        dep_list.append(round(curr_dep, 1))
        det_list.append(round(curr_det, 2))
        tr_list.append(round(curr_treso, 2))
        mtr_list.append(round(curr_mois, 1))

    df_proj = pd.DataFrame({
        "Année": ans_p,
        "Produits (M€)": p_list,
        "Charges (M€)": c_list,
        "CAF (M€)": caf_list,
        "Fonds Propres (M€)": fp_list,
        "Dépendance Pub. (%)": dep_list,
        "Dette (M€)": det_list,
        "Trésorerie (M€)": tr_list,
        "Trésorerie (mois)": mtr_list
    })

    kp1, kp2, kp3, kp4 = st.columns(4)
    tcaf_30 = (caf_list[-1]/p_list[-1]*100)
    kp1.metric("CAF Projetée 2030", f"+{caf_list[-1]:.2f} M€", f"Taux : {tcaf_30:.1f} % (Cible {normes['taux_caf']} %)")
    kp2.metric("Fonds Propres 2030", f"{fp_list[-1]:.2f} M€", f"Variation : {(fp_list[-1]-fp_list[0]):+.2f} M€")
    kp3.metric("Dépendance Pub. 2030", f"{dep_list[-1]:.1f} %", f"Seuil max : {normes['dep_pub_max']} %", delta_color="normal" if dep_list[-1] <= normes["dep_pub_max"] else "inverse")
    kp4.metric("Trésorerie 2030", f"{mtr_list[-1]:.1f} mois", f"{tr_list[-1]:.2f} M€ (Cible {normes['mois_treso']} mois)", delta_color="normal" if mtr_list[-1] >= 3.0 else "inverse")

    gp_c1, gp_c2 = st.columns(2)
    with gp_c1:
        st.markdown("**Trajectoire Produits vs Charges (M€)**")
        st.line_chart(df_proj.set_index("Année")[["Produits (M€)", "Charges (M€)"]])
    with gp_c2:
        st.markdown("**Trajectoire Fonds Propres & Trésorerie (M€)**")
        st.line_chart(df_proj.set_index("Année")[["Fonds Propres (M€)", "Trésorerie (M€)"]])

    st.dataframe(df_proj, hide_index=True, use_container_width=True)

    feuille_route = [
        f"1. Équilibre & CAF : Viser une CAF de {caf_list[-1]:.2f} M€ en 2030 ({tcaf_30:.1f} % vs {normes['taux_caf']} % attendus) pour consolider les fonds propres à {fp_list[-1]:.2f} M€.",
        f"2. Diversification des ressources : Compenser la faible indexation des subventions (+{tx_subv_proj:.1f} %/an) par une hausse de +{tx_prop_proj:.1f} %/an des prestations et du mécénat, ramenant la dépendance publique à {dep_list[-1]:.1f} %.",
        f"3. Maîtrise salariale : Plafonner la dérive de la masse salariale à +{tx_ms_proj:.1f} %/an en alignant les recrutements sur les financements acquis.",
        f"4. Investissements & Liquidité : Calibrer les investissements à {ppi_annuel:.2f} M€/an ({taux_subv_inv} % subventionnés) en maintenant une réserve de trésorerie de {mtr_list[-1]:.1f} mois."
    ]

    pdf_pro_bytes = build_pdf_asso_prospective(asso_nom, normes, ppi_annuel, df_proj, feuille_route)
    st.download_button("📑 Télécharger la Synthèse Prospective Associative (PDF 1 page)", data=pdf_pro_bytes, file_name=f"Synthese_Prospective_Asso_{asso_nom.replace(' ','_')}.pdf", mime="application/pdf", type="primary")

    afficher_synthese_unique(
        "Prospective : Soutenabilité du Projet Associatif à 5 ans",
        f"À l'horizon 2030, le taux de dépendance publique évolue vers {dep_list[-1]:.1f} % (seuil : {normes['dep_pub_max']} %), les fonds propres atteignent {fp_list[-1]:.2f} M€ et la trésorerie s'établit à {mtr_list[-1]:.1f} mois de fonctionnement (cible : {normes['mois_treso']} mois).",
        f"Le scénario repose sur un développement de +{tx_prop_proj:.1f} %/an des ressources propres pour absorber le différentiel entre l'évolution des subventions (+{tx_subv_proj:.1f} %/an) et la hausse des salaires (+{tx_ms_proj:.1f} %/an).",
        f"{'La trajectoire financière est soutenable et permet de sécuriser le projet associatif sur 5 ans.' if (mtr_list[-1] >= 3.0 and caf_list[-1] > 0) else 'Alerte de soutenabilité : renforcer le plan de développement des ressources propres ou ajuster le rythme des recrutements pour préserver au moins 3 mois de trésorerie.'}"
    )
