import html
import io
import re
from pathlib import Path

import streamlit as st
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── Configuration ────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Contrats de Travail",
    page_icon="📄",
    layout="wide",
)

st.markdown("""
<style>
[data-testid="stSidebar"] { min-width: 380px; max-width: 380px; }
.contract-view {
    background: white;
    border: 1px solid #d0d0d0;
    border-radius: 8px;
    padding: 48px 60px;
    font-family: "Times New Roman", Times, serif;
    font-size: 12.5px;
    line-height: 1.9;
    color: #111;
    white-space: pre-wrap;
    max-height: 88vh;
    overflow-y: auto;
    box-shadow: 0 2px 10px rgba(0,0,0,.08);
}
.stDownloadButton > button { width: 100%; }
</style>
""", unsafe_allow_html=True)

TEMPLATE_DIR = Path(__file__).parent / "templates"

# ── Core helpers ─────────────────────────────────────────────────────────────

def load_template(ctype: str) -> str:
    return (TEMPLATE_DIR / f"contrat_{ctype.lower()}.txt").read_text("utf-8")


def render(template: str, data: dict) -> str:
    for k, v in data.items():
        template = template.replace(f"{{{{{k}}}}}", v)
    return template


# ── PDF generation ────────────────────────────────────────────────────────────

def make_pdf(text: str) -> bytes:
    # encode to cp1252 (fpdf built-in fonts) — replaces any unsupported char with ?
    safe = text.encode("cp1252", errors="replace").decode("cp1252")

    pdf = FPDF(format="A4")
    pdf.set_margins(22, 28, 22)
    pdf.set_auto_page_break(auto=True, margin=22)
    pdf.add_page()

    # Pre-compute usable width once — avoids "not enough horizontal space" when
    # fpdf2 2.8+ calculates w=0 relative to a drifted X position after ln().
    w = pdf.w - pdf.l_margin - pdf.r_margin

    for raw in safe.split("\n"):
        line = raw.rstrip()

        if not line:
            pdf.ln(3)
            continue

        stripped = line.strip()
        pdf.set_x(pdf.l_margin)  # always reset to left margin before each cell

        # Big centred titles
        if stripped in ("CONTRAT DE TRAVAIL", "A DUREE INDETERMINEE", "A DUREE DETERMINEE (CDD)"):
            pdf.set_font("Helvetica", "B", 13)
            pdf.multi_cell(w, 8, stripped, align="C")
            pdf.ln(2)
            continue

        # Article headers  (ARTICLE X : ...)
        if re.match(r"^ARTICLE \d+\s*:", stripped):
            pdf.set_font("Helvetica", "BU", 10)
            pdf.multi_cell(w, 6, stripped)
            pdf.ln(1)
            continue

        # Bullet items
        if stripped.startswith("- "):
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(w, 5.5, "   * " + stripped[2:])
            continue

        # Normal paragraph
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(w, 5.5, stripped)

    return bytes(pdf.output())


# ── DOCX generation ───────────────────────────────────────────────────────────

def make_docx(text: str) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.left_margin  = Inches(1.1)
        section.right_margin = Inches(1.1)
        section.top_margin   = Inches(1.3)
        section.bottom_margin = Inches(1.3)

    for raw in text.split("\n"):
        line = raw.strip()

        if not line:
            doc.add_paragraph()
            continue

        if line in ("CONTRAT DE TRAVAIL", "A DUREE INDETERMINEE", "A DUREE DETERMINEE (CDD)"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line)
            run.bold = True
            run.font.size = Pt(14)
            continue

        if re.match(r"^ARTICLE \d+\s*:", line):
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.bold = True
            run.underline = True
            run.font.size = Pt(11)
            continue

        if line.startswith("- "):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            run = p.add_run("•  " + line[2:])
            run.font.size = Pt(10)
            continue

        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(10)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Sidebar — form ────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Paramètres du contrat")
    ctype = st.radio("Type de contrat", ["CDI", "CDD"], horizontal=True)

    # Entreprise
    with st.expander("🏢 Entreprise", expanded=True):
        nom_societe  = st.text_input("Raison sociale",    "AU JONGLEUR NOTRE DAME")
        siege_social = st.text_input("Siège social",       "8 RUE DU CLOITRE DE NOTRE DAME, 75004 PARIS")
        siret        = st.text_input("SIRET",              "672 014 727 00032")
        urssaf_val   = st.text_input("URSSAF",             "117 000001565566658")
        representant = st.text_input("Représentant légal", "Madame BORJI Mohamed")
        titre_rep    = st.text_input("Titre",              "président")

    # Salarié
    with st.expander("👤 Salarié", expanded=True):
        civilite        = st.selectbox("Civilité", ["Monsieur", "Madame"])
        nom_salarie     = st.text_input("Nom complet",        "DORE SALEEM Mohamed Ismail")
        adresse         = st.text_input("Adresse",            "7 RESIDENCE DU PLATEAU, 94500 CHAMPIGNY SUR MARNE")
        date_naissance  = st.text_input("Date de naissance",  "02/09/1978")
        lieu_naissance  = st.text_input("Lieu de naissance",  "INDE")
        nationalite     = st.text_input("Nationalité",        "INDIENNE")
        has_sejour      = st.checkbox("Titre de séjour", value=True)
        num_titre       = st.text_input("N° titre de séjour", "F943114362")  if has_sejour else "N/A"
        date_val_sejour = st.text_input("Date de validité",   "20/04/2019")  if has_sejour else "N/A"
        num_ss          = st.text_input("N° Sécurité Sociale", "1 78 09 99 223 055 46")

    # Contrat
    with st.expander("📅 Contrat", expanded=True):
        date_debut = st.text_input("Date de début", "14 avril 2026")
        date_fin   = motif_cdd = ""
        if ctype == "CDD":
            date_fin  = st.text_input("Date de fin",  "31 décembre 2026")
            motif_cdd = st.text_input("Motif du CDD", "accroissement temporaire d'activité")
        poste  = st.text_input("Poste",  "vendeur")
        niveau = st.text_input("Niveau", "2")

    # Temps de travail
    with st.expander("⏱️ Temps de travail"):
        heures_semaine   = st.text_input("Heures / semaine",             "35")
        has_medical      = st.checkbox("Aménagement médical du temps de travail", value=True)
        heures_medicales = st.text_input("H/semaine après avis médecin", "17.50") if has_medical else ""
        lieu_travail     = st.text_input("Lieu de travail", "8 rue du cloitre Notre Dame, 75004 PARIS")

    # Ancienneté
    with st.expander("📈 Ancienneté"):
        has_anciennete   = st.checkbox("Reprise d'ancienneté", value=True)
        date_anciennete  = taux_anciennete = ancienne_societe = ""
        if has_anciennete:
            date_anciennete  = st.text_input("Date d'ancienneté", "01/06/2018")
            taux_anciennete  = st.text_input("Taux prime (%)",     "6")
            ancienne_societe = st.text_input("Ancienne société",   "ARCADES SOUVENIRS")

    # Organismes sociaux
    with st.expander("🏥 Organismes sociaux"):
        caisse_retraite = st.text_input("Caisse retraite", "Réunica AG2R")
        mutuelle        = st.text_input("Mutuelle",        "MMA")

    # Signature
    with st.expander("✍️ Signature"):
        date_signature = st.text_input("Date de signature", "14 avril 2026")
        date_dpae      = st.text_input("Date DPAE",          "13 avril 2026")
        ville_sig      = st.text_input("Ville",              "PARIS")

# ── Build optional text blocks ────────────────────────────────────────────────

anciennete_block = ""
if has_anciennete and date_anciennete:
    anciennete_block = (
        f"\n\nIl conservera son ancienneté du {date_anciennete} ainsi que sa prime "
        f"d'ancienneté de {taux_anciennete}% qu'il avait chez {ancienne_societe}."
    )

medical_block = ""
if has_medical and heures_medicales:
    medical_block = (
        f" Suite à l'avis de la médecine du travail, le salarié effectuera "
        f"{heures_medicales} heures par semaine. Ce temps de travail sera modifié "
        f"en fonction de l'avis du médecin du travail."
    )

# ── Data dict ─────────────────────────────────────────────────────────────────

data: dict[str, str] = {
    "NOM_SOCIETE":          nom_societe,
    "SIEGE_SOCIAL":         siege_social,
    "SIRET":                siret,
    "URSSAF":               urssaf_val,
    "REPRESENTANT_LEGAL":   representant,
    "TITRE_REPRESENTANT":   titre_rep,
    "CIVILITE":             civilite,
    "NOM_SALARIE":          nom_salarie,
    "ADRESSE_SALARIE":      adresse,
    "DATE_NAISSANCE":       date_naissance,
    "LIEU_NAISSANCE":       lieu_naissance,
    "NATIONALITE":          nationalite,
    "NUM_TITRE_SEJOUR":     num_titre,
    "DATE_VALIDITE_SEJOUR": date_val_sejour,
    "NUM_SS":               num_ss,
    "DATE_DEBUT":           date_debut,
    "DATE_FIN":             date_fin,
    "MOTIF_CDD":            motif_cdd,
    "POSTE":                poste,
    "NIVEAU":               niveau,
    "HEURES_SEMAINE":       heures_semaine,
    "AVIS_MEDICAL":         medical_block,
    "ANCIENNETE":           anciennete_block,
    "LIEU_TRAVAIL":         lieu_travail,
    "CAISSE_RETRAITE":      caisse_retraite,
    "MUTUELLE":             mutuelle,
    "DATE_SIGNATURE":       date_signature,
    "DATE_DPAE":            date_dpae,
    "VILLE_SIGNATURE":      ville_sig,
}

# ── Render template ───────────────────────────────────────────────────────────

template = load_template(ctype)
rendered = render(template, data)

# ── Header + download buttons ─────────────────────────────────────────────────

col_title, col_pdf, col_docx, col_txt = st.columns([5, 1, 1, 1])

with col_title:
    st.title(f"📄 Contrat de Travail — {ctype}")

safe_name = re.sub(r"\s+", "_", nom_salarie)

with col_pdf:
    st.download_button(
        "⬇️ PDF",
        data=make_pdf(rendered),
        file_name=f"contrat_{ctype}_{safe_name}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with col_docx:
    st.download_button(
        "⬇️ DOCX",
        data=make_docx(rendered),
        file_name=f"contrat_{ctype}_{safe_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

with col_txt:
    st.download_button(
        "⬇️ TXT",
        data=rendered.encode("utf-8"),
        file_name=f"contrat_{ctype}_{safe_name}.txt",
        mime="text/plain",
        use_container_width=True,
    )

# ── Contract preview ──────────────────────────────────────────────────────────

st.markdown(
    f'<div class="contract-view">{html.escape(rendered)}</div>',
    unsafe_allow_html=True,
)
