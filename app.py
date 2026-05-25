import base64
import html
import io
import json
import os
import re
from pathlib import Path

import streamlit as st
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Configuration ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Contrats de Travail",
    page_icon="📄",
    layout="wide",
)

st.markdown("""
<style>
[data-testid="stSidebar"] { min-width: 400px; max-width: 400px; }
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
.import-box {
    background: #f0f7ff;
    border: 1px solid #b3d4f5;
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 8px;
}
</style>
""", unsafe_allow_html=True)

TEMPLATE_DIR = Path(__file__).parent / "templates"
if not TEMPLATE_DIR.exists():
    TEMPLATE_DIR = Path(__file__).parent

# ── Session-state defaults ─────────────────────────────────────────────────────

_DEFAULTS: dict = {
    "fi_nom_societe":      "AU JONGLEUR NOTRE DAME",
    "fi_siege_social":     "8 RUE DU CLOITRE DE NOTRE DAME, 75004 PARIS",
    "fi_siret":            "672 014 727 00032",
    "fi_urssaf":           "117 000001565566658",
    "fi_representant":     "Madame BORJI Mohamed",
    "fi_titre_rep":        "président",
    "fi_civilite":         "Monsieur",
    "fi_nom_salarie":      "DORE SALEEM Mohamed Ismail",
    "fi_adresse":          "7 RESIDENCE DU PLATEAU, 94500 CHAMPIGNY SUR MARNE",
    "fi_date_naissance":   "02/09/1978",
    "fi_lieu_naissance":   "INDE",
    "fi_nationalite":      "INDIENNE",
    "fi_has_sejour":       True,
    "fi_num_titre":        "F943114362",
    "fi_date_val_sejour":  "20/04/2019",
    "fi_num_ss":           "1 78 09 99 223 055 46",
    "fi_date_debut":       "14 avril 2026",
    "fi_date_fin":         "31 décembre 2026",
    "fi_motif_cdd":        "accroissement temporaire d'activité",
    "fi_poste":            "vendeur",
    "fi_niveau":           "2",
    "fi_has_essai":        False,
    "fi_duree_essai":      "13 jours",
    "fi_heures_semaine":   "35",
    "fi_has_medical":      True,
    "fi_heures_medicales": "17.50",
    "fi_lieu_travail":     "8 rue du cloitre Notre Dame, 75004 PARIS",
    "fi_has_anciennete":   True,
    "fi_date_anciennete":  "01/06/2018",
    "fi_taux_anciennete":  "6",
    "fi_ancienne_societe": "ARCADES SOUVENIRS",
    "fi_caisse_retraite":  "Réunica AG2R",
    "fi_mutuelle":         "MMA",
    "fi_date_signature":   "14 avril 2026",
    "fi_date_dpae":        "13 avril 2026",
    "fi_ville_sig":        "PARIS",
}

for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ── Core helpers ──────────────────────────────────────────────────────────────

def load_template(ctype: str) -> str:
    return (TEMPLATE_DIR / f"contrat_{ctype.lower()}.txt").read_text("utf-8")


def render(template: str, data: dict) -> str:
    for k, v in data.items():
        template = template.replace(f"{{{{{k}}}}}", v)
    return template


# ── Claude auto-fill ──────────────────────────────────────────────────────────

_EXTRACT_PROMPT = """\
Tu es un expert en lecture de documents officiels français \
(contrat de travail, carte d'identité, titre de séjour, passeport, fiche de paie, email RH, etc.).
Extrais TOUTES les informations disponibles dans le document.
Réponds UNIQUEMENT avec un objet JSON valide, sans aucun autre texte ni markdown.
Si une information est absente, laisse la valeur à chaîne vide "".

{
  "civilite":             "Monsieur ou Madame selon le genre",
  "nom_salarie":          "NOM ET PRENOM EN MAJUSCULES",
  "adresse":              "adresse complète du salarié",
  "date_naissance":       "JJ/MM/AAAA",
  "lieu_naissance":       "ville et/ou pays de naissance",
  "nationalite":          "NATIONALITE EN MAJUSCULES",
  "num_titre_sejour":     "numéro titre de séjour",
  "date_validite_sejour": "JJ/MM/AAAA date validité titre séjour",
  "num_ss":               "numéro sécurité sociale avec espaces",
  "nom_societe":          "raison sociale de l'entreprise",
  "siege_social":         "adresse complète du siège social",
  "siret":                "numéro SIRET",
  "urssaf":               "numéro URSSAF",
  "representant":         "civilité + nom du représentant légal",
  "titre_rep":            "titre du représentant (gérant, président…)",
  "poste":                "intitulé du poste / emploi",
  "niveau":               "niveau ou coefficient",
  "date_debut":           "date de début du contrat en toutes lettres ou JJ/MM/AAAA",
  "date_fin":             "date de fin du contrat (CDD) en toutes lettres ou JJ/MM/AAAA",
  "motif_cdd":            "motif du CDD",
  "heures_semaine":       "nombre d'heures hebdomadaires (ex: 35 ou 17.50)",
  "lieu_travail":         "adresse du lieu de travail",
  "caisse_retraite":      "caisse de retraite complémentaire",
  "mutuelle":             "mutuelle santé",
  "date_signature":       "date de signature du contrat",
  "date_dpae":            "date de la DPAE",
  "ville_sig":            "ville de signature"
}"""

_FIELD_MAP = {
    "civilite":             "fi_civilite",
    "nom_salarie":          "fi_nom_salarie",
    "adresse":              "fi_adresse",
    "date_naissance":       "fi_date_naissance",
    "lieu_naissance":       "fi_lieu_naissance",
    "nationalite":          "fi_nationalite",
    "num_titre_sejour":     "fi_num_titre",
    "date_validite_sejour": "fi_date_val_sejour",
    "num_ss":               "fi_num_ss",
    "nom_societe":          "fi_nom_societe",
    "siege_social":         "fi_siege_social",
    "siret":                "fi_siret",
    "urssaf":               "fi_urssaf",
    "representant":         "fi_representant",
    "titre_rep":            "fi_titre_rep",
    "poste":                "fi_poste",
    "niveau":               "fi_niveau",
    "date_debut":           "fi_date_debut",
    "date_fin":             "fi_date_fin",
    "motif_cdd":            "fi_motif_cdd",
    "heures_semaine":       "fi_heures_semaine",
    "lieu_travail":         "fi_lieu_travail",
    "caisse_retraite":      "fi_caisse_retraite",
    "mutuelle":             "fi_mutuelle",
    "date_signature":       "fi_date_signature",
    "date_dpae":            "fi_date_dpae",
    "ville_sig":            "fi_ville_sig",
}


def _apply_extracted(info: dict) -> int:
    """Pousse les infos extraites dans session_state. Retourne le nb de champs remplis."""
    filled = 0
    for src, dst in _FIELD_MAP.items():
        val = info.get(src, "")
        if val:
            st.session_state[dst] = val
            wk = f"w_{dst}"
            if wk in st.session_state:
                st.session_state[wk] = val
            filled += 1
    if info.get("num_titre_sejour"):
        st.session_state["fi_has_sejour"] = True
        if "w_fi_has_sejour" in st.session_state:
            st.session_state["w_fi_has_sejour"] = True
    return filled


def extract_from_text(text: str, api_key: str) -> dict:
    from groq import Groq
    client = Groq(api_key=api_key)
    prompt = (
        f"{_EXTRACT_PROMPT}\n\nVoici le texte à analyser :\n\n{text}"
    )
    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=512,
    )
    raw = resp.choices[0].message.content.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


def extract_from_images(files: list, api_key: str) -> dict:
    from groq import Groq

    client = Groq(api_key=api_key)
    merged: dict = {}

    for f in files:
        raw = f.read()
        ext = f.name.rsplit(".", 1)[-1].lower()
        mime = f"image/{'jpeg' if ext in ('jpg', 'jpeg') else ext}"
        b64 = base64.standard_b64encode(raw).decode()

        resp = client.chat.completions.create(
            model="meta-llama/llama-4-scout-17b-16e-instruct",
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                    {"type": "text",      "text": _EXTRACT_PROMPT},
                ],
            }],
            max_tokens=512,
        )

        raw_text = resp.choices[0].message.content.strip()
        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
        raw_text = re.sub(r"\s*```$", "", raw_text)
        info = json.loads(raw_text)
        # Les valeurs non vides écrasent les précédentes
        for k, v in info.items():
            if v:
                merged[k] = v

    return merged


# ── PDF generation ────────────────────────────────────────────────────────────

def make_pdf(text: str) -> bytes:
    safe = text.encode("cp1252", errors="replace").decode("cp1252")

    pdf = FPDF(format="A4")
    pdf.set_margins(25, 30, 25)
    pdf.set_auto_page_break(auto=True, margin=25)
    pdf.add_page()

    w = pdf.w - pdf.l_margin - pdf.r_margin
    _lv = re.compile(r'^([A-Z][A-Z°/\s]{1,27})\s*:\s+(.+)$')
    LBL_W = 68

    for raw in safe.split("\n"):
        line = raw.rstrip()

        if not line.strip():
            pdf.ln(4)
            continue

        s = line.strip()
        pdf.set_x(pdf.l_margin)

        if s == "CONTRAT DE TRAVAIL":
            pdf.set_font("Helvetica", "B", 14)
            y0 = pdf.get_y()
            pdf.rect(pdf.l_margin, y0, w, 27)
            pdf.set_y(y0 + 4)
            pdf.multi_cell(w, 9, s, align="C")
            continue

        if s in ("A DUREE INDETERMINEE", "A DUREE DETERMINEE (CDD)"):
            pdf.set_font("Helvetica", "B", 12)
            pdf.multi_cell(w, 9, s, align="C")
            pdf.ln(8)
            continue

        if re.match(r"^ARTICLE \d+\s*:", s):
            pdf.set_font("Helvetica", "BU", 10)
            pdf.multi_cell(w, 7, s)
            pdf.ln(2)
            continue

        if s.startswith("- "):
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(w, 5.5, "      *  " + s[2:])
            continue

        m = _lv.match(s)
        if m:
            label = m.group(1).strip() + " :"
            value = m.group(2)
            y = pdf.get_y()
            pdf.set_font("Helvetica", "", 10)
            pdf.set_xy(pdf.l_margin, y)
            pdf.cell(LBL_W, 6, label)
            pdf.set_xy(pdf.l_margin + LBL_W, y)
            pdf.multi_cell(w - LBL_W, 6, value)
            continue

        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(w, 5.5, s, align="J")

    return bytes(pdf.output())


# ── DOCX generation ───────────────────────────────────────────────────────────

def _cell_borders(cell, sz: int = 12, color: str = "000000", fill: str = "") -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  fill)
        tcPr.append(shd)


def _no_table_borders(tbl) -> None:
    tbl_elem = tbl._tbl
    tblPr = tbl_elem.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl_elem.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        tblBorders.append(el)
    tblPr.append(tblBorders)


def make_docx(text: str) -> bytes:  # noqa: C901
    doc = Document()

    # ── Mise en page A4 ───────────────────────────────────────────────────────
    for section in doc.sections:
        section.page_width    = Cm(21)
        section.page_height   = Cm(29.7)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.0)

    # Style par défaut — Times New Roman 11 justifié
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after  = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.JUSTIFY

    F = "Times New Roman"   # police corps
    S = Pt(11)              # taille corps

    def _fmt(p,
             before: float = 0, after: float = 0,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             line_spacing=None):
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after  = Pt(after)
        p.paragraph_format.alignment    = align
        if line_spacing:
            p.paragraph_format.line_spacing = line_spacing

    def _run(p, txt, bold=False, italic=False, underline=False,
             font=None, size=None):
        r = p.add_run(txt)
        r.bold      = bold
        r.italic    = italic
        r.underline = underline
        r.font.name = font or F
        r.font.size = size or S
        return r

    def _tab_stop(p, pos_twips: int, align: str = "left"):
        pPr = p._p.get_or_add_pPr()
        tabs = OxmlElement("w:tabs")
        tab  = OxmlElement("w:tab")
        tab.set(qn("w:val"), align)
        tab.set(qn("w:pos"), str(pos_twips))
        tabs.append(tab)
        pPr.append(tabs)

    # Regex
    _lv      = re.compile(r"^([A-Z][A-Z°/\s]{1,27})\s*:\s+(.+)$")
    _agissant = re.compile(
        r"^(Agissant par l.interm.diaire de son repr.sentant l.gal, )(.+?)(,\s*.+\.)$"
    )

    lines = text.split("\n")
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        # ── Ligne vide — hauteur minimale pour ne pas gonfler la page ───────────
        if not line:
            p = doc.add_paragraph()
            _fmt(p, before=0, after=0)
            p.paragraph_format.line_spacing = Pt(4)
            i += 1
            continue

        # ── Titre encadré ─────────────────────────────────────────────────────
        if line == "CONTRAT DE TRAVAIL":
            subtitle = ""
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if nxt in ("A DUREE INDETERMINEE", "A DUREE DETERMINEE (CDD)"):
                    subtitle = nxt
                    i += 1

            tbl  = doc.add_table(rows=1, cols=1)
            _no_table_borders(tbl)
            cell = tbl.cell(0, 0)
            _cell_borders(cell, sz=18, color="2E74B5", fill="BDD7EE")

            p1 = cell.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _fmt(p1, before=8, after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
            r1 = p1.add_run("CONTRAT DE TRAVAIL")
            r1.bold = True; r1.font.name = "Arial"; r1.font.size = Pt(16)

            if subtitle:
                p2 = cell.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _fmt(p2, before=2, after=8, align=WD_ALIGN_PARAGRAPH.CENTER)
                r2 = p2.add_run(subtitle)
                r2.bold = True; r2.font.name = "Arial"; r2.font.size = Pt(13)

            # Espace après le tableau
            pa = doc.add_paragraph()
            _fmt(pa, before=4, after=0)
            pa.paragraph_format.line_spacing = Pt(4)
            i += 1
            continue

        # ── ARTICLE n : ───────────────────────────────────────────────────────
        if re.match(r"^ARTICLE \d+\s*:", line):
            p = doc.add_paragraph()
            _fmt(p, before=14, after=6, align=WD_ALIGN_PARAGRAPH.LEFT)
            _run(p, line, bold=True, underline=True)
            i += 1
            continue

        # ── Puces ─────────────────────────────────────────────────────────────
        if line.startswith("- "):
            p = doc.add_paragraph()
            _fmt(p, before=2, after=2)
            p.paragraph_format.left_indent       = Cm(0.8)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            _run(p, "•  " + line[2:])
            i += 1
            continue

        # ── LABEL : valeur ────────────────────────────────────────────────────
        m = _lv.match(line)
        if m:
            label = m.group(1).strip() + " :"
            value = m.group(2).strip()
            p = doc.add_paragraph()
            _fmt(p, before=0, after=2, align=WD_ALIGN_PARAGRAPH.LEFT)
            _tab_stop(p, 3600)            # ≈ 6.35 cm
            r_lbl = _run(p, label, underline=True)
            r_lbl.font.size = Pt(10)
            p.add_run("\t")
            r_val = _run(p, value, bold=True)
            r_val.font.size = Pt(10)
            i += 1
            continue

        # ── "Agissant par..." → nom en gras ───────────────────────────────────
        m2 = _agissant.match(line)
        if m2:
            p = doc.add_paragraph()
            _fmt(p, before=4, after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
            _run(p, m2.group(1))
            _run(p, m2.group(2), bold=True)
            _run(p, m2.group(3))
            i += 1
            continue

        # ── "Entre les soussignés" ────────────────────────────────────────────
        if line.startswith("Entre les soussign"):
            p = doc.add_paragraph()
            _fmt(p, before=4, after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
            _run(p, line)
            i += 1
            continue

        # ── "Et" ──────────────────────────────────────────────────────────────
        if line == "Et":
            p = doc.add_paragraph()
            _fmt(p, before=2, after=2, align=WD_ALIGN_PARAGRAPH.LEFT)
            _run(p, line)
            i += 1
            continue

        # ── Signature "LE SALARIE … L'EMPLOYEUR" ─────────────────────────────
        if "LE SALARIE" in line and "L'EMPLOYEUR" in line:
            p = doc.add_paragraph()
            _fmt(p, before=40, after=0, align=WD_ALIGN_PARAGRAPH.LEFT)
            _tab_stop(p, 9072, align="right")   # tabulation droite à ~16 cm
            _run(p, "LE SALARIE", bold=True)
            p.add_run("\t")
            _run(p, "L'EMPLOYEUR", bold=True)
            i += 1
            continue

        # ── "Fait le..." / "A VILLE" → gauche ────────────────────────────────
        if re.match(r"^(Fait le |A [A-Z])", line):
            p = doc.add_paragraph()
            _fmt(p, before=4, after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
            _run(p, line)
            i += 1
            continue

        # ── "Signature précédée..." → italique ────────────────────────────────
        if line.startswith("Signature pr"):
            p = doc.add_paragraph()
            _fmt(p, before=6, after=0, align=WD_ALIGN_PARAGRAPH.LEFT)
            _run(p, line, italic=True)
            i += 1
            continue

        # ── Texte courant justifié ────────────────────────────────────────────
        p = doc.add_paragraph()
        _fmt(p, before=0, after=4)
        _run(p, line)
        i += 1

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Paramètres du contrat")

    # ── Import automatique ─────────────────────────────────────────────────
    with st.expander("🤖 Import automatique", expanded=True):

        st.markdown(
            "🔑 Clé **gratuite** sur [console.groq.com](https://console.groq.com) "
            "→ **API Keys** → **Create API key**"
        )
        _api_key = st.text_input(
            "Clé API Groq",
            type="password",
            value=os.environ.get("GROQ_API_KEY", ""),
            placeholder="gsk_...",
            help="100% gratuite, sans carte bancaire.",
        )

        _tab_img, _tab_txt = st.tabs(["📸 Par image", "📋 Par texte"])

        # ── Onglet image ───────────────────────────────────────────────
        with _tab_img:
            st.caption("Glissez les documents du salarié (CNI, titre de séjour, passeport…)")
            _docs = st.file_uploader(
                "Documents",
                type=["jpg", "jpeg", "png", "webp"],
                accept_multiple_files=True,
                label_visibility="collapsed",
            )
            if _docs:
                st.info(f"📎 {len(_docs)} document(s) chargé(s)")
            if st.button("🔍 Extraire les images", type="primary",
                         use_container_width=True, disabled=not bool(_docs)):
                if not _api_key:
                    st.error("⚠️ Entrez votre clé API Groq.")
                else:
                    with st.spinner("Analyse des images en cours…"):
                        try:
                            _info = extract_from_images(_docs, _api_key)
                            _n = _apply_extracted(_info)
                            st.success(f"✅ {_n} champ(s) rempli(s) — vérifiez et corrigez si besoin.")
                        except json.JSONDecodeError:
                            st.error("Réponse illisible. Essayez avec des images plus nettes.")
                        except Exception as _e:
                            st.error(f"Erreur : {_e}")

        # ── Onglet texte ───────────────────────────────────────────────
        with _tab_txt:
            st.caption("Collez n'importe quel texte contenant les infos du salarié.")
            _pasted = st.text_area(
                "Texte à analyser",
                placeholder="Copiez-collez ici un email, un document scanné (OCR), une fiche RH…",
                height=160,
                label_visibility="collapsed",
            )
            if st.button("📋 Extraire le texte", type="primary",
                         use_container_width=True, disabled=not bool(_pasted.strip())):
                if not _api_key:
                    st.error("⚠️ Entrez votre clé API Groq.")
                else:
                    with st.spinner("Analyse du texte en cours…"):
                        try:
                            _info = extract_from_text(_pasted, _api_key)
                            _n = _apply_extracted(_info)
                            st.success(f"✅ {_n} champ(s) rempli(s) — vérifiez et corrigez si besoin.")
                        except json.JSONDecodeError:
                            st.error("Réponse illisible. Reformulez ou ajoutez plus de contexte.")
                        except Exception as _e:
                            st.error(f"Erreur : {_e}")

        # ── Réinitialiser ──────────────────────────────────────────────
        def _reset_all():
            for _k, _v in _DEFAULTS.items():
                st.session_state[_k] = _v
        st.button("🔄 Réinitialiser tous les champs", use_container_width=True,
                  on_click=_reset_all)

    ctype = st.radio("Type de contrat", ["CDI", "CDD"], horizontal=True)

    # ── Helper : vide une section (on_click → s'exécute avant le rerun) ──
    def _rst(btn_key: str, data_keys: list) -> None:
        def _cb():
            for k in data_keys:
                st.session_state[k] = False if isinstance(_DEFAULTS.get(k), bool) else ""
        st.button("🗑️ Vider cette section", key=btn_key,
                  on_click=_cb, use_container_width=True)

    # ── 🏢 Entreprise ─────────────────────────────────────────────────────
    with st.expander("🏢 Entreprise", expanded=True):
        _rst("rst_ent", ["fi_nom_societe","fi_siege_social","fi_siret",
                         "fi_urssaf","fi_representant","fi_titre_rep"])
        nom_societe  = st.text_input("Raison sociale",    key="fi_nom_societe")
        siege_social = st.text_input("Siège social",      key="fi_siege_social")
        siret        = st.text_input("SIRET",             key="fi_siret")
        urssaf_val   = st.text_input("URSSAF",            key="fi_urssaf")
        representant = st.text_input("Représentant légal",key="fi_representant")
        titre_rep    = st.text_input("Titre",             key="fi_titre_rep")

    # ── 👤 Salarié ────────────────────────────────────────────────────────
    with st.expander("👤 Salarié", expanded=True):
        _rst("rst_sal", ["fi_civilite","fi_nom_salarie","fi_adresse",
                         "fi_date_naissance","fi_lieu_naissance","fi_nationalite",
                         "fi_has_sejour","fi_num_titre","fi_date_val_sejour","fi_num_ss"])
        civilite = st.selectbox("Civilité", ["Monsieur", "Madame"], key="fi_civilite")
        nom_salarie     = st.text_input("Nom complet",       key="fi_nom_salarie")
        adresse         = st.text_input("Adresse",           key="fi_adresse")
        date_naissance  = st.text_input("Date de naissance", key="fi_date_naissance")
        lieu_naissance  = st.text_input("Lieu de naissance", key="fi_lieu_naissance")
        nationalite     = st.text_input("Nationalité",       key="fi_nationalite")
        has_sejour      = st.checkbox("Titre de séjour",     key="fi_has_sejour")
        if has_sejour:
            num_titre       = st.text_input("N° titre de séjour", key="fi_num_titre")
            date_val_sejour = st.text_input("Date de validité",   key="fi_date_val_sejour")
        else:
            num_titre = "N/A"; date_val_sejour = "N/A"
        num_ss = st.text_input("N° Sécurité Sociale", key="fi_num_ss")

    # ── 📅 Contrat ────────────────────────────────────────────────────────
    with st.expander("📅 Contrat", expanded=True):
        _rst("rst_con", ["fi_date_debut","fi_date_fin","fi_motif_cdd",
                         "fi_poste","fi_niveau","fi_has_essai","fi_duree_essai"])
        date_debut = st.text_input("Date de début", key="fi_date_debut")
        date_fin = motif_cdd = ""
        if ctype == "CDD":
            date_fin  = st.text_input("Date de fin",  key="fi_date_fin")
            motif_cdd = st.text_input("Motif du CDD", key="fi_motif_cdd")
        poste  = st.text_input("Poste",  key="fi_poste")
        niveau = st.text_input("Niveau", key="fi_niveau")
        st.markdown("**Période d'essai**")
        has_essai   = st.checkbox("Avec période d'essai", key="fi_has_essai")
        duree_essai = st.text_input("Durée de la période d'essai", key="fi_duree_essai") if has_essai else ""

    # ── ⏱️ Temps de travail ───────────────────────────────────────────────
    JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    with st.expander("⏱️ Temps de travail", expanded=True):
        _rst("rst_tmp", ["fi_heures_semaine","fi_has_medical",
                         "fi_heures_medicales","fi_lieu_travail"])
        est_temps_partiel = st.radio(
            "Régime horaire", ["Temps plein", "Temps partiel"],
            horizontal=True, key="fi_regime"
        ) == "Temps partiel"
        heures_semaine   = st.text_input("Heures / semaine", key="fi_heures_semaine")
        has_medical      = st.checkbox("Aménagement médical du temps de travail", key="fi_has_medical")
        heures_medicales = st.text_input("H/semaine après avis médecin", key="fi_heures_medicales") if has_medical else ""

        planning: dict[str, list[tuple[str, str]]] = {}
        if est_temps_partiel:
            st.markdown("**Planning hebdomadaire**")
            for jour in JOURS:
                actif = st.checkbox(jour, key=f"j_{jour}")
                if actif:
                    c1, c2 = st.columns(2)
                    d1 = c1.text_input("De",  "09:00", key=f"d1_{jour}")
                    f1 = c2.text_input("À",   "13:00", key=f"f1_{jour}")
                    slots: list[tuple[str, str]] = [(d1, f1)]
                    if st.checkbox("+ Coupure", key=f"cp_{jour}"):
                        c3, c4 = st.columns(2)
                        d2 = c3.text_input("De ",  "15:00", key=f"d2_{jour}")
                        f2 = c4.text_input("À ",   "18:00", key=f"f2_{jour}")
                        slots.append((d2, f2))
                    planning[jour] = slots

        lieu_travail = st.text_input("Lieu de travail", key="fi_lieu_travail")

    # ── 📈 Ancienneté ─────────────────────────────────────────────────────
    with st.expander("📈 Ancienneté"):
        _rst("rst_anc", ["fi_has_anciennete","fi_date_anciennete",
                         "fi_taux_anciennete","fi_ancienne_societe"])
        has_anciennete  = st.checkbox("Reprise d'ancienneté", key="fi_has_anciennete")
        date_anciennete = taux_anciennete = ancienne_societe = ""
        if has_anciennete:
            date_anciennete  = st.text_input("Date d'ancienneté", key="fi_date_anciennete")
            taux_anciennete  = st.text_input("Taux prime (%)",    key="fi_taux_anciennete")
            ancienne_societe = st.text_input("Ancienne société",  key="fi_ancienne_societe")

    # ── 🏥 Organismes sociaux ─────────────────────────────────────────────
    with st.expander("🏥 Organismes sociaux"):
        _rst("rst_org", ["fi_caisse_retraite","fi_mutuelle"])
        caisse_retraite = st.text_input("Caisse retraite", key="fi_caisse_retraite")
        mutuelle        = st.text_input("Mutuelle",        key="fi_mutuelle")

    # ── ✍️ Signature ──────────────────────────────────────────────────────
    with st.expander("✍️ Signature"):
        _rst("rst_sig", ["fi_date_signature","fi_date_dpae","fi_ville_sig"])
        date_signature = st.text_input("Date de signature", key="fi_date_signature")
        date_dpae      = st.text_input("Date DPAE",         key="fi_date_dpae")
        ville_sig      = st.text_input("Ville",             key="fi_ville_sig")

# ── Build optional text blocks ────────────────────────────────────────────────

if has_essai and duree_essai:
    article_3_block = (
        f"Le présent engagement est conclu sous réserve d'une période d'essai de "
        f"{duree_essai} au cours de laquelle il pourra être mis fin au contrat, "
        f"à tout moment, par l'une ou l'autre des parties, en respectant le délai "
        f"de prévenance prévu aux articles L. 1221-25 et L. 1221-26 du Code du travail.\n\n"
        f"La période d'essai s'entend d'une période de travail effectif. Toute suspension "
        f"de l'exécution du contrat, quel qu'en soit le motif, entrainera une prolongation "
        f"de la période d'essai d'une durée équivalente à celle de la suspension.\n\n"
        f"Toute rupture de période d'essai, quel qu'en soit l'auteur, sera notifiée par "
        f"écrit. Celle-ci sera remise en main propre contre décharge ou adressée en "
        f"recommandé avec AR."
    )
else:
    article_3_block = "Le présent engagement est conclu sans aucune période d'essai."

if est_temps_partiel:
    horaire_intro = "selon la répartition hebdomadaire suivante"
    if planning:
        lignes = ["\n\nLa répartition hebdomadaire des horaires de travail est la suivante :\n"]
        for jour, slots in planning.items():
            if len(slots) == 1:
                d, f = slots[0]
                lignes.append(f"- {jour} : de {d} à {f}")
            else:
                parties = " et de ".join(f"{d} à {f}" for d, f in slots)
                lignes.append(f"- {jour} : de {parties}")
        lignes.append(
            "\n\nCette répartition pourra être modifiée sous réserve d'un délai "
            "de prévenance de 7 jours ouvrés, conformément aux dispositions légales."
        )
        planning_block = "\n".join(lignes)
    else:
        planning_block = ""
else:
    horaire_intro  = "selon les horaires en vigueur de la société"
    planning_block = ""

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
    "ARTICLE_3_CONTENU":    article_3_block,
    "HEURES_SEMAINE":       heures_semaine,
    "HORAIRE_INTRO":        horaire_intro,
    "AVIS_MEDICAL":         medical_block,
    "PLANNING_HORAIRE":     planning_block,
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
