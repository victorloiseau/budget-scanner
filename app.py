import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pytesseract
import re, io
from PIL import Image, ImageEnhance, ImageFilter
from datetime import date

# ── Config ────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Tickets → Budget", page_icon="🧾", layout="centered")

SHEET_NAME = "Tracker mensuel"
REEL_COL   = 4
NOTE_COL   = 6

CATEGORIES = [
    ("Epicerie",                  5),
    ("Restaurants / cafeteria",   6),
    ("Cafe / boissons",           7),
    ("STM (abonnement etudiant)", 8),
    ("Telephone canadien",        9),
    ("Internet",                 10),
    ("Hygiene & soin",           11),
    ("Pharmacie",                12),
    ("Livres & materiel etudes", 13),
    ("Impression / fournitures", 14),
    ("Sorties & loisirs",        15),
    ("Streaming",                16),
    ("Sport / gym",              17),
    ("Frais etudiants HEC",      18),
    ("Imprevus",                 19),
]
CAT_NAMES = [n for n, _ in CATEGORIES]
CAT_ROW   = {n: r for n, r in CATEGORIES}

BUDGET_SERRE    = [200,  0,  0, 55, 20,  0, 25, 10, 15, 10,  30,  0,  0, 20,  50]
BUDGET_REALISTE = [270, 80, 20, 55, 40,  0, 40, 20, 50, 20, 100, 20, 15, 20, 100]

AUTO_DETECT = {
    "Epicerie":                 ["maxi","iga","metro","provigo","costco","walmart",
                                 "supermarche","marche","grocery","loblaws","alimentation"],
    "Restaurants / cafeteria":  ["restaurant","cafeteria","mcdonald","subway","burger",
                                 "pizza","sushi","bistro","brasserie","tim horton",
                                 "popeyes","wendy","five guys"],
    "Cafe / boissons":          ["starbucks","second cup","van houtte","cafe","coffee"],
    "Pharmacie":                ["pharmaprix","jean coutu","uniprix","shoppers","pharmacie"],
    "Hygiene & soin":           ["sephora","beauty","cosmetique","parfum"],
    "Sorties & loisirs":        ["cinema","musee","theatre","spectacle","concert"],
    "Livres & materiel etudes": ["librairie","indigo","chapitres","parascolaire"],
    "Sport / gym":              ["gym","sport","fitness","athletic"],
}

SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# ── Google Sheets ─────────────────────────────────────────────────────────────
@st.cache_resource
def get_gsheet():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["SPREADSHEET_ID"])
    try:
        return sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        return None

def init_sheet(ws):
    ws.update("A1:A2", [
        ["Budget Montreal - Victor Loiseau"],
        ["Logement + chauffage + electricite inclus dans Darlington"],
    ])
    ws.update("A4:F4", [["Poste","Budget serre ($ CA)","Budget realiste ($ CA)",
                          "REEL ($ CA)","Ecart vs Realiste","Notes / Date"]])
    rows = []
    for (label, _), s, r in zip(CATEGORIES, BUDGET_SERRE, BUDGET_REALISTE):
        n = 5 + len(rows)
        rows.append([label, s, r, 0, f"=D{n}-C{n}", ""])
    ws.update(f"A5:F{4+len(rows)}", rows)
    t = 5 + len(rows)
    ws.update(f"A{t}:F{t}", [["TOTAL", sum(BUDGET_SERRE), sum(BUDGET_REALISTE),
                               f"=SUM(D5:D{t-1})", f"=D{t}-C{t}", ""]])

def read_current(ws, row):
    val = ws.cell(row, REEL_COL).value
    try:
        return float(str(val).replace(",", ".")) if val else 0.0
    except ValueError:
        return 0.0

def save_expense(ws, row, amount, note):
    cur = read_current(ws, row)
    ws.update_cell(row, REEL_COL, round(cur + amount, 2))
    if note:
        ex = ws.cell(row, NOTE_COL).value or ""
        ws.update_cell(row, NOTE_COL, (ex + " | " if ex else "") + note)

# ── OCR ───────────────────────────────────────────────────────────────────────
def preprocess(img: Image.Image) -> Image.Image:
    img = img.convert("L")
    img = ImageEnhance.Contrast(img).enhance(2.5)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = img.filter(ImageFilter.SHARPEN)
    return img

def ocr_receipt(img: Image.Image) -> str:
    processed = preprocess(img)
    config    = "--oem 3 --psm 6 -l fra+eng"
    return pytesseract.image_to_string(processed, config=config)

def extract_total(text: str):
    lines = text.lower().split("\n")
    for line in reversed(lines):
        if any(k in line for k in ["total","montant","a payer","due","balance","amount"]):
            amounts = re.findall(r"\d{1,4}[.,]\d{2}", line)
            if amounts:
                return float(amounts[-1].replace(",", "."))
    all_a = re.findall(r"\d{1,4}[.,]\d{2}", text)
    return max((float(a.replace(",", ".")) for a in all_a), default=None)

def detect_store(text: str) -> str:
    return " ".join(text.split("\n")[:2]).strip() or "Non détecté"

def detect_category(text: str) -> str:
    tl = text.lower()
    for cat, kws in AUTO_DETECT.items():
        if any(k in tl for k in kws):
            return cat
    return "Imprevus"

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("🧾 Tickets de caisse → Budget Montréal")
st.caption("Photo du ticket → dépense ajoutée automatiquement dans Google Sheets.")

try:
    ws = get_gsheet()
except Exception as e:
    st.error(f"Connexion Google Sheets impossible : {e}")
    st.stop()

if ws is None:
    st.warning(f"La feuille '{SHEET_NAME}' n'existe pas encore.")
    if st.button("Initialiser le budget dans Google Sheets", type="primary"):
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
        sh = gspread.authorize(creds).open_by_key(st.secrets["SPREADSHEET_ID"])
        ws = sh.add_worksheet(SHEET_NAME, rows=30, cols=6)
        init_sheet(ws)
        st.cache_resource.clear()
        st.success("Budget initialisé !")
        st.rerun()
    st.stop()

sheet_url = f"https://docs.google.com/spreadsheets/d/{st.secrets['SPREADSHEET_ID']}"
st.markdown(f"[📊 Voir le budget en direct]({sheet_url})")
st.divider()

uploaded = st.file_uploader("📷 Photo du ticket", type=["jpg","jpeg","png","webp","heic"])

if not uploaded:
    st.markdown("""
**Mode d'emploi :**
1. Prends le ticket en photo (bien éclairé, bien à plat)
2. Dépose-le ici
3. Vérifie montant et catégorie
4. Clique Ajouter → Google Sheets se met à jour
    """)
    st.stop()

img = Image.open(uploaded)
st.image(img, caption="Ticket reçu", use_column_width=True)

if st.button("🔍 Analyser", type="primary", use_container_width=True):
    with st.spinner("Lecture du ticket..."):
        try:
            text     = ocr_receipt(img)
            total    = extract_total(text)
            store    = detect_store(text)
            category = detect_category(text)
            st.session_state.update({
                "text": text, "total": total,
                "store": store, "category": category, "analyzed": True,
            })
        except Exception as e:
            st.error(f"Erreur OCR : {e}")
            st.session_state["analyzed"] = False

if not st.session_state.get("analyzed"):
    st.stop()

st.divider()
st.subheader("Vérification avant ajout")

col1, col2 = st.columns(2)
with col1:
    st.markdown(f"**Magasin :** {st.session_state['store']}")
    t = st.session_state["total"]
    if t:
        amount = st.number_input("Montant ($ CA)", value=float(t), min_value=0.01, step=0.01, format="%.2f")
    else:
        st.warning("Montant non détecté — entre-le manuellement.")
        amount = st.number_input("Montant ($ CA)", min_value=0.01, step=0.01, format="%.2f")

with col2:
    cat      = st.session_state["category"]
    def_idx  = CAT_NAMES.index(cat) if cat in CAT_NAMES else len(CAT_NAMES)-1
    category = st.selectbox("Catégorie", CAT_NAMES, index=def_idx)
    note     = st.text_input("Note (optionnel)", placeholder=date.today().isoformat())

with st.expander("Texte brut lu sur le ticket"):
    st.text(st.session_state.get("text",""))

row    = CAT_ROW[category]
before = read_current(ws, row)
after  = round(before + amount, 2)

st.markdown(f"| **{category}** | Avant : {before:.2f} $ | Après : **{after:.2f} $** |")

if st.button("✅ Ajouter au budget", type="primary", use_container_width=True):
    try:
        save_expense(ws, row, amount, note)
        st.success(f"**{amount:.2f} $ CA** ajouté à **{category}** — total : **{after:.2f} $ CA**")
        st.markdown(f"[Voir dans Google Sheets]({sheet_url})")
        st.balloons()
        st.session_state["analyzed"] = False
    except Exception as e:
        st.error(f"Erreur : {e}")
