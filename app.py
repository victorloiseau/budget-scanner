import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from google.cloud import vision
import re, io
from PIL import Image
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

BUDGET_SERRE    = [200, 0, 0, 55, 20, 0, 25, 10, 15, 10, 30, 0, 0, 20, 50]
BUDGET_REALISTE = [270,80,20,55, 40, 0, 40, 20, 50, 20,100,20,15, 20,100]

AUTO_DETECT = {
    "Epicerie":                 ["maxi","iga","metro","provigo","costco","walmart",
                                 "supermarche","marche","grocery","alimentation","loblaws"],
    "Restaurants / cafeteria":  ["restaurant","cafeteria","mcdonald","mcdo","subway",
                                 "burger","pizza","sushi","bistro","brasserie","tim horton",
                                 "popeyes","wendy","five guys"],
    "Cafe / boissons":          ["starbucks","second cup","van houtte","cafe","coffee","tim horton"],
    "Pharmacie":                ["pharmaprix","jean coutu","uniprix","shoppers","pharmacie"],
    "Hygiene & soin":           ["sephora","beauty","cosmetique","parfum"],
    "Sorties & loisirs":        ["cinema","musee","theatre","spectacle","concert"],
    "Livres & materiel etudes": ["librairie","indigo","chapitres","parascolaire"],
    "Sport / gym":              ["gym","sport","fitness","running","athletic"],
}

SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/cloud-vision",
]

# ── Credentials ───────────────────────────────────────────────────────────────
@st.cache_resource
def get_credentials():
    return Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES,
    )

# ── Google Sheets ─────────────────────────────────────────────────────────────
@st.cache_resource
def get_gsheet():
    creds = get_credentials()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(st.secrets["SPREADSHEET_ID"])
    try:
        return sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        return None

def init_sheet(ws):
    headers = ["Poste de depense","Budget serre ($ CA)","Budget realiste ($ CA)",
               "REEL ($ CA)","Ecart vs Realiste","Notes / Date"]
    ws.update("A4:F4", [headers])
    rows = []
    for (label, _), s, r in zip(CATEGORIES, BUDGET_SERRE, BUDGET_REALISTE):
        row_num = 5 + len(rows)
        rows.append([label, s, r, 0, f"=D{row_num}-C{row_num}", ""])
    ws.update(f"A5:F{4+len(rows)}", rows)
    total_row = 5 + len(rows)
    ws.update(f"A{total_row}:F{total_row}", [[
        "TOTAL", sum(BUDGET_SERRE), sum(BUDGET_REALISTE),
        f"=SUM(D5:D{total_row-1})", f"=D{total_row}-C{total_row}", ""
    ]])
    ws.update("A1:A2", [
        ["Budget Montreal - Victor Loiseau"],
        ["Logement + chauffage + electricite inclus dans Darlington — non comptes"],
    ])

def read_current(ws, row: int) -> float:
    val = ws.cell(row, REEL_COL).value
    try:
        return float(str(val).replace(",", ".")) if val else 0.0
    except ValueError:
        return 0.0

def save_expense(ws, row: int, amount: float, note: str):
    current = read_current(ws, row)
    ws.update_cell(row, REEL_COL, round(current + amount, 2))
    if note:
        existing = ws.cell(row, NOTE_COL).value or ""
        ws.update_cell(row, NOTE_COL, (existing + " | " if existing else "") + note)

# ── Google Vision OCR ─────────────────────────────────────────────────────────
def ocr_receipt(image: Image.Image) -> str:
    creds  = get_credentials()
    client = vision.ImageAnnotatorClient(credentials=creds)
    buf    = io.BytesIO()
    image.convert("RGB").save(buf, format="JPEG", quality=90)
    response = client.text_detection(image=vision.Image(content=buf.getvalue()))
    if response.error.message:
        raise Exception(response.error.message)
    annotations = response.text_annotations
    return annotations[0].description if annotations else ""

def extract_total(text: str) -> float | None:
    lines = text.lower().split("\n")
    kws   = ["total","montant","a payer","due","balance","sous-total","grand total","amount"]
    for line in reversed(lines):
        if any(k in line for k in kws):
            amounts = re.findall(r"\d{1,4}[.,]\d{2}", line)
            if amounts:
                return float(amounts[-1].replace(",", "."))
    all_amounts = re.findall(r"\d{1,4}[.,]\d{2}", text)
    return max((float(a.replace(",", ".")) for a in all_amounts), default=None)

def detect_store(text: str) -> str:
    first_lines = " ".join(text.split("\n")[:3])
    return first_lines.strip() or "Non détecté"

def detect_category(text: str) -> str:
    tl = text.lower()
    for cat, keywords in AUTO_DETECT.items():
        if any(k in tl for k in keywords):
            return cat
    return "Imprevus"

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("🧾 Tickets de caisse → Budget Montréal")
st.caption("Prends une photo de ton ticket — la dépense s'ajoute automatiquement dans Google Sheets.")

# Connexion
try:
    ws = get_gsheet()
except Exception as e:
    st.error(f"Connexion Google Sheets impossible : {e}")
    st.stop()

if ws is None:
    st.warning(f"La feuille '{SHEET_NAME}' n'existe pas encore.")
    if st.button("Initialiser le budget dans Google Sheets", type="primary"):
        creds = get_credentials()
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(st.secrets["SPREADSHEET_ID"])
        ws = sh.add_worksheet(SHEET_NAME, rows=30, cols=6)
        init_sheet(ws)
        st.cache_resource.clear()
        st.success("Budget initialisé !")
        st.rerun()
    st.stop()

sheet_url = f"https://docs.google.com/spreadsheets/d/{st.secrets['SPREADSHEET_ID']}"
st.markdown(f"[📊 Voir le budget en direct]({sheet_url})")

st.divider()

# Upload
uploaded = st.file_uploader("📷 Photo du ticket de caisse", type=["jpg","jpeg","png","webp","heic"])

if not uploaded:
    st.markdown("""
**Comment ça marche :**
1. Prends ton ticket en photo
2. Dépose-le ici
3. Vérifie montant et catégorie
4. Clique Ajouter — c'est dans Google Sheets en 3 secondes
    """)
    st.stop()

img = Image.open(uploaded)
st.image(img, caption="Ticket reçu", use_column_width=True)

if st.button("🔍 Analyser", type="primary", use_container_width=True):
    with st.spinner("Lecture du ticket en cours..."):
        try:
            text     = ocr_receipt(img)
            total    = extract_total(text)
            store    = detect_store(text)
            category = detect_category(text)
            st.session_state.update({
                "text": text, "total": total,
                "store": store, "category": category,
                "analyzed": True,
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
        st.warning("Montant non détecté, entre-le manuellement.")
        amount = st.number_input("Montant ($ CA)", min_value=0.01, step=0.01, format="%.2f")

with col2:
    detected_cat = st.session_state["category"]
    default_idx  = CAT_NAMES.index(detected_cat) if detected_cat in CAT_NAMES else len(CAT_NAMES)-1
    category     = st.selectbox("Catégorie", CAT_NAMES, index=default_idx)
    note         = st.text_input("Note (optionnel)", placeholder=f"{date.today().isoformat()}")

with st.expander("Texte brut lu sur le ticket"):
    st.text(st.session_state.get("text", ""))

row    = CAT_ROW[category]
before = read_current(ws, row)
after  = round(before + amount, 2)

st.markdown(f"""
| Poste | Avant | Après |
|---|---|---|
| **{category}** | {before:.2f} $ | **{after:.2f} $** |
""")

if st.button("✅ Ajouter au budget", type="primary", use_container_width=True):
    try:
        save_expense(ws, row, amount, note)
        st.success(f"**{amount:.2f} $ CA** ajouté à **{category}** — total : **{after:.2f} $ CA**")
        st.markdown(f"[Voir dans Google Sheets]({sheet_url})")
        st.balloons()
        st.session_state["analyzed"] = False
    except Exception as e:
        st.error(f"Erreur : {e}")
