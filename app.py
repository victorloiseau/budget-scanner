import streamlit as st
import anthropic
import gspread
from google.oauth2.service_account import Credentials
import json, base64, re, io
from PIL import Image
from datetime import date

# ── Config ────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Tickets → Budget", page_icon="🧾", layout="centered")

SHEET_NAME = "Tracker mensuel"
REEL_COL   = 4   # colonne D (1-based)
NOTE_COL   = 6   # colonne F

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
BUDGET_REALISTE = [270,80,20, 55, 40, 0, 40, 20, 50, 20,100,20,15, 20,100]

# ── Connexions ────────────────────────────────────────────────────────────────
def get_gsheet() -> gspread.Worksheet:
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"],
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["SPREADSHEET_ID"])
    try:
        return sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        return None

def init_sheet(ws: gspread.Worksheet):
    """Initialise la structure du tracker si la feuille est vide."""
    headers = ["Poste de depense", "Budget serre ($ CA)", "Budget realiste ($ CA)",
               "REEL ($ CA)", "Ecart vs Realiste", "Notes / Date"]
    ws.update("A4:F4", [headers])
    rows = []
    for (label, _), s, r in zip(CATEGORIES, BUDGET_SERRE, BUDGET_REALISTE):
        rows.append([label, s, r, 0, f"=D{5+len(rows)}-C{5+len(rows)}", ""])
    ws.update(f"A5:F{4+len(rows)}", rows)
    total_row = 5 + len(rows)
    ws.update(f"A{total_row}:F{total_row}",
              [["TOTAL",
                sum(BUDGET_SERRE),
                sum(BUDGET_REALISTE),
                f"=SUM(D5:D{total_row-1})",
                f"=D{total_row}-C{total_row}",
                ""]])
    ws.update("A1", [["Budget Montreal - Victor Loiseau"]])
    ws.update("A2", [["Logement + chauffage + electricite inclus dans Darlington - non comptes"]])

def get_api_key() -> str | None:
    if "ANTHROPIC_API_KEY" in st.secrets:
        return st.secrets["ANTHROPIC_API_KEY"]
    return None

# ── Analyse du ticket ─────────────────────────────────────────────────────────
def analyze_receipt(image: Image.Image, api_key: str) -> dict:
    buf = io.BytesIO()
    image.convert("RGB").save(buf, format="JPEG", quality=90)
    b64 = base64.standard_b64encode(buf.getvalue()).decode()

    client = anthropic.Anthropic(api_key=api_key)
    prompt = f"""Analyse ce ticket de caisse. Reponds UNIQUEMENT avec un JSON valide, sans markdown.

Format :
{{
  "store": "nom du magasin",
  "total": 12.34,
  "date": "{date.today().isoformat()}",
  "category": "categorie exacte parmi la liste",
  "items_summary": "resume bref en 1 ligne"
}}

Categories disponibles (texte exact) :
Epicerie | Restaurants / cafeteria | Cafe / boissons | Pharmacie | Hygiene & soin |
Livres & materiel etudes | Sorties & loisirs | Sport / gym | Streaming |
Impression / fournitures | Imprevus

Si total illisible : null. Si date illisible : "{date.today().isoformat()}"."""

    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=300,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text", "text": prompt},
            ],
        }],
    )
    raw = re.sub(r"^```(?:json)?\n?|\n?```$", "", msg.content[0].text.strip())
    return json.loads(raw)

# ── Lecture / écriture Google Sheets ─────────────────────────────────────────
def read_current(ws: gspread.Worksheet, row: int) -> float:
    val = ws.cell(row, REEL_COL).value
    try:
        return float(str(val).replace(",", ".")) if val else 0.0
    except ValueError:
        return 0.0

def save_expense(ws: gspread.Worksheet, row: int, amount: float, note: str):
    current = read_current(ws, row)
    ws.update_cell(row, REEL_COL, round(current + amount, 2))
    if note:
        existing = ws.cell(row, NOTE_COL).value or ""
        ws.update_cell(row, NOTE_COL, (existing + " | " if existing else "") + note)

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("🧾 Tickets de caisse → Budget Montréal")
st.caption("Prends une photo de ton ticket, la dépense s'ajoute automatiquement dans Google Sheets.")

api_key = get_api_key()
if not api_key:
    st.error("Clé API Anthropic manquante dans les secrets Streamlit.")
    st.stop()

# Connexion Google Sheets
try:
    ws = get_gsheet()
except Exception as e:
    st.error(f"Impossible de se connecter à Google Sheets : {e}")
    st.stop()

if ws is None:
    st.warning(f"La feuille '{SHEET_NAME}' n'existe pas encore.")
    if st.button("Initialiser le budget dans Google Sheets"):
        sh = gspread.authorize(
            Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=["https://spreadsheets.google.com/feeds",
                        "https://www.googleapis.com/auth/drive"],
            )
        ).open_by_key(st.secrets["SPREADSHEET_ID"])
        ws = sh.add_worksheet(SHEET_NAME, rows=30, cols=6)
        init_sheet(ws)
        st.success("Budget initialisé dans Google Sheets !")
        st.rerun()
    st.stop()

# Lien vers le Google Sheet
sheet_url = f"https://docs.google.com/spreadsheets/d/{st.secrets['SPREADSHEET_ID']}"
st.markdown(f"[Voir le budget en direct sur Google Sheets]({sheet_url})")

st.divider()

# Upload ticket
uploaded = st.file_uploader("📷 Photo du ticket de caisse", type=["jpg","jpeg","png","webp","heic"])

if not uploaded:
    st.markdown("""
**Astuce** : ticket bien à plat, bonne lumière.
Après chaque scan, ta dépense s'ajoute directement dans Google Sheets — accessible depuis ton téléphone en temps réel.
    """)
    st.stop()

img = Image.open(uploaded)
st.image(img, caption="Ticket reçu", use_column_width=True)

if st.button("🔍 Analyser", type="primary", use_container_width=True):
    with st.spinner("Claude lit ton ticket..."):
        try:
            result = analyze_receipt(img, api_key)
            st.session_state.update({"result": result, "analyzed": True})
        except json.JSONDecodeError:
            st.error("Ticket illisible — essaie avec une photo plus nette.")
            st.session_state["analyzed"] = False
        except anthropic.AuthenticationError:
            st.error("Clé API invalide.")
            st.session_state["analyzed"] = False
        except Exception as e:
            st.error(f"Erreur : {e}")
            st.session_state["analyzed"] = False

if not st.session_state.get("analyzed"):
    st.stop()

result = st.session_state["result"]

st.divider()
st.subheader("Vérification avant ajout")

col1, col2 = st.columns(2)

with col1:
    st.markdown(f"**Magasin :** {result.get('store','Non détecté')}")
    st.markdown(f"**Date :** {result.get('date', date.today().isoformat())}")
    if result.get("items_summary"):
        st.caption(result["items_summary"])

    detected = result.get("total")
    if detected:
        amount = st.number_input("Montant ($ CA)", value=float(detected),
                                  min_value=0.01, step=0.01, format="%.2f")
    else:
        st.warning("Montant non détecté.")
        amount = st.number_input("Montant ($ CA)", min_value=0.01, step=0.01, format="%.2f")

with col2:
    detected_cat = result.get("category", "Imprevus")
    default_idx = CAT_NAMES.index(detected_cat) if detected_cat in CAT_NAMES else len(CAT_NAMES)-1
    category = st.selectbox("Catégorie", CAT_NAMES, index=default_idx)
    note = st.text_input("Note (optionnel)",
                          placeholder=f"{result.get('store','')} {result.get('date','')}")

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
