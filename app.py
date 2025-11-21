
# ==========================================================
# EY GIS RECONCILER — FINAL VERSION (GIS → ORG CHART FUZZY)
# ==========================================================

import streamlit as st
st.set_page_config(page_title="EY GIS Reconciler — Final", layout="wide")

import pandas as pd
import json
import base64
import tempfile
import io
import re
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side

# Azure Vision
try:
    from openai import AzureOpenAI
except:
    AzureOpenAI = None

try:
    import fitz  # PyMuPDF
except:
    fitz = None

# OCR fallback
try:
    from PIL import Image
    import pytesseract
except:
    Image = None
    pytesseract = None


# -------------------------
# SIDEBAR
# -------------------------
st.sidebar.header("Controls")
FUZZY_THRESHOLD = st.sidebar.slider("Fuzzy threshold", 60, 100, 80, 1)


# -------------------------
# AZURE CONFIG
# -------------------------
AZURE_ENDPOINT = st.secrets.get("AZURE_OPENAI_ENDPOINT", None)
AZURE_KEY = st.secrets.get("AZURE_OPENAI_KEY", None)
AZURE_DEPLOY = st.secrets.get("AZURE_OPENAI_DEPLOYMENT", "gpt4o")
AZURE_API_VER = st.secrets.get("AZURE_OPENAI_API_VERSION", "2025-03-01-preview")

client = None
if AZURE_KEY and AzureOpenAI:
    client = AzureOpenAI(api_key=AZURE_KEY,
                         api_version=AZURE_API_VER,
                         azure_endpoint=AZURE_ENDPOINT)


# -------------------------
# NORMALIZATION
# -------------------------
def normalize_gis(raw):
    """Use only columns Entity Name + Parent Name."""
    cols = [c.lower().strip() for c in raw.columns]

    if "entity name" in cols and "parent name" in cols:
        ent = raw.columns[cols.index("entity name")]
        par = raw.columns[cols.index("parent name")]
        df = pd.DataFrame({
            "entity": raw[ent].astype(str).str.strip(),
            "parent": raw[par].astype(str).str.strip()
        })
        return df

    st.error("GIS file must contain 'Entity Name' and 'Parent Name' columns.")
    return pd.DataFrame(columns=["entity", "parent"])


def normalize_client(df):
    """Convert extraction JSON/CSV to entity + parent."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["entity", "parent"])

    cols = [c.lower() for c in df.columns]

    ent_col = df.columns[0]
    par_col = df.columns[1] if len(df.columns) > 1 else None

    df_out = pd.DataFrame({
        "entity": df[ent_col].astype(str).str.strip(),
        "parent": df[par_col].astype(str).str.strip() if par_col else ""
    })

    return df_out


# -------------------------
# PDF → Images
# -------------------------
def pdf_to_images(pdf_bytes, dpi=250):
    if fitz is None:
        st.error("PyMuPDF not installed.")
        return []
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        path = tmp.name
    pdf = fitz.open(path)

    images = []
    for p in pdf:
        pix = p.get_pixmap(dpi=dpi)
        images.append(pix.tobytes("png"))
    return images


# -------------------------
# Azure Vision Extractor
# -------------------------
def call_gpt4o_extract(image_bytes):
    if client is None:
        return None

    img64 = base64.b64encode(image_bytes).decode()

    try:
        resp = client.responses.create(
            model=AZURE_DEPLOY,
            input=[{
                "role": "user",
                "content": [
                    {"type": "input_text",
                     "text": "Extract ALL entities and parents. "
                             "Return STRICT JSON array: "
                             "[{\"entity\":\"...\",\"parent\":\"...\"}, ...]"},
                    {"type": "input_image",
                     "image_url": f"data:image/png;base64,{img64}"}
                ]
            }],
            max_output_tokens=1400
        )
        text_out = []
        for msg in resp.output:
            for c in msg.content:
                if hasattr(c, "text"):
                    text_out.append(c.text)
        return "\n".join(text_out)
    except Exception as e:
        st.warning(f"Vision error: {e}")
        return None


# -------------------------
# Parse JSON from model
# -------------------------
def parse_model_json(raw_text):
    if not raw_text:
        return None

    cleaned = raw_text.replace("```json", "").replace("```", "").strip()

    try:
        obj = json.loads(cleaned)
        df = pd.DataFrame(obj)
        return normalize_client(df)
    except:
        pass

    try:
        obj = eval(cleaned.replace(": null", ": None"))
        df = pd.DataFrame(obj)
        return normalize_client(df)
    except:
        pass

    return None


# -------------------------
# *** FINAL RECONCILIATION LOGIC ***
# GIS → Org Chart only (as requested)
# -------------------------
def build_reconciliation(df_client, df_gis, fuzzy_threshold=80):
    """
    NEW VERSION:
    - Fuzzy Matching is now applied ONLY on ORG CHART (df_client)
    - GIS fuzzy matching removed completely
    - Output contains fuzzy match entity + fuzzy match parent from ORG CHART
    """

    df_c = df_client.copy()
    df_g = df_gis.copy()

    df_c["entity_l"] = df_c["entity"].astype(str).str.lower().str.strip()
    df_c["parent_l"] = df_c["parent"].astype(str).str.lower().str.strip()
    df_g["entity_l"] = df_g["entity"].astype(str).str.lower().str.strip()
    df_g["parent_l"] = df_g["parent"].astype(str).str.lower().str.strip()

    all_entities = sorted(set(df_c["entity_l"]).union(set(df_g["entity_l"])))

    rows = []
    for ent in all_entities:
        crows = df_c[df_c["entity_l"] == ent]
        grows = df_g[df_g["entity_l"] == ent]

        # extract names from both sides
        client_name = crows.iloc[0]["entity"] if not crows.empty else ""
        client_parent = crows.iloc[0]["parent"] if not crows.empty else ""

        gis_name = grows.iloc[0]["entity"] if not grows.empty else ""
        gis_parent = grows.iloc[0]["parent"] if not grows.empty else ""

        # ----------------------------
        # NEW → fuzzy match AGAINST ORG CHART ONLY
        # ----------------------------
        best_ent_match = ""
        best_ent_score = 0

        best_par_match = ""
        best_par_score = 0

        for _, r in df_c.iterrows():

            # entity name fuzzy
            e_score = fuzz.token_sort_ratio(str(gis_name).lower(), str(r["entity"]).lower())
            if e_score > best_ent_score:
                best_ent_score = e_score
                best_ent_match = r["entity"]

            # parent fuzzy
            p_score = fuzz.token_sort_ratio(str(gis_parent).lower(), str(r["parent"]).lower())
            if p_score > best_par_score:
                best_par_score = p_score
                best_par_match = r["parent"]

        # ----------------------------
        # ACTION LOGIC (unchanged)
        # ----------------------------
        in_client = bool(client_name.strip())
        in_gis = bool(gis_name.strip())

        if in_client and not in_gis:
            action = "To be added in GIS."
            color = "yellow"
        elif in_gis and not in_client:
            action = "To be removed from GIS."
            color = "red"
        else:
            if str(client_parent).lower().strip() == str(gis_parent).lower().strip():
                action = "Matched"
                color = "white"
            else:
                action = "Mismatch noted, please further check."
                color = "red"

        rows.append({
            "Entity Name in GIS": gis_name,
            "Entity Name in Org Chart": client_name,
            "Parent Name in GIS": gis_parent,
            "Parent Name in Org Chart": client_parent,

            # NEW fuzzy columns using ORG CHART only
            "Fuzzy Match (Org Entity)": best_ent_match,
            "Fuzzy Match (Org Parent)": best_par_match,
            "Fuzzy Score (Entity)": best_ent_score,
            "Fuzzy Score (Parent)": best_par_score,

            "Action point": action,
            "_color": color
        })

    return pd.DataFrame(rows)

# -------------------------
# Excel Export
# -------------------------
def export_excel(recon_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"

    headers = [
        "Entity Name in GIS",
        "Entity Name in Org Chart",
        "Parent Name in GIS",
        "Parent Name in Org Chart",
        "Fuzzy Match (Org Entity)",
        "Fuzzy Match (Org Parent)",
        "Fuzzy Score (Entity)",
        "Fuzzy Score (Parent)",
        "Action point"
    ]
    ws.append(headers)

    yellow = "FFF2CC"
    red = "F8CBCC"
    white = "FFFFFF"
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))

    for _, r in recon_df.iterrows():
        ws.append([
            r.get("Entity Name in GIS", ""),
            r.get("Entity Name in Org Chart", ""),
            r.get("Parent Name in GIS", ""),
            r.get("Parent Name in Org Chart", ""),
            r.get("Fuzzy Match (Org Entity)", ""),
            r.get("Fuzzy Match (Org Parent)", ""),
            r.get("Fuzzy Score (Entity)", ""),
            r.get("Fuzzy Score (Parent)", ""),
            r.get("Action point", "")
        ])

    # header style
    for cell in ws[1]:
        cell.fill = PatternFill(start_color="FFD700", fill_type="solid")
        cell.font = Font(bold=True)
        cell.border = border

    # row coloring
    for i in range(2, len(recon_df) + 2):
        color = recon_df.iloc[i-2].get("_color", "white")
        fill_color = white if color == "white" else red if color == "red" else yellow
        for cell in ws[i]:
            cell.fill = PatternFill(start_color=fill_color, fill_type="solid")
            cell.border = border

    # save
    path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    wb.save(path)
    return path

# -------------------------
# UI
# -------------------------
st.title("EY GIS Reconciler — GIS → Org Chart (Fuzzy Only Within Org Chart)")

client_file = st.file_uploader("Upload Org Chart (PNG/JPG/PDF/CSV/JSON)", type=["png","jpg","jpeg","pdf","csv","json"])
gis_file = st.file_uploader("Upload GIS Extract", type=["xlsx","csv"])

if not gis_file:
    st.stop()

# Load GIS
if gis_file.name.endswith(".xlsx"):
    raw_gis = pd.read_excel(gis_file)
else:
    raw_gis = pd.read_csv(gis_file)

df_gis = normalize_gis(raw_gis)


# -------------------------
# Extract client org chart
# -------------------------
df_client = pd.DataFrame(columns=["entity", "parent"])
pages = []

if client_file:
    data = client_file.read()

    # PDF
    if client_file.name.endswith(".pdf"):
        images = pdf_to_images(data, dpi=250)
        for img in images:
            raw = call_gpt4o_extract(img)
            df = parse_model_json(raw)
            if df is not None:
                pages.append(df)

    # Images
    elif client_file.name.lower().endswith(("png","jpg","jpeg")):
        raw = call_gpt4o_extract(data)
        df = parse_model_json(raw)
        if df is not None:
            pages.append(df)

    # CSV / JSON
    else:
        try:
            if client_file.name.endswith(".json"):
                df_client_raw = pd.DataFrame(json.loads(data))
            else:
                df_client_raw = pd.read_csv(io.BytesIO(data))
            pages.append(normalize_client(df_client_raw))
        except:
            pass

if pages:
    df_client = pd.concat(pages, ignore_index=True)


# Show preview
st.subheader("Extracted Org Chart (Client)")
st.dataframe(df_client)


# -------------------------
# RUN RECONCILIATION
# -------------------------
st.subheader("Reconciliation (GIS → Org Chart)")
recon = build_reconciliation(df_client, df_gis, fuzzy_threshold=FUZZY_THRESHOLD)

st.dataframe(recon.drop(columns=["_color"]), height=500)


# -------------------------
# EXPORT BUTTON
# -------------------------
path = export_excel(recon)
with open(path, "rb") as f:
    st.download_button("Download Reconciliation Excel", f, "reconciliation.xlsx")

