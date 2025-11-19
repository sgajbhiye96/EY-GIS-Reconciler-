
import streamlit as st
st.set_page_config(page_title="EY GIS Reconciler", layout="wide")

import pandas as pd
import base64
import tempfile
import json
import fitz  # PyMuPDF
from openai import AzureOpenAI
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from rapidfuzz import fuzz

# ---------------------------
# Azure configuration
# ---------------------------
AZURE_ENDPOINT = "https://azureopenaids2025.openai.azure.com"
DEPLOYMENT_NAME = "gpt4o"
API_VERSION = "2025-03-01-preview"
API_KEY = st.secrets["AZURE_OPENAI_KEY"]

client = AzureOpenAI(
    api_key=API_KEY,
    api_version=API_VERSION,
    azure_endpoint=AZURE_ENDPOINT
)

# ---------------------------
# PDF to Images
# ---------------------------
def pdf_to_images(uploaded_file):
    images = []
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
    pdf = fitz.open(tmp_path)
    for page in pdf:
        pix = page.get_pixmap(dpi=200)
        images.append(pix.tobytes("png"))
    return images

# ---------------------------
# GPT-4o Vision Extraction
# ---------------------------
def call_gpt4o_extract(image_bytes, max_output_tokens=1400):
    try:
        img_b64 = base64.b64encode(image_bytes).decode()

        response = client.responses.create(
            model=DEPLOYMENT_NAME,
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "Extract ALL entities and parent-child relationships "
                                "from this organisation chart. "
                                "Return STRICT JSON = [{\"entity\":\"\",\"parent\":\"\"}, ...]. "
                                "Top-level → parent=null. No hallucinations."
                            )
                        },
                        {
                            "type": "input_image",
                            "image_url": f"data:image/png;base64,{img_b64}"
                        }
                    ]
                }
            ],
            max_output_tokens=max_output_tokens
        )

        extracted = []
        if not response.output:
            return None

        for msg in response.output:
            for c in msg.content:
                if hasattr(c, "text") and c.text:
                    extracted.append(c.text)

        return "\n".join(extracted)

    except Exception as e:
        st.error(f"GPT Vision error: {e}")
        return None

# ---------------------------
# Robust JSON Parsing
# ---------------------------
def parse_model_json(raw_text):
    if not raw_text:
        return None
    cleaned = raw_text.strip().replace("```json", "").replace("```", "").strip()

    # JSON attempt
    try:
        return pd.DataFrame(json.loads(cleaned))
    except:
        pass

    # Python fallback
    try:
        py = cleaned.replace(": null", ": None")
        return pd.DataFrame(eval(py))
    except Exception as e:
        st.error(f"JSON parse failed: {e}")
        st.code(raw_text[:3000])
        return None

# ---------------------------
# Normalization FIX (IMPORTANT)
# ---------------------------
def normalize_entity_parent(df):
    df = df.copy()
    # Normalize column names to lower
    df.columns = [c.lower().strip() for c in df.columns]

    # Ensure entity exists
    if "entity" not in df.columns:
        df = df.rename(columns={df.columns[0]: "entity"})

    # Ensure parent exists
    if "parent" not in df.columns:
        if len(df.columns) > 1:
            df = df.rename(columns={df.columns[1]: "parent"})
        else:
            df["parent"] = ""

    df["entity"] = df["entity"].astype(str).str.strip()
    df["parent"] = (
        df["parent"]
        .astype(str)
        .str.strip()
        .replace(["None", "none", "nan", "NaN", "NULL", "null"], "")
    )
    return df[["entity", "parent"]]

# ---------------------------
# Collapse duplicate entities using fuzzy + GPT-confidence
# ---------------------------
def collapse_duplicates(df, conf_df):
    df = df.copy()

    groups = {}
    for _, r in conf_df.iterrows():
        rep = r["representative"]
        for m in r["merged_items"]:
            groups.setdefault(rep, []).append(m)

    dedup_rows = []
    for rep, members in groups.items():
        parents = df[df["entity"].isin(members)]["parent"].unique().tolist()
        parent = parents[0] if parents else ""
        dedup_rows.append({"entity": rep, "parent": parent})

    dedup = pd.DataFrame(dedup_rows)
    return dedup[["entity", "parent"]]

# ---------------------------
# Reconciliation with separate fuzzy columns & parent comparison
# ---------------------------
def build_reconciliation(df_client, df_gis):
    # Defensive copies
    df_client_cmp = df_client.copy()
    df_gis_cmp = df_gis.copy()

    # lower-case helpers for comparison
    df_client_cmp["entity_l"] = df_client_cmp["entity"].astype(str).str.lower()
    df_client_cmp["parent_l"] = df_client_cmp["parent"].astype(str).str.lower()

    df_gis_cmp["entity_l"] = df_gis_cmp["entity"].astype(str).str.lower()
    df_gis_cmp["parent_l"] = df_gis_cmp["parent"].astype(str).str.lower()

    entities = sorted(
        set(df_client_cmp["entity_l"]).union(set(df_gis_cmp["entity_l"]))
    )

    rows = []
    for ent in entities:
        client_row = df_client_cmp[df_client_cmp["entity_l"] == ent]
        gis_row = df_gis_cmp[df_gis_cmp["entity_l"] == ent]

        entity_disp = (
            client_row.iloc[0]["entity"] if not client_row.empty else gis_row.iloc[0]["entity"]
        )

        org_parent = client_row.iloc[0]["parent"] if not client_row.empty else ""
        gis_parent = gis_row.iloc[0]["parent"] if not gis_row.empty else ""

        # -----------------------------
        # 1. Exact match comparison
        # -----------------------------
        if client_row.empty:
            exact_status = "MISSING IN CLIENT"
        elif gis_row.empty:
            exact_status = "MISSING IN GIS"
        else:
            if client_row.iloc[0]["parent_l"] == gis_row.iloc[0]["parent_l"]:
                exact_status = "EXACT MATCH"
            else:
                exact_status = "PARENT MISMATCH"

        # -----------------------------
        # 2. Fuzzy Matching (best candidate in GIS for entity)
        # -----------------------------
        fuzzy_best = None
        fuzzy_best_parent = ""
        fuzzy_score = 0

        for _, g in df_gis_cmp.iterrows():
            # ensure strings
            g_ent = str(g["entity"])
            score = fuzz.token_sort_ratio(entity_disp.lower(), g_ent.lower())
            if score > fuzzy_score:
                fuzzy_score = score
                fuzzy_best = g_ent
                fuzzy_best_parent = g.get("parent", "")

        # -----------------------------
        # 3. Final parent comparison logic
        # -----------------------------
        if exact_status == "EXACT MATCH":
            final_parent_compare = "Exact Parent Match"
        elif exact_status == "PARENT MISMATCH":
            final_parent_compare = "Exact Parent Mismatch"
        elif fuzzy_score >= 85:
            # compare org_parent vs fuzzy_best_parent
            if str(org_parent).strip().lower() == str(fuzzy_best_parent).strip().lower():
                final_parent_compare = "Fuzzy Parent Match"
            else:
                final_parent_compare = "Fuzzy Parent Mismatch"
        else:
            final_parent_compare = "No Suitable Match"

        # Row: keep consistent column names
        rows.append({
            "Entity": entity_disp,
            "Org Chart Parent": org_parent,
            "GIS Parent (Exact)": gis_parent,
            "Exact Status": exact_status,
            "Fuzzy Best Match (GIS)": fuzzy_best,
            "Fuzzy Best Match Parent (GIS)": fuzzy_best_parent,
            "Fuzzy Score": fuzzy_score,
            "Final Parent Comparison": final_parent_compare
        })

    recon = pd.DataFrame(rows)

    # Split outputs as needed (use Exact Status)
    only_client = recon[recon["Exact Status"] == "MISSING IN GIS"].reset_index(drop=True)
    only_gis    = recon[recon["Exact Status"] == "MISSING IN CLIENT"].reset_index(drop=True)
    mismatch    = recon[recon["Final Parent Comparison"].isin(["Exact Parent Mismatch", "Fuzzy Parent Mismatch"])].reset_index(drop=True)

    return recon, only_client, only_gis, mismatch

# ---------------------------
# Excel Workbook Builder (robust)
# ---------------------------
def apply_style(ws):
    yellow = "FFD700"
    gray = "E0E0E0"
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    # style header row (first row)
    for cell in ws[1]:
        cell.fill = PatternFill(start_color=yellow, fill_type="solid")
        cell.font = Font(bold=True)
        cell.border = border

    # style other rows
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.fill = PatternFill(start_color=gray, fill_type="solid")
            cell.border = border

def create_excel(df_client, df_gis, recon, only_client, only_gis, mismatch):
    wb = Workbook()

    # -- Reconciliation Table (main)
    ws = wb.active
    ws.title = "Reconciliation Table"

    main_headers = [
        "Entity",
        "Org Chart Parent",
        "GIS Parent (Exact)",
        "Exact Status",
        "Fuzzy Best Match (GIS)",
        "Fuzzy Best Match Parent (GIS)",
        "Fuzzy Score",
        "Final Parent Comparison"
    ]
    ws.append(main_headers)
    for _, r in recon.iterrows():
        ws.append([
            r.get("Entity", ""),
            r.get("Org Chart Parent", ""),
            r.get("GIS Parent (Exact)", r.get("GIS Parent", "")),
            r.get("Exact Status", ""),
            r.get("Fuzzy Best Match (GIS)", ""),
            r.get("Fuzzy Best Match Parent (GIS)", ""),
            r.get("Fuzzy Score", ""),
            r.get("Final Parent Comparison", "")
        ])
    apply_style(ws)

    # -- Only in Client Org
    ws = wb.create_sheet("Only in Client Org")
    ws.append(["Entity", "Org Chart Parent"])
    for _, r in only_client.iterrows():
        ws.append([
            r.get("Entity", ""),
            r.get("Org Chart Parent", "")
        ])
    apply_style(ws)

    # -- Only in GIS
    ws = wb.create_sheet("Only in GIS")
    ws.append([
        "Entity",
        "GIS Parent (Exact)",
        "Fuzzy Best Match (GIS)",
        "Fuzzy Best Match Parent (GIS)",
        "Fuzzy Score"
    ])
    for _, r in only_gis.iterrows():
        ws.append([
            r.get("Entity", ""),
            r.get("GIS Parent (Exact)", r.get("GIS Parent", "")),
            r.get("Fuzzy Best Match (GIS)", ""),
            r.get("Fuzzy Best Match Parent (GIS)", ""),
            r.get("Fuzzy Score", ""),
        ])
    apply_style(ws)

    # -- Parent Mismatch
    ws = wb.create_sheet("Parent Mismatch")
    ws.append([
        "Entity",
        "Client Parent",
        "GIS Parent (Exact)",
        "Fuzzy Best Match (GIS)",
        "Fuzzy Best Match Parent (GIS)",
        "Fuzzy Score",
        "Final Parent Comparison"
    ])
    for _, r in mismatch.iterrows():
        ws.append([
            r.get("Entity", ""),
            r.get("Org Chart Parent", ""),
            r.get("GIS Parent (Exact)", r.get("GIS Parent", "")),
            r.get("Fuzzy Best Match (GIS)", ""),
            r.get("Fuzzy Best Match Parent (GIS)", ""),
            r.get("Fuzzy Score", ""),
            r.get("Final Parent Comparison", "")
        ])
    apply_style(ws)

    # -- Full Extracted Client Tree
    ws = wb.create_sheet("Full Extracted Client Tree")
    ws.append(["Entity", "Parent"])
    for _, r in df_client.iterrows():
        # df_client columns are 'entity' and 'parent' (lowercase)
        ws.append([r.get("entity", ""), r.get("parent", "")])
    apply_style(ws)

    # Save and return path
    path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    wb.save(path)
    return path

# ---------------------------
# UI
# ---------------------------
st.title("📊 EY GIS Reconciler — Final Version")

with st.expander("Required New Process"):
    st.markdown("""
- Audit team uploads the latest client organisation chart into the tool  
- GIS reconciler tool will scan the file uploaded by user and compares with the GIS data  
- If the GIS recon tool cannot fetch GIS data automatically, user uploads GIS extract  
- The tool gives a detailed report highlighting differences  
- Risk team updates GIS after validation  
""")

client_file = st.file_uploader("Upload Client Org Chart (PDF/JPG/PNG)", type=["pdf","jpg","jpeg","png"])
gis_file = st.file_uploader("Upload GIS Extract (Excel/CSV)", type=["xlsx","csv"])

if not client_file or not gis_file:
    st.stop()

# GIS load
if gis_file.name.endswith(".xlsx"):
    df_gis_raw = pd.read_excel(gis_file)
else:
    df_gis_raw = pd.read_csv(gis_file)

df_gis = normalize_entity_parent(df_gis_raw)

# Extract hierarchy
st.subheader("Extracting from org chart...")
images = pdf_to_images(client_file) if client_file.name.endswith(".pdf") else [client_file.read()]
pages = []

for i, img in enumerate(images, start=1):
    st.write(f"Page {i}")
    raw = call_gpt4o_extract(img)
    if not raw:
        continue
    st.code(raw[:2500])
    df = parse_model_json(raw)
    if df is not None:
        df = normalize_entity_parent(df)
        pages.append(df)

if not pages:
    st.error("No hierarchy extracted.")
    st.stop()

df_client = pd.concat(pages, ignore_index=True)

# Reconciliation
st.subheader("Building reconciliation...")
recon, only_client, only_gis, mismatch = build_reconciliation(df_client, df_gis)

st.dataframe(recon)

# Debug output for assurance
st.write("RECON COLUMNS:", recon.columns.tolist())
st.write("ONLY_GIS COLUMNS:", only_gis.columns.tolist())

# Export
excel = create_excel(df_client, df_gis, recon, only_client, only_gis, mismatch)

with open(excel, "rb") as f:
    st.download_button("Download GIS_Reconciliation.xlsx", f, "GIS_Reconciliation.xlsx")

st.success("Done.")

