# app.py
import streamlit as st
st.set_page_config(page_title="EY GIS Reconciler — Entity/Parent (Azure Vision)", layout="wide")

import pandas as pd
import json
import base64
import tempfile
import io
import re
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side

# Azure Responses client
try:
    from openai import AzureOpenAI
except Exception:
    AzureOpenAI = None

# PDF image conversion
try:
    import fitz  # pymupdf
except Exception:
    fitz = None

# Optional OCR fallback
try:
    from PIL import Image
    import pytesseract
except Exception:
    Image = None
    pytesseract = None

# ---------------------------
# Sidebar controls
# ---------------------------
st.sidebar.header("Controls")
FUZZY_DEFAULT = 80
FUZZY_THRESHOLD = st.sidebar.slider("Fuzzy threshold", 60, 100, FUZZY_DEFAULT, 1)
st.sidebar.write("Use Azure Vision (Responses) keys in Streamlit secrets for image/PDF extraction.")

# ---------------------------
# Azure/OpenAI config (optional)
# ---------------------------
AZURE_ENDPOINT = st.secrets.get("AZURE_OPENAI_ENDPOINT", None)
AZURE_KEY = st.secrets.get("AZURE_OPENAI_KEY", None)
AZURE_DEPLOY = st.secrets.get("AZURE_OPENAI_DEPLOYMENT", "gpt4o")
AZURE_API_VER = st.secrets.get("AZURE_OPENAI_API_VERSION", "2025-03-01-preview")

client = None
if AZURE_KEY and AzureOpenAI is not None and AZURE_ENDPOINT:
    client = AzureOpenAI(api_key=AZURE_KEY, api_version=AZURE_API_VER, azure_endpoint=AZURE_ENDPOINT)

# ---------------------------
# Helpers: normalize GIS & client data (KEEP parentheses)
# ---------------------------
def normalize_gis_dataframe(raw):
    """Map GIS file to DataFrame with columns: entity, parent (keeps parentheses)."""
    if raw is None or raw.empty:
        return pd.DataFrame(columns=["entity", "parent"])
    cols = [str(c).strip() for c in raw.columns]
    lower = [c.lower() for c in cols]
    if "entity name" in lower and "parent name" in lower:
        ent = cols[lower.index("entity name")]
        par = cols[lower.index("parent name")]
        out = pd.DataFrame({
            "entity": raw[ent].astype(str).str.strip(),
            "parent": raw[par].astype(str).str.strip()
        })
        return out[["entity", "parent"]]
    # fallback detection
    ent_key = None
    par_key = None
    for k in ["entity", "name", "company"]:
        if k in lower:
            ent_key = cols[lower.index(k)]
            break
    for k in ["parent", "parent name", "parent_name"]:
        if k in lower:
            par_key = cols[lower.index(k)]
            break
    if ent_key is None:
        ent_key = cols[0]
    if par_key is None:
        par_key = cols[1] if len(cols) > 1 else None
    ent_series = raw[ent_key].astype(str).str.strip()
    par_series = raw[par_key].astype(str).str.strip() if par_key else pd.Series([""] * len(ent_series))
    # basic cleaning (KEEP parentheses)
    def clean_keep(x):
        x = str(x).replace("\n", " ").strip()
        x = re.sub(r"\s+", " ", x)
        return x
    ent_series = ent_series.apply(clean_keep)
    par_series = par_series.apply(clean_keep)
    return pd.DataFrame({"entity": ent_series, "parent": par_series})

def normalize_client_df(df):
    """Normalize extracted client dataframe to columns entity,parent (keeps parentheses)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["entity", "parent"])
    cols = [str(c).strip() for c in df.columns]
    lower = [c.lower() for c in cols]
    ent_col = None
    par_col = None
    for k in ["entity", "name", "child", "node", "company"]:
        if k in lower:
            ent_col = cols[lower.index(k)]
            break
    for k in ["parent", "parent name", "parent_name"]:
        if k in lower:
            par_col = cols[lower.index(k)]
            break
    if ent_col is None:
        ent_col = cols[0]
    if par_col is None:
        par_col = cols[1] if len(cols) > 1 else None
    ent_series = df[ent_col].astype(str).str.strip()
    par_series = df[par_col].astype(str).str.strip() if par_col else pd.Series([""] * len(ent_series))
    def clean_keep(x):
        x = str(x).replace("\n", " ").strip()
        x = re.sub(r"\s+", " ", x)
        return x
    ent_series = ent_series.apply(clean_keep)
    par_series = par_series.apply(clean_keep)
    return pd.DataFrame({"entity": ent_series, "parent": par_series})

# ---------------------------
# PDF -> images
# ---------------------------
def pdf_to_images_bytes(pdf_bytes, dpi=200):
    if fitz is None:
        raise RuntimeError("pymupdf (fitz) required for PDF processing. Install pymupdf.")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        path = tmp.name
    pdf = fitz.open(path)
    images = []
    for page in pdf:
        pix = page.get_pixmap(dpi=dpi)
        images.append(pix.tobytes("png"))
    return images

# ---------------------------
# Azure Vision call (Responses)
# ---------------------------
def call_gpt4o_extract(image_bytes, max_output_tokens=1500):
    """Call Azure Responses (vision) to extract JSON-like entity/parent array."""
    if client is None:
        return None
    img_b64 = base64.b64encode(image_bytes).decode()
    try:
        resp = client.responses.create(
            model=AZURE_DEPLOY,
            input=[{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": (
                        "Extract ALL entities and parent-child relationships from this organisation chart image. "
                        "Return STRICT JSON array EXACTLY like: [{\"entity\":\"<entity name>\", \"parent\":\"<parent name>\"}, ...]. "
                        "Top-level entities should have parent as an empty string. Return ONLY the JSON array." 
                    )},
                    {"type": "input_image", "image_url": f"data:image/png;base64,{img_b64}"}
                ]
            }],
            max_output_tokens=max_output_tokens
        )
        texts = []
        if resp.output:
            for msg in resp.output:
                for c in msg.content:
                    if hasattr(c, "text") and c.text:
                        texts.append(c.text)
        return "\n".join(texts)
    except Exception as e:
        st.warning(f"Vision API error: {e}")
        return None

# ---------------------------
# Robust model output parser
# ---------------------------
def parse_model_json(raw_text):
    """Try to parse strict JSON from model response, fallback to heuristics."""
    if not raw_text or not raw_text.strip():
        return None
    cleaned = raw_text.strip().replace("```json", "").replace("```", "").strip()
    # try json loads
    try:
        obj = json.loads(cleaned)
        if isinstance(obj, dict):
            for v in obj.values():
                if isinstance(v, list):
                    obj = v
                    break
        if isinstance(obj, list):
            df = pd.DataFrame(obj)
            return normalize_client_df(df)
    except Exception:
        pass
    # try eval for python-like output
    try:
        py = cleaned.replace(": null", ": None")
        obj = eval(py, {"__builtins__": {}})
        if isinstance(obj, list):
            df = pd.DataFrame(obj)
            return normalize_client_df(df)
    except Exception:
        pass
    # heuristic: parse lines "entity: X" "parent: Y" or "Entity - Parent" or CSV
    pairs = []
    lines = [ln.strip() for ln in cleaned.splitlines() if ln.strip()]
    i = 0
    while i < len(lines):
        ln = lines[i]
        # entity: value
        m = re.match(r'^(?:entity|name)\s*[:\-]\s*(.+)$', ln, flags=re.I)
        if m:
            ent = m.group(1).strip()
            parent = ""
            if i+1 < len(lines):
                m2 = re.match(r'^(?:parent)\s*[:\-]\s*(.+)$', lines[i+1], flags=re.I)
                if m2:
                    parent = m2.group(1).strip()
                    i += 1
            pairs.append({"entity": ent, "parent": parent})
            i += 1
            continue
        # csv-like
        if "," in ln:
            parts = [p.strip() for p in ln.split(",")]
            if len(parts) >= 2:
                pairs.append({"entity": parts[0], "parent": parts[1]})
                i += 1
                continue
        # dash split
        parts = re.split(r'\s+[—–-]\s+', ln)
        if len(parts) == 2:
            pairs.append({"entity": parts[0].strip(), "parent": parts[1].strip()})
            i += 1
            continue
        # treat as entity line
        pairs.append({"entity": ln, "parent": ""})
        i += 1
    if pairs:
        df = pd.DataFrame(pairs)
        return normalize_client_df(df)
    return None

# ---------------------------
# Reconciliation: uses ONLY entity + parent names
# ---------------------------
def build_reconciliation(df_client, df_gis, fuzzy_threshold=80):
    """
    Return reconciliation DataFrame with columns:
    Entity Name in GIS, Entity Name in Org Chart, Parent Name in GIS, Parent Name in Org Chart,
    Fuzzy Best Match (GIS), Fuzzy Best Match Parent (GIS), Fuzzy Score, Action point, _color
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

        client_name = crows.iloc[0]["entity"] if not crows.empty else ""
        client_parent = crows.iloc[0]["parent"] if not crows.empty else ""

        gis_name = grows.iloc[0]["entity"] if not grows.empty else ""
        gis_parent = grows.iloc[0]["parent"] if not grows.empty else ""

        base_to_match = (client_name or gis_name or ent).strip()

        # fuzzy best candidate from GIS names
        best_name = ""
        best_parent = ""
        best_score = 0
        for _, g in df_g.iterrows():
            score = fuzz.token_sort_ratio(base_to_match.lower(), str(g["entity"]).lower())
            if score > best_score:
                best_score = score
                best_name = g["entity"]
                best_parent = g["parent"]

        in_client = bool(str(client_name).strip())
        in_gis = bool(str(gis_name).strip())

        # action rules exactly as specified
        if in_client and not in_gis:
            action = "To be added in GIS."
            color = "red"
        elif in_gis and not in_client:
            action = "To be removed from GIS."
            color = "yellow"
        elif in_gis and in_client:
            # both present: compare parents case-insensitive (keep parentheses in values)
            if str(client_parent).strip().lower() == str(gis_parent).strip().lower():
                action = "Matched"
                color = "white"
            else:
                action = "Mismatch noted, please further check."
                color = "red"
        else:
            action = "Matched."
            color = "yellow"

        rows.append({
            "Entity Name in GIS": gis_name,
            "Entity Name in Org Chart": client_name,
            "Parent Name in GIS": gis_parent,
            "Parent Name in Org Chart": client_parent,
            "Fuzzy Best Match (GIS)": best_name,
            "Fuzzy Best Match Parent (GIS)": best_parent,
            "Fuzzy Score": int(best_score),
            "Action point": action,
            "_color": color
        })

    recon = pd.DataFrame(rows)
    cols_order = ["Entity Name in GIS","Entity Name in Org Chart","Parent Name in GIS","Parent Name in Org Chart",
                  "Fuzzy Best Match (GIS)","Fuzzy Best Match Parent (GIS)","Fuzzy Score","Action point","_color"]
    for c in cols_order:
        if c not in recon.columns:
            recon[c] = ""
    return recon[cols_order]

# ---------------------------
# Excel export styled
# ---------------------------
def export_styled_excel(recon_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"

    headers = [
        "Entity Name in GIS","Entity Name in Org Chart","Parent Name in GIS","Parent Name in Org Chart",
        "Fuzzy Best Match (GIS)","Fuzzy Best Match Parent (GIS)","Fuzzy Score","Action point"
    ]
    ws.append(headers)

    yellow = "FFF2CC"
    red = "F8CBCC"
    white = "FFFFFF"
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for _, r in recon_df.iterrows():
        ws.append([
            r["Entity Name in GIS"],
            r["Entity Name in Org Chart"],
            r["Parent Name in GIS"],
            r["Parent Name in Org Chart"],
            r["Fuzzy Best Match (GIS)"],
            r["Fuzzy Best Match Parent (GIS)"],
            r["Fuzzy Score"],
            r["Action point"]
        ])

    for cell in ws[1]:
        cell.fill = PatternFill(start_color="FFD700", fill_type="solid")
        cell.font = Font(bold=True)
        cell.border = border

    for idx, row in enumerate(recon_df.itertuples(index=False), start=2):
        color = getattr(row, "_color", "white") if hasattr(row, "_color") else "white"
        fill = white
        if color == "yellow":
            fill = yellow
        elif color == "red":
            fill = red
        for cell in ws[idx]:
            cell.fill = PatternFill(start_color=fill, fill_type="solid")
            cell.border = border

    path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    wb.save(path)
    return path

# ---------------------------
# Streamlit UI
# ---------------------------
st.title("📊 EY GIS Reconciler — Entity & Parent (Azure Vision)")

st.markdown("""
**Instructions**
- Upload the **GIS extract (Excel or CSV)** — it must contain **'Entity Name'** and **'Parent Name'** columns (case-insensitive).
- Upload the **Client org chart (PNG/JPG/PDF)**. Azure Vision (Responses API) will attempt to extract `entity,parent` JSON from each page.
- If Azure Vision is not configured, upload a pre-extracted `client.csv` or `client.json` containing `entity,parent`.
""")

client_file = st.file_uploader("Client org chart (PNG/JPG/PDF) OR pre-extracted CSV/JSON", type=["png","jpg","jpeg","pdf","csv","json"])
gis_file = st.file_uploader("GIS Extract (Excel/CSV) — must contain 'Entity Name' & 'Parent Name'", type=["xlsx","csv"])

if not gis_file:
    st.info("Please upload the GIS extract.")
    st.stop()

# Load GIS
try:
    if gis_file.name.lower().endswith(".xlsx"):
        raw_gis = pd.read_excel(gis_file)
    else:
        raw_gis = pd.read_csv(gis_file)
except Exception as e:
    st.error(f"Failed to read GIS file: {e}")
    st.stop()

df_gis = normalize_gis_dataframe(raw_gis)

# Extract client entities:
st.subheader("Extracting from org chart (Azure Vision)")

pages = []
raw_outputs = []

if client_file:
    data = client_file.read()
    # image or pdf or csv/json
    if client_file.name.lower().endswith(".pdf"):
        if fitz is None:
            st.error("pymupdf not installed — cannot process PDF. Install pymupdf.")
        else:
            try:
                images = pdf_to_images_bytes(data, dpi=250)
            except Exception as e:
                st.error(f"PDF -> images conversion failed: {e}")
                images = []
            for i, img in enumerate(images, start=1):
                st.write(f"Page {i}")
                raw = None
                # Azure Vision
                if client is not None:
                    raw = call_gpt4o_extract(img)
                if raw:
                    st.code(raw[:2500])
                    raw_outputs.append(raw)
                    df_parsed = parse_model_json(raw)
                    if df_parsed is not None:
                        pages.append(df_parsed)
                    else:
                        st.warning(f"Vision output parsed to nothing on page {i}.")
                else:
                    st.warning(f"No Vision output page {i}. Trying OCR fallback.")
                    if Image and pytesseract:
                        pil = Image.open(io.BytesIO(img))
                        ocr_text = pytesseract.image_to_string(pil)
                        if ocr_text:
                            st.code(ocr_text[:1000])
                            raw_outputs.append(ocr_text)
                            df_parsed = parse_model_json(ocr_text)
                            if df_parsed is not None:
                                pages.append(df_parsed)
                    else:
                        st.info("OCR not available (pytesseract/Pillow missing).")
    elif client_file.name.lower().endswith((".png","jpg","jpeg")):
        img = data
        raw = None
        if client is not None:
            raw = call_gpt4o_extract(img)
        if raw:
            st.code(raw[:2500])
            raw_outputs.append(raw)
            df_parsed = parse_model_json(raw)
            if df_parsed is not None:
                pages.append(df_parsed)
            else:
                st.warning("Vision output parsed to nothing.")
        else:
            st.warning("No Vision output for image. Trying OCR fallback.")
            if Image and pytesseract:
                pil = Image.open(io.BytesIO(img))
                ocr_text = pytesseract.image_to_string(pil)
                if ocr_text:
                    st.code(ocr_text[:1000])
                    raw_outputs.append(ocr_text)
                    df_parsed = parse_model_json(ocr_text)
                    if df_parsed is not None:
                        pages.append(df_parsed)
            else:
                st.info("OCR not available (pytesseract/Pillow missing).")
    else:
        # csv or json uploaded
        try:
            if client_file.name.lower().endswith(".json"):
                decoded = data.decode("utf-8")
                obj = json.loads(decoded)
                df_client_raw = pd.DataFrame(obj)
            else:
                df_client_raw = pd.read_csv(io.BytesIO(data))
            df_client = normalize_client_df(df_client_raw)
            pages.append(df_client)
        except Exception as e:
            st.error(f"Failed to read client CSV/JSON: {e}")
else:
    st.info("No client file uploaded. You can upload a PNG/JPG/PDF or a pre-extracted CSV/JSON.")

if pages:
    df_client = pd.concat(pages, ignore_index=True)
    df_client = normalize_client_df(df_client)
else:
    df_client = pd.DataFrame(columns=["entity", "parent"])
    st.warning("No extracted client entities — proceeding with empty client tree.")

# show preview
st.subheader("Preview – extracted entities (client)")
if df_client.empty:
    st.write("No client entities extracted.")
else:
    st.dataframe(df_client.head(500))

if raw_outputs:
    with st.expander("Raw Vision outputs"):
        for i, r in enumerate(raw_outputs, start=1):
            st.write(f"--- Page {i} ---")
            st.code(r[:3000])

# persist session state for approvals
if "df_client" not in st.session_state:
    st.session_state.df_client = df_client.copy()
if "df_gis" not in st.session_state:
    st.session_state.df_gis = df_gis.copy()

# Reconciliation
st.subheader("Reconciliation table")
recon = build_reconciliation(st.session_state.df_client, st.session_state.df_gis, fuzzy_threshold=FUZZY_THRESHOLD)

display_cols = ["Entity Name in GIS","Entity Name in Org Chart","Parent Name in GIS","Parent Name in Org Chart",
                "Fuzzy Best Match (GIS)","Fuzzy Best Match Parent (GIS)","Fuzzy Score","Action point"]
st.dataframe(recon[display_cols], height=500)

# Human approval (fuzzy suggestions) — only for GIS entities missing in client (Entity Name in Org Chart empty)
st.subheader("Approve fuzzy suggestions (client missing)")
candidates = recon[(recon["Entity Name in Org Chart"].str.strip() == "") & (recon["Fuzzy Score"] >= FUZZY_THRESHOLD)]

if not candidates.empty:
    cand_display = candidates.apply(lambda r: f"{r['Fuzzy Best Match (GIS)']} — parent: {r['Fuzzy Best Match Parent (GIS)']} (score {r['Fuzzy Score']})", axis=1).tolist()
    sel = st.selectbox("Choose fuzzy candidate", options=cand_display)
    if st.button("Accept fuzzy candidate and add to client tree"):
        idx = cand_display.index(sel)
        row = candidates.iloc[idx]
        new_row = {"entity": row["Fuzzy Best Match (GIS)"], "parent": row["Fuzzy Best Match Parent (GIS)"]}
        st.session_state.df_client = pd.concat([st.session_state.df_client, pd.DataFrame([new_row])], ignore_index=True)
        st.success(f"Added '{new_row['entity']}' to client tree (in-session).")
        st.experimental_rerun()
else:
    st.info("No fuzzy candidates meeting threshold and missing in client.")

# Export styled Excel
st.subheader("Export")
excel_path = export_styled_excel(recon)
with open(excel_path, "rb") as f:
    st.download_button("Download reconciliation (styled Excel)", f, file_name="reconciliation.xlsx")

# Export updated client CSV (after approvals)
st.subheader("Download updated client after approvals")
if not st.session_state.df_client.empty:
    tmp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv").name
    st.session_state.df_client.to_csv(tmp_csv, index=False)
    with open(tmp_csv, "rb") as f:
        st.download_button("Download updated client CSV", f, file_name="client_updated.csv")

st.write("Done. This tool compares ONLY Entity Name & Parent Name (no IDs). Parent matching keeps parentheses unchanged.")
