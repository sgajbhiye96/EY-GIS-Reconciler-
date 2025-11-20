

import streamlit as st
st.set_page_config(page_title="EY GIS Reconciler — Entity/Parent Focus", layout="wide")

import pandas as pd
import json
import base64
import tempfile
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side

# Optional Azure client (only used if secrets present)
try:
    from openai import AzureOpenAI
except Exception:
    AzureOpenAI = None

# -----------------------
# Sidebar / Controls
# -----------------------
st.sidebar.header("Controls")
FUZZY_DEFAULT = 80
FUZZY_THRESHOLD = st.sidebar.slider("Fuzzy threshold", 60, 100, FUZZY_DEFAULT, 1)
st.sidebar.markdown("**GIS columns used:** `Entity Name` and `Parent Name` (confirmed).")

# Azure/OpenAI secrets (optional)
AZURE_ENDPOINT = st.secrets.get("AZURE_OPENAI_ENDPOINT", None)
AZURE_KEY = st.secrets.get("AZURE_OPENAI_KEY", None)
AZURE_DEPLOY = st.secrets.get("AZURE_OPENAI_DEPLOYMENT", "gpt4o")
AZURE_API_VER = st.secrets.get("AZURE_OPENAI_API_VERSION", "2025-03-01-preview")

client = None
if AZURE_KEY and AzureOpenAI is not None and AZURE_ENDPOINT:
    client = AzureOpenAI(api_key=AZURE_KEY, api_version=AZURE_API_VER, azure_endpoint=AZURE_ENDPOINT)

# -----------------------
# Helper functions
# -----------------------
def normalize_entity_parent_from_generic(df):
    """
    Normalize a DataFrame to exactly columns ['entity','parent'] (strings)
    If GIS file has 'Entity Name' and 'Parent Name', map them.
    If it's generic (json/csv from extraction), attempt to find columns.
    """
    df = df.copy()
    # standardize column names for detection
    cols_l = [str(c).strip() for c in df.columns]
    lower = [c.lower() for c in cols_l]

    # Try GIS explicit mapping first
    if "entity name" in [c.lower() for c in cols_l] and "parent name" in [c.lower() for c in cols_l]:
        # preserve original case column names by finding indices
        ent_col = cols_l[[c.lower() for c in cols_l].index("entity name")]
        parent_col = cols_l[[c.lower() for c in cols_l].index("parent name")]
        df2 = pd.DataFrame({
            "entity": df[ent_col].astype(str).str.strip(),
            "parent": df[parent_col].astype(str).str.strip()
        })
        return df2

    # Otherwise fallback heuristics
    # If 'entity' and 'parent' present
    if "entity" in lower and "parent" in lower:
        ent_col = cols_l[lower.index("entity")]
        parent_col = cols_l[lower.index("parent")]
        return pd.DataFrame({
            "entity": df[ent_col].astype(str).str.strip(),
            "parent": df[parent_col].astype(str).str.strip()
        })

    # Try first two columns
    if len(cols_l) >= 2:
        ent_col = cols_l[0]
        parent_col = cols_l[1]
        return pd.DataFrame({
            "entity": df[ent_col].astype(str).str.strip(),
            "parent": df[parent_col].astype(str).str.strip()
        })

    # If only one column, parent empty
    ent_col = cols_l[0] if cols_l else "entity"
    df2 = pd.DataFrame({
        "entity": df[ent_col].astype(str).str.strip(),
        "parent": ["" for _ in range(len(df))]
    })
    return df2

def compute_reconciliation(df_client, df_gis, fuzzy_threshold=80):
    """
    Compute reconciliation rows.
    Returns DataFrame with:
    ['Entity Name in GIS','Entity Name in Org Chart','Parent Name in GIS','Parent Name in Org Chart',
     'Fuzzy Best Match (GIS)','Fuzzy Best Match Parent (GIS)','Fuzzy Score','Action point','_color']
    """
    df_c = df_client.copy()
    df_g = df_gis.copy()

    # lowercase helper columns for comparison
    df_c["entity_l"] = df_c["entity"].str.lower().str.strip()
    df_c["parent_l"] = df_c["parent"].str.lower().str.strip()
    df_g["entity_l"] = df_g["entity"].str.lower().str.strip()
    df_g["parent_l"] = df_g["parent"].str.lower().str.strip()

    # union of names
    all_entities = sorted(set(df_c["entity_l"]).union(set(df_g["entity_l"])))

    rows = []
    for ent in all_entities:
        c_rows = df_c[df_c["entity_l"] == ent]
        g_rows = df_g[df_g["entity_l"] == ent]

        client_name = c_rows.iloc[0]["entity"] if not c_rows.empty else ""
        gis_name = g_rows.iloc[0]["entity"] if not g_rows.empty else ""

        client_parent = c_rows.iloc[0]["parent"] if not c_rows.empty else ""
        gis_parent = g_rows.iloc[0]["parent"] if not g_rows.empty else ""

        # Base string to fuzzy-match: prefer visible client name, otherwise GIS, otherwise raw ent key
        base_to_match = (client_name or gis_name or ent).strip()

        # Find best fuzzy candidate from GIS (search only GIS entity names)
        best_name = ""
        best_parent = ""
        best_score = 0
        for _, g in df_g.iterrows():
            score = fuzz.token_sort_ratio(base_to_match.lower(), str(g["entity"]).lower())
            if score > best_score:
                best_score = score
                best_name = g["entity"]
                best_parent = g["parent"]

        # Determine presence of text names (we treat empty strings as missing)
        in_client = bool(client_name and client_name.strip())
        in_gis = bool(gis_name and gis_name.strip())

        # ACTION RULES (as requested)
        # If entity exists in Client and missing in GIS -> "To be removed from GIS tree."
        # If entity exists in GIS and missing in Client -> "To be added in GIS tree."
        # If exists in both but parent mismatch -> "Mismatch noted, please further check."
        if in_client and not in_gis:
            action = "To be removed from GIS tree."
            color = "red"
        elif in_gis and not in_client:
            action = "To be added in GIS tree."
            color = "yellow"
        elif in_gis and in_client:
            # compare parents case-insensitive
            if client_parent.strip().lower() == gis_parent.strip().lower():
                action = "Matched"
                color = "white"
            else:
                action = "Mismatch noted, please further check."
                color = "red"
        else:
            # both missing - shouldn't happen normally; treat as add if client has non-empty base
            if base_to_match.strip():
                action = "To be added in GIS tree."
                color = "yellow"
            else:
                action = "Matched"
                color = "white"

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
    return recon

def export_styled_excel(recon_df):
    """Export a styled Excel with the exact columns and coloring."""
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

    # header style
    for cell in ws[1]:
        cell.fill = PatternFill(start_color="FFD700", fill_type="solid")
        cell.font = Font(bold=True)
        cell.border = border

    # rows style
    for i, _ in enumerate(recon_df.itertuples(index=False), start=2):
        color = recon_df.iloc[i-2].get("_color", "white")
        fill_color = white
        if color == "yellow":
            fill_color = yellow
        elif color == "red":
            fill_color = red
        for cell in ws[i]:
            cell.fill = PatternFill(start_color=fill_color, fill_type="solid")
            cell.border = border

    path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    wb.save(path)
    return path

# -----------------------
# UI
# -----------------------
st.title("EY GIS Reconciler — Entity & Parent Name Matching (Option B)")
st.markdown("**Note:** This tool uses only `Entity Name` and `Parent Name` from the GIS extract and the client org chart.")

st.subheader("1) Upload inputs")
st.write("- GIS extract MUST contain columns: `Entity Name` and `Parent Name`.")
st.write("- Client extraction may be a JSON/CSV with `entity`/`parent` or you can upload PDF if Azure vision is configured.")

client_file = st.file_uploader("Client Org Chart (PDF / JSON / CSV)", type=["pdf","json","csv"])
gis_file = st.file_uploader("GIS Extract (Excel/CSV) — must contain 'Entity Name' & 'Parent Name'", type=["xlsx","csv"])

if not gis_file:
    st.stop()

# Load GIS using explicit column names "Entity Name" and "Parent Name"
try:
    if gis_file.name.lower().endswith(".xlsx"):
        raw_gis = pd.read_excel(gis_file)
    else:
        raw_gis = pd.read_csv(gis_file)
except Exception as e:
    st.error(f"Failed to read GIS file: {e}")
    st.stop()

# Normalize and map to entity,parent (explicitly reading 'Entity Name' and 'Parent Name')
df_gis = normalize_entity_parent_from_generic(raw_gis)

# Load client (either via PDF extraction or JSON/CSV)
df_client = None
if client_file:
    if client_file.name.lower().endswith(".pdf"):
        if client is None:
            st.warning("Azure OpenAI not configured. Upload pre-extracted JSON/CSV for client instead of PDF.")
        else:
            # extract using fitted approach - conservative (may require valid Azure/OpenAI)
            import fitz
            images = []
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(client_file.read())
                tmp_path = tmp.name
            pdf = fitz.open(tmp_path)
            pages = []
            for i, page in enumerate(pdf, start=1):
                pix = page.get_pixmap(dpi=300)
                img_bytes = pix.tobytes("png")
                # call Azure OpenAI Vision (Responses) similar to prior examples
                try:
                    img_b64 = base64.b64encode(img_bytes).decode()
                    resp = client.responses.create(
                        model=AZURE_DEPLOY,
                        input=[{
                            "role":"user",
                            "content":[
                                {"type":"input_text","text":(
                                    "Extract ALL entities and parent-child relationships from this org chart. "
                                    "Return STRICT JSON array like: [{\"entity\":\"<name>\",\"parent\":\"<parent name>\"}, ...]."
                                )},
                                {"type":"input_image","image_url":f"data:image/png;base64,{img_b64}"}
                            ]
                        }],
                        max_output_tokens=1500
                    )
                    text_acc = []
                    if resp.output:
                        for msg in resp.output:
                            for c in msg.content:
                                if hasattr(c, "text") and c.text:
                                    text_acc.append(c.text)
                    raw_text = "\n".join(text_acc)
                    # parse robustly
                    cleaned = raw_text.strip().replace("```json","").replace("```","").strip()
                    try:
                        parsed = pd.DataFrame(json.loads(cleaned))
                    except Exception:
                        try:
                            parsed = pd.DataFrame(eval(cleaned.replace(": null", ": None")))
                        except Exception:
                            parsed = None
                    if parsed is not None:
                        parsed = normalize_entity_parent_from_generic(parsed)
                        pages.append(parsed)
                except Exception as e:
                    st.error(f"Vision error on page {i}: {e}")
            if pages:
                df_client = pd.concat(pages, ignore_index=True)
    else:
        # JSON/CSV
        try:
            if client_file.name.lower().endswith(".json"):
                obj = json.load(client_file)
                df_client = pd.DataFrame(obj)
            else:
                df_client = pd.read_csv(client_file)
        except Exception as e:
            st.error(f"Failed to read client file: {e}")
            df_client = None

if df_client is None:
    st.info("No client extraction loaded (or extraction failed). Proceeding with empty client tree.")
    df_client = pd.DataFrame(columns=["entity","parent"])

df_client = normalize_entity_parent_from_generic(df_client)

# Persist session state for approvals
if "df_client" not in st.session_state:
    st.session_state.df_client = df_client.copy()
if "df_gis" not in st.session_state:
    st.session_state.df_gis = df_gis.copy()

# Compute reconciliation
recon = compute_reconciliation(st.session_state.df_client, st.session_state.df_gis, fuzzy_threshold=FUZZY_THRESHOLD)

st.subheader("2) Reconciliation table (Entity & Parent name focus)")
cols_to_show = ["Entity Name in GIS","Entity Name in Org Chart","Parent Name in GIS","Parent Name in Org Chart","Fuzzy Best Match (GIS)","Fuzzy Best Match Parent (GIS)","Fuzzy Score","Action point"]
st.dataframe(recon[cols_to_show], height=450)

# Human approval: accept fuzzy candidate (score >= threshold & missing in client)
st.subheader("3) Human approve fuzzy suggestion")
candidates = recon[(recon["Fuzzy Score"] >= FUZZY_THRESHOLD) & (recon["Entity Name in Org Chart"].str.strip() == "")]
if not candidates.empty:
    cand_display = candidates.apply(lambda r: f"{r['Fuzzy Best Match (GIS)']} (score: {r['Fuzzy Score']}) — parent: {r['Fuzzy Best Match Parent (GIS)']}", axis=1).tolist()
    sel = st.selectbox("Choose fuzzy candidate to accept", options=cand_display)
    if st.button("Accept fuzzy match and add to Client org chart"):
        idx = cand_display.index(sel)
        row = candidates.iloc[idx]
        new_row = {"entity": row["Fuzzy Best Match (GIS)"], "parent": row["Fuzzy Best Match Parent (GIS)"]}
        st.session_state.df_client = pd.concat([st.session_state.df_client, pd.DataFrame([new_row])], ignore_index=True)
        st.success(f"Accepted fuzzy match '{new_row['entity']}' — added to Client org chart (in-session).")
        st.experimental_rerun()
else:
    st.info("No fuzzy candidates meeting threshold AND missing in Client.")

# Export
st.subheader("4) Export styled reconciliation")
excel_path = export_styled_excel(recon)
with open(excel_path, "rb") as f:
    st.download_button("Download reconciliation (styled Excel)", f, file_name="reconciliation.xlsx")

# Optional: export updated client CSV after approvals
st.subheader("5) Export updated client (after approvals)")
csv_path = None
if not st.session_state.df_client.empty:
    csv_path = tempfile.NamedTemporaryFile(delete=False, suffix=".csv").name
    st.session_state.df_client.to_csv(csv_path, index=False)
    with open(csv_path, "rb") as f:
        st.download_button("Download updated client CSV", f, file_name="client_updated.csv")

st.write("Done. Matching relies ONLY on Entity Name and Parent Name from GIS and Org Chart extraction.")
