import io
import json
import os
import smtplib
import time
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import (
    AgGrid,
    DataReturnMode,
    GridOptionsBuilder,
    GridUpdateMode,
    JsCode,
)
import streamlit as st
from streamlit_quill import st_quill

import db

# --- PAGE CONFIG & CSS ---
st.set_page_config(
    page_title="RMS Procurement Shipment Tracker",
    layout="wide",
    initial_sidebar_state="collapsed",
)


@st.cache_resource(show_spinner=False)
def _boot(day):
  """Schema/index setup + daily backup run once per day, not on every click."""
  db.init_db()
  migration = None
  try:
    import seed  # one-time, non-destructive merge of the tracker data

    migration = seed.apply_pending()
  except Exception as exc:  # never block the app
    migration = {"error": str(exc)}
  db.recompute_all_formulas()  # delays / warnings depend on today's date
  db.auto_backup_daily()
  return migration or {}


_migration = _boot(datetime.now().strftime("%Y%m%d"))
if _migration and not st.session_state.get("_mig_seen"):
  st.session_state["_mig_seen"] = True
  if _migration.get("error"):
    st.error(
        "Data migration could not be applied and was rolled back - no data was"
        f" changed. Backup: {_migration.get('backup', '')}. Details: {_migration['error']}"
    )
_DBV = db.db_version()  # changes whenever anything is written to the database


@st.cache_data(show_spinner=False, max_entries=16)
def cached_dropdowns(category, ver):
  return db.get_dropdowns(category)


@st.cache_data(show_spinner=False, max_entries=16)
def cached_load(fiscal_year, ver):
  return db.load_data(fiscal_year)


@st.cache_data(show_spinner=False, max_entries=8)
def build_export_bytes(_export_df, cache_key):
  buf = io.BytesIO()
  with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    _export_df.to_excel(writer, index=False, sheet_name="Shipments")
  return buf.getvalue()


@st.cache_data(show_spinner=False)
def build_template_bytes(cols, help_map):
  buf = io.BytesIO()
  with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    pd.DataFrame(columns=list(cols)).to_excel(
        writer, index=False, sheet_name="Shipments"
    )
    pd.DataFrame(
        {"Column": list(cols), "What to enter": [help_map.get(c, "") for c in cols]}
    ).to_excel(writer, index=False, sheet_name="Column guide")
    ws = writer.sheets["Shipments"]
    for i, name in enumerate(cols, start=1):
      ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = max(
          16, min(34, len(name) + 4)
      )
  return buf.getvalue()


st.markdown(
    """
    <style>
        header {visibility: hidden;} #MainMenu {visibility: hidden;}
        [data-testid="stSidebar"] { display: none !important; }
        
        .block-container { padding-top: 0.5rem !important; padding-bottom: 0.5rem !important; background-color: #f4f7f6;}
        div[data-testid="stVerticalBlock"] > div { padding-bottom: 0.1rem !important; }
        div[data-testid="stForm"] { margin-bottom: 0px !important; }
        
        .top-banner {
            background: linear-gradient(90deg, #117a8b 0%, #0056b3 100%);
            padding: 6px 15px; border-radius: 6px; color: white;
            display: flex; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 8px;
        }
        .top-banner h1 { margin:0; font-weight: 800; font-size: 19px; color: white; letter-spacing: 0.5px;}
        .top-banner img { border-radius: 4px; margin-right: 15px; background: white; padding: 2px; }
        
        .eff-card {
            padding: 10px 15px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); height: 75px;
            display: flex; flex-direction: column; justify-content: center; border: 1px solid #eee; margin-top: 2px; margin-bottom: 5px;
        }
        .eff-title { font-size: 10px; font-weight: 800; text-transform: uppercase; margin-bottom: 2px; color: #495057; }
        .eff-value { font-size: 22px; font-weight: 900; margin: 0; line-height: 1.1; }
        
        div.stTabs { margin-top: -5px; }
        div[data-baseweb="tab-list"] { background-color: #e0f2f1 !important; border-radius: 8px; padding: 5px 10px; gap: 5px; }
        div[data-baseweb="tab"] { background-color: transparent !important; border-radius: 6px !important; padding-top: 8px !important; padding-bottom: 8px !important; margin: 0px !important; }
        div[data-baseweb="tab"][aria-selected="true"] { background-color: #ffffff !important; box-shadow: 0 2px 4px rgba(0,0,0,0.05) !important; border: 1px solid #b2dfdb !important; }
        
        .align-btn { display: flex; align-items: flex-end; height: 100%; padding-bottom: 2px;}

        /* ACTION TOOLBAR UNIFORM BUTTON STYLING & PERFECT ALIGNMENT */
        .toolbar-container {
            display: flex;
            align-items: center;
            width: 100%;
            margin-top: 4px;
            margin-bottom: 8px;
        }
        .toolbar-container div[data-testid="stColumn"] {
            display: flex;
            align-items: center;
        }
        .toolbar-container div[data-testid="stColumn"] > div {
            width: 100%;
        }
        .toolbar-container button,
        .toolbar-container div[data-testid="stDownloadButton"] button,
        .toolbar-container div[data-testid="stPopover"] button,
        .toolbar-container div[data-testid="stButton"] button {
            background: linear-gradient(90deg, #117a8b 0%, #0056b3 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 6px !important;
            font-weight: 800 !important;
            font-size: 13px !important;
            height: 42px !important;
            min-height: 42px !important;
            width: 100% !important;
            padding: 0px 12px !important;
            box-shadow: 0 2px 4px rgba(0,0,0,0.12) !important;
            transition: all 0.2s ease-in-out !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            margin: 0 !important;
        }
        .toolbar-container button:hover,
        .toolbar-container div[data-testid="stDownloadButton"] button:hover,
        .toolbar-container div[data-testid="stPopover"] button:hover,
        .toolbar-container div[data-testid="stButton"] button:hover {
            background: linear-gradient(90deg, #0d6575 0%, #004085 100%) !important;
            color: #ffffff !important;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2) !important;
            transform: translateY(-1px) !important;
        }
        .toolbar-container button p,
        .toolbar-container button span,
        .toolbar-container button div,
        .toolbar-container button svg {
            color: #ffffff !important;
            fill: #ffffff !important;
            font-weight: 800 !important;
            font-size: 13px !important;
            margin: 0 !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)



st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        html, body, [class*="css"], .stMarkdown, button, input, textarea, select {
            font-family: 'Plus Jakarta Sans', 'Segoe UI', system-ui, -apple-system, sans-serif !important;
        }
        .stApp { background: linear-gradient(180deg, #f4f7ff 0%, #ffffff 380px) !important; }
        .block-container { max-width: 1680px; padding: 0.9rem 2.2rem 3rem !important; background: transparent !important; }

        /* TOP BAR */
        .app-bar { display: flex; align-items: center; justify-content: space-between; padding: 12px 24px; margin-bottom: 18px; background: #fff; border-radius: 22px; box-shadow: 0 1px 2px rgba(15,23,42,0.04), 0 14px 32px -18px rgba(15,23,42,0.2); }
        .app-bar-left { display: flex; align-items: center; gap: 16px; }
        .app-bar-left img { height: 50px !important; width: auto !important; border-radius: 0 !important; background: transparent !important; padding: 0 !important; margin: 0 !important; }
        .brand { font-size: 30px; font-weight: 800; letter-spacing: -1px; color: #0f172a; }
        .app-bar-sep { width: 1px; height: 34px; background: #dde4f0; }
        .app-bar-title { font-size: 26px; font-weight: 800; color: #0f172a; letter-spacing: -0.6px; }
        .app-bar-date { font-size: 13.5px; font-weight: 600; color: #64748b; background: #fff; border: 1px solid #e8edf5; padding: 9px 18px; border-radius: 999px; }

        /* KPI CARDS - coloured, borderless, text always stays inside */
        .kpi-card { container-type: inline-size; border: none !important; border-radius: 20px; padding: 18px 20px; min-height: 98px; overflow: hidden; margin-bottom: 18px; box-shadow: 0 1px 2px rgba(15,23,42,0.04); }
        .kpi-label { font-size: clamp(11px, 7cqw, 14px); font-weight: 700; margin-bottom: 10px; line-height: 1.25; overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow-wrap: anywhere; }
        .kpi-value { font-size: clamp(14px, 17cqw, 32px); font-weight: 800; letter-spacing: -1px; line-height: 1.1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

        /* KPI row wraps instead of squeezing on narrow screens */
        div[data-testid="stHorizontalBlock"]:has(.kpi-card) { flex-wrap: wrap !important; gap: 14px !important; }
        div[data-testid="stHorizontalBlock"]:has(.kpi-card) > div { min-width: 150px !important; flex: 1 1 150px !important; }

        /* PIPELINE */
        .pipe-wrap { background: #fff; border: none; border-radius: 22px; padding: 20px 24px; margin: 16px 0 18px 0;
            box-shadow: 0 1px 2px rgba(15,23,42,0.04), 0 14px 32px -16px rgba(15,23,42,0.18); }
        .pipe-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; flex-wrap: wrap; gap: 8px; }
        .pipe-title { font-size: 15px; font-weight: 700; color: #0f172a; }
        .chip { display: inline-block; font-size: 12.5px; font-weight: 700; padding: 6px 13px; border-radius: 999px; margin-left: 8px; border: none; }
        .chip.warn { background: #fff1de; color: #b45309; }
        .chip.bad { background: #fee2e2; color: #b91c1c; }
        .chip.good { background: #dcfce7; color: #166534; }
        .pipe-bar { display: flex; gap: 3px; height: 10px; background: transparent; }
        .pipe-seg { height: 100%; border-radius: 999px; }
        .pipe-legend { display: flex; flex-wrap: wrap; gap: 8px 22px; margin-top: 14px; font-size: 13px; color: #475569; }
        .pipe-legend i { display: inline-block; width: 9px; height: 9px; border-radius: 50%; margin-right: 7px; vertical-align: 0; }

        /* TABS - pill navigation (works for both Streamlit tab implementations) */
        [role="tablist"], div[data-baseweb="tab-list"] { display: inline-flex !important; width: auto !important; gap: 4px; background: #fff !important;
            border: 1px solid #e8edf5 !important; border-radius: 999px !important; padding: 6px !important; box-shadow: 0 10px 26px -16px rgba(15,23,42,0.25) !important; }
        [role="tablist"]::after, [role="tablist"]::before, [role="tab"]::after, [role="tab"]::before { display: none !important; content: none !important; }
        [role="tab"], div[data-baseweb="tab"] { border-radius: 999px !important; border: none !important; border-bottom: none !important; padding: 10px 22px !important; height: auto !important; background: transparent !important; box-shadow: none !important; }
        [role="tab"] p, [role="tab"] span, div[data-baseweb="tab"] p { color: #475569 !important; font-weight: 600 !important; }
        [role="tab"]:hover, div[data-baseweb="tab"]:hover { background: #f1f5f9 !important; }
        [role="tab"][aria-selected="true"], div[data-baseweb="tab"][aria-selected="true"] { background: linear-gradient(90deg, #0d8a9d 0%, #1a63d6 60%, #6b46f0 100%) !important; box-shadow: 0 10px 22px -10px rgba(26,99,214,0.7) !important; }
        [role="tab"][aria-selected="true"] p, [role="tab"][aria-selected="true"] span, div[data-baseweb="tab"][aria-selected="true"] p { color: #ffffff !important; }
        div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] { display: none !important; }
        [data-testid="stTabs"] { border-bottom: none !important; }
        .react-aria-SelectionIndicator, [class*="SelectionIndicator"] { display: none !important; }

        /* INPUTS, BUTTONS, PANELS */
        div[data-baseweb="select"] > div, .stTextInput input, .stTextArea textarea, .stNumberInput input { border-radius: 14px !important; border: 1px solid #e2e8f0 !important; background: #fff !important; box-shadow: none !important; }
        .stButton > button, .stDownloadButton > button, [data-testid="stPopover"] button { border-radius: 999px !important; border: 1px solid #e2e8f0 !important; background: #fff !important; color: #0f172a !important; font-weight: 600 !important; box-shadow: none !important; padding: 0.55rem 1.1rem !important; transition: box-shadow .15s ease, border-color .15s ease; }
        .stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stPopover"] button:hover { border-color: #b8c7f5 !important; box-shadow: 0 10px 22px -14px rgba(26,99,214,0.6) !important; color: #1a63d6 !important; }
        .stButton > button[kind="primary"] { background: linear-gradient(90deg, #0d8a9d 0%, #1a63d6 60%, #6b46f0 100%) !important; color: #fff !important; border: none !important; }
        div[data-testid="stExpander"] { border: none !important; border-radius: 22px !important; background: #fff; box-shadow: 0 1px 2px rgba(15,23,42,0.04), 0 14px 32px -16px rgba(15,23,42,0.18); }
        div[data-testid="stAlert"] { border-radius: 16px !important; }
        h2, h3 { font-weight: 800 !important; color: #0f172a; letter-spacing: -0.4px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# RMS logo (embedded fallback). Drop your own file next to app.py as logo_2.jpg,
# logo.png or logo.jpg and it is used instead.
_DEFAULT_LOGO_B64 = "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"


@st.cache_resource(show_spinner=False)
def _logo_html():
  import base64

  for name, mime in (
      ("logo_2.jpg", "jpeg"), ("logo_2.png", "png"), ("logo.png", "png"),
      ("logo.jpg", "jpeg"), ("assets/logo.png", "png"),
  ):
    if os.path.exists(name):
      with open(name, "rb") as image_file:
        encoded = base64.b64encode(image_file.read()).decode()
      return f"<img src='data:image/{mime};base64,{encoded}' alt='RMS'/>"
  return f"<img src='data:image/png;base64,{_DEFAULT_LOGO_B64}' alt='RMS'/>"


logo_html = _logo_html()

_brand = logo_html
st.markdown(
    f"<div class='app-bar'><div class='app-bar-left'>{_brand}<span class='app-bar-sep'></span>"
    "<span class='app-bar-title'>Procurement Shipment Execution &amp; Tracking Dashboard</span></div>"
    f"<div class='app-bar-date'>{datetime.now().strftime('%A, %d %B %Y')}</div></div>",
    unsafe_allow_html=True,
)


_KPI_PALETTE = {
    "total": ("#e9eff8", "#1e3a5f"),
    "sea": ("#e0f5ea", "#14683a"),
    "air": ("#e1ecff", "#1d4ed8"),
    "road": ("#fff0cf", "#8a4b0c"),
    "cargo": ("#eee6ff", "#5b21b6"),
    "clear": ("#dcf4f7", "#0e7490"),
}


def create_kpi_card(title, value, bg_color, border_color, val_color):
  clean = "".join(ch for ch in title if ord(ch) < 0x2000).strip()
  bg, fg = bg_color, val_color
  for key, (pbg, pfg) in _KPI_PALETTE.items():
    if key in clean.lower():
      bg, fg = pbg, pfg
      break
  text = str(value)
  if text.endswith(" USD"):  # the label already says (USD) and the figure has $
    text = text[:-4]
  return (
      f"<div class='kpi-card' style='background: {bg};'>"
      f"<div class='kpi-label' title='{clean}' style='color: {fg};'>{clean}</div>"
      f"<div class='kpi-value' title='{text}' style='color: {fg};'>{text}</div></div>"
  )


def format_currency(val, symbol="USD"):
  try:
    val = float(val)
    if val >= 1_000_000:
      return f"${val/1_000_000:.2f}M {symbol}"
    elif val >= 1_000:
      return f"${val/1_000:.2f}K {symbol}"
    return f"${val:,.0f} {symbol}"
  except:
    return f"0 {symbol}"


def send_email_smtp(
    host, port, user, password, recipients, subject, body_html, attachment_list=None
):
  try:
    msg = MIMEMultipart()
    msg["From"] = user
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html, "html"))

    if attachment_list:
      for fname, fbytes in attachment_list:
        part = MIMEApplication(fbytes, Name=fname)
        part["Content-Disposition"] = f'attachment; filename="{fname}"'
        msg.attach(part)

    with smtplib.SMTP(host, int(port)) as server:
      server.starttls()
      server.login(user, password)
      server.sendmail(user, recipients, msg.as_string())
    return True, "Email sent successfully!"
  except Exception as e:
    return False, f"Failed to send email: {str(e)}"


ALL_DATE_COLS = [
    "ETA POD",
    "Actual arrival at port",
    "Cargo ETA RMS",
    "Actual arrival at RMS",
    "DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS",
    "Deadline to clear the shipment",
]

# Focus input columns (Col M, N, P, R, S, W, Y - EXCLUDING Current status)
FOCUS_COLUMNS = [
    "BILL OF LADING OR AIRWAY BILL",
    "POD",
    "Actual arrival at port",
    "Actual arrival at RMS",
    "DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS",
    "delays",
    "MAH letter link",
]

db_status_opts = cached_dropdowns("status", _DBV)
db_mode_opts = cached_dropdowns("mode", _DBV)
db_incoterm_opts = cached_dropdowns("incoterm", _DBV)

STATUS_OPTIONS = [""] + [opt["label"] for opt in db_status_opts]
MODE_OPTIONS = [""] + sorted(
    list({opt["label"].upper() for opt in db_mode_opts if opt["label"]})
)
INCOTERM_OPTIONS = [""] + [opt["label"] for opt in db_incoterm_opts]

status_color_map = {opt["label"]: opt["color"] for opt in db_status_opts}
mode_color_map = {
    opt["label"].upper(): opt["color"] for opt in db_mode_opts if opt["label"]
}


@st.dialog("⚙️ Manage Dropdown Options", width="large")
def manage_dropdowns_dialog():
  st.markdown("Edit existing dropdown options or add new ones.")
  cat = st.selectbox(
      "Select List to Edit", ["Current status", "Shipping mode", "Incoterm"]
  )
  cat_map = {
      "Current status": "status",
      "Shipping mode": "mode",
      "Incoterm": "incoterm",
  }
  active_cat = cat_map[cat]

  options = db.get_dropdowns(active_cat)
  st.markdown("---")

  for opt in options:
    c1, c2, c3, c4 = st.columns([1, 4, 1.5, 1.5])
    with c1:
      new_c = st.color_picker(
          "Color",
          opt["color"],
          key=f"c_{opt['id']}",
          label_visibility="collapsed",
      )
    with c2:
      new_l = st.text_input(
          "Label",
          opt["label"],
          key=f"l_{opt['id']}",
          label_visibility="collapsed",
      )
    with c3:
      if st.button(
          "💾 Update", key=f"save_{opt['id']}", use_container_width=True
      ):
        db.update_dropdown(opt["id"], new_l.strip(), new_c)
        st.rerun()
    with c4:
      if st.button(
          "❌ Del", key=f"del_{opt['id']}", use_container_width=True
      ):
        db.delete_dropdown(opt["id"])
        st.rerun()

  st.markdown("### ➕ Add New Option")
  nc1, nc2, nc3 = st.columns([1, 4, 3])
  with nc1:
    new_color = st.color_picker("Pick Color", "#ffffff", key="new_col")
  with nc2:
    new_label = st.text_input("Option Label", key="new_lab")
  with nc3:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Add Option", type="primary", use_container_width=True):
      if new_label:
        db.add_dropdown(active_cat, new_label.strip(), new_color)
        st.rerun()


# --- DIALOG 1: TAKE ACTION, EMAIL & FILE ATTACHMENTS WORKFLOW ---
@st.dialog("⚡ Take Action & Shipment Workflow", width="large")
def take_action_dialog(row_data):
  shipment_id = int(row_data.get("id", 0))
  po = str(row_data.get("PURCHASE ORDER", "N/A")).strip()
  supplier = str(row_data.get("SUPPLIER /MANUFACTURER", "N/A")).strip()
  mode = str(row_data.get("Shipping mode", "N/A")).strip().upper()
  carrier = str(row_data.get("Carrier", "N/A")).strip()
  bl = str(row_data.get("BILL OF LADING OR AIRWAY BILL", "N/A")).strip()
  pod = str(row_data.get("POD", "N/A")).strip()
  port_arr = str(row_data.get("Actual arrival at port", "N/A")).strip()
  doc_sharing = str(
      row_data.get(
          "DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS",
          "N/A",
      )
  ).strip()
  eta_rms = str(row_data.get("Cargo ETA RMS", "N/A")).strip()
  rms_arr = str(row_data.get("Actual arrival at RMS", "N/A")).strip()
  deadline = str(row_data.get("Deadline to clear the shipment", "N/A")).strip()
  delays = str(row_data.get("delays", "N/A")).strip()
  warning = str(row_data.get("Action warning", "N/A")).strip()

  st.markdown(
      f"**Target Shipment:** PO `# {po}` | **Supplier:** `{supplier}` |"
      f" **BL/AWB:** `{bl}`"
  )
  st.divider()

  act_tab1, act_tab2, act_tab3 = st.tabs([
      "✉️ Compose Email Alert",
      "📁 Attached Documents Repository",
      "📜 Change History Trail",
  ])

  with act_tab1:
    st.subheader("✉️ Compose Email Alert to RMS Team & Clearing Agents")
    rms_df = db.get_rms_emails()
    email_options = (
        rms_df["Email"].tolist()
        if not rms_df.empty
        else ["procurement@rms.rw", "clearing.agent@company.com"]
    )

    selected_recipients = st.multiselect(
        "Select RMS Recipient(s)*",
        options=email_options,
        default=email_options[:1],
    )
    custom_cc = st.text_input("Additional CC Email(s) (comma separated)")
    email_subject = st.text_input(
        "Email Subject*",
        value=f"URGENT: RMS Shipment Escalation - PO #{po} ({carrier})",
    )

    default_body = f"""
        <p>Dear Clearing & Logistics Team,</p>
        <p>Please review the shipment notification details for <b>Purchase Order #{po}</b>:</p>
        <ul>
            <li><b>PO Number:</b> {po}</li>
            <li><b>Supplier:</b> {supplier}</li>
            <li><b>Shipping Mode:</b> {mode}</li>
            <li><b>Carrier:</b> {carrier}</li>
            <li><b>BL / Airway Bill:</b> {bl}</li>
            <li><b>Port of Discharge:</b> {pod}</li>
            <li><b>Actual Port Arrival:</b> {port_arr}</li>
            <li><b>Doc Sharing Date:</b> {doc_sharing}</li>
            <li><b>ETA RMS:</b> {eta_rms}</li>
            <li><b>Actual RMS Arrival:</b> {rms_arr}</li>
            <li><b>Clearance Deadline:</b> {deadline}</li>
            <li><b>Delay Days:</b> {delays}</li>
            <li><b>Action Warning:</b> {warning}</li>
        </ul>
        <p>Kindly expedite document processing with RFDA / Customs and provide an updated status immediately.</p>
        <p>Best regards,<br><b>RMS Procurement Team</b></p>
        """

    st.markdown("**Compose Email Message:**")
    email_body_html = st_quill(
        value=default_body, html=True, key=f"quill_email_action_{shipment_id}"
    )

    st.markdown("### 📎 Select File Attachments")
    existing_docs_df = db.get_row_documents(shipment_id)
    selected_doc_ids = []
    if not existing_docs_df.empty:
      st.markdown("Select from previously uploaded shipment documents:")
      for _, d_row in existing_docs_df.iterrows():
        if st.checkbox(
            f"Doc #{d_row['Doc #']}: {d_row['File Name']} ({d_row['Uploaded By']})",
            key=f"att_chk_{d_row['id']}",
        ):
          selected_doc_ids.append(d_row["id"])

    new_att_files = st.file_uploader(
        "Or attach new document(s)",
        accept_multiple_files=True,
        key=f"new_email_att_{shipment_id}",
    )

    smtp_host = os.getenv("SMTP_HOST", "smtp.office365.com")
    try:
      smtp_port = int(os.getenv("SMTP_PORT", 587))
    except Exception:
      smtp_port = 587
    smtp_user = os.getenv("SMTP_USER", "alerts@rms.rw")
    smtp_pass = os.getenv("SMTP_PASSWORD", "")

    if st.button(
        "✉️ Send Email Alert Now", type="primary", use_container_width=True
    ):
      all_recipients = selected_recipients.copy()
      if custom_cc.strip():
        all_recipients.extend(
            [e.strip() for e in custom_cc.split(",") if e.strip()]
        )

      if not all_recipients:
        st.error("Please select at least one recipient email.")
      else:
        attachments = []
        for d_id in selected_doc_ids:
          fname, ftype, fdata = db.get_document_blob(d_id)
          if fdata:
            attachments.append((fname, fdata))
        if new_att_files:
          for nf in new_att_files:
            attachments.append((nf.name, nf.read()))

        if not smtp_user or not smtp_pass:
          # Simulation mode if credentials not in env
          db.log_action(
              f"Email alert logged for PO #{po} to {', '.join(all_recipients)}"
          )
          st.success(
              f"✅ Email alert queued & logged for {', '.join(all_recipients)}!"
          )
          time.sleep(1.0)
          st.rerun()
        else:
          with st.spinner("Sending Email..."):
            success, msg = send_email_smtp(
                smtp_host,
                smtp_port,
                smtp_user,
                smtp_pass,
                all_recipients,
                email_subject,
                email_body_html,
                attachments,
            )
            if success:
              db.log_action(
                  f"Email alert sent for PO #{po} to"
                  f" {', '.join(all_recipients)}"
              )
              st.success(msg)
            else:
              st.error(msg)

  with act_tab2:
    st.subheader("📁 Attached Documents & Repository")
    uploader_name = st.text_input(
        "Your Name / Officer Title*",
        value="Procurement Officer",
        key=f"uploader_name_{shipment_id}",
    )
    uploaded_files = st.file_uploader(
        "Upload Document(s) for this shipment line item",
        accept_multiple_files=True,
        key=f"tab_doc_uploader_{shipment_id}",
    )

    if st.button(
        "⬆️ Upload & Assign Doc Numbers",
        key=f"btn_upload_doc_{shipment_id}",
        type="primary",
    ):
      if not uploader_name.strip():
        st.error("Please enter your name to upload.")
      elif not uploaded_files:
        st.warning("Please select files first.")
      else:
        count = db.save_row_documents(
            shipment_id, uploaded_files, uploader_name.strip()
        )
        st.success(f"Successfully uploaded and numbered {count} document(s)!")
        st.rerun()

    st.divider()
    st.markdown("### 📋 Shipment Line Item Document Repository")
    docs_df = db.get_row_documents(shipment_id)
    if docs_df.empty:
      st.info("No documents uploaded for this shipment item yet.")
    else:
      for _, doc in docs_df.iterrows():
        c1, c2, c3, c4, c5 = st.columns([1, 4, 3, 2, 2])
        with c1:
          st.markdown(f"**Doc #{doc['Doc #']}**")
        with c2:
          st.markdown(f"📄 **{doc['File Name']}**")
        with c3:
          st.markdown(
              f"👤 {doc['Uploaded By']} | 🕒 {doc['Uploaded At']}"
          )
        with c4:
          fname, ftype, fdata = db.get_document_blob(doc["id"])
          if fdata:
            st.download_button(
                "💾 Download",
                data=fdata,
                file_name=fname,
                mime=ftype,
                key=f"dl_btn_{doc['id']}",
            )
        with c5:
          if st.button("🗑️ Delete File", key=f"del_doc_{doc['id']}"):
            db.delete_row_document(doc["id"], user_name=uploader_name)
            st.success("Document deleted.")
            st.rerun()

  with act_tab3:
    st.subheader("📜 Shipment Line Item Change History Trail")
    trail_df = db.get_row_change_trail(shipment_id)
    if trail_df.empty:
      st.info(
          "No cell modifications or document uploads recorded for this row"
          " yet."
      )
    else:
      st.dataframe(trail_df, use_container_width=True, hide_index=True)


# --- POPUP ADVANCED EDITOR ---
@st.dialog("✏️ Edit Shipment Details", width="large")
def edit_row_dialog(row_data, active_fy):
  s_id = row_data.get("id", "new")

  status_idx = (
      STATUS_OPTIONS.index(row_data.get("Current status", ""))
      if row_data.get("Current status") in STATUS_OPTIONS
      else 0
  )
  raw_mode = str(row_data.get("Shipping mode", "")).strip().upper()
  mode_idx = MODE_OPTIONS.index(raw_mode) if raw_mode in MODE_OPTIONS else 0
  inco_idx = (
      INCOTERM_OPTIONS.index(row_data.get("Incoterm", ""))
      if row_data.get("Incoterm") in INCOTERM_OPTIONS
      else 0
  )

  c1, c2, c3, c4 = st.columns(4)
  with c1:
    new_po = st.text_input("PO Number", value=row_data.get("PURCHASE ORDER", ""))
    new_status = st.selectbox(
        "Current status", STATUS_OPTIONS, index=status_idx
    )
    new_supplier = st.text_input(
        "Supplier / Manufacturer",
        value=row_data.get("SUPPLIER /MANUFACTURER", ""),
    )
  with c2:
    new_inv_no = st.text_input(
        "Invoice Number", value=row_data.get("Invoice number", "")
    )
    try:
      def_inv_val = float(row_data.get("Inv value", 0))
    except:
      def_inv_val = 0.0
    new_inv_val = st.number_input("Invoice Value", value=def_inv_val)
    new_curr = st.text_input(
        "Currency", value=row_data.get("Currency", "USD")
    )
  with c3:
    new_mode = st.selectbox("Shipping mode", MODE_OPTIONS, index=mode_idx)
    new_inco = st.selectbox("Incoterm", INCOTERM_OPTIONS, index=inco_idx)
    new_carrier = st.text_input("Carrier", value=row_data.get("Carrier", ""))
  with c4:
    new_bl = st.text_input(
        "BL / Airway Bill",
        value=row_data.get("BILL OF LADING OR AIRWAY BILL", ""),
    )
    new_pod = st.text_input("Port of Discharge (POD)", value=row_data.get("POD", ""))
    new_container = st.text_input(
        "Container Size", value=row_data.get("Container Size", "")
    )

  st.markdown("### 📅 Milestones & Tracking Dates")
  date_vals = {}
  dc1, dc2, dc3, dc4, dc5 = st.columns(5)
  for idx, col_name in enumerate(ALL_DATE_COLS[:5]):
    target_col = [dc1, dc2, dc3, dc4, dc5][idx]
    with target_col:
      raw_d = pd.to_datetime(row_data.get(col_name, ""), errors="coerce")
      def_d = raw_d.date() if pd.notna(raw_d) else None
      selected_d = st.date_input(
          col_name, value=def_d, key=f"d_{col_name}_{s_id}"
      )
      date_vals[col_name] = (
          selected_d.strftime("%Y-%m-%d") if selected_d else ""
      )

  st.divider()
  def_str = lambda val: "" if pd.isna(val) else str(val)
  st.markdown("**Cargo Description**")
  new_desc = st_quill(
      value=def_str(row_data.get("Cargo description")),
      html=True,
      key=f"q_desc_{s_id}",
  )

  new_mah = st.text_input(
      "MAH Letter Link", value=row_data.get("MAH letter link", "")
  )

  def _collect_row():
    updated_row = {
        "id": row_data.get("id"),
        "fiscal_year": active_fy,
        "PURCHASE ORDER": new_po,
        "SUPPLIER /MANUFACTURER": new_supplier,
        "Cargo description": new_desc,
        "Invoice number": new_inv_no,
        "Inv value": new_inv_val,
        "Currency": new_curr,
        "Quantity": row_data.get("Quantity", ""),
        "Shipping mode": new_mode.upper() if new_mode else "",
        "Incoterm": new_inco,
        "Container Size": new_container,
        "Carrier": new_carrier,
        "BILL OF LADING OR AIRWAY BILL": new_bl,
        "POD": new_pod,
        "Current status": new_status,
        "MAH letter link": new_mah,
    }
    updated_row.update(date_vals)
    return updated_row

  b_col1, b_col2, b_col3 = st.columns([1, 1, 1])
  with b_col1:
    if st.button(
        "💾 Save Shipment Details", type="primary", use_container_width=True
    ):
      updated_row = _collect_row()
      db.upsert_shipment(updated_row)
      st.success("✅ Data saved!")
      st.cache_data.clear()

      # Reset selection state so checkbox is unchecked when exiting
      if "shipments_aggrid_main" in st.session_state:
        del st.session_state["shipments_aggrid_main"]

      time.sleep(0.8)
      st.rerun()

  with b_col2:
    if st.button(
        "⚡ Take Action (Email & Files)",
        type="secondary",
        use_container_width=True,
    ):
      st.session_state["trigger_dialog"] = ("take_action", row_data)
      st.rerun()

  with b_col3:
    if st.button(
        "➕ Save as New Shipment",
        use_container_width=True,
        help="Keeps the original untouched and adds these values as a new row.",
    ):
      new_row = _collect_row()
      new_row.pop("id", None)
      if not str(new_row.get("PURCHASE ORDER", "")).strip():
        st.error("⚠️ PO Number is required to create a new shipment.")
      else:
        db.upsert_shipments_bulk([new_row], force_new=True)
        st.success("✅ Saved as a new shipment!")
        st.cache_data.clear()
        if "shipments_aggrid_main" in st.session_state:
          del st.session_state["shipments_aggrid_main"]
        time.sleep(0.8)
        st.rerun()


# --- POPUP: DUPLICATE ROW (quick way to add a similar shipment) ---
@st.dialog("📑 Duplicate Shipment", width="large")
def duplicate_row_dialog(row_data, active_fy):
  st.markdown(
      f"Create new shipment(s) from PO `{row_data.get('PURCHASE ORDER', 'N/A')}`"
      f" | Supplier: `{row_data.get('SUPPLIER /MANUFACTURER', 'N/A')}`"
  )
  d1, d2 = st.columns([1, 3])
  with d1:
    n_copies = st.number_input(
        "Number of copies", min_value=1, max_value=50, value=1, step=1
    )
  with d2:
    clear_tracking = st.checkbox(
        "Clear shipment-specific fields on the copy (BL/AWB, invoice number,"
        " actual dates, MAH link)",
        value=True,
    )
  st.caption(
      "The copies appear in the table straight away. Click any cell to edit it,"
      " or tick the row to open the Advanced Editor."
  )
  c1, c2 = st.columns(2)
  with c1:
    if st.button("Cancel", use_container_width=True, key="dup_cancel"):
      st.rerun()
  with c2:
    if st.button(
        "📑 Create Copies",
        type="primary",
        use_container_width=True,
        key="dup_go",
    ):
      made = db.duplicate_shipment(
          int(row_data["id"]),
          copies=int(n_copies),
          clear_tracking=clear_tracking,
      )
      st.cache_data.clear()
      if "shipments_aggrid_main" in st.session_state:
        del st.session_state["shipments_aggrid_main"]
      st.success(f"✅ {made} new shipment row(s) created.")
      time.sleep(0.8)
      st.rerun()


@st.dialog("🗑️ Confirm Deletion")
def delete_row_dialog(row_data):
  st.warning(
      f"Are you sure you want to delete PO: `{row_data.get('PURCHASE ORDER', 'N/A')}`"
      f" | AWB/BL: `{row_data.get('BILL OF LADING OR AIRWAY BILL', 'N/A')}`?"
  )
  user_name = st.text_input("Enter your name to confirm deletion:")
  c1, c2 = st.columns(2)
  with c1:
    if st.button("Cancel", use_container_width=True):
      st.rerun()
  with c2:
    if st.button("Delete Permanently", type="primary", use_container_width=True):
      if not user_name.strip():
        st.error("⚠️ Your name is required.")
      else:
        db.delete_shipment(row_data["id"], user_name.strip())
        st.rerun()


# --- FISCAL YEAR ---
fiscal_years = db.get_fiscal_years()
col_fy1, col_fy2 = st.columns([1, 1])
with col_fy1:
  selected_fy = st.selectbox(
      "📂 Active Tracking Year", fiscal_years, label_visibility="collapsed"
  )
with col_fy2:
  with st.expander("Initialize new tracking year"):
    new_fy = st.text_input("Enter Year (e.g., 2026-2027)")
    if st.button("Create Year"):
      if new_fy and new_fy not in fiscal_years:
        db.add_fiscal_year(new_fy.strip())
        st.rerun()

tab_dash, tab_emails, tab_upload, tab_logs, tab_deleted = st.tabs([
    "Shipment tracker",
    "RMS email directory",
    "Upload Excel data",
    "Edit trail",
    "Deleted records",
])


# ==========================================
# FRAGMENT ISOLATED HIGH-SPEED GRID (ZERO FLICKER)
# ==========================================
@st.fragment
def render_tracker_grid(df, selected_fy, f_mode, f_status, f_carrier, search_query):
  f_df = df.copy()

  # Issue 2 Fix: Normalize Shipping Mode case so SEA, Sea, sea combine into 'SEA'
  if "Shipping mode" in f_df.columns:
    f_df["Shipping mode"] = (
        f_df["Shipping mode"].fillna("").astype(str).str.strip().str.upper()
    )

  # Search Filter
  if search_query:
    q = search_query.strip().lower()
    str_cols = f_df.select_dtypes(include=["object", "string"]).columns
    # One vectorised pass: join every text column per row, then a single search.
    haystack = pd.Series("", index=f_df.index)
    for _c in str_cols:
      haystack = haystack + "\x1f" + f_df[_c].fillna("").astype(str)
    f_df = f_df[haystack.str.lower().str.contains(q, regex=False)]

  # Dropdown Filters
  if f_mode:
    norm_f_mode = [m.upper() for m in f_mode]
    f_df = f_df[f_df["Shipping mode"].isin(norm_f_mode)]
  if f_status:
    f_df = f_df[f_df["Current status"].isin(f_status)]
  if f_carrier:
    f_df = f_df[f_df["Carrier"].isin(f_carrier)]

  # Click-to-drill status filter (set by the status buttons below the toolbar)
  f_df_all = f_df
  _sel = st.session_state.get("quick_status", "All")
  if "Current status" in f_df_all.columns:
    _status_txt = f_df_all["Current status"].fillna("").astype(str).str.strip().replace("", "No status")
    if _sel != "All" and _sel not in set(_status_txt):
      _sel = "All"
    if _sel != "All":
      f_df = f_df_all[_status_txt == _sel]

  # Issue 1 Fix: Move "Received in Warehouse" shipments to the VERY BOTTOM of the list
  if not f_df.empty and "Current status" in f_df.columns:
    f_df["_is_received"] = (
        f_df["Current status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .apply(lambda s: 1 if ("received" in s or "warehouse" in s) else 0)
    )
    f_df = f_df.sort_values(
        by=["_is_received", "id"], ascending=[True, False]
    ).drop(columns=["_is_received"])

  # --- ACTIONS & EXPORT ROW WITH UNIFORM TOOLBAR STYLING ---
  st.markdown("<div class='toolbar-container'>", unsafe_allow_html=True)
  col_export, col_custom_col, col_config, col_spacer = st.columns(
      [3, 3, 3, 3]
  )
  with col_export:
    # EXPORT TO EXCEL BUTTON
    export_df = f_df.drop(columns=["id", "fiscal_year"], errors="ignore")
    excel_data = build_export_bytes(
        export_df,
        (
            selected_fy,
            _DBV,
            tuple(f_mode),
            tuple(f_status),
            tuple(f_carrier),
            search_query or "",
        ),
    )

    st.download_button(
        label="Export to Excel (.xlsx)",
        data=excel_data,
        file_name=f"RMS_Shipment_Tracker_{selected_fy}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

  with col_custom_col:
    with st.popover("Add custom column", use_container_width=True):
      new_col_name = st.text_input("Column Name")
      if st.button("Add to Grid") and new_col_name:
        ok, msg = db.add_custom_column(new_col_name.strip())
        if ok:
          st.success(msg)
        else:
          st.error(msg)
        st.rerun()

  with col_config:
    if st.button("Settings (colors & dropdowns)", use_container_width=True):
      manage_dropdowns_dialog()

  with col_spacer:
    with st.popover("Backup database", use_container_width=True):
      st.caption(
          "A full copy of the database (all fiscal years, edit trail, emails)."
          " Keep it somewhere safe off the server."
      )
      if st.button("Prepare backup file", key="prep_backup"):
        try:
          _bk_path = db.create_backup("manual")
          with open(_bk_path, "rb") as _fh:
            st.session_state["_bk_bytes"] = _fh.read()
          st.session_state["_bk_name"] = os.path.basename(_bk_path)
          os.remove(_bk_path)  # the download is the copy; don't pile up files
        except Exception as _exc:
          st.error(f"Could not create the backup: {_exc}")
      if st.session_state.get("_bk_bytes"):
        st.download_button(
            f"Download {st.session_state['_bk_name']}",
            data=st.session_state["_bk_bytes"],
            file_name=st.session_state["_bk_name"],
            mime="application/octet-stream",
            use_container_width=True,
        )

  st.markdown("</div>", unsafe_allow_html=True)

  # --- STATUS PIPELINE: click a status to drill into its shipments ---
  _stat_col = "Current status"
  if _stat_col in f_df_all.columns and len(f_df_all):
    _txt = f_df_all[_stat_col].fillna("").astype(str).str.strip().replace("", "No status")
    _counts = _txt.value_counts()
    _known = [o["label"] for o in db_status_opts]
    _order = [k for k in _known if k in _counts.index] + [k for k in _counts.index if k not in _known]
    _colors = {o["label"]: o["color"] for o in db_status_opts}
    _total = int(_counts.sum())
    _segs = "".join(
        f"<div class='pipe-seg' title='{k}: {_counts[k]}' style='flex:{_counts[k]}; background:{_colors.get(k, '#9aa5b1')};"
        f" opacity:{1 if _sel in ('All', k) else 0.22};'></div>"
        for k in _order
    )
    _open = f_df_all[_txt != "Received in Warehouse"]
    _warn = int((_open.get("Action warning", pd.Series(dtype=str)).fillna("").astype(str).str.strip() != "").sum())
    _late = int((pd.to_numeric(_open.get("delays", pd.Series(dtype=float)), errors="coerce") > 0).sum())
    _chips = (
        f"<span class='chip {'warn' if _warn else 'good'}'>{_warn} need action</span>"
        f"<span class='chip {'bad' if _late else 'good'}'>{_late} overdue</span>"
    )
    st.markdown(
        "<div class='pipe-wrap'><div class='pipe-head'><div class='pipe-title'>Shipment pipeline"
        " <span style='font-weight:500;color:#64748b'>&middot; click a status to see its shipments</span></div>"
        f"<div>{_chips}</div></div><div class='pipe-bar'>{_segs}</div></div>",
        unsafe_allow_html=True,
    )

    def _pick_status(label):
      st.session_state["quick_status"] = label

    _btn_cols = st.columns(len(_order) + 1)
    with _btn_cols[0]:
      st.button(f"All  {_total}", key="qs_all", type="primary" if _sel == "All" else "secondary",
                on_click=_pick_status, args=("All",), use_container_width=True)
    _dot_css = ""
    for _i, _k in enumerate(_order):
      with _btn_cols[_i + 1]:
        st.button(f"{_k}  {_counts[_k]}", key=f"qs_{_i}", type="primary" if _sel == _k else "secondary",
                  on_click=_pick_status, args=(_k,), use_container_width=True)
      _dot_css += (
          f".st-key-qs_{_i} button p::before {{content:''; display:inline-block; width:9px; height:9px; border-radius:50%;"
          f" background:{_colors.get(_k, '#9aa5b1')}; margin-right:8px; box-shadow:0 0 0 2px rgba(255,255,255,.85);}}"
      )
    st.markdown(
        "<style>[class*='st-key-qs_'] button p {white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}"
        + _dot_css + "</style>",
        unsafe_allow_html=True,
    )

  # --- PRIMARY KPI MATH (CARGO VALUE WITH EUR STATIC RATE) ---
  if f_df.empty:
    total_val_usd = 0.0
  else:
    _vals = pd.to_numeric(f_df["Inv value"], errors="coerce").fillna(0.0)
    _is_eur = (
        f_df["Currency"].fillna("").astype(str).str.strip().str.upper()
    ).isin(["EURO", "EUR"])
    # Static exchange rate: 1 EUR = 1.08 USD
    total_val_usd = float(_vals.where(~_is_eur, _vals * 1.08).sum())

  _modes = (
      f_df["Shipping mode"].fillna("").astype(str).str.upper()
      if "Shipping mode" in f_df.columns
      else pd.Series(dtype=str)
  )
  n_sea = int((_modes == "SEA").sum())
  n_air = int((_modes == "AIR").sum())
  n_road = int((_modes == "ROAD").sum())

  warn_mask = f_df.get("Action warning", pd.Series()).fillna("").str.strip() != ""
  warn_count = int(warn_mask.sum()) if not f_df.empty else 0

  # Issue 3 Fix: Specific KPI Calculations
  delays_series = pd.to_numeric(f_df.get("delays", None), errors="coerce")
  avg_delays = (
      round(delays_series.mean(), 1)
      if not delays_series.dropna().empty
      else 0.0
  )

  days_clear_series = pd.to_numeric(
      f_df.get("DAYS TO CLEAR THE SHIPMENT", None), errors="coerce"
  )
  avg_days_clear = (
      round(days_clear_series.mean(), 1)
      if not days_clear_series.dropna().empty
      else 0.0
  )

  met1, met2, met3, met4, met5, met6, met7 = st.columns(7)
  with met1:
    st.markdown(
        create_kpi_card(
            "Total Shipments", len(f_df), "#f8f9fa", "#adb5bd", "#2c3e50"
        ),
        unsafe_allow_html=True,
    )
  with met2:
    st.markdown(
        create_kpi_card(
            "Sea freight",
            n_sea,
            "#e8f5e9",
            "#28a745",
            "#155724",
        ),
        unsafe_allow_html=True,
    )
  with met3:
    st.markdown(
        create_kpi_card(
            "Air freight",
            n_air,
            "#e3f2fd",
            "#007bff",
            "#004085",
        ),
        unsafe_allow_html=True,
    )
  with met4:
    st.markdown(
        create_kpi_card(
            "Road freight",
            n_road,
            "#fff3cd",
            "#ffc107",
            "#856404",
        ),
        unsafe_allow_html=True,
    )
  with met5:
    st.markdown(
        create_kpi_card(
            "Cargo Value (USD)",
            format_currency(total_val_usd),
            "#f3e5f5",
            "#6f42c1",
            "#4a148c",
        ),
        unsafe_allow_html=True,
    )
  with met6:
    st.markdown(
        create_kpi_card(
            "Avg Delay (days, + = late)",
            avg_delays,
            "#ffebee" if avg_delays > 0 else "#e8f5e9",
            "#dc3545" if avg_delays > 0 else "#28a745",
            "#721c24" if avg_delays > 0 else "#155724",
        ),
        unsafe_allow_html=True,
    )
  with met7:
    st.markdown(
        create_kpi_card(
            "Avg Days to Clear",
            avg_days_clear,
            "#e3f2fd",
            "#007bff",
            "#004085",
        ),
        unsafe_allow_html=True,
    )

  # --- ADVANCED EXPANDER ---
  with st.expander(
      "Analytics and performance charts",
      expanded=True,
  ):
    tab_carrier, tab_status, tab_clearance = st.tabs([
        "Carrier distribution",
        "Status breakdown",
        "Clearance performance",
    ])

    _UNSET = {"", "tba", "pending", "to be advised", "no", "na", "n/a", "none", "unknown", "-"}
    _MODE_COLORS = {"SEA": "#1a63d6", "AIR": "#0d8a9d", "ROAD": "#f59e0b"}

    def _style(fig, height):
      fig.update_layout(
          height=height, showlegend=False, margin=dict(l=8, r=36, t=8, b=8),
          font=dict(family="Plus Jakarta Sans, Segoe UI, sans-serif", size=13, color="#334155"),
          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
      )
      return fig

    def _show(fig):
      try:
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
      except TypeError:  # older Streamlit
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with tab_carrier:
      _car = f_df.get("Carrier", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
      _key = _car.str.lower().str.replace(r"\s+", " ", regex=True).replace({"et": "ethiopian", "ethiopian airlines": "ethiopian"})
      _known = ~_key.isin(_UNSET)
      if _known.any():
        _g = pd.DataFrame({"key": _key[_known], "raw": _car[_known]})
        _tot = _g.groupby("key").size().sort_values(ascending=False)
        _lab = _g.groupby("key")["raw"].agg(lambda s: s.value_counts().index[0])
        if "ethiopian" in _lab.index:
          _lab["ethiopian"] = "Ethiopian Airlines"
        _top = _tot.head(12)
        _rows = [(_lab[k], int(n), "#1a63d6") for k, n in _top.items()]
        if len(_tot) > 12:
          _rows.append((f"Other ({len(_tot) - 12} carriers)", int(_tot.iloc[12:].sum()), "#94a3b8"))
        _names = [(n if len(n) <= 34 else n[:33] + "…") for n, _, _ in _rows]
        fig_c = go.Figure(go.Bar(
            x=[n for _, n, _ in _rows], y=_names, orientation="h",
            marker=dict(color=[c for _, _, c in _rows], cornerradius=6),
            text=[n for _, n, _ in _rows], textposition="outside", cliponaxis=False,
            customdata=[r[0] for r in _rows], hovertemplate="%{customdata}: %{x} shipments<extra></extra>",
        ))
        fig_c.update_yaxes(autorange="reversed", showgrid=False, ticks="", automargin=True)
        fig_c.update_xaxes(showgrid=True, gridcolor="#eef2f7", zeroline=False, title="Shipments")
        _show(_style(fig_c, 80 + 36 * len(_rows)))
        if (~_known).any():
          st.caption(f"{int((~_known).sum())} shipments have no carrier assigned yet (TBA / pending) and are not shown above.")
      else:
        st.info("No carrier recorded for the shipments shown.")

    with tab_status:
      _st = f_df.get("Current status", pd.Series(dtype=str)).fillna("").astype(str).str.strip().replace("", "No status")
      if len(_st):
        _cnt = _st.value_counts()
        _ord = [o["label"] for o in db_status_opts if o["label"] in _cnt.index] + [k for k in _cnt.index if k not in {o["label"] for o in db_status_opts}]
        _col = {o["label"]: o["color"] for o in db_status_opts}
        sc1, sc2 = st.columns([1, 1.5])
        with sc1:
          fig_s = go.Figure(go.Pie(
              labels=_ord, values=[int(_cnt[k]) for k in _ord], hole=0.62, sort=False,
              marker=dict(colors=[_col.get(k, "#9aa5b1") for k in _ord], line=dict(color="#ffffff", width=3)),
              textinfo="percent", textfont=dict(color="#ffffff", size=13),
              hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
          ))
          fig_s.add_annotation(text=f"<b>{int(_cnt.sum())}</b><br><span style='font-size:12px;color:#64748b'>shipments</span>",
                               showarrow=False, font=dict(size=30, color="#0f172a"))
          _show(_style(fig_s, 380))
        with sc2:
          _md = f_df.get("Shipping mode", pd.Series(dtype=str)).fillna("").astype(str).str.strip().str.upper().replace("", "NOT SET")
          _x = pd.DataFrame({"mode": _md.values, "status": _st.values}).groupby(["mode", "status"]).size().reset_index(name="n")
          fig_m = go.Figure()
          for k in _ord:
            _d = _x[_x["status"] == k]
            if len(_d):
              fig_m.add_bar(y=_d["mode"], x=_d["n"], name=k, orientation="h", marker=dict(color=_col.get(k, "#9aa5b1")),
                            hovertemplate=k + ": %{x}<extra></extra>")
          fig_m.update_layout(barmode="stack")
          fig_m.update_yaxes(autorange="reversed", showgrid=False, ticks="")
          fig_m.update_xaxes(showgrid=True, gridcolor="#eef2f7", zeroline=False, title="Shipments by transport mode and status")
          _style(fig_m, 380)
          fig_m.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.28, x=0))
          _show(fig_m)

    with tab_clearance:
      df_cl = f_df.dropna(subset=["DAYS TO CLEAR THE SHIPMENT"]).copy()
      df_cl["DAYS TO CLEAR THE SHIPMENT"] = pd.to_numeric(df_cl["DAYS TO CLEAR THE SHIPMENT"], errors="coerce")
      df_cl = df_cl.dropna(subset=["DAYS TO CLEAR THE SHIPMENT"])
      if not df_cl.empty:
        df_cl["mode"] = df_cl["Shipping mode"].fillna("").astype(str).str.strip().str.upper().replace("", "NOT SET")
        _s = df_cl.groupby("mode")["DAYS TO CLEAR THE SHIPMENT"].agg(median="median", mean="mean", n="size").reset_index()
        _s = _s.sort_values("median", ascending=False)
        fig_cl = go.Figure(go.Bar(
            x=_s["median"], y=_s["mode"], orientation="h",
            marker=dict(color=[_MODE_COLORS.get(m, "#94a3b8") for m in _s["mode"]], cornerradius=6),
            text=[f"{m:.0f} days  ·  {n} shipments" for m, n in zip(_s["median"], _s["n"])], textposition="outside", cliponaxis=False,
            customdata=_s[["mean", "n"]].values,
            hovertemplate="%{y}: median %{x:.0f} days, average %{customdata[0]:.1f} days (%{customdata[1]} shipments)<extra></extra>",
        ))
        fig_cl.update_yaxes(autorange="reversed", showgrid=False, ticks="")
        fig_cl.update_xaxes(showgrid=True, gridcolor="#eef2f7", zeroline=False, range=[0, float(_s["median"].max()) * 1.45 + 1], title="Median days to clear (document sharing to arrival at RMS)")
        _show(_style(fig_cl, 90 + 70 * len(_s)))
      else:
        st.info("No clearance times to show yet for the shipments selected.")

  # --- AG-GRID TRACKER ---
  action_container_top = st.empty()

  display_df = f_df.drop(columns=["fiscal_year"], errors="ignore").copy()

  for c in ALL_DATE_COLS:
    if c in display_df:
      display_df[c] = (
          pd.to_datetime(display_df[c], errors="coerce")
          .dt.strftime("%Y-%m-%d")
          .fillna("")
      )

  gb = GridOptionsBuilder.from_dataframe(display_df)

  gb.configure_default_column(
      wrapText=True,
      autoHeight=True,
      editable=True,
      minWidth=140,
      tooltipValueGetter=JsCode(
          "function(params) { return params.value ? String(params.value) :"
          " ''; }"
      ),
  )

  for col in display_df.columns:
    gb.configure_column(col, headerTooltip=col)

  dynamic_rowstyle = JsCode("""
    function(params) {
        if (!params || !params.data) return null;
        const rowStatus = params.data['Current status'] ? String(params.data['Current status']).trim().toLowerCase() : '';
        if (rowStatus === 'received in warehouse' || (rowStatus.includes('received') && rowStatus.includes('warehouse'))) {
            return {'backgroundColor': '#e8f5e9', 'color': '#155724', 'fontWeight': 'bold'};
        }
        return null;
    }
    """)

  dynamic_cellstyle = JsCode(f"""
    function(params) {{
        if (!params || !params.colDef || !params.colDef.field) return null;
        
        const focusCols = {json.dumps(FOCUS_COLUMNS)};
        const statusMap = {json.dumps(status_color_map)};
        const modeMap = {json.dumps(mode_color_map)};
        const field = params.colDef.field;
        
        const rowStatus = params.data && params.data['Current status'] ? String(params.data['Current status']).trim().toLowerCase() : '';
        if (rowStatus === 'received in warehouse' || (rowStatus.includes('received') && rowStatus.includes('warehouse'))) {{
            return {{'backgroundColor': '#e8f5e9', 'color': '#155724', 'fontWeight': 'bold'}};
        }}

        if (field === 'Action warning' && params.value && String(params.value).trim() !== '') {{
            return {{'backgroundColor': '#ffebee', 'color': '#721c24', 'fontWeight': 'bold'}};
        }}
        
        if (field === 'Current status' && statusMap[params.value]) {{
            let bgColor = statusMap[params.value];
            let hexcolor = bgColor.replace("#", "");
            if (hexcolor.length === 3) hexcolor = hexcolor.split('').map(function(hex){{return hex+hex}}).join('');
            var r = parseInt(hexcolor.substr(0,2),16);
            var g = parseInt(hexcolor.substr(2,2),16);
            var b = parseInt(hexcolor.substr(4,2),16);
            var yiq = ((r*299)+(g*587)+(b*114))/1000;
            let tColor = (yiq >= 128) ? 'black' : 'white';
            return {{'backgroundColor': bgColor, 'color': tColor, 'fontWeight': 'bold'}};
        }}

        if (field === 'Shipping mode' && modeMap[params.value]) {{
            return {{'backgroundColor': modeMap[params.value], 'color': 'black', 'fontWeight': 'bold'}};
        }}

        const portArr = params.data ? params.data['Actual arrival at port'] : null;
        if (portArr && String(portArr).trim() !== '' && String(portArr).trim().toLowerCase() !== 'nat' && String(portArr).trim().toLowerCase() !== 'tba') {{
            if (field === 'Actual arrival at port') {{
                return {{'backgroundColor': '#cce5ff', 'color': '#004085', 'fontWeight': 'bold'}};
            }}
        }}

        if (focusCols.includes(field)) {{
            return {{'backgroundColor': '#fff3cd', 'color': '#856404', 'fontWeight': '500'}};
        }}
        
        return null;
    }}
    """)

  gb.configure_grid_options(
      singleClickEdit=True,
      domLayout="normal",
      getRowStyle=dynamic_rowstyle,
      enableBrowserTooltips=True,
      animateRows=False,
      suppressAnimationFrame=True,
  )
  gb.configure_selection(selection_mode="single", use_checkbox=True)

  custom_date_editor = JsCode("""
    class DatePickerEditor {
        init(params) {
            this.eInput = document.createElement('input');
            this.eInput.type = 'date';
            
            let val = params.value ? String(params.value).trim() : '';
            if (val) {
                let match = val.match(/^(\\d{4}-\\d{2}-\\d{2})/);
                if (match) {
                    this.eInput.value = match[1];
                } else {
                    let d = new Date(val);
                    if (!isNaN(d.getTime())) {
                        let yyyy = d.getFullYear();
                        let mm = String(d.getMonth() + 1).padStart(2, '0');
                        let dd = String(d.getDate()).padStart(2, '0');
                        this.eInput.value = yyyy + '-' + mm + '-' + dd;
                    } else {
                        this.eInput.value = '';
                    }
                }
            } else {
                this.eInput.value = '';
            }

            this.eInput.style.width = '100%';
            this.eInput.style.height = '100%';
            this.eInput.style.border = 'none';
            this.eInput.style.padding = '4px';
            this.eInput.style.boxSizing = 'border-box';
        }
        getGui() {
            return this.eInput;
        }
        afterGuiAttached() {
            this.eInput.focus();
            if (this.eInput.showPicker) {
                try { this.eInput.showPicker(); } catch(e) {}
            }
        }
        getValue() {
            return this.eInput.value;
        }
        destroy() {}
        isPopup() {
            return false;
        }
    }
    """)

  if "id" in display_df.columns:
    gb.configure_column("id", hide=True)

  for col in ALL_DATE_COLS:
    if col in display_df.columns:
      gb.configure_column(
          col,
          editable=True,
          width=160,
          type=["dateColumnFilter", "dateColumn"],
          cellEditor=custom_date_editor,
          cellStyle=dynamic_cellstyle,
          headerTooltip=col,
      )

  if "PURCHASE ORDER" in display_df.columns:
    gb.configure_column(
        "PURCHASE ORDER", width=140, pinned="left", checkboxSelection=True
    )
  if "Cargo description" in display_df.columns:
    gb.configure_column(
        "Cargo description", width=380, minWidth=350, wrapText=True
    )
  if "Current status" in display_df.columns:
    gb.configure_column(
        "Current status",
        editable=True,
        cellEditor="agSelectCellEditor",
        cellEditorParams={"values": STATUS_OPTIONS},
        width=160,
        pinned="right",
        cellStyle=dynamic_cellstyle,
    )
  if "Shipping mode" in display_df.columns:
    gb.configure_column(
        "Shipping mode",
        editable=True,
        cellEditor="agSelectCellEditor",
        cellEditorParams={"values": MODE_OPTIONS},
        width=130,
        cellStyle=dynamic_cellstyle,
    )
  if "Incoterm" in display_df.columns:
    gb.configure_column(
        "Incoterm",
        editable=True,
        cellEditor="agSelectCellEditor",
        cellEditorParams={"values": INCOTERM_OPTIONS},
        width=110,
    )

  # DERIVED CALCULATED COLUMNS (REAL-TIME INSTANT GETTERS)
  js_days_to_clear_getter = JsCode("""
  function(params) {
      if (!params || !params.data) return null;
      let rmsStr = params.data['Actual arrival at RMS'];
      let docStr = params.data['DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS'];
      
      if (!rmsStr || !docStr || String(rmsStr).trim() === '' || String(docStr).trim() === '') return params.data['DAYS TO CLEAR THE SHIPMENT'] || null;
      let rmsDate = new Date(String(rmsStr).trim());
      let docDate = new Date(String(docStr).trim());
      if (isNaN(rmsDate.getTime()) || isNaN(docDate.getTime())) return params.data['DAYS TO CLEAR THE SHIPMENT'] || null;
      
      rmsDate.setHours(0,0,0,0);
      docDate.setHours(0,0,0,0);
      
      let diffTime = rmsDate.getTime() - docDate.getTime();
      let diffDays = Math.round(diffTime / (1000 * 3600 * 24));
      return diffDays;
  }
  """)

  js_delays_getter = JsCode("""
  function(params) {
      if (!params || !params.data) return null;
      let rmsStr = params.data['Actual arrival at RMS'];
      let manualDeadline = params.data['Deadline to clear the shipment'];
      let docStr = params.data['DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS'];
      
      let deadlineDate = null;
      if (manualDeadline && String(manualDeadline).trim() !== '') {
          deadlineDate = new Date(String(manualDeadline).trim());
      } else if (docStr && String(docStr).trim() !== '') {
          let d = new Date(String(docStr).trim());
          if (!isNaN(d.getTime())) {
              deadlineDate = new Date(d);
              deadlineDate.setDate(deadlineDate.getDate() + 2);
          }
      }
      
      if (!deadlineDate || isNaN(deadlineDate.getTime())) return params.data['delays'] || null;
      
      let rmsDate = null;
      if (rmsStr && String(rmsStr).trim() !== '') {
          rmsDate = new Date(String(rmsStr).trim());
      } else {
          rmsDate = new Date(); // Today
      }
      
      if (!rmsDate || isNaN(rmsDate.getTime())) return params.data['delays'] || null;
      
      deadlineDate.setHours(0,0,0,0);
      rmsDate.setHours(0,0,0,0);
      
      let diffTime = rmsDate.getTime() - deadlineDate.getTime();
      let diffDays = Math.round(diffTime / (1000 * 3600 * 24));
      return diffDays;
  }
  """)

  js_action_warning_getter = JsCode("""
  function(params) {
      if (!params || !params.data) return '';
      let today = new Date();
      today.setHours(0,0,0,0);
      
      let etaStr = params.data['Cargo ETA RMS'];
      let portStr = params.data['Actual arrival at port'];
      let docStr = params.data['DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS'];
      let rmsStr = params.data['Actual arrival at RMS'];
      
      let warnings = [];
      
      if (etaStr && String(etaStr).trim() !== '') {
          let etaD = new Date(String(etaStr).trim());
          if (!isNaN(etaD.getTime())) {
              etaD.setHours(0,0,0,0);
              let diffDays = Math.round((etaD.getTime() - today.getTime()) / (1000 * 3600 * 24));
              if (diffDays <= 3) warnings.push("ACTION: ETA RMS due/reached");
          }
      }
      
      if (portStr && String(portStr).trim() !== '') {
          let portD = new Date(String(portStr).trim());
          if (!isNaN(portD.getTime())) {
              portD.setHours(0,0,0,0);
              let diffDays = Math.round((portD.getTime() - today.getTime()) / (1000 * 3600 * 24));
              if (diffDays <= 3) warnings.push("ACTION: Port arrival due/reached");
              if (!docStr || String(docStr).trim() === '') warnings.push("ACTION: Share shipping documents");
          }
      }
      
      if (docStr && String(docStr).trim() !== '' && (!rmsStr || String(rmsStr).trim() === '')) {
          let docD = new Date(String(docStr).trim());
          if (!isNaN(docD.getTime())) {
              docD.setHours(0,0,0,0);
              let daysPassed = Math.round((today.getTime() - docD.getTime()) / (1000 * 3600 * 24));
              if (daysPassed >= 2) warnings.push("WARNING: Shipment clearance delay - push clearing agent to expedite");
          }
      }
      
      return warnings.join("; ");
  }
  """)

  if "DAYS TO CLEAR THE SHIPMENT" in display_df.columns:
    gb.configure_column(
        "DAYS TO CLEAR THE SHIPMENT",
        editable=False,
        width=130,
        valueGetter=js_days_to_clear_getter,
    )
  if "delays" in display_df.columns:
    gb.configure_column(
        "delays",
        editable=False,
        width=100,
        cellStyle=dynamic_cellstyle,
        valueGetter=js_delays_getter,
    )
  if "Action warning" in display_df.columns:
    gb.configure_column(
        "Action warning",
        editable=False,
        width=280,
        cellStyle=dynamic_cellstyle,
        valueGetter=js_action_warning_getter,
    )

  custom_css = {
      ".ag-header": {
          "background-color": "#117a8b !important",
          "border-bottom": "2px solid #0056b3 !important",
      },
      ".ag-header-row": {
          "background-color": "#117a8b !important",
          "color": "white !important",
      },
      ".ag-header-cell-label": {
          "color": "#ffffff !important",
          "font-weight": "bold !important",
      },
  }

  # Fixed-height grid: the header row stays visible while rows scroll inside it
  # (and only the visible rows are drawn, which keeps big tables fast).
  grid_h = int(min(640, max(240, 80 + 56 * len(display_df))))

  grid_response = AgGrid(
      display_df,
      height=grid_h,
      gridOptions=gb.build(),
      update_mode=GridUpdateMode.VALUE_CHANGED
      | GridUpdateMode.SELECTION_CHANGED,
      data_return_mode=DataReturnMode.AS_INPUT,
      fit_columns_on_grid_load=False,
      theme="streamlit",
      allow_unsafe_jscode=True,
      custom_css=custom_css,
      reload_data=False,
      key="shipments_aggrid_main",
  )

  # --- INSTANT AUTO-SAVE & BACKGROUND COMPUTATION ENGINE ---
  edited_df = grid_response["data"]
  if (
      isinstance(edited_df, pd.DataFrame)
      and not edited_df.empty
      and "id" in edited_df.columns
  ):
    skip_cols = {"_selectedRowNodeInfo", "id", "fiscal_year"}
    cmp_cols = [
        c for c in edited_df.columns if c not in skip_cols and c in df.columns
    ]

    def _as_text(frame):
      txt = frame.astype(object).where(frame.notna(), "").astype(str)
      return txt.apply(lambda col: col.str.strip())

    new_ids = pd.to_numeric(edited_df["id"], errors="coerce")
    valid = new_ids.notna()
    changed_rows = []
    if valid.any():
      edited_valid = edited_df.loc[valid].copy()
      edited_valid.index = new_ids[valid].astype(int).values
      edited_valid = edited_valid[~edited_valid.index.duplicated()]

      old_part = df.drop_duplicates("id").set_index("id")
      common = edited_valid.index.intersection(old_part.index)
      if len(common) and cmp_cols:
        new_txt = _as_text(edited_valid.loc[common, cmp_cols])
        for _c in cmp_cols:
          if _c in ALL_DATE_COLS:
            new_txt[_c] = new_txt[_c].map(db.norm_date_str)
        old_txt = _as_text(old_part.loc[common, cmp_cols])
        diff_mask = (new_txt != old_txt).any(axis=1)
        changed_ids = list(common[diff_mask.values])

        for rid in changed_ids:
          new_row = edited_valid.loc[rid]
          row_dict = {"id": int(rid), "fiscal_year": selected_fy}
          for col_name in edited_df.columns:
            if col_name in skip_cols:
              continue
            raw = new_row.get(col_name, "")
            new_val = str(raw).strip() if pd.notna(raw) else ""
            if col_name in ALL_DATE_COLS:
              new_val = db.norm_date_str(new_val)
            row_dict[col_name] = new_val
          changed_rows.append(row_dict)

    if changed_rows:
      db.upsert_shipments_bulk(changed_rows, changed_by="Officer")
      st.cache_data.clear()
      st.rerun()

  # --- ACTION BAR & AUTOMATIC DIALOG TRIGGER ON CHECKBOX SELECTION ---
  selected_rows = grid_response.get("selected_rows")
  selected_data = None
  if isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
    selected_data = selected_rows.iloc[0].to_dict()
  elif isinstance(selected_rows, list) and len(selected_rows) > 0:
    selected_data = selected_rows[0]

  queued_dlg = st.session_state.pop("trigger_dialog", None)
  bar_dialog_opened = False

  if selected_data:
    row_id_sel = selected_data.get("id")
    if pd.isna(row_id_sel):
      po_sel = str(selected_data.get("PURCHASE ORDER", "")).strip()
      match_orig = df[df["PURCHASE ORDER"] == po_sel]
      if not match_orig.empty:
        selected_data["id"] = match_orig.iloc[0]["id"]
    else:
      selected_data["id"] = int(row_id_sel)

    with action_container_top.container():
      st.markdown("### 🛠️ Selected Shipment Actions")
      b1, b2, b3, b4, _ = st.columns([2.5, 3.2, 2, 2.2, 2.1])
      with b1:
        if st.button(
            "⚡ Take Action (Email & Files)",
            type="primary",
            use_container_width=True,
        ):
          bar_dialog_opened = True
          take_action_dialog(selected_data)
      with b2:
        if st.button(
            "✏️ Edit Selected Row (Advanced Editor)", use_container_width=True
        ):
          bar_dialog_opened = True
          edit_row_dialog(selected_data, selected_fy)
      with b3:
        if st.button("🗑️ Delete Row", use_container_width=True):
          bar_dialog_opened = True
          delete_row_dialog(selected_data)
      with b4:
        if st.button("📑 Duplicate Row", use_container_width=True):
          bar_dialog_opened = True
          duplicate_row_dialog(selected_data, selected_fy)

    if queued_dlg:
      dlg_type, dlg_data = queued_dlg
      if dlg_type == "take_action":
        take_action_dialog(dlg_data)
      elif dlg_type == "edit":
        edit_row_dialog(dlg_data, selected_fy)
      elif dlg_type == "delete":
        delete_row_dialog(dlg_data)
    elif not bar_dialog_opened:
      # Automatically open Advanced Editor on checkbox selection
      edit_row_dialog(selected_data, selected_fy)
  elif queued_dlg:
    dlg_type, dlg_data = queued_dlg
    if dlg_type == "take_action":
      take_action_dialog(dlg_data)
    elif dlg_type == "edit":
      edit_row_dialog(dlg_data, selected_fy)
    elif dlg_type == "delete":
      delete_row_dialog(dlg_data)


# --- TRACKER TAB RENDER ---
with tab_dash:
  df = cached_load(selected_fy, _DBV)
  # --- SEARCH & FILTERS ---
  f1, f2, f3, f4 = st.columns([2.5, 1, 1, 1])
  with f1:
    search_query = st.text_input(
        "Search shipments",
        placeholder="Type PO#, Supplier, Cargo, BL/AWB, Carrier, POD...",
    )
  with f2:
    f_mode = st.multiselect("Shipping Mode", options=MODE_OPTIONS)
  with f3:
    f_status = st.multiselect("Current Status", options=STATUS_OPTIONS)
  with f4:
    f_carrier = st.multiselect(
        "Carrier",
        options=sorted(df["Carrier"].dropna().unique())
        if not df.empty
        else [],
    )

  render_tracker_grid(df, selected_fy, f_mode, f_status, f_carrier, search_query)

# --- EMAIL DIRECTORY TAB ---
with tab_emails:
  st.subheader("📧 RMS Team & Clearing Email Directory Manager")
  st.markdown(
      "Register RMS department emails and clearing agents that will appear in"
      " the **Take Action** modal."
  )

  ec_col1, ec_col2 = st.columns([1, 2])
  with ec_col1:
    st.markdown("### ➕ Add New Email Entry")
    with st.form("add_email_form_shipment", clear_on_submit=True):
      e_name = st.text_input("Name / Officer Title*")
      e_email = st.text_input("Email Address*")
      e_dept = st.text_input("Department / Entity", value="Logistics")
      e_role = st.text_input("Role / Designation", value="Clearing Agent")

      if st.form_submit_button("Register Email", type="primary"):
        if not e_name or not e_email:
          st.error("Name and Email are required.")
        else:
          ok, msg = db.add_rms_email(e_name, e_email, e_dept, e_role)
          if ok:
            st.success(msg)
          else:
            st.error(msg)
          st.rerun()

  with ec_col2:
    st.markdown("### 📋 Registered Directory")
    rms_emails_df = db.get_rms_emails()
    if rms_emails_df.empty:
      st.info("No emails registered yet.")
    else:
      for _, erow in rms_emails_df.iterrows():
        ec1, ec2, ec3 = st.columns([3, 3, 1])
        with ec1:
          st.markdown(f"**{erow['Name']}** ({erow['Department']})")
        with ec2:
          st.markdown(f"`{erow['Email']}`")
        with ec3:
          if st.button("🗑️", key=f"del_email_{erow['ID']}"):
            db.delete_rms_email(erow["ID"])
            st.rerun()

with tab_upload:
  _import_msg = st.session_state.pop("import_msg", None)
  if _import_msg:
    st.success(_import_msg)

  tmpl_cols = db.template_columns()

  def run_import(import_df, update_existing):
    """Shared by all three add-data modes. Returns True on success."""
    rows, skipped, unmatched = db.prepare_import_rows(
        import_df, selected_fy, STATUS_OPTIONS, sparse=True
    )
    if not rows:
      st.error(
          "No valid rows found. Make sure there is a **PURCHASE ORDER** column"
          " (or first column) with values."
      )
      return False
    created, updated = db.upsert_shipments_bulk(
        rows, force_new=not update_existing, skip_blank=True
    )
    db.log_action(
        f"Bulk add to {selected_fy}: {created} new, {updated} updated"
    )
    msg = (
        f"Success! {created} new and {updated} updated shipment records synced"
        f" into {selected_fy} with recalculated deadlines and action warnings."
    )
    if skipped:
      msg += f" ({skipped} rows without a PO were skipped.)"
    if unmatched:
      msg += " Ignored columns: " + ", ".join(f"`{c}`" for c in unmatched[:8])
    st.session_state["import_msg"] = msg
    st.cache_data.clear()
    return True

  st.subheader(f"Add shipments — {selected_fy}")
  st.markdown(
      "Upload an Excel sheet, paste rows copied from Excel, or type straight"
      f" into a table to add or update shipments in **{selected_fy}**."
  )
  st.info(
      "**Column order:** "
      + " → ".join(tmpl_cols)
      + "\n\nOnly **PURCHASE ORDER** is required; the rest are optional. Sheets"
      " with other headers (e.g. 'PO', 'Vendor', 'AWB', 'Qty') are matched"
      " automatically, and the computed columns (days to clear, delays, action"
      " warning) are calculated for you."
  )

  tp1, tp2 = st.columns([1, 3])
  with tp1:
    st.download_button(
        "Download Excel template",
        data=build_template_bytes(tuple(tmpl_cols), db.COLUMN_HELP_UI()),
        file_name="Shipment_Import_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
  with tp2:
    with st.expander("What each column means"):
      st.dataframe(
          pd.DataFrame({
              "Column": tmpl_cols,
              "What to enter": [db.COLUMN_HELP_UI().get(c, "") for c in tmpl_cols],
          }),
          hide_index=True,
          use_container_width=True,
      )

  add_mode = st.radio(
      "How would you like to add shipments?",
      ["📁 Upload Excel", "📋 Paste from Excel", "⌨️ Type in a table"],
      horizontal=True,
      key="add_mode",
  )
  update_existing = st.checkbox(
      "Update existing shipments when the same PO + BL/AWB already exists"
      " (untick to always add new rows)",
      value=True,
      key="update_existing",
  )

  if add_mode == "📁 Upload Excel":
    with st.form("upload_form", clear_on_submit=True):
      uploaded_file = st.file_uploader(
          "Upload Excel File (Shipment Tracker.xlsx)", type=["xlsx", "xls"]
      )
      if (
          st.form_submit_button("Process and Sync Data", type="primary")
          and uploaded_file
      ):
        with st.spinner("Processing & Calculating Formulas..."):
          try:
            import_df = pd.read_excel(uploaded_file)
            if run_import(import_df, update_existing):
              time.sleep(1)
              st.rerun()
          except Exception as e:
            st.error(f"Error processing file: {e}")

  elif add_mode == "📋 Paste from Excel":
    import csv

    st.caption(
        "Copy cells in Excel (Ctrl+C) and paste below (Ctrl+V). Include the"
        " header row, or paste only data in the column order shown above."
    )
    pasted = st.text_area(
        "Paste rows here",
        height=220,
        key="paste_box",
        label_visibility="collapsed",
        placeholder="PO\tSupplier\tCargo description\t...",
    )
    if pasted.strip():
      try:
        grid_rows = [
            r for r in csv.reader(io.StringIO(pasted), delimiter="\t") if any(c.strip() for c in r)
        ]
        width = max(len(r) for r in grid_rows)
        grid_rows = [r + [""] * (width - len(r)) for r in grid_rows]
        if db.headers_look_known(grid_rows[0]):
          paste_df = pd.DataFrame(grid_rows[1:], columns=grid_rows[0])
        else:
          names = [
              tmpl_cols[i] if i < len(tmpl_cols) else f"Extra {i + 1}"
              for i in range(width)
          ]
          paste_df = pd.DataFrame(grid_rows, columns=names)
        st.dataframe(paste_df, hide_index=True, use_container_width=True, height=240)
        st.caption(f"{len(paste_df)} row(s) detected.")
        if st.button(
            f"➕ Add {len(paste_df)} shipment(s) to {selected_fy}",
            type="primary",
            key="paste_go",
        ):
          if run_import(paste_df, update_existing):
            st.session_state.pop("paste_box", None)
            st.rerun()
      except Exception as e:
        st.error(f"Could not read the pasted data: {e}")

  else:  # Type in a table
    st.caption(
        "Fill in the rows below (use the + at the bottom to add more, or paste"
        " cells straight into the table). Blank rows are ignored."
    )
    date_cfg = {
        c: st.column_config.DateColumn(c, format="YYYY-MM-DD")
        for c in db.DATE_UI_COLS
        if c in tmpl_cols
    }
    type_cfg = {
        **date_cfg,
        db.RENAME_MAP["inv_value"]: st.column_config.NumberColumn(
            db.RENAME_MAP["inv_value"], min_value=0.0
        ),
        db.RENAME_MAP["shipping_mode"]: st.column_config.SelectboxColumn(
            db.RENAME_MAP["shipping_mode"], options=MODE_OPTIONS[1:]
        ),
        db.RENAME_MAP["incoterm"]: st.column_config.SelectboxColumn(
            db.RENAME_MAP["incoterm"], options=INCOTERM_OPTIONS[1:]
        ),
        db.RENAME_MAP["current_status"]: st.column_config.SelectboxColumn(
            db.RENAME_MAP["current_status"], options=STATUS_OPTIONS[1:]
        ),
    }
    blank = {}
    for c in tmpl_cols:
      if c in db.DATE_UI_COLS:
        blank[c] = pd.Series([pd.NaT] * 5, dtype="datetime64[ns]")
      elif c == db.RENAME_MAP["inv_value"]:
        blank[c] = pd.Series([None] * 5, dtype="float64")
      else:
        blank[c] = pd.Series([None] * 5, dtype="object")
    typed_df = st.data_editor(
        pd.DataFrame(blank),
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config=type_cfg,
        key="type_table",
    )
    po_col_name = db.RENAME_MAP["po_number"]
    n_ready = int(
        typed_df[po_col_name].fillna("").astype(str).str.strip().ne("").sum()
    )
    if st.button(
        f"➕ Add {n_ready} shipment(s) to {selected_fy}",
        type="primary",
        key="type_go",
        disabled=n_ready == 0,
    ):
      if run_import(typed_df, update_existing):
        st.session_state.pop("type_table", None)
        st.rerun()

with tab_logs:
  st.subheader("📝 Edit Trail & Activity Logs")
  logs_df = db.get_logs()
  if not logs_df.empty:
    st.dataframe(logs_df, hide_index=True, width="stretch")

with tab_deleted:
  st.subheader(f"🗑️ Deleted Shipments for [{selected_fy}]")
  deleted_df = db.get_deleted_shipments(selected_fy)
  if not deleted_df.empty:
    st.dataframe(
        deleted_df[[
            "PURCHASE ORDER",
            "SUPPLIER /MANUFACTURER",
            "BILL OF LADING OR AIRWAY BILL",
            "Deleted By",
            "Deleted At",
        ]],
        hide_index=True,
        width="stretch",
    )
