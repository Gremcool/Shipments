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
  db.auto_backup_daily()
  return True


_boot(datetime.now().strftime("%Y%m%d"))
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



@st.cache_resource(show_spinner=False)
def _logo_html():
  if not os.path.exists("logo_2.jpg"):
    return ""
  import base64

  with open("logo_2.jpg", "rb") as image_file:
    encoded_string = base64.b64encode(image_file.read()).decode()
  return f"<img src='data:image/jpeg;base64,{encoded_string}' height='35'/>"


logo_html = _logo_html()

st.markdown(
    f"<div class='top-banner'>{logo_html}<h1>🚢 RMS Procurement Shipment"
    " Execution & Tracking Dashboard</h1></div>",
    unsafe_allow_html=True,
)


def create_kpi_card(title, value, bg_color, border_color, val_color):
  return f"""<div class='eff-card' style='background-color: {bg_color}; border-left: 5px solid {border_color};'>
        <div class='eff-title'>{title}</div>
        <div class='eff-value' style='color: {val_color};'>{value}</div>
        </div>"""


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
  with st.expander("➕ Initialize New Tracking Year"):
    new_fy = st.text_input("Enter Year (e.g., 2026-2027)")
    if st.button("Create Year"):
      if new_fy and new_fy not in fiscal_years:
        db.add_fiscal_year(new_fy.strip())
        st.rerun()

tab_dash, tab_emails, tab_upload, tab_logs, tab_deleted = st.tabs([
    "📊 Shipment Tracker & Analytics",
    "📧 RMS Email Directory",
    "📂 Upload Excel Data",
    "📝 Edit Trail",
    "🗑️ Deleted Records",
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
        label="📥 Export to Excel (.xlsx)",
        data=excel_data,
        file_name=f"RMS_Shipment_Tracker_{selected_fy}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

  with col_custom_col:
    with st.popover("➕ Add Custom Column", use_container_width=True):
      new_col_name = st.text_input("Column Name")
      if st.button("Add to Grid") and new_col_name:
        ok, msg = db.add_custom_column(new_col_name.strip())
        if ok:
          st.success(msg)
        else:
          st.error(msg)
        st.rerun()

  with col_config:
    if st.button("⚙️ Settings (Colors & Dropdowns)", use_container_width=True):
      manage_dropdowns_dialog()

  st.markdown("</div>", unsafe_allow_html=True)

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
            "Sea Freight 🚢",
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
            "Air Freight ✈️",
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
            "Road Freight 🚛",
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
            "Avg Days for Delay",
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
      "📈 View Procurement Shipment Analytics & Performance Charts",
      expanded=True,
  ):
    tab_carrier, tab_status, tab_clearance = st.tabs([
        "📊 Carrier Distribution",
        "📊 Status Breakdown",
        "📊 Clearance Performance",
    ])

    with tab_carrier:
      df_c = f_df[
          f_df.get("Carrier", pd.Series()).astype(str).str.strip() != ""
      ]
      if not df_c.empty:
        fig_c = px.bar(
            df_c.groupby("Carrier")
            .size()
            .reset_index(name="Shipment Count"),
            x="Carrier",
            y="Shipment Count",
            color="Carrier",
            text="Shipment Count",
        )
        fig_c.update_layout(
            xaxis_title="Carrier",
            yaxis_title="Total Shipments",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            height=280,
        )
        st.plotly_chart(fig_c, use_container_width=True)

    with tab_status:
      df_s = f_df[
          f_df.get("Current status", pd.Series()).astype(str).str.strip() != ""
      ]
      if not df_s.empty:
        fig_s = px.pie(
            df_s, names="Current status", title="Shipments Status Proportion"
        )
        fig_s.update_layout(height=280)
        st.plotly_chart(fig_s, use_container_width=True)

    with tab_clearance:
      df_cl = f_df.dropna(subset=["DAYS TO CLEAR THE SHIPMENT"]).copy()
      if not df_cl.empty:
        # Issue 2 Fix: Normalize Shipping Mode case so Plotly combines SEA/sea into one single box
        df_cl["Shipping mode"] = (
            df_cl["Shipping mode"].astype(str).str.strip().str.upper()
        )
        fig_cl = px.box(
            df_cl,
            x="Shipping mode",
            y="DAYS TO CLEAR THE SHIPMENT",
            color="Shipping mode",
            title="Days to Clear Shipment by Transport Mode",
        )
        fig_cl.update_layout(height=280)
        st.plotly_chart(fig_cl, use_container_width=True)

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
      
      let diffTime = deadlineDate.getTime() - rmsDate.getTime();
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
        "🔍 Search Shipments",
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
        import_df, selected_fy, STATUS_OPTIONS
    )
    if not rows:
      st.error(
          "No valid rows found. Make sure there is a **PURCHASE ORDER** column"
          " (or first column) with values."
      )
      return False
    created, updated = db.upsert_shipments_bulk(
        rows, force_new=not update_existing
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

  st.subheader(f"📥 Add Shipments — {selected_fy}")
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
        "⬇️ Download Excel Template",
        data=build_template_bytes(tuple(tmpl_cols), db.COLUMN_HELP_UI()),
        file_name="Shipment_Import_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
  with tp2:
    with st.expander("📋 See what each column means"):
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
