from datetime import datetime
import os
import sqlite3
import pandas as pd

DB_FILE = os.getenv("DATABASE_PATH", "shipment_tracker.db")

RENAME_MAP = {
    'po_number': 'PURCHASE ORDER',
    'supplier': 'SUPPLIER /MANUFACTURER',
    'cargo_desc': 'Cargo description',
    'invoice_no': 'Invoice number',
    'inv_value': 'Inv value',
    'currency': 'Currency',
    'quantity': 'Quantity',
    'shipping_mode': 'Shipping mode',
    'incoterm': 'Incoterm',
    'container_size': 'Container Size',
    'carrier': 'Carrier',
    'bl_awb': 'BILL OF LADING OR AIRWAY BILL',
    'pod': 'POD',
    'eta_pod': 'ETA POD',
    'actual_port_arrival': 'Actual arrival at port',
    'cargo_eta_rms': 'Cargo ETA RMS',
    'actual_rms_arrival': 'Actual arrival at RMS',
    'doc_sharing_date': (
        'DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS'
    ),
    'current_status': 'Current status',
    'days_to_clear': 'DAYS TO CLEAR THE SHIPMENT',
    'clearance_deadline': 'Deadline to clear the shipment',
    'delays': 'delays',
    'action_warning': 'Action warning',
    'mah_letter_link': 'MAH letter link',
}


def get_conn():
  db_dir = os.path.dirname(DB_FILE)
  if db_dir and not os.path.exists(db_dir):
    os.makedirs(db_dir, exist_ok=True)
  return sqlite3.connect(DB_FILE, check_same_thread=False)


def get_rename_map():
  conn = get_conn()
  c = conn.cursor()
  c.execute("PRAGMA table_info(shipments)")
  columns = [row[1] for row in c.fetchall()]
  conn.close()

  current_map = RENAME_MAP.copy()
  for col in columns:
    if col not in current_map and col not in [
        "id",
        "is_deleted",
        "deleted_by",
        "deleted_at",
        "fiscal_year",
    ]:
      current_map[col] = col
  return current_map


def init_db():
  conn = get_conn()
  c = conn.cursor()

  # 1. Shipments Table
  c.execute("""CREATE TABLE IF NOT EXISTS shipments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fiscal_year TEXT,
        po_number TEXT,
        supplier TEXT,
        cargo_desc TEXT,
        invoice_no TEXT,
        inv_value REAL DEFAULT 0.0,
        currency TEXT,
        quantity TEXT,
        shipping_mode TEXT,
        incoterm TEXT,
        container_size TEXT,
        carrier TEXT,
        bl_awb TEXT,
        pod TEXT,
        eta_pod TEXT,
        actual_port_arrival TEXT,
        cargo_eta_rms TEXT,
        actual_rms_arrival TEXT,
        doc_sharing_date TEXT,
        current_status TEXT,
        days_to_clear INTEGER,
        clearance_deadline TEXT,
        delays INTEGER,
        action_warning TEXT,
        mah_letter_link TEXT,
        is_deleted INTEGER DEFAULT 0,
        deleted_by TEXT,
        deleted_at TEXT
    )""")

  # 2. Dropdowns Table
  c.execute("""CREATE TABLE IF NOT EXISTS dropdowns (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT,
        label TEXT,
        color TEXT
    )""")

  c.execute("SELECT count(*) FROM dropdowns")
  if c.fetchone()[0] == 0:
    defaults = [
        ("status", "Awaiting Departure", "#6c757d"),
        ("status", "In Transit", "#007bff"),
        ("status", "Arrived at Port", "#ffc107"),
        ("status", "Under Clearance", "#17a2b8"),
        ("status", "Received in Warehouse", "#28a745"),
        ("status", "Cancelled/Delayed", "#dc3545"),
        ("mode", "SEA", "#e3f2fd"),
        ("mode", "AIR", "#f3e5f5"),
        ("mode", "ROAD", "#fff3cd"),
        ("incoterm", "CIP", "#ffffff"),
        ("incoterm", "EX-W", "#ffffff"),
        ("incoterm", "FOB", "#ffffff"),
        ("incoterm", "DDP", "#ffffff"),
    ]
    c.executemany(
        "INSERT INTO dropdowns (category, label, color) VALUES (?, ?, ?)",
        defaults,
    )

  # 3. Documents Repository per shipment
  c.execute("""CREATE TABLE IF NOT EXISTS shipment_documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        shipment_id INTEGER,
        doc_number INTEGER,
        file_name TEXT,
        file_type TEXT,
        file_data BLOB,
        uploaded_by TEXT,
        uploaded_at TEXT
    )""")

  # 4. RMS Email Directory
  c.execute("""CREATE TABLE IF NOT EXISTS rms_emails (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        email TEXT UNIQUE,
        department TEXT,
        role TEXT
    )""")

  c.execute("SELECT count(*) FROM rms_emails")
  if c.fetchone()[0] == 0:
    default_emails = [
        (
            "Procurement Team",
            "procurement@rms.rw",
            "Procurement",
            "Department Email",
        ),
        (
            "Logistics & Clearing Agent",
            "clearing.agent@company.com",
            "Logistics",
            "Clearing Agent",
        ),
        (
            "Warehouse Receiving",
            "warehouse@rms.rw",
            "Warehouse",
            "Receiving Officer",
        ),
    ]
    c.executemany(
        "INSERT INTO rms_emails (name, email, department, role) VALUES (?, ?,"
        " ?, ?)",
        default_emails,
    )

  # 5. Row History Trail
  c.execute("""CREATE TABLE IF NOT EXISTS shipment_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        shipment_id INTEGER,
        field_changed TEXT,
        old_value TEXT,
        new_value TEXT,
        changed_by TEXT,
        changed_at TEXT
    )""")

  c.execute(
      "CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT,"
      " timestamp TEXT, message TEXT)"
  )
  c.execute("CREATE TABLE IF NOT EXISTS fiscal_years (year TEXT PRIMARY KEY)")
  c.execute('INSERT OR IGNORE INTO fiscal_years (year) VALUES ("2025-2026")')
  c.execute('INSERT OR IGNORE INTO fiscal_years (year) VALUES ("2026-2027")')

  conn.commit()
  conn.close()


# --- DROPDOWNS MANAGEMENT ---
def get_dropdowns(category):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      "SELECT id, label, color FROM dropdowns WHERE category = ? ORDER BY label",
      (category,),
  )
  data = [
      {"id": row[0], "label": row[1], "color": row[2]} for row in c.fetchall()
  ]
  conn.close()
  return data


def add_dropdown(category, label, color):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      "INSERT INTO dropdowns (category, label, color) VALUES (?, ?, ?)",
      (category, label, color),
  )
  conn.commit()
  conn.close()


def update_dropdown(db_id, new_label, new_color):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      "UPDATE dropdowns SET label = ?, color = ? WHERE id = ?",
      (new_label, new_color, db_id),
  )
  conn.commit()
  conn.close()


def delete_dropdown(db_id):
  conn = get_conn()
  c = conn.cursor()
  c.execute("DELETE FROM dropdowns WHERE id = ?", (db_id,))
  conn.commit()
  conn.close()


# --- RMS EMAIL DIRECTORY ---
def get_rms_emails():
  conn = get_conn()
  df = pd.read_sql(
      "SELECT id as 'ID', name as 'Name', email as 'Email', department as"
      " 'Department', role as 'Role' FROM rms_emails ORDER BY name ASC",
      conn,
  )
  conn.close()
  return df


def add_rms_email(name, email, department="Procurement", role="Officer"):
  conn = get_conn()
  c = conn.cursor()
  try:
    c.execute(
        "INSERT INTO rms_emails (name, email, department, role) VALUES (?, ?,"
        " ?, ?)",
        (name, email, department, role),
    )
    conn.commit()
    log_action(f"Registered new RMS Email: {email} ({name})")
    return True, "Email registered successfully!"
  except sqlite3.IntegrityError:
    return False, "This email address is already registered."
  finally:
    conn.close()


def delete_rms_email(email_id):
  conn = get_conn()
  c = conn.cursor()
  c.execute("DELETE FROM rms_emails WHERE id = ?", (email_id,))
  conn.commit()
  conn.close()


# --- ROW DOCUMENT REPOSITORY ---
def save_row_documents(shipment_id, uploaded_files, uploader_name="Officer"):
  conn = get_conn()
  c = conn.cursor()

  c.execute(
      "SELECT COALESCE(MAX(doc_number), 0) FROM shipment_documents WHERE"
      " shipment_id = ?",
      (shipment_id,),
  )
  max_num = c.fetchone()[0]

  now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  count = 0
  for f in uploaded_files:
    max_num += 1
    fname = f.name
    ftype = f.type if hasattr(f, "type") else "application/octet-stream"
    fbytes = f.read()

    c.execute(
        """INSERT INTO shipment_documents 
            (shipment_id, doc_number, file_name, file_type, file_data, uploaded_by, uploaded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (
            shipment_id,
            max_num,
            fname,
            ftype,
            sqlite3.Binary(fbytes),
            uploader_name,
            now_str,
        ),
    )
    log_row_change(
        shipment_id,
        "Document Uploaded",
        "",
        f"Doc #{max_num}: {fname}",
        uploader_name,
        conn=conn,
    )
    count += 1

  conn.commit()
  conn.close()
  return count


def get_row_documents(shipment_id):
  conn = get_conn()
  df = pd.read_sql(
      "SELECT id, doc_number as 'Doc #', file_name as 'File Name', uploaded_by"
      " as 'Uploaded By', uploaded_at as 'Uploaded At' FROM shipment_documents"
      " WHERE shipment_id = ? ORDER BY doc_number ASC",
      conn,
      params=(shipment_id,),
  )
  conn.close()
  return df


def get_document_blob(doc_id):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      "SELECT file_name, file_type, file_data FROM shipment_documents WHERE id"
      " = ?",
      (int(doc_id),),
  )
  row = c.fetchone()
  conn.close()
  if row:
    return row[0], row[1], row[2]
  return None, None, None


def delete_row_document(doc_id, user_name="Officer"):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      "SELECT shipment_id, file_name FROM shipment_documents WHERE id = ?",
      (int(doc_id),),
  )
  row = c.fetchone()
  if row:
    s_id, fname = row[0], row[1]
    c.execute("DELETE FROM shipment_documents WHERE id = ?", (int(doc_id),))
    log_row_change(
        s_id, "Document Deleted", fname, "", user_name, conn=conn
    )
    conn.commit()
  conn.close()


# --- ROW CHANGE HISTORY TRAIL ---
def log_row_change(
    shipment_id, field_name, old_val, new_val, changed_by="System", conn=None
):
  close_at_end = False
  if conn is None:
    conn = get_conn()
    close_at_end = True
  c = conn.cursor()
  now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  c.execute(
      """INSERT INTO shipment_history 
        (shipment_id, field_changed, old_value, new_value, changed_by, changed_at)
        VALUES (?, ?, ?, ?, ?, ?)""",
      (shipment_id, field_name, str(old_val), str(new_val), changed_by, now_str),
  )
  if close_at_end:
    conn.commit()
    conn.close()


def get_row_change_trail(shipment_id):
  conn = get_conn()
  df = pd.read_sql(
      "SELECT field_changed as 'Field Changed', old_value as 'Old Value',"
      " new_value as 'New Value', changed_by as 'Modified By', changed_at as"
      " 'Timestamp' FROM shipment_history WHERE shipment_id = ? ORDER BY id"
      " DESC",
      conn,
      params=(shipment_id,),
  )
  conn.close()
  return df


# --- CALCULATIONS ENGINE ---
def compute_formulas(row_data):
  today = datetime.now().date()

  # Helper to retrieve field values whether row_data key is DB column name or UI column name
  def get_field_val(db_key):
    ui_key = RENAME_MAP.get(db_key, db_key)
    val = row_data.get(ui_key)
    if val is None or str(val).strip() == "":
      val = row_data.get(db_key)
    return "" if val is None else str(val).strip()

  mode = get_field_val("shipping_mode").upper()
  doc_date_str = get_field_val("doc_sharing_date")
  rms_arr_str = get_field_val("actual_rms_arrival")
  port_arr_str = get_field_val("actual_port_arrival")
  cargo_eta_str = get_field_val("cargo_eta_rms")
  deadline_str = get_field_val("clearance_deadline")

  doc_d = (
      pd.to_datetime(doc_date_str, errors="coerce").date()
      if doc_date_str and str(doc_date_str).lower() not in ["nat", "none", "nan"]
      else None
  )
  rms_arr_d = (
      pd.to_datetime(rms_arr_str, errors="coerce").date()
      if rms_arr_str and str(rms_arr_str).lower() not in ["nat", "none", "nan"]
      else None
  )
  port_arr_d = (
      pd.to_datetime(port_arr_str, errors="coerce").date()
      if port_arr_str and str(port_arr_str).lower() not in ["nat", "none", "nan"]
      else None
  )
  cargo_eta_d = (
      pd.to_datetime(cargo_eta_str, errors="coerce").date()
      if cargo_eta_str
      and str(cargo_eta_str).lower() not in ["nat", "none", "nan"]
      else None
  )
  manual_deadline_d = (
      pd.to_datetime(deadline_str, errors="coerce").date()
      if deadline_str
      and str(deadline_str).lower() not in ["nat", "none", "nan"]
      else None
  )

  # Calculate Days to Clear (Actual RMS Arrival - Doc Sharing Date)
  days_to_clear_val = (rms_arr_d - doc_d).days if rms_arr_d and doc_d else None

  # Standard SLA clearance window = 2 days
  sla_offset = 2

  # Determine Clearance Deadline
  if manual_deadline_d:
    deadline_d = manual_deadline_d
    clearance_deadline = manual_deadline_d.strftime("%Y-%m-%d")
  elif doc_d:
    deadline_d = doc_d + pd.Timedelta(days=sla_offset)
    clearance_deadline = deadline_d.strftime("%Y-%m-%d")
  else:
    deadline_d = None
    clearance_deadline = ""

  # Delays = Deadline to clear the shipment - Actual arrival at RMS
  if deadline_d and rms_arr_d:
    delays_val = (deadline_d - rms_arr_d).days
  elif deadline_d and not rms_arr_d:
    delays_val = (deadline_d - today).days
  else:
    delays_val = None

  # Dynamic Action Warnings
  warnings = []
  if cargo_eta_d and cargo_eta_d <= today + pd.Timedelta(days=3):
    warnings.append("ACTION: ETA RMS due/reached")
  if port_arr_d and port_arr_d <= today + pd.Timedelta(days=3):
    warnings.append("ACTION: Port arrival due/reached")
  if port_arr_d and not doc_d:
    warnings.append("ACTION: Share shipping documents")

  if doc_d and not rms_arr_d and (today - doc_d).days >= sla_offset:
    warnings.append(
        "WARNING: Shipment clearance delay - push clearing agent to expedite"
    )

  # Write computed fields to both UI and DB keys
  row_data[RENAME_MAP["days_to_clear"]] = days_to_clear_val
  row_data["days_to_clear"] = days_to_clear_val

  row_data[RENAME_MAP["clearance_deadline"]] = clearance_deadline
  row_data["clearance_deadline"] = clearance_deadline

  row_data[RENAME_MAP["delays"]] = delays_val
  row_data["delays"] = delays_val

  row_data[RENAME_MAP["action_warning"]] = "; ".join(warnings)
  row_data["action_warning"] = "; ".join(warnings)

  if RENAME_MAP["shipping_mode"] in row_data and row_data[RENAME_MAP["shipping_mode"]]:
    row_data[RENAME_MAP["shipping_mode"]] = str(row_data[RENAME_MAP["shipping_mode"]]).strip().upper()
  if "shipping_mode" in row_data and row_data["shipping_mode"]:
    row_data["shipping_mode"] = str(row_data["shipping_mode"]).strip().upper()

  return row_data


# --- GENERAL DATA CRUD ---
def get_fiscal_years():
  conn = get_conn()
  c = conn.cursor()
  c.execute("SELECT year FROM fiscal_years ORDER BY year DESC")
  years = [row[0] for row in c.fetchall()]
  conn.close()
  return years


def add_fiscal_year(year):
  conn = get_conn()
  c = conn.cursor()
  c.execute("INSERT OR IGNORE INTO fiscal_years (year) VALUES (?)", (year,))
  conn.commit()
  conn.close()
  log_action(f"Created new Fiscal Year: {year}")


def add_custom_column(col_name):
  conn = get_conn()
  c = conn.cursor()
  try:
    c.execute(f'ALTER TABLE shipments ADD COLUMN "{col_name}" TEXT')
    conn.commit()
    log_action(f"Added custom column: '{col_name}'")
    return True, f"Column '{col_name}' added successfully!"
  except sqlite3.OperationalError:
    return False, f"Column '{col_name}' already exists or name is invalid."
  finally:
    conn.close()


def log_action(message):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      "INSERT INTO logs (timestamp, message) VALUES (?, ?)",
      (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), message),
  )
  conn.commit()
  conn.close()


def get_logs():
  conn = get_conn()
  df = pd.read_sql(
      "SELECT timestamp, message FROM logs ORDER BY id DESC LIMIT 100", conn
  )
  conn.close()
  return df


def load_data(fiscal_year):
  conn = get_conn()
  df = pd.read_sql(
      "SELECT * FROM shipments WHERE fiscal_year = ? AND (is_deleted = 0 OR"
      " is_deleted IS NULL)",
      conn,
      params=(fiscal_year,),
  )
  conn.close()
  df.rename(columns=get_rename_map(), inplace=True)
  df.drop(
      columns=["is_deleted", "deleted_by", "deleted_at"],
      errors="ignore",
      inplace=True,
  )

  # Normalize Shipping Mode case to prevent duplicate chart/filter entries
  if "Shipping mode" in df.columns:
    df["Shipping mode"] = (
        df["Shipping mode"].fillna("").astype(str).str.strip().str.upper()
    )

  return df


def get_deleted_shipments(fiscal_year):
  conn = get_conn()
  df = pd.read_sql(
      "SELECT * FROM shipments WHERE fiscal_year = ? AND is_deleted = 1",
      conn,
      params=(fiscal_year,),
  )
  conn.close()
  if not df.empty:
    df.rename(columns=get_rename_map(), inplace=True)
    df.rename(
        columns={"deleted_by": "Deleted By", "deleted_at": "Deleted At"},
        inplace=True,
    )
  return df


def delete_shipment(db_id, user_name="Officer"):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      "UPDATE shipments SET is_deleted = 1, deleted_by = ?, deleted_at = ?"
      " WHERE id = ?",
      (user_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), db_id),
  )
  conn.commit()
  conn.close()
  log_action(f"Shipment ID {db_id} deleted by {user_name}")


def upsert_shipments_bulk(rows_data, changed_by="Officer"):
  conn = get_conn()
  c = conn.cursor()
  reverse_map = {v: k for k, v in get_rename_map().items()}

  for row_data in rows_data:
    row_data = compute_formulas(row_data)
    db_row = {"fiscal_year": row_data["fiscal_year"], "is_deleted": 0}

    for ui_col, val in row_data.items():
      if ui_col in reverse_map:
        db_row[reverse_map[ui_col]] = val

    po = db_row.get("po_number", "")
    bl = db_row.get("bl_awb", "")

    # Normalize shipping mode case for DB write
    if "shipping_mode" in db_row and db_row["shipping_mode"]:
      db_row["shipping_mode"] = str(db_row["shipping_mode"]).strip().upper()

    # Match by ID first, then PO
    if "id" in row_data and row_data["id"]:
      c.execute("SELECT id FROM shipments WHERE id = ?", (row_data["id"],))
      existing = c.fetchone()
    else:
      c.execute(
          "SELECT id FROM shipments WHERE po_number = ? AND (bl_awb = ? OR"
          " bl_awb IS NULL) AND fiscal_year = ?",
          (po, bl, db_row["fiscal_year"]),
      )
      existing = c.fetchone()

    cols = list(db_row.keys())
    vals = [db_row[k] for k in cols]

    if existing:
      s_id = existing[0]
      c.execute("SELECT * FROM shipments WHERE id = ?", (s_id,))
      old_vals_row = c.fetchone()
      col_names = [description[0] for description in c.description]
      old_dict = dict(zip(col_names, old_vals_row)) if old_vals_row else {}

      set_clause = ", ".join([f'"{col}" = ?' for col in cols])
      vals.append(s_id)
      c.execute(f"UPDATE shipments SET {set_clause} WHERE id = ?", vals)

      for col_k, new_v in db_row.items():
        old_v = old_dict.get(col_k, "")
        if str(old_v).strip() != str(new_v).strip():
          ui_header = RENAME_MAP.get(col_k, col_k)
          log_row_change(s_id, ui_header, old_v, new_v, changed_by, conn=conn)
    else:
      quoted_cols = [f'"{col}"' for col in cols]
      col_str = ", ".join(quoted_cols)
      placeholders = ", ".join(["?"] * len(cols))
      c.execute(
          f"INSERT INTO shipments ({col_str}) VALUES ({placeholders})", vals
      )
      new_id = c.lastrowid
      log_row_change(
          new_id, "Shipment Created", "", f"PO #{po}", changed_by, conn=conn
      )

  conn.commit()
  conn.close()


def upsert_shipment(row_data, changed_by="Officer"):
  upsert_shipments_bulk([row_data], changed_by)


def update_single_cell(db_id, db_col_name, new_val, user_name="Officer"):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      f'UPDATE shipments SET "{db_col_name}" = ? WHERE id = ?',
      (str(new_val), db_id),
  )
  conn.commit()
  conn.close()