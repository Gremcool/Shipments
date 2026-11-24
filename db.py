from datetime import datetime, timedelta
import os
import sqlite3
import pandas as pd

DB_NAME = "shipment_tracker.db"
VOLUME_DIR = os.getenv("RAILWAY_VOLUME_MOUNT_PATH", "").strip()


def _resolve_db_file():
  """DATABASE_PATH wins, then the Railway volume, then a local file."""
  explicit = os.getenv("DATABASE_PATH", "").strip()
  if explicit:
    return explicit
  if VOLUME_DIR:
    return os.path.join(VOLUME_DIR, DB_NAME)
  return DB_NAME


DB_FILE = _resolve_db_file()
BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(DB_FILE)), "backups")
MAX_AUTO_BACKUPS = 7


def is_persistent():
  """True when the database lives on a mounted volume (or explicit path)."""
  if VOLUME_DIR and os.path.abspath(DB_FILE).startswith(
      os.path.abspath(VOLUME_DIR)
  ):
    return True
  return bool(os.getenv("DATABASE_PATH", "").strip())


def running_on_railway():
  return bool(os.getenv("RAILWAY_ENVIRONMENT") or os.getenv("RAILWAY_SERVICE_ID"))


def storage_info():
  size = os.path.getsize(DB_FILE) if os.path.exists(DB_FILE) else 0
  return {
      "db_file": os.path.abspath(DB_FILE),
      "volume_dir": VOLUME_DIR or "(none attached)",
      "persistent": is_persistent(),
      "on_railway": running_on_railway(),
      "size_bytes": size,
  }


def table_counts():
  conn = get_conn()
  out = {}
  for label, q in [
      ("Shipments (active)", "SELECT count(*) FROM shipments WHERE is_deleted=0"),
      ("Shipments (deleted)", "SELECT count(*) FROM shipments WHERE is_deleted=1"),
      ("Documents", "SELECT count(*) FROM shipment_documents"),
      ("Email contacts", "SELECT count(*) FROM rms_emails"),
      ("History entries", "SELECT count(*) FROM shipment_history"),
  ]:
    try:
      out[label] = conn.execute(q).fetchone()[0]
    except sqlite3.Error:
      out[label] = 0
  conn.close()
  return out


def create_backup(tag="manual"):
  """Consistent snapshot using SQLite's online backup API."""
  os.makedirs(BACKUP_DIR, exist_ok=True)
  stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
  path = os.path.join(BACKUP_DIR, f"{tag}_{stamp}.db")
  src = get_conn()
  dst = sqlite3.connect(path)
  src.backup(dst)
  dst.close()
  src.close()
  return path


def list_backups():
  if not os.path.isdir(BACKUP_DIR):
    return []
  files = [f for f in os.listdir(BACKUP_DIR) if f.endswith(".db")]
  files.sort(reverse=True)
  return [
      (f, os.path.getsize(os.path.join(BACKUP_DIR, f)),
       datetime.fromtimestamp(os.path.getmtime(os.path.join(BACKUP_DIR, f))
                              ).strftime("%Y-%m-%d %H:%M"))
      for f in files
  ]


def auto_backup_daily():
  """One automatic snapshot per day (only if there is data); keep the last 7."""
  try:
    if not os.path.exists(DB_FILE):
      return
    today = datetime.now().strftime("%Y%m%d")
    if any(f.startswith(f"auto_{today}") for f, _, _ in list_backups()):
      return
    if table_counts().get("Shipments (active)", 0) == 0:
      return
    create_backup("auto")
    autos = [f for f, _, _ in list_backups() if f.startswith("auto_")]
    for old in autos[MAX_AUTO_BACKUPS:]:
      os.remove(os.path.join(BACKUP_DIR, old))
  except Exception:
    pass


def read_backup(name):
  path = os.path.join(BACKUP_DIR, os.path.basename(name))
  with open(path, "rb") as f:
    return f.read()


def restore_from_bytes(data):
  """Validate an uploaded .db file, snapshot the current one, then replace it."""
  tmp = DB_FILE + ".restore_tmp"
  with open(tmp, "wb") as f:
    f.write(data)
  try:
    chk = sqlite3.connect(tmp)
    tables = {r[0] for r in chk.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")}
    chk.close()
    if "shipments" not in tables:
      os.remove(tmp)
      return False, "That file is not a valid shipment tracker database."
  except sqlite3.Error as e:
    if os.path.exists(tmp):
      os.remove(tmp)
    return False, f"Invalid database file: {e}"
  if os.path.exists(DB_FILE):
    create_backup("pre_restore")
  for ext in ("-wal", "-shm"):
    try:
      os.remove(DB_FILE + ext)
    except OSError:
      pass
  os.replace(tmp, DB_FILE)
  return True, "Database restored successfully."

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
    'import_license_stage': 'Import license stage',
    'missing_cargo_doc': 'Missing cargo document',
    'cargo_dispatched': 'Cargo dispatched',
    'clearing_agent': 'Clearing agent',
    'arrival_notice_ready': 'Arrival notice ready',
    'actual_kigali_arrival': 'Actual arrival in Kigali',
    'remarks': 'Remarks',
    'reason_for_delay': 'Reason for delay',
    'responsible_staff': 'Responsible staff',
}

# Extra tracker fields kept as plain text columns (added automatically).
EXTRA_DB_COLUMNS = [
    'import_license_stage', 'missing_cargo_doc', 'cargo_dispatched',
    'clearing_agent', 'arrival_notice_ready', 'actual_kigali_arrival',
    'remarks', 'reason_for_delay', 'responsible_staff',
]


def get_conn():
  db_dir = os.path.dirname(DB_FILE)
  if db_dir and not os.path.exists(db_dir):
    os.makedirs(db_dir, exist_ok=True)
  conn = sqlite3.connect(DB_FILE, check_same_thread=False, timeout=30)
  try:
    # WAL = readers never block the writer; NORMAL sync is safe with WAL and
    # much faster on network/volume disks.
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA temp_store=MEMORY")
    conn.execute("PRAGMA cache_size=-20000")
  except sqlite3.Error:
    pass
  return conn


def db_version():
  """Cheap change-detector (file stamps) used as a cache key by the app."""
  stamp = []
  for path in (DB_FILE, DB_FILE + "-wal"):
    try:
      st_ = os.stat(path)
      stamp.append((st_.st_mtime_ns, st_.st_size))
    except OSError:
      stamp.append(None)
  return tuple(stamp)


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

  # 1b. Extra tracker columns + one-off migrations
  existing = {row[1] for row in c.execute("PRAGMA table_info(shipments)")}
  for col in EXTRA_DB_COLUMNS:
    if col not in existing:
      c.execute(f'ALTER TABLE shipments ADD COLUMN "{col}" TEXT')
  c.execute(
      "CREATE TABLE IF NOT EXISTS app_meta (key TEXT PRIMARY KEY, value TEXT)"
  )
  conn.commit()
  if not c.execute(
      "SELECT 1 FROM app_meta WHERE key = 'delay_sign_v2'"
  ).fetchone():
    recompute_all_formulas(conn)
    c.execute("INSERT INTO app_meta (key, value) VALUES ('delay_sign_v2', ?)",
              (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))
    conn.commit()

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
  # Indexes: make year loads, row history and document lookups instant.
  c.execute(
      "CREATE INDEX IF NOT EXISTS idx_shipments_fy ON shipments"
      " (fiscal_year, is_deleted)"
  )
  c.execute(
      "CREATE INDEX IF NOT EXISTS idx_shipments_po ON shipments"
      " (fiscal_year, po_number)"
  )
  c.execute(
      "CREATE INDEX IF NOT EXISTS idx_history_ship ON shipment_history"
      " (shipment_id)"
  )
  c.execute(
      "CREATE INDEX IF NOT EXISTS idx_docs_ship ON shipment_documents"
      " (shipment_id)"
  )
  c.execute("CREATE TABLE IF NOT EXISTS fiscal_years (year TEXT PRIMARY KEY)")
  c.execute('INSERT OR IGNORE INTO fiscal_years (year) VALUES ("2025-2026")')
  c.execute('INSERT OR IGNORE INTO fiscal_years (year) VALUES ("2026-2027")')

  conn.commit()
  conn.close()


# --- DROPDOWNS MANAGEMENT ---
def recompute_all_formulas(conn=None):
  """Refresh deadline / delays / days-to-clear / warnings for every shipment.

  Delays and warnings depend on today's date, so this runs once a day at
  start-up (and once after the delay-sign fix).
  """
  own = conn is None
  if own:
    conn = get_conn()
  c = conn.cursor()
  c.execute("SELECT * FROM shipments")
  names = [x[0] for x in c.description]
  for rec in c.fetchall():
    row = dict(zip(names, rec))
    calc = compute_formulas(dict(row))
    c.execute(
        "UPDATE shipments SET days_to_clear = ?, clearance_deadline = ?,"
        " delays = ?, action_warning = ? WHERE id = ?",
        (calc["days_to_clear"], calc["clearance_deadline"], calc["delays"],
         calc["action_warning"], row["id"]),
    )
  conn.commit()
  if own:
    conn.close()


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


# --- FAST VALUE HELPERS ---
_NULL_TOKENS = {"", "nat", "none", "nan", "null"}


def to_date(val):
  """Parse many date shapes to a date, or None. ISO strings take a fast path."""
  if val is None:
    return None
  try:
    if pd.isna(val):
      return None
  except (TypeError, ValueError):
    pass
  if isinstance(val, datetime):
    return val.date()
  if hasattr(val, "year") and hasattr(val, "month") and hasattr(val, "day"):
    try:
      return datetime(val.year, val.month, val.day).date()
    except (ValueError, TypeError):
      return None
  sval = str(val).strip()
  if sval.lower() in _NULL_TOKENS:
    return None
  try:
    return datetime.strptime(sval[:10], "%Y-%m-%d").date()
  except ValueError:
    pass
  try:
    ts = pd.to_datetime(sval, errors="coerce")
  except Exception:
    return None
  return None if pd.isna(ts) else ts.date()


def norm_date_str(val):
  d = to_date(val)
  return d.strftime("%Y-%m-%d") if d else ""


def clean_text(val):
  """NaN -> '', integral floats -> '4500123' (not '4500123.0'), strip."""
  if val is None:
    return ""
  try:
    if pd.isna(val):
      return ""
  except (TypeError, ValueError):
    pass
  if isinstance(val, float) and val.is_integer():
    return str(int(val))
  return str(val).strip()


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

  doc_d = to_date(doc_date_str)
  rms_arr_d = to_date(rms_arr_str)
  port_arr_d = to_date(port_arr_str)
  cargo_eta_d = to_date(cargo_eta_str)
  manual_deadline_d = to_date(deadline_str)

  # Calculate Days to Clear (Actual RMS Arrival - Doc Sharing Date)
  days_to_clear_val = (rms_arr_d - doc_d).days if rms_arr_d and doc_d else None

  # Standard SLA clearance window = 2 days
  sla_offset = 2

  # Determine Clearance Deadline
  if manual_deadline_d:
    deadline_d = manual_deadline_d
    clearance_deadline = manual_deadline_d.strftime("%Y-%m-%d")
  elif doc_d:
    deadline_d = doc_d + timedelta(days=sla_offset)
    clearance_deadline = deadline_d.strftime("%Y-%m-%d")
  else:
    deadline_d = None
    clearance_deadline = ""

  # Delays = Actual arrival at RMS - Deadline (positive = late, same as the
  # Excel tracker). Open shipments are measured against today.
  if deadline_d and rms_arr_d:
    delays_val = (rms_arr_d - deadline_d).days
  elif deadline_d and not rms_arr_d:
    delays_val = (today - deadline_d).days
  else:
    delays_val = None

  # Dynamic Action Warnings
  warnings = []
  if cargo_eta_d and cargo_eta_d <= today + timedelta(days=3):
    warnings.append("ACTION: ETA RMS due/reached")
  if port_arr_d and port_arr_d <= today + timedelta(days=3):
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


_PLACEHOLDERS = {"", "tba", "pending", "na", "n/a", "none", "nan", "null"}


def _is_placeholder(val):
  return str("" if val is None else val).strip().lower() in _PLACEHOLDERS


def _norm_key(val):
  return " ".join(str("" if val is None else val).split()).upper()


def pick_existing(cands, bl, inv):
  """Choose which existing row (same PO + fiscal year) an upload row updates.

  cands: list of (id, bl_awb, invoice_no) that are not already claimed by an
  earlier row of the same upload. Returns an id, or None (= add as new row).
  1) same real BL/AWB
  2) same real invoice no. (a corrected BL on the same invoice is an update)
  3) BL is still a placeholder (TBA/pending/blank) on either side, the
     invoice does not contradict, and exactly one candidate is left.
  Anything ambiguous is added as new, never guessed.
  """
  bl_n, inv_n = _norm_key(bl), _norm_key(inv)
  if not _is_placeholder(bl_n):
    for cid, cbl, _ in cands:
      if _norm_key(cbl) == bl_n:
        return cid
  if not _is_placeholder(inv_n):
    hits = [c for c in cands if _norm_key(c[2]) == inv_n]
    if len(hits) == 1:
      return hits[0][0]
    if hits and _is_placeholder(bl_n):
      return hits[0][0]
  # a candidate whose invoice is real and different is another shipment
  compat = [
      c for c in cands
      if (_is_placeholder(bl_n) or _is_placeholder(c[1]))
      and (
          _is_placeholder(inv_n)
          or _is_placeholder(c[2])
          or _norm_key(c[2]) == inv_n
      )
  ]
  if len(compat) == 1:
    return compat[0][0]
  return None


def _find_existing(c, db_row, claimed):
  # PO text is compared ignoring extra spaces / line breaks (stored POs often
  # contain them), so "A AND  B" and "A AND B" are the same PO.
  want = _norm_key(db_row.get("po_number", ""))
  c.execute(
      "SELECT id, bl_awb, invoice_no, po_number FROM shipments WHERE"
      " fiscal_year = ? ORDER BY COALESCE(is_deleted, 0), id",
      (db_row["fiscal_year"],),
  )
  cands = [
      r[:3] for r in c.fetchall()
      if r[0] not in claimed and _norm_key(r[3]) == want
  ]
  if not cands:
    return None
  return pick_existing(cands, db_row.get("bl_awb"), db_row.get("invoice_no"))


def upsert_shipments_bulk(
    rows_data, changed_by="Officer", force_new=False, skip_blank=False
):
  """Insert/update many shipments in one transaction.

  force_new=True always inserts a brand-new row (used for duplicates and for
  "add as new" imports). skip_blank=True (bulk imports) never overwrites an
  existing value with an empty one, so partial sheets are safe.
  Returns (created_count, updated_count).
  """
  conn = get_conn()
  c = conn.cursor()
  reverse_map = {v: k for k, v in get_rename_map().items()}
  created = updated = 0
  claimed = set()  # existing rows already updated by this upload

  def _translate(row):
    out = {}
    for ui_col, val in row.items():
      if ui_col in reverse_map:
        out[reverse_map[ui_col]] = val
    return out

  try:
    for row_data in rows_data:
      row_data = dict(row_data)
      fy = row_data["fiscal_year"]
      probe = {"fiscal_year": fy, **_translate(row_data)}

      old_dict = None
      if not force_new:
        if row_data.get("id"):
          c.execute("SELECT * FROM shipments WHERE id = ?", (row_data["id"],))
        else:
          found = _find_existing(c, probe, claimed)
          c.execute("SELECT * FROM shipments WHERE id = ?", (found,))
        old_row = c.fetchone()
        if old_row:
          old_dict = dict(zip([d[0] for d in c.description], old_row))
          claimed.add(old_dict["id"])

      if skip_blank and old_dict:
        db_row = {
            k: v for k, v in probe.items()
            if k in ("fiscal_year",) or str("" if v is None else v).strip() != ""
        }
        db_row["is_deleted"] = 0
        merged = {**old_dict, **db_row}
        if "clearance_deadline" not in db_row:
          # an auto deadline (doc date + 2 days) must follow a new doc date
          od = to_date(old_dict.get("doc_sharing_date"))
          odl = to_date(old_dict.get("clearance_deadline"))
          if odl is None or (od and odl == od + timedelta(days=2)):
            merged["clearance_deadline"] = ""
        calc = compute_formulas(merged)
        for k in ("days_to_clear", "clearance_deadline", "delays",
                  "action_warning"):
          db_row[k] = calc[k]
      else:
        calc = compute_formulas(row_data)
        db_row = {"fiscal_year": fy, "is_deleted": 0, **_translate(calc)}
        if skip_blank:
          db_row = {
              k: v for k, v in db_row.items()
              if v is not None and (str(v).strip() != "" or k == "fiscal_year")
              or k in ("days_to_clear", "delays", "action_warning",
                       "clearance_deadline")
          }
          if not str(db_row.get("current_status", "")).strip():
            db_row["current_status"] = "In Transit"

      po = db_row.get("po_number", "")

      # Normalize shipping mode case for DB write
      if "shipping_mode" in db_row and db_row["shipping_mode"]:
        db_row["shipping_mode"] = str(db_row["shipping_mode"]).strip().upper()

      if force_new:
        db_row.pop("id", None)

      cols = list(db_row.keys())
      vals = [db_row[k] for k in cols]

      if old_dict:
        s_id = old_dict["id"]
        set_clause = ", ".join([f'"{col}" = ?' for col in cols])
        vals.append(s_id)
        c.execute(f"UPDATE shipments SET {set_clause} WHERE id = ?", vals)
        updated += 1

        for col_k, new_v in db_row.items():
          old_v = old_dict.get(col_k, "")
          old_s = "" if old_v is None else str(old_v).strip()
          new_s = "" if new_v is None else str(new_v).strip()
          if old_s != new_s:
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
        created += 1
        log_row_change(
            new_id,
            "Shipment Created",
            "",
            f"PO #{po}" + (" (duplicated)" if force_new else ""),
            changed_by,
            conn=conn,
        )

    conn.commit()
  finally:
    conn.close()
  return created, updated


def upsert_shipment(row_data, changed_by="Officer"):
  return upsert_shipments_bulk([row_data], changed_by)


def duplicate_shipment(db_id, copies=1, clear_tracking=True, changed_by="Officer"):
  """Create `copies` new rows cloned from an existing shipment."""
  conn = get_conn()
  df = pd.read_sql("SELECT * FROM shipments WHERE id = ?", conn, params=(int(db_id),))
  conn.close()
  if df.empty:
    return 0
  src = df.iloc[0].to_dict()
  rename = get_rename_map()
  skip = {"id", "is_deleted", "deleted_by", "deleted_at"}
  reset_keys = {
      "bl_awb", "actual_port_arrival", "actual_rms_arrival", "doc_sharing_date",
      "mah_letter_link", "invoice_no",
  }
  rows = []
  for _ in range(max(1, int(copies))):
    row = {}
    for db_key, val in src.items():
      if db_key in skip:
        continue
      if db_key == "fiscal_year":
        row["fiscal_year"] = val
        continue
      if clear_tracking and db_key in reset_keys:
        val = ""
      row[rename.get(db_key, db_key)] = "" if pd.isna(val) else val
    rows.append(row)
  created, _ = upsert_shipments_bulk(rows, changed_by, force_new=True)
  log_action(f"Shipment ID {db_id} duplicated x{created} by {changed_by}")
  return created


# --- IMPORT HELPERS (Excel upload / paste / type-in table) ---
COMPUTED_UI_COLS = {
    RENAME_MAP["days_to_clear"],
    RENAME_MAP["delays"],
    RENAME_MAP["action_warning"],
}
DATE_UI_COLS = [
    RENAME_MAP[k]
    for k in (
        "eta_pod",
        "actual_port_arrival",
        "cargo_eta_rms",
        "actual_rms_arrival",
        "doc_sharing_date",
        "clearance_deadline",
        "actual_kigali_arrival",
    )
]

# Extra header spellings people commonly use (normalised: lowercase a-z0-9).
_HEADER_ALIASES = {
    "po": "po_number", "ponumber": "po_number", "pono": "po_number",
    "ponum": "po_number", "purchaseorderno": "po_number",
    "supplier": "supplier", "manufacturer": "supplier", "vendor": "supplier",
    "suppliermanufacturer": "supplier",
    "cargo": "cargo_desc", "description": "cargo_desc", "goods": "cargo_desc",
    "itemdescription": "cargo_desc",
    "invoice": "invoice_no", "invoiceno": "invoice_no",
    "invoicenumber": "invoice_no", "invno": "invoice_no",
    "invoicevalue": "inv_value", "value": "inv_value", "amount": "inv_value",
    "curr": "currency",
    "qty": "quantity",
    "mode": "shipping_mode", "transportmode": "shipping_mode",
    "modeoftransport": "shipping_mode", "shippingmode": "shipping_mode",
    "incoterms": "incoterm",
    "container": "container_size",
    "shippingline": "carrier", "airline": "carrier",
    "bl": "bl_awb", "awb": "bl_awb", "blawb": "bl_awb", "blno": "bl_awb",
    "awbno": "bl_awb", "billoflading": "bl_awb", "airwaybill": "bl_awb",
    "portofdischarge": "pod",
    "portarrival": "actual_port_arrival",
    "etarms": "cargo_eta_rms",
    "rmsarrival": "actual_rms_arrival",
    "docsharingdate": "doc_sharing_date",
    "status": "current_status",
    "deadline": "clearance_deadline", "clearancedeadline": "clearance_deadline",
    "mahletter": "mah_letter_link", "mahlink": "mah_letter_link",
    "comments": "remarks", "additionalremarks": "remarks", "notes": "remarks",
    "forwarder": "clearing_agent", "importlicense": "import_license_stage",
    "licensestage": "import_license_stage", "dispatched": "cargo_dispatched",
    "arrivalnotice": "arrival_notice_ready",
    "missingdocument": "missing_cargo_doc", "staff": "responsible_staff",
}

# What the template / "column meaning" panel shows (input columns only).
TEMPLATE_DB_KEYS = [
    "po_number", "supplier", "cargo_desc", "invoice_no", "inv_value",
    "currency", "quantity", "shipping_mode", "incoterm", "container_size",
    "carrier", "bl_awb", "pod", "eta_pod", "actual_port_arrival",
    "cargo_eta_rms", "actual_rms_arrival", "doc_sharing_date",
    "current_status", "clearance_deadline", "mah_letter_link",
    "import_license_stage", "missing_cargo_doc", "cargo_dispatched",
    "clearing_agent", "arrival_notice_ready", "actual_kigali_arrival",
    "remarks",
]
COLUMN_HELP = {
    "po_number": "Purchase order number (REQUIRED - rows without it are skipped).",
    "supplier": "Supplier / manufacturer name.",
    "cargo_desc": "What is being shipped.",
    "invoice_no": "Supplier invoice number.",
    "inv_value": "Invoice amount (number).",
    "currency": "USD, EUR ... (EUR is converted at 1.08 for the KPI card).",
    "quantity": "Quantity shipped (free text, e.g. '200 boxes').",
    "shipping_mode": "SEA, AIR or ROAD.",
    "incoterm": "CIP, EX-W, FOB, DDP ...",
    "container_size": "e.g. 20ft, 40ft, LCL.",
    "carrier": "Shipping line / airline / transporter.",
    "bl_awb": "Bill of lading or airway bill number.",
    "pod": "Port of discharge.",
    "eta_pod": "Planned arrival at port (YYYY-MM-DD).",
    "actual_port_arrival": "Actual arrival at port (YYYY-MM-DD).",
    "cargo_eta_rms": "Planned arrival at RMS (YYYY-MM-DD).",
    "actual_rms_arrival": "Actual arrival at RMS (YYYY-MM-DD).",
    "doc_sharing_date": "Date shipping documents were shared with warehouse/FLS.",
    "current_status": "One of your status options; blank = In Transit.",
    "clearance_deadline": "Optional manual deadline; blank = doc sharing date + 2 days.",
    "mah_letter_link": "Link to the MAH letter.",
    "import_license_stage": "Not started, Drafted, Submitted, Approved, Rejected, Queried ...",
    "missing_cargo_doc": "Which cargo document is still missing (blank = none).",
    "cargo_dispatched": "Yes / No - has the supplier dispatched the cargo?",
    "clearing_agent": "Clearing agent or forwarder handling the shipment.",
    "arrival_notice_ready": "Yes / No - is the arrival notice ready?",
    "actual_kigali_arrival": "Actual arrival in Kigali (YYYY-MM-DD).",
    "remarks": "Free-text notes (freight quote status, issues, next step).",
}


def COLUMN_HELP_UI():
  return {RENAME_MAP[k]: v for k, v in COLUMN_HELP.items()}


def template_columns():
  return [RENAME_MAP[k] for k in TEMPLATE_DB_KEYS]


def _hnorm(text):
  return "".join(ch for ch in str(text).lower() if ch.isalnum())


def match_columns(columns):
  """Map uploaded headers -> app column names (exact, DB-key or alias match)."""
  lookup = {}
  current = get_rename_map()  # includes custom columns
  for db_key, ui in current.items():
    lookup[_hnorm(ui)] = ui
    lookup[_hnorm(db_key)] = ui
  for alias, db_key in _HEADER_ALIASES.items():
    lookup.setdefault(alias, RENAME_MAP[db_key])
  result = {}
  for col in columns:
    ui = lookup.get(_hnorm(col))
    if ui and ui not in COMPUTED_UI_COLS and ui not in result.values():
      result[col] = ui
  return result


def headers_look_known(values):
  """True if at least 2 of the given cells are recognisable column headers."""
  return len(match_columns([v for v in values if str(v).strip()])) >= 2


def prepare_import_rows(import_df, fiscal_year, status_options=None, sparse=False):
  """Turn any sheet-like DataFrame into row dicts ready for upsert.

  Returns (rows, skipped_blank_po, unmatched_headers).
  sparse=True: columns missing from the sheet are left out of each row (so an
  update never touches them) and unusable status / value cells become blank
  instead of "In Transit" / 0.
  """
  import_df = import_df.copy()
  import_df.columns = [str(c).strip().replace("\n", " ") for c in import_df.columns]
  mapping = match_columns(import_df.columns)
  unmatched = [c for c in import_df.columns if c not in mapping and str(c).strip()]
  ui_to_src = {ui: src for src, ui in mapping.items()}

  status_lookup = {str(o).strip().lower(): o for o in (status_options or []) if o}
  po_col = RENAME_MAP["po_number"]
  if po_col not in ui_to_src:
    return [], len(import_df), unmatched

  extra_cols = [u for u in ui_to_src if u != po_col]
  rows, skipped = [], 0
  records = import_df.to_dict("records")
  for rec in records:
    raw_po = clean_text(rec.get(ui_to_src[po_col]))
    if not raw_po or raw_po.lower() == "nan":
      skipped += 1
      continue
    row = {"fiscal_year": fiscal_year, po_col: raw_po}
    for ui_col in template_columns() + [
        u for u in extra_cols if u not in template_columns()
    ]:
      if ui_col == po_col:
        continue
      src = ui_to_src.get(ui_col)
      if sparse and not src:
        continue
      val = rec.get(src, "") if src else ""
      if ui_col in DATE_UI_COLS:
        val = norm_date_str(val)
      elif ui_col == RENAME_MAP["inv_value"]:
        try:
          val = float(str(val).replace(",", "").replace("$", "").strip())
          if pd.isna(val):
            val = "" if sparse else 0.0
        except (TypeError, ValueError):
          val = "" if sparse else 0.0
      elif ui_col == RENAME_MAP["current_status"]:
        val = status_lookup.get(
            clean_text(val).lower(), "" if sparse else "In Transit"
        )
      else:
        val = clean_text(val)
      row[ui_col] = val
    rows.append(row)
  return rows, skipped, unmatched


def update_single_cell(db_id, db_col_name, new_val, user_name="Officer"):
  conn = get_conn()
  c = conn.cursor()
  c.execute(
      f'UPDATE shipments SET "{db_col_name}" = ? WHERE id = ?',
      (str(new_val), db_id),
  )
  conn.commit()
  conn.close()