"""Build seed/seed_data.json - the data the app merges into the live database
once, automatically, the first time it starts after deployment.

  python scripts/build_seed.py
"""
import os, re, sys, json, runpy, warnings
warnings.filterwarnings("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT); sys.path.insert(0, ROOT); os.makedirs("seed", exist_ok=True)
import pandas as pd
import db

runpy.run_path("scripts/prepare_update_file.py")          # merge tracker into the 26-Oct export
m = pd.read_pickle("seed/_merged.pkl"); os.remove("seed/_merged.pkl")
A, EX = m["APP_COLS"], m["EXTRA_UI"]
COMPUTED = {A[19], A[21], A[22]}
REV = {v: k for k, v in db.RENAME_MAP.items()}
DATE_UI = {db.RENAME_MAP[k] for k in ("eta_pod", "actual_port_arrival", "cargo_eta_rms", "actual_rms_arrival", "doc_sharing_date", "clearance_deadline", "actual_kigali_arrival")}
CURRENT_FY = "2026-2027"
nz = lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else re.sub(r"[ \t]+", " ", str(v)).strip()

def norm(v):
  if v is None or v == "": return ""
  if isinstance(v, float): return f"{v:.4f}"
  return " ".join(str(v).split())

def clean(ui, v):
  if ui == A[4]:
    if v in ("", None): return None
    if isinstance(v, (int, float)): return float(v)
    t = re.sub(r"[^\d,.\-]", "", str(v))
    t = t.replace(".", "").replace(",", ".") if re.search(r",\d{1,2}$", t) else t.replace(",", "")
    try: return float(t)
    except ValueError: return None
  v = db.norm_date_str(v) if ui in DATE_UI else (v.strip() if isinstance(v, str) else v)
  return v if v not in ("", None) else None

def vals(x, cols):
  out = {}
  for ui in cols:
    if ui in COMPUTED or ui not in REV: continue
    v = clean(ui, x.get(ui))
    if v is not None: out[REV[ui]] = v
  return out

records = []
for x in m["out"]:
  if x["__new"]:
    records.append({"kind": "new", "fy": CURRENT_FY, "src": x["__src"], "bl": x[A[11]], "inv": x[A[3]], "values": vals(x, A + EX)})
    continue
  o = x["__orig"]
  changed = {ui: x[ui] for ui in A + EX if ui not in COMPUTED and norm(x[ui]) != norm(o.get(ui, ""))}
  records.append({
    "kind": "update", "fy": CURRENT_FY, "src": x["__src"],
    "base_bl": o[A[11]], "base_inv": o[A[3]], "bl": x[A[11]], "inv": x[A[3]], "po": x["__raw"][A[0]] if x.get("__raw") else x[A[0]],
    "values": vals(changed, changed), "base": vals(o, changed), "full": vals(x, A + EX),
  })

# ---------- older tracker sheets ----------
TRK = "original_data/Shipment_tracker_last_update.xlsx"
PH = {"", "tba", "pending", "na", "n/a", "none", "nan", "null", "-"}
today = pd.Timestamp.today().date()
def fy_of(*dates, default):
  for v in dates:
    d = db.to_date(v)
    if d and d <= today:
      y = d.year if d.month >= 7 else d.year - 1
      return f"{y}-{y + 1}"
  return default
def infer_mode(pod, carrier, bl):
  t = f"{nz(pod)} {nz(carrier)}".lower()
  if re.search(r"\b(kia|kai|airport)\b|rwandair|ethiopian|kenya airways|emirates|qatar|turkish|lufthansa|klm|dhl", t): return "AIR"
  if re.search(r"gatuna|rusumo|kagitumba|nemba|cyanika|truck|road", t): return "ROAD"
  if re.search(r"dar ?es|mombasa|djibouti|durban|beira|maersk|msc|cma|hapag|evergreen|cosco", t): return "SEA"
  if re.fullmatch(r"\d{3}[- ]?\d{7,8}", nz(bl).replace(" ", "")): return "AIR"
  return ""
def status_of(arrived, text):
  t = nz(text).lower()
  if re.search(r"pending|not yet|awaiting|under clearance|in transit", t) and not arrived:
    return "Under Clearance" if "clear" in t else "In Transit"
  return "Received in Warehouse"
def load_old(name):
  d = pd.read_excel(TRK, sheet_name=name, dtype=object, keep_default_na=False, na_values=[""])
  d.columns = [" ".join(str(c).split()) for c in d.columns]
  return d
def nonempty(d):
  keep = [c for c in d.columns if not c.startswith("Unnamed")]
  return d[d[keep].apply(lambda r: any(nz(v) for v in r), axis=1)]

d2 = nonempty(load_old("Delivered-2nd_old_sheet")).drop(columns=["SN"], errors="ignore")
d2 = d2[[c for c in d2.columns if not c.startswith("Unnamed")]].copy()
d2["PURCHASE ORDER"] = d2["PURCHASE ORDER"].map(lambda v: nz(v) if nz(v).lower() not in PH else "NO PO")
d2["Shipping mode"] = [mo if nz(mo) else infer_mode(p, c, b) for mo, p, c, b in zip(d2["Shipping mode"], d2["POD"], d2["Carrier"], d2["BILL OF LADING OR AIRWAY BILL"])]
d2["Current status"] = [status_of(nz(a), t) for a, t in zip(d2["Actual arrival at RMS"], d2["Current status"])]
d2["__fy"] = [fy_of(a, p, default="2025-2026") for a, p in zip(d2["Actual arrival at RMS"], d2["Actual arrival at port"])]
d2 = d2.rename(columns={"QTY": "Quantity", " Inv value ": "Inv value"})

d1 = nonempty(load_old("Delivered-old_sheet"))
unn = [c for c in d1.columns if c.startswith("Unnamed")]
def remarks(r):
  p = []
  if nz(r.get("RFDA FEES PAID")): p.append(f"RFDA fees paid: {nz(r['RFDA FEES PAID'])}")
  if nz(r.get("E-mail reference")): p.append(f"E-mail ref: {nz(r['E-mail reference'])}")
  if nz(r.get("FULL SHIPPING DOCUMENTS RECEIPT DATE")): p.append(f"Full docs received: {db.norm_date_str(r['FULL SHIPPING DOCUMENTS RECEIPT DATE']) or nz(r['FULL SHIPPING DOCUMENTS RECEIPT DATE'])}")
  if nz(r.get("Comments")): p.append(nz(r["Comments"]))
  ex = [nz(r.get(c)) for c in unn if nz(r.get(c))]
  if ex: p.append("Notes: " + " | ".join(ex))
  return " | ".join(p)
leg = pd.DataFrame({
  "PURCHASE ORDER": d1["PO"].map(lambda v: nz(v) if nz(v).lower() not in PH else "NO PO"),
  "SUPPLIER /MANUFACTURER": d1["SUPPLIER"], "Cargo description": d1["PRODUCT"], "Invoice number": d1["INV NO"],
  "Quantity": d1["QUANTITY"], "Container Size": d1["No of containers"], "BILL OF LADING OR AIRWAY BILL": d1["BL/AWB No"],
  "POD": d1["POD"], "ETA POD": d1["ETA PORTS"], "Cargo ETA RMS": d1["ETA RMS"], "Actual arrival at RMS": d1["date of receipt at RMS"],
  "DATE OF SHARING SHIPPING DOCUMENTS WITH RMS WAREHOUSE RECEIVING/FLS": d1["SHARED DATE WITH CLEARING AGENT/SUPPLIER /DP WORLD"],
  "Responsible staff": d1["Responsible staff"], "Remarks": [remarks(r) for _, r in d1.iterrows()],
})
leg["Shipping mode"] = [infer_mode(p, "", b) for p, b in zip(leg["POD"], leg["BILL OF LADING OR AIRWAY BILL"])]
leg["Current status"] = [status_of(nz(a), t) for a, t in zip(leg["Actual arrival at RMS"], d1["Comments"])]
leg["__fy"] = [fy_of(a, e, default="2024-2025") for a, e in zip(leg["Actual arrival at RMS"], leg["Cargo ETA RMS"])]

cur = {(re.sub(r"\s+", " ", nz(x[A[0]])).upper(), nz(x[A[3]]).upper()) for x in m["rows"]}
for name, frame in (("Delivered-2nd_old sheet", d2), ("Delivered-old sheet", leg)):
  skipped = 0
  for _, r in frame.iterrows():
    if r["__fy"] == CURRENT_FY and (re.sub(r"\s+", " ", nz(r["PURCHASE ORDER"])).upper(), nz(r["Invoice number"]).upper()) in cur:
      skipped += 1; continue
    v = {}
    for ui in frame.columns:
      if ui == "__fy" or ui in COMPUTED or ui not in REV: continue
      cv = clean(ui, r[ui] if ui == A[4] and nz(r[ui]) == "" else (nz(r[ui]) if ui not in DATE_UI and ui != A[4] else r[ui]))
      if cv is not None and cv != "": v[REV[ui]] = cv
    records.append({"kind": "history", "fy": r["__fy"], "src": name, "bl": nz(r.get("BILL OF LADING OR AIRWAY BILL")), "inv": nz(r.get("Invoice number")), "values": v})
  if skipped: print(f"{name}: {skipped} rows already in {CURRENT_FY}, not duplicated")

seed = {"version": 1, "fiscal_years": sorted({r["fy"] for r in records}), "records": records}
os.makedirs("seed", exist_ok=True)
json.dump(seed, open("seed/seed_data.json", "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"), default=str)
import collections
print(dict(collections.Counter((r["fy"], r["kind"]) for r in records)))
