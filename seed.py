"""One-time, non-destructive data migration, applied automatically at start-up.

It merges the Excel tracker data (seed/seed_data.json) into whatever database
is live at that moment:

  * a full backup is taken first (backups/pre_migration_*.db)
  * nothing is ever deleted; shipments you deleted in the app stay deleted
  * a cell the tracker would change is only overwritten when it is blank (or
    "TBA"/"pending") in the live database, or still holds the value it had in
    the 26-Oct snapshot. A value somebody edited since then is kept, and the
    tracker's value is written to that shipment's Edit Trail instead.
  * every change is written to the Edit Trail with the old value
  * it runs once (recorded in app_meta) and is safe to re-run
"""
import json
import os
import re
from datetime import datetime

import db

SEED_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "seed", "seed_data.json")
META_KEY = "seed_data_v1"
WHO = "Data migration"


def _norm(v):
  if v is None:
    return ""
  if isinstance(v, (int, float)):
    return f"{float(v):.4f}"
  return " ".join(str(v).split()).upper()


def _blank(v):
  return db._is_placeholder(v) or (isinstance(v, str) and not v.strip())


def _po_key(p):
  q = " ".join(str(p or "").split()).upper()
  return "NO PO" if q in ("", "NO PO", "NOPO", "NA", "N/A", "TBA", "NONE", "PENDING", "-") else q


def _already_applied(conn):
  return conn.execute("SELECT 1 FROM app_meta WHERE key = ?", (META_KEY,)).fetchone() is not None


def apply_pending():
  """Returns None when there is nothing to do, else a summary dict."""
  if not os.path.exists(SEED_FILE):
    return None
  conn = db.get_conn()
  try:
    if _already_applied(conn):
      return None
  finally:
    conn.close()

  with open(SEED_FILE, encoding="utf-8") as fh:
    seed = json.load(fh)

  backup = None
  try:
    backup = db.create_backup("pre_migration")
  except Exception:
    pass

  conn = db.get_conn()
  c = conn.cursor()
  st = {"updated_rows": 0, "cells_updated": 0, "cells_kept": 0, "added": 0, "history_added": 0,
        "skipped_deleted": 0, "already_present": 0, "backup": os.path.basename(backup) if backup else ""}
  try:
    cols_ok = {r[1] for r in c.execute("PRAGMA table_info(shipments)")}
    ui = db.RENAME_MAP
    for fy in seed["fiscal_years"]:
      c.execute("INSERT OR IGNORE INTO fiscal_years (year) VALUES (?)", (fy,))

    # live rows per fiscal year (as they are right now)
    live = {}
    def load_fy(fy):
      if fy not in live:
        c.execute("SELECT * FROM shipments WHERE fiscal_year = ?", (fy,))
        names = [d[0] for d in c.description]
        live[fy] = [dict(zip(names, r)) for r in c.fetchall()]
      return live[fy]
    claimed = set()

    def candidates(fy, po):
      k = _po_key(po)
      return [(r["id"], r["bl_awb"], r["invoice_no"]) for r in load_fy(fy)
              if r["id"] not in claimed and _po_key(r["po_number"]) == k]

    def find(rec, first_bl, first_inv, second_bl=None, second_inv=None):
      po = rec["values"].get("po_number") or rec.get("full", {}).get("po_number") or rec.get("po", "")
      cands = candidates(rec["fy"], po)
      if not cands:
        return None
      hit = db.pick_existing(cands, first_bl, first_inv)
      if hit is None and second_bl is not None:
        hit = db.pick_existing(cands, second_bl, second_inv)
      return hit

    def row_by_id(fy, rid):
      return next(r for r in load_fy(fy) if r["id"] == rid)

    def insert(rec, values):
      d = {k: v for k, v in values.items() if k in cols_ok and v not in (None, "")}
      d["fiscal_year"] = rec["fy"]
      d["is_deleted"] = 0
      d.setdefault("current_status", "In Transit")
      if d.get("shipping_mode"):
        d["shipping_mode"] = str(d["shipping_mode"]).strip().upper()
      keys = list(d)
      c.execute(f"INSERT INTO shipments ({','.join(chr(34) + k + chr(34) for k in keys)}) VALUES ({','.join('?' * len(keys))})",
                [d[k] for k in keys])
      rid = c.lastrowid
      db.log_row_change(rid, "Shipment Created", "", f"PO #{d.get('po_number', '')} (tracker import)", WHO, conn=conn)
      return rid

    def apply_cells(row, rec, new_vals, base_vals, fill_only=False):
      """Write tracker values into a live row, protecting later live edits."""
      sets = {}
      for col, nv in new_vals.items():
        if col not in cols_ok or col == "po_number":
          continue
        cur = row.get(col)
        base = base_vals.get(col)
        take = _blank(cur) or (not fill_only and _norm(cur) == _norm(base))
        if _norm(cur) == _norm(nv):
          continue
        if take:
          sets[col] = nv
        else:
          st["cells_kept"] += 1
          db.log_row_change(row["id"], f"MIGRATION kept your value - {ui.get(col, col)}", cur, nv, WHO, conn=conn)
      if sets:
        c.execute(f"UPDATE shipments SET {', '.join(chr(34) + k + chr(34) + ' = ?' for k in sets)} WHERE id = ?",
                  list(sets.values()) + [row["id"]])
        for col, nv in sets.items():
          db.log_row_change(row["id"], ui.get(col, col), row.get(col), nv, WHO, conn=conn)
          row[col] = nv
        st["cells_updated"] += len(sets)
        st["updated_rows"] += 1

    recs = seed["records"]
    # 1) existing shipments the tracker updates
    for rec in [r for r in recs if r["kind"] == "update"]:
      hit = find(rec, rec["base_bl"], rec["base_inv"], rec["bl"], rec["inv"])
      if hit is None:
        insert(rec, rec["full"]); st["added"] += 1
        continue
      claimed.add(hit)
      row = row_by_id(rec["fy"], hit)
      if row.get("is_deleted"):
        st["skipped_deleted"] += 1
        continue
      apply_cells(row, rec, rec["values"], rec["base"])
    # 2) shipments only in the tracker
    for rec in [r for r in recs if r["kind"] == "new"]:
      hit = find(rec, rec["bl"], rec["inv"])
      if hit is None:
        rid = insert(rec, rec["values"]); st["added"] += 1
        claimed.add(rid)
        continue
      claimed.add(hit)
      row = row_by_id(rec["fy"], hit)
      if row.get("is_deleted"):
        st["skipped_deleted"] += 1
      else:
        st["already_present"] += 1
        apply_cells(row, rec, rec["values"], {}, fill_only=True)
    # 3) older sheets (separate fiscal years) - added as their own rows, never merged.
    #    Rows already in the database from an earlier load are skipped by count, so
    #    genuine repeated lines (same PO/invoice/BL) are all kept.
    from collections import Counter
    def hkey(fy, po, inv, bl, eta):
      return (fy, _po_key(po), _norm(inv), _norm(bl), _norm(eta))
    have = Counter()
    for fy in {r["fy"] for r in recs if r["kind"] == "history"}:
      for r in load_fy(fy):
        have[hkey(fy, r["po_number"], r["invoice_no"], r["bl_awb"], r["eta_pod"])] += 1
    for rec in [r for r in recs if r["kind"] == "history"]:
      v = rec["values"]
      k = hkey(rec["fy"], v.get("po_number"), v.get("invoice_no"), v.get("bl_awb"), v.get("eta_pod"))
      if have[k] > 0:
        have[k] -= 1
        st["already_present"] += 1
        continue
      insert(rec, v); st["history_added"] += 1

    db.recompute_all_formulas(conn)
    c.execute("INSERT OR REPLACE INTO app_meta (key, value) VALUES (?, ?)",
              (META_KEY, json.dumps({**st, "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})))
    conn.commit()
  except Exception as exc:
    conn.rollback()
    conn.close()
    try:
      db.log_action(f"Data migration FAILED and was rolled back (nothing changed): {exc}")
    except Exception:
      pass
    return {"error": str(exc), "backup": st["backup"]}
  conn.close()
  try:
    db.log_action(
        f"Data migration applied: {st['updated_rows']} shipments updated ({st['cells_updated']} cells), "
        f"{st['added']} added, {st['history_added']} historic rows, {st['cells_kept']} newer live values kept")
  except Exception:
    pass
  return st
