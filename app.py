from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, session, g, has_request_context
import random
import os
import json
import uuid
import calendar
import shutil
import zipfile
import tempfile
import io
from datetime import datetime
from collections import defaultdict
from pathlib import Path
from contextlib import contextmanager
import time
import copy

import pandas as pd
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or "finance-schedule-secret-9527"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_DUTY = os.path.join(BASE_DIR, "members.xlsx")
FILE_LEAVE = os.path.join(BASE_DIR, "attendance.xlsx")
WORK_SCHEME_FILE = os.path.join(BASE_DIR, "work_schemes.json")
HISTORY_FILE = os.path.join(BASE_DIR, "work_assignment_history.xlsx")
DUTY_HISTORY_FILE = os.path.join(BASE_DIR, "duty_schedule_history.xlsx")
PLAN_STATUS_FILE = os.path.join(BASE_DIR, "monthly_plan_status.json")
OPERATION_LOG_FILE = os.path.join(BASE_DIR, "operation_log.xlsx")
ADMIN_PASSWORD = "9527"
SYSTEM_SETTINGS_FILE = os.path.join(BASE_DIR, "system_settings.json")
AUTO_BACKUP_DIR = os.path.join(BASE_DIR, "auto_backups")
WEEKLY_CLEAN_FILE = os.path.join(BASE_DIR, "weekly_clean_status.json")
WORK_ITEMS_FILE = os.path.join(BASE_DIR, "work_items.json")

# ======== 資料庫層（Neon PostgreSQL）========
# 設定環境變數 DATABASE_URL 即啟用資料庫；未設定時自動沿用原本的 Excel/JSON 檔案。
DATABASE_URL = os.environ.get("DATABASE_URL", "")
_PG = None


def db_enabled():
    return bool(DATABASE_URL)


def _pg():
    """延遲載入 psycopg2，未安裝或本機未用 DB 時不影響其他功能。"""
    global _PG
    if _PG is None:
        import psycopg2
        import psycopg2.extras  # noqa: F401
        _PG = psycopg2
    return _PG


def get_conn():
    return _pg().connect(DATABASE_URL)


@contextmanager
def pg_conn():
    """同一個請求內共用一條連線（請求結束才關），大幅減少跨區連線握手次數。
    非請求情境（啟動時建表／遷移）則開完即關。"""
    in_request = False
    try:
        in_request = has_request_context()
    except Exception:
        in_request = False

    if in_request:
        conn = getattr(g, "_db_conn", None)
        if conn is None or getattr(conn, "closed", 1):
            conn = get_conn()
            g._db_conn = conn
        try:
            yield conn
            conn.commit()
        except Exception:
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        # 不在這裡關閉，交由 teardown_request 在請求結束時統一關閉
    else:
        conn = get_conn()
        try:
            yield conn
            conn.commit()
        except Exception:
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        finally:
            try:
                conn.close()
            except Exception:
                pass


@app.teardown_request
def _close_request_db_conn(exc):
    conn = getattr(g, "_db_conn", None)
    if conn is not None:
        try:
            conn.close()
        except Exception:
            pass
        try:
            g._db_conn = None
        except Exception:
            pass


def init_db():
    """建立所有資料表（冪等），並在資料庫為空時自動從現有檔案匯入一次。"""
    if not db_enabled():
        return
    try:
        with pg_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS kv_store (
                        k TEXT PRIMARY KEY,
                        v JSONB NOT NULL
                    );
                    CREATE TABLE IF NOT EXISTS members (
                        name TEXT NOT NULL, dept TEXT NOT NULL, shift TEXT NOT NULL,
                        status TEXT DEFAULT '啟用', note TEXT DEFAULT '',
                        PRIMARY KEY (name, dept, shift)
                    );
                    CREATE TABLE IF NOT EXISTS leaves (
                        year INT NOT NULL, month INT NOT NULL, name TEXT NOT NULL,
                        day INT NOT NULL, leave_type TEXT DEFAULT '',
                        PRIMARY KEY (year, month, name, day)
                    );
                    CREATE TABLE IF NOT EXISTS duty_members (
                        name TEXT NOT NULL, shift TEXT NOT NULL,
                        PRIMARY KEY (name, shift)
                    );
                    CREATE TABLE IF NOT EXISTS duty_history (
                        id BIGSERIAL PRIMARY KEY,
                        date TEXT, year INT, month INT, day INT, weekday TEXT,
                        shift TEXT, duty_person TEXT, created_at TEXT, batch_id TEXT
                    );
                    CREATE TABLE IF NOT EXISTS operation_log (
                        id BIGSERIAL PRIMARY KEY,
                        ts TEXT, role TEXT, action TEXT, dept TEXT, shift TEXT,
                        month TEXT, content TEXT, reason TEXT
                    );
                    CREATE TABLE IF NOT EXISTS work_history (
                        id BIGSERIAL PRIMARY KEY,
                        date TEXT, year INT, month INT, day INT,
                        dept TEXT, shift TEXT, person TEXT, task TEXT, scheme TEXT,
                        method TEXT, created_at TEXT, batch_id TEXT
                    );
                    CREATE INDEX IF NOT EXISTS idx_work_history_ym ON work_history (year, month, dept, shift);
                    CREATE INDEX IF NOT EXISTS idx_leaves_ym ON leaves (year, month);
                    CREATE INDEX IF NOT EXISTS idx_duty_history_ym ON duty_history (year, month, shift);
                """)
            conn.commit()
        migrate_all_if_needed()
    except Exception as e:
        print(f"初始化資料庫失敗：{e}")


def kv_get(key, default=None):
    if not db_enabled():
        return default
    try:
        with pg_conn() as conn, conn.cursor() as cur:
            cur.execute("SELECT v FROM kv_store WHERE k=%s", (key,))
            row = cur.fetchone()
        if not row or row[0] is None:
            return default
        val = row[0]
        if isinstance(val, (dict, list)):
            return val
        if isinstance(val, str):
            try:
                return json.loads(val)
            except Exception:
                return default
        return default
    except Exception as e:
        print(f"讀取設定（{key}）失敗：{e}")
        return default


def kv_put(key, data):
    if not db_enabled():
        return False
    try:
        with pg_conn() as conn, conn.cursor() as cur:
            cur.execute(
                "INSERT INTO kv_store (k, v) VALUES (%s, %s::jsonb) ON CONFLICT (k) DO UPDATE SET v = EXCLUDED.v",
                (key, json.dumps(data, ensure_ascii=False)),
            )
            conn.commit()
        return True
    except Exception as e:
        print(f"儲存設定（{key}）失敗：{e}")
        return False


def _table_empty(table):
    try:
        with pg_conn() as conn, conn.cursor() as cur:
            cur.execute(f"SELECT COUNT(*) FROM {table}")
            return int(cur.fetchone()[0]) == 0
    except Exception:
        return False


def _migrate_attendance_from_excel():
    with pd.ExcelFile(FILE_LEAVE, engine="openpyxl") as xls:
        df_basic = pd.read_excel(xls, sheet_name=SHEET_BASIC) if SHEET_BASIC in xls.sheet_names else pd.DataFrame(columns=["姓名", "部門", "班別", "狀態", "備註"])
        leave_sheets = [s for s in xls.sheet_names if str(s).startswith("Leave_")]
        if not leave_sheets:
            now = datetime.now()
            save_attendance_frames(df_basic, pd.DataFrame({"姓名": df_basic.get("姓名", pd.Series(dtype=str))}), now.year, now.month)
            return
        for s in leave_sheets:
            try:
                parts = str(s).split("_")
                yy, mm = int(parts[1]), int(parts[2])
            except Exception:
                continue
            df_leave = pd.read_excel(xls, sheet_name=s)
            save_attendance_frames(df_basic, df_leave, yy, mm)


def migrate_all_if_needed():
    """各資料表為空、且本機尚有對應舊檔時，匯入一次（首次切換到資料庫用）。"""
    if not db_enabled():
        return
    try:
        if _table_empty("work_history") and os.path.exists(HISTORY_FILE):
            df = pd.read_excel(HISTORY_FILE, sheet_name=SHEET_HISTORY)
            if df is not None and not df.empty:
                _save_history_df_to_db(df)
                print(f"已匯入工作分配歷史 {len(df)} 筆")
    except Exception as e:
        print(f"工作分配歷史遷移失敗：{e}")
    try:
        if _table_empty("members") and os.path.exists(FILE_LEAVE):
            _migrate_attendance_from_excel()
            print("已匯入假表與人員")
    except Exception as e:
        print(f"假表遷移失敗：{e}")
    try:
        if _table_empty("duty_members") and os.path.exists(FILE_DUTY):
            df = pd.read_excel(FILE_DUTY)
            if "姓名" in df.columns:
                if "班別" not in df.columns:
                    df["班別"] = default_duty_shift()
                save_duty_members_df(df)
                print("已匯入值日生名單")
    except Exception as e:
        print(f"值日生名單遷移失敗：{e}")
    try:
        if _table_empty("duty_history") and os.path.exists(DUTY_HISTORY_FILE):
            df = pd.read_excel(DUTY_HISTORY_FILE)
            if df is not None and not df.empty:
                save_duty_history_df(df)
                print("已匯入值日生歷史")
    except Exception as e:
        print(f"值日生歷史遷移失敗：{e}")
    try:
        if _table_empty("operation_log") and os.path.exists(OPERATION_LOG_FILE):
            df = pd.read_excel(OPERATION_LOG_FILE)
            if df is not None and not df.empty:
                save_operation_log_df(df)
                print("已匯入操作紀錄")
    except Exception as e:
        print(f"操作紀錄遷移失敗：{e}")
    for key, path in [("system_settings", SYSTEM_SETTINGS_FILE), ("work_schemes", WORK_SCHEME_FILE), ("plan_status", PLAN_STATUS_FILE), ("weekly_clean", WEEKLY_CLEAN_FILE)]:
        try:
            if kv_get(key) is None and os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                kv_put(key, data)
                print(f"已匯入設定 {key}")
        except Exception as e:
            print(f"{key} 遷移失敗：{e}")
# ======== 資料庫層結束 ========



SHEET_BASIC = "Sheet1"
SHEET_LEAVE = "LeaveSheet"
SHEET_HISTORY = "History"

DEPARTMENTS = ["財務", "BBP", "四方"]
SHIFTS = ["早班", "中班", "晚班"]
LEAVE_TYPES = ["休", "事", "病", "特", "年", "加", "公", "產", "檢", "育", "婚", "喪", "理", "福"]


def default_system_settings():
    return {
        "system_title": "財務部門智能排班系統",
        "admin_password": "9527",
        "departments": ["財務", "BBP", "四方"],
        "shifts": [
            {"name": "早班", "start": "08:00", "end": "16:00", "duty_count": 1},
            {"name": "中班", "start": "16:00", "end": "24:00", "duty_count": 1},
            {"name": "晚班", "start": "00:00", "end": "08:00", "duty_count": 1}
        ],
        "weekly_cleaner_count": 3,
        "auto_backup_enabled": True,
        "auto_backup_keep": 10
    }


_SETTINGS_CACHE = {"data": None, "ts": 0.0}
_SETTINGS_TTL = 30.0


def load_system_settings():
    if _SETTINGS_CACHE["data"] is not None and (time.time() - _SETTINGS_CACHE["ts"]) < _SETTINGS_TTL:
        return copy.deepcopy(_SETTINGS_CACHE["data"])
    data = default_system_settings()
    saved = None
    if db_enabled():
        saved = kv_get("system_settings")
    elif os.path.exists(SYSTEM_SETTINGS_FILE):
        try:
            with open(SYSTEM_SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
        except Exception as e:
            print(f"讀取系統設定失敗：{e}")
    if isinstance(saved, dict):
        data.update(saved)
    # 清理資料
    data["system_title"] = clean_text(data.get("system_title", "")) if 'clean_text' in globals() else str(data.get("system_title", "")).strip()
    if not data["system_title"]:
        data["system_title"] = "財務部門智能排班系統"
    data["admin_password"] = str(data.get("admin_password", "9527")).strip() or "9527"
    departments = []
    raw_depts = data.get("departments", [])
    if isinstance(raw_depts, str):
        raw_depts = raw_depts.replace("，", ",").replace("\n", ",").split(",")
    for d in raw_depts:
        d = str(d).strip()
        if d and d not in departments:
            departments.append(d)
    data["departments"] = departments or ["財務", "BBP", "四方"]
    shifts = []
    for item in data.get("shifts", []):
        if isinstance(item, dict):
            name = str(item.get("name", "")).strip()
            start = str(item.get("start", "00:00")).strip() or "00:00"
            end = str(item.get("end", "08:00")).strip() or "08:00"
            dc = item.get("duty_count", 1)
            cc = item.get("cleaner_count", data.get("weekly_cleaner_count", 3))
        else:
            name, start, end, dc, cc = str(item).strip(), "00:00", "08:00", 1, 3
        try:
            dc = max(0, int(dc))
        except Exception:
            dc = 1
        try:
            cc = max(1, int(cc))
        except Exception:
            cc = 3
        if name and name not in [s.get("name") for s in shifts]:
            shifts.append({"name": name, "start": start, "end": end, "duty_count": dc, "cleaner_count": cc})
    data["shifts"] = shifts or default_system_settings()["shifts"]
    try:
        data["weekly_cleaner_count"] = max(1, int(data.get("weekly_cleaner_count", 3)))
    except Exception:
        data["weekly_cleaner_count"] = 3
    data["auto_backup_enabled"] = bool(data.get("auto_backup_enabled", True))
    try:
        data["auto_backup_keep"] = max(3, int(data.get("auto_backup_keep", 10)))
    except Exception:
        data["auto_backup_keep"] = 10
    _SETTINGS_CACHE["data"] = copy.deepcopy(data)
    _SETTINGS_CACHE["ts"] = time.time()
    return data


def save_system_settings(data):
    try:
        if db_enabled():
            kv_put("system_settings", data)
        else:
            with open(SYSTEM_SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        _SETTINGS_CACHE["data"] = None
        apply_system_settings(data)
        return True, "系統設定已儲存"
    except Exception as e:
        return False, str(e)


def apply_system_settings(data=None):
    global ADMIN_PASSWORD, DEPARTMENTS, SHIFTS
    data = data or load_system_settings()
    ADMIN_PASSWORD = str(data.get("admin_password", "9527")).strip() or "9527"
    DEPARTMENTS = list(data.get("departments", ["財務", "BBP", "四方"]))
    SHIFTS = [s.get("name") for s in data.get("shifts", []) if s.get("name")] or ["早班", "中班", "晚班"]
    return data


def verify_admin_password(pwd):
    """相容舊明碼與新雜湊：設定值是雜湊就用 check_password_hash，否則明碼比對。"""
    stored = ADMIN_PASSWORD
    pwd = "" if pwd is None else str(pwd)
    if isinstance(stored, str) and (stored.startswith("pbkdf2:") or stored.startswith("scrypt:")):
        try:
            return check_password_hash(stored, pwd)
        except Exception:
            return False
    return pwd == stored


def shift_time_label(shift_name):
    data = load_system_settings()
    for s in data.get("shifts", []):
        if s.get("name") == shift_name:
            return f"{s.get('start','')}～{s.get('end','')}"
    return ""


def daily_duty_count_for_shift(shift):
    """該班別每日要排幾位值日生（系統設定，預設 1）。"""
    for s in load_system_settings().get("shifts", []):
        if s.get("name") == shift:
            try:
                return max(0, int(s.get("duty_count", 1)))
            except Exception:
                return 1
    return 1


def make_auto_backup(reason="auto"):
    settings = load_system_settings()
    if not settings.get("auto_backup_enabled", True):
        return ""
    os.makedirs(AUTO_BACKUP_DIR, exist_ok=True)
    path = os.path.join(AUTO_BACKUP_DIR, f"auto_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{reason}.zip")
    files = [FILE_DUTY, FILE_LEAVE, WORK_SCHEME_FILE, HISTORY_FILE, DUTY_HISTORY_FILE, PLAN_STATUS_FILE, OPERATION_LOG_FILE, SYSTEM_SETTINGS_FILE, WEEKLY_CLEAN_FILE]
    try:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in files:
                if os.path.exists(f):
                    zf.write(f, arcname=os.path.basename(f))
            zf.writestr("backup_info.txt", f"自動備份\n原因：{reason}\n建立時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        keep = int(settings.get("auto_backup_keep", 10))
        backups = sorted(Path(AUTO_BACKUP_DIR).glob("auto_backup_*.zip"), key=lambda x: x.stat().st_mtime, reverse=True)
        for old in backups[keep:]:
            try: old.unlink()
            except Exception: pass
        return path
    except Exception as e:
        print(f"自動備份失敗：{e}")
        return ""


def auto_backup_summary():
    os.makedirs(AUTO_BACKUP_DIR, exist_ok=True)
    backups = sorted(Path(AUTO_BACKUP_DIR).glob("auto_backup_*.zip"), key=lambda x: x.stat().st_mtime, reverse=True)
    latest = backups[0] if backups else None
    return {
        "count": len(backups),
        "latest_name": latest.name if latest else "",
        "latest_time": datetime.fromtimestamp(latest.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S") if latest else ""
    }


apply_system_settings(load_system_settings())

def is_admin():
    return session.get("admin_logged_in") is True


def current_role():
    return "管理員" if is_admin() else "查看者"


def require_admin_json():
    if not is_admin():
        return jsonify({"status": "error", "message": "需要管理員權限，請先輸入管理員密碼 9527 登入。"})
    return None


def require_admin_redirect():
    """表單／連結類操作用：未登入時導回首頁並顯示提示，而不是回傳 JSON 蓋掉畫面。"""
    if not is_admin():
        return redirect(url_for("index", login_msg="需要管理員權限，請先在右上角用密碼登入後再操作。"))
    return None


def operation_log_columns():
    return ["時間", "使用者權限", "操作類型", "部門", "班別", "月份", "內容", "原因"]


def load_operation_log_df():
    cols = operation_log_columns()
    if db_enabled():
        try:
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("SELECT ts, role, action, dept, shift, month, content, reason FROM operation_log ORDER BY id")
                rows = cur.fetchall()
            df = pd.DataFrame(rows, columns=cols)
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
            return df[cols]
        except Exception as e:
            print(f"讀取操作紀錄失敗（DB）：{e}")
            return pd.DataFrame(columns=cols)
    if os.path.exists(OPERATION_LOG_FILE):
        try:
            df = pd.read_excel(OPERATION_LOG_FILE)
        except Exception:
            df = pd.DataFrame(columns=cols)
    else:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


def save_operation_log_df(df):
    if db_enabled():
        try:
            import psycopg2.extras as extras
            rows = []
            for _, r in df.iterrows():
                rows.append((
                    clean_text(r.get("時間", "")), clean_text(r.get("使用者權限", "")), clean_text(r.get("操作類型", "")),
                    clean_text(r.get("部門", "")), clean_text(r.get("班別", "")), clean_text(r.get("月份", "")),
                    clean_text(r.get("內容", "")), clean_text(r.get("原因", "")),
                ))
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("TRUNCATE operation_log")
                if rows:
                    extras.execute_values(cur, "INSERT INTO operation_log (ts, role, action, dept, shift, month, content, reason) VALUES %s", rows)
                conn.commit()
            return True, ""
        except Exception as e:
            return False, f"寫入資料庫失敗：{e}"
    try:
        df.to_excel(OPERATION_LOG_FILE, index=False)
        return True, ""
    except Exception as e:
        return False, ensure_excel_not_open_error(e) if 'ensure_excel_not_open_error' in globals() else str(e)


def log_operation(action, department="", shift="", year=None, month=None, content="", reason=""):
    try:
        selected, _ = now_info(year, month) if 'now_info' in globals() else (datetime.now(), 0)
        row = {
            "時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "使用者權限": current_role(),
            "操作類型": clean_text(action) if 'clean_text' in globals() else str(action),
            "部門": clean_text(department) if 'clean_text' in globals() else str(department),
            "班別": clean_text(shift) if 'clean_text' in globals() else str(shift),
            "月份": f"{selected.year}-{selected.month:02d}",
            "內容": clean_text(content) if 'clean_text' in globals() else str(content),
            "原因": clean_text(reason) if 'clean_text' in globals() else str(reason),
        }
        df = load_operation_log_df()
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        save_operation_log_df(df)
    except Exception as e:
        print(f"寫入操作紀錄失敗：{e}")


def now_info(year=None, month=None):
    now = datetime.now()
    y = int(year or now.year)
    m = int(month or now.month)
    if m < 1 or m > 12:
        y, m = now.year, now.month
    _, num_days = calendar.monthrange(y, m)
    # 保留目前時間，但把年月切成使用者選擇的月份
    selected = now.replace(year=y, month=m, day=min(now.day, num_days))
    return selected, num_days


def leave_sheet_name(year=None, month=None):
    selected, _ = now_info(year, month)
    return f"Leave_{selected.year}_{selected.month:02d}"



def plan_key(year, month, department, shift):
    selected, _ = now_info(year, month)
    return f"{selected.year}_{selected.month:02d}__{department}__{shift}"


def load_plan_status():
    if db_enabled():
        d = kv_get("plan_status")
        return d if isinstance(d, dict) else {}
    if os.path.exists(PLAN_STATUS_FILE):
        try:
            with open(PLAN_STATUS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}
    return {}


def save_plan_status(data):
    try:
        if db_enabled():
            kv_put("plan_status", data)
        else:
            with open(PLAN_STATUS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        return True, ""
    except Exception as e:
        return False, str(e)


def load_work_items_map():
    """工作項目庫：dict { group_key: [ {"name": str}, ... ] }，每個部門×班別各自一份。"""
    raw = None
    if db_enabled():
        raw = kv_get("work_items")
    elif os.path.exists(WORK_ITEMS_FILE):
        try:
            with open(WORK_ITEMS_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except Exception:
            raw = None
    result = {}
    if isinstance(raw, dict):
        for k, v in raw.items():
            items, seen = [], set()
            if isinstance(v, list):
                for it in v:
                    name = clean_text(it.get("name", "")) if isinstance(it, dict) else clean_text(it)
                    if name and name not in seen:
                        seen.add(name); items.append({"name": name})
            result[str(k)] = items
    return result


def load_work_items(department, shift):
    return load_work_items_map().get(group_key(department, shift), [])


def build_item_stats(department, shift, year, month):
    """本月該部門班別：人員 × 工作項目標籤 的接到次數。
    欄＝項目庫標籤（固定列出，含 0 次）＋實際出現過但不在項目庫的標籤（附在後面）。
    列＝本月該組實際有被排到工作的人員。
    """
    items = load_work_items(department, shift)
    item_names = []
    for it in items:
        nm = clean_text(it.get("name", "")) if isinstance(it, dict) else clean_text(it)
        if nm and nm not in item_names:
            item_names.append(nm)
    stats = {}
    persons = []
    try:
        df = load_history_df()
    except Exception:
        df = None
    if df is not None and not df.empty:
        try:
            sub = df[
                (pd.to_numeric(df["年份"], errors="coerce") == year)
                & (pd.to_numeric(df["月份"], errors="coerce") == month)
                & (df["部門"].astype(str) == str(department))
                & (df["班別"].astype(str) == str(shift))
            ]
            for _, row in sub.iterrows():
                person = clean_text(row.get("人員", ""))
                work = clean_text(row.get("工作", ""))
                if not person:
                    continue
                if person not in stats:
                    stats[person] = {}
                    persons.append(person)
                for tag in work.split():
                    tag = tag.strip()
                    if tag:
                        stats[person][tag] = stats[person].get(tag, 0) + 1
        except Exception as e:
            print(f"統計工作項目失敗：{e}")
    persons.sort()
    extra_tags = []
    for p in stats:
        for tag in stats[p]:
            if tag not in item_names and tag not in extra_tags:
                extra_tags.append(tag)
    extra_tags.sort()
    columns = item_names + extra_tags
    return {"columns": columns, "persons": persons, "stats": stats, "extra": extra_tags}


def build_duty_stats(shift, year, month):
    """本月該班別：每人當值日生的次數（值日生名單裡 0 次的人也列出）。
    回傳 [(姓名, 次數), ...]，依次數高到低、同次數依姓名。"""
    counts = {}
    try:
        df = load_duty_history_df()
    except Exception:
        df = None
    if df is not None and not df.empty:
        try:
            sub = df[
                (pd.to_numeric(df["年份"], errors="coerce") == year)
                & (pd.to_numeric(df["月份"], errors="coerce") == month)
                & (df["班別"].astype(str) == str(shift))
            ]
            for _, row in sub.iterrows():
                person = clean_text(row.get("值日生", ""))
                if person:
                    counts[person] = counts.get(person, 0) + 1
        except Exception as e:
            print(f"統計值日生失敗：{e}")
    try:
        for m in load_duty_members(shift):
            m = clean_text(m)
            if m and m not in counts:
                counts[m] = 0
    except Exception:
        pass
    return sorted(counts.items(), key=lambda x: (-x[1], x[0]))


def build_combined_stats(department, shift, year, month):
    """合併『工作項目統計』與『值日生次數』成同一張表。
    列＝該部門班別有被排工作的人（只算本部門班別，不混入其他部門）。
    欄＝工作項目們＋工作合計＋值日生次數。
    值日生雖按班別輪，但這裡只對「本部門班別的人」顯示其值日次數，
    不會把別部門的晚班值日生塞進本部門的表。"""
    item = build_item_stats(department, shift, year, month)
    duty = dict(build_duty_stats(shift, year, month))
    persons = sorted(item["persons"])
    return {
        "columns": item["columns"],
        "persons": persons,
        "stats": item["stats"],
        "duty": duty,
        "extra": item["extra"],
    }


def save_work_items_map(m):
    clean = {}
    for k, v in (m or {}).items():
        items, seen = [], set()
        for it in (v or []):
            name = clean_text(it.get("name", "")) if isinstance(it, dict) else clean_text(it)
            if name and name not in seen:
                seen.add(name); items.append({"name": name})
        clean[str(k)] = items
    try:
        if db_enabled():
            kv_put("work_items", clean)
        else:
            with open(WORK_ITEMS_FILE, "w", encoding="utf-8") as f:
                json.dump(clean, f, ensure_ascii=False, indent=2)
        return True, ""
    except Exception as e:
        return False, str(e)


def get_plan_status(year, month, department, shift):
    data = load_plan_status()
    item = data.get(plan_key(year, month, department, shift), {})
    status = clean_text(item.get("status", "草稿中")) if isinstance(item, dict) else "草稿中"
    if status not in ["草稿中", "已發布", "已鎖定"]:
        status = "草稿中"
    return {
        "status": status,
        "updated_at": clean_text(item.get("updated_at", "")) if isinstance(item, dict) else "",
        "note": clean_text(item.get("note", "")) if isinstance(item, dict) else ""
    }


def set_plan_status(year, month, department, shift, status, note=""):
    if status not in ["草稿中", "已發布", "已鎖定"]:
        return False, "狀態錯誤"
    data = load_plan_status()
    data[plan_key(year, month, department, shift)] = {
        "status": status,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "note": clean_text(note)
    }
    return save_plan_status(data)


def is_plan_locked(year, month, department, shift):
    return get_plan_status(year, month, department, shift).get("status") == "已鎖定"


def clean_text(value):
    val = str(value).strip()
    if val in ["", "nan", "NaN", "None", "NaT"]:
        return ""
    if val.endswith(".0"):
        val = val[:-2]
    return val


def clean_leave_value(value):
    val = clean_text(value)
    return val if val in LEAVE_TYPES else val


def parse_tasks(raw_tasks):
    tasks = []
    for line in str(raw_tasks).replace("，", ",").splitlines():
        for part in line.split(","):
            task = part.strip()
            if task:
                tasks.append(task)
    return tasks


def ensure_excel_not_open_error(e):
    return f"寫入失敗，請確認 Excel 檔案沒有開著：{str(e)}"


def default_duty_shift():
    return SHIFTS[0] if SHIFTS else "晚班"


def load_duty_members_df():
    """值日生名單：姓名 + 班別兩欄。舊版只有姓名時，補上預設班別。"""
    cols = ["姓名", "班別"]
    if db_enabled():
        try:
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("SELECT name, shift FROM duty_members ORDER BY shift, name")
                rows = cur.fetchall()
            if rows:
                df = pd.DataFrame(rows, columns=cols)
                df["姓名"] = df["姓名"].fillna("").astype(str).str.strip()
                df["班別"] = df["班別"].fillna("").astype(str).str.strip()
                df = df[df["姓名"] != ""].copy()
                df.loc[~df["班別"].isin(SHIFTS), "班別"] = default_duty_shift()
                if not df.empty:
                    return df[cols].drop_duplicates(subset=["姓名", "班別"], keep="first").reset_index(drop=True)
            ds = default_duty_shift()
            df = pd.DataFrame([{"姓名": n, "班別": ds} for n in ["大明", "小明", "明明"]])
            save_duty_members_df(df)
            return df
        except Exception as e:
            print(f"讀取值日生名單失敗（DB）：{e}")
            return pd.DataFrame(columns=cols)
    if os.path.exists(FILE_DUTY):
        try:
            df = pd.read_excel(FILE_DUTY)
            if "姓名" in df.columns:
                if "班別" not in df.columns:
                    df["班別"] = default_duty_shift()
                df["姓名"] = df["姓名"].fillna("").astype(str).str.strip()
                df["班別"] = df["班別"].fillna("").astype(str).str.strip()
                df = df[df["姓名"] != ""].copy()
                df.loc[~df["班別"].isin(SHIFTS), "班別"] = default_duty_shift()
                if not df.empty:
                    return df[cols].drop_duplicates(subset=["姓名", "班別"], keep="first").reset_index(drop=True)
        except Exception as e:
            print(f"讀取值日生名單失敗：{e}")
    ds = default_duty_shift()
    df = pd.DataFrame([{"姓名": n, "班別": ds} for n in ["大明", "小明", "明明"]])
    save_duty_members_df(df)
    return df


def save_duty_members_df(df):
    if db_enabled():
        try:
            import psycopg2.extras as extras
            rows = []
            for _, r in df.iterrows():
                nm = clean_text(r.get("姓名", ""))
                if not nm:
                    continue
                sh = clean_text(r.get("班別", "")) or default_duty_shift()
                rows.append((nm, sh))
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("TRUNCATE duty_members")
                if rows:
                    extras.execute_values(cur, "INSERT INTO duty_members (name, shift) VALUES %s ON CONFLICT (name, shift) DO NOTHING", rows)
                conn.commit()
            return True
        except Exception as e:
            print(f"保存值日生名單失敗（DB）：{e}")
            return False
    try:
        df[["姓名", "班別"]].to_excel(FILE_DUTY, index=False)
        return True
    except Exception as e:
        print(f"保存值日生名單失敗：{e}")
        return False


def load_duty_members(shift=None):
    """回傳值日生姓名清單。指定 shift 則只回該班別，否則回全部。"""
    df = load_duty_members_df()
    if shift:
        names = df[df["班別"].astype(str).str.strip() == shift]["姓名"].tolist()
    else:
        names = df["姓名"].tolist()
    out = []
    for n in names:
        n = clean_text(n)
        if n and n not in out:
            out.append(n)
    return out


def duty_members_by_shift():
    """回傳 {班別: [姓名]}，供前端分班別顯示。"""
    df = load_duty_members_df()
    result = {sh: [] for sh in SHIFTS}
    for _, row in df.iterrows():
        nm = clean_text(row.get("姓名", ""))
        sh = clean_text(row.get("班別", "")) or default_duty_shift()
        result.setdefault(sh, [])
        if nm and nm not in result[sh]:
            result[sh].append(nm)
    return result


def empty_attendance_frames(year=None, month=None):
    now, num_days = now_info(year, month)
    df_basic = pd.DataFrame(columns=["姓名", "部門", "班別", "狀態", "備註"])
    df_leave = pd.DataFrame(columns=["姓名"] + [str(d) for d in range(1, num_days + 1)])
    return df_basic, df_leave


def save_attendance_frames(df_basic, df_leave, year=None, month=None):
    """儲存人員基本資料與指定月份假表；其他月份假表會保留。"""
    if db_enabled():
        try:
            import psycopg2.extras as extras
            selected, _ = now_info(year, month)
            y, m = selected.year, selected.month
            mrows = []
            for _, r in df_basic.iterrows():
                nm = clean_text(r.get("姓名", ""))
                if not nm:
                    continue
                mrows.append((nm, clean_text(r.get("部門", "")) or "財務", clean_text(r.get("班別", "晚班")) or "晚班", clean_text(r.get("狀態", "啟用")) or "啟用", clean_text(r.get("備註", ""))))
            lrows = []
            for _, r in df_leave.iterrows():
                nm = clean_text(r.get("姓名", ""))
                if not nm:
                    continue
                for col in df_leave.columns:
                    cs = str(col)
                    if cs.isdigit():
                        lv = clean_leave_value(r.get(col, ""))
                        if lv:
                            lrows.append((y, m, nm, int(cs), lv))
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("TRUNCATE members")
                if mrows:
                    extras.execute_values(cur, "INSERT INTO members (name, dept, shift, status, note) VALUES %s ON CONFLICT (name, dept, shift) DO UPDATE SET status=EXCLUDED.status, note=EXCLUDED.note", mrows)
                cur.execute("DELETE FROM leaves WHERE year=%s AND month=%s", (y, m))
                if lrows:
                    extras.execute_values(cur, "INSERT INTO leaves (year, month, name, day, leave_type) VALUES %s ON CONFLICT (year, month, name, day) DO UPDATE SET leave_type=EXCLUDED.leave_type", lrows)
                conn.commit()
            return True, ""
        except Exception as e:
            return False, f"寫入資料庫失敗：{e}"
    target_leave_sheet = leave_sheet_name(year, month)
    existing_sheets = {}

    if os.path.exists(FILE_LEAVE):
        try:
            with pd.ExcelFile(FILE_LEAVE, engine="openpyxl") as xls:
                for sheet in xls.sheet_names:
                    if sheet not in [SHEET_BASIC, SHEET_LEAVE, target_leave_sheet]:
                        existing_sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
        except Exception:
            existing_sheets = {}

    try:
        with pd.ExcelWriter(FILE_LEAVE, engine="openpyxl", mode="w") as writer:
            df_basic.to_excel(writer, sheet_name=SHEET_BASIC, index=False)
            df_leave.to_excel(writer, sheet_name=target_leave_sheet, index=False)
            # 兼容舊版：目前月份也同步寫入 LeaveSheet，避免舊資料讀不到
            now = datetime.now()
            if int(year or now.year) == now.year and int(month or now.month) == now.month:
                df_leave.to_excel(writer, sheet_name=SHEET_LEAVE, index=False)
            for sheet, df in existing_sheets.items():
                safe_name = str(sheet)[:31]
                if safe_name not in [SHEET_BASIC, SHEET_LEAVE, target_leave_sheet]:
                    df.to_excel(writer, sheet_name=safe_name, index=False)
        return True, ""
    except Exception as e:
        return False, ensure_excel_not_open_error(e)


def init_attendance_sheets(year=None, month=None):
    selected, num_days = now_info(year, month)
    target_leave_sheet = leave_sheet_name(selected.year, selected.month)

    if db_enabled():
        try:
            y, m = selected.year, selected.month
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("SELECT name, dept, shift, status, note FROM members ORDER BY name")
                mrows = cur.fetchall()
                cur.execute("SELECT name, day, leave_type FROM leaves WHERE year=%s AND month=%s", (y, m))
                lrows = cur.fetchall()
            df_basic = pd.DataFrame(mrows, columns=["姓名", "部門", "班別", "狀態", "備註"])
            for col in ["姓名", "部門", "班別", "狀態", "備註"]:
                if col not in df_basic.columns:
                    df_basic[col] = "啟用" if col == "狀態" else ""
            df_basic["姓名"] = df_basic["姓名"].fillna("").astype(str).str.strip()
            df_basic["部門"] = df_basic["部門"].fillna("").astype(str).str.strip()
            df_basic["班別"] = df_basic["班別"].fillna("晚班").astype(str).str.strip().replace("", "晚班")
            df_basic["狀態"] = df_basic["狀態"].fillna("啟用").astype(str).str.strip().replace("", "啟用")
            df_basic["備註"] = df_basic["備註"].fillna("").astype(str).str.strip()
            df_basic = df_basic[df_basic["姓名"] != ""].drop_duplicates(subset=["姓名", "部門", "班別"], keep="first").reset_index(drop=True)
            df_basic.loc[~df_basic["部門"].isin(DEPARTMENTS), "部門"] = "財務"
            df_basic.loc[~df_basic["班別"].isin(SHIFTS), "班別"] = "晚班"
            df_basic.loc[~df_basic["狀態"].isin(["啟用", "停用"]), "狀態"] = "啟用"
            leave_map = {}
            for (nm, day, lt) in lrows:
                nm2 = clean_text(nm)
                lv = clean_leave_value(lt)
                if nm2 and lv:
                    try:
                        leave_map.setdefault(nm2, {})[str(int(day))] = lv
                    except Exception:
                        pass
            names = []
            for nm in df_basic["姓名"].tolist():
                nm = clean_text(nm)
                if nm and nm not in names:
                    names.append(nm)
            day_cols = [str(d) for d in range(1, num_days + 1)]
            rows = []
            for nm in names:
                row = {"姓名": nm}
                lm = leave_map.get(nm, {})
                for d in day_cols:
                    row[d] = lm.get(d, "")
                rows.append(row)
            df_leave = pd.DataFrame(rows, columns=["姓名"] + day_cols) if rows else pd.DataFrame(columns=["姓名"] + day_cols)
            return df_basic, df_leave
        except Exception as e:
            print(f"讀取假表失敗（DB）：{e}")
            return empty_attendance_frames(selected.year, selected.month)

    if not os.path.exists(FILE_LEAVE):
        df_basic, df_leave = empty_attendance_frames(selected.year, selected.month)
        save_attendance_frames(df_basic, df_leave, selected.year, selected.month)
        return df_basic, df_leave

    try:
        with pd.ExcelFile(FILE_LEAVE, engine="openpyxl") as xls:
            df_basic = pd.read_excel(xls, sheet_name=SHEET_BASIC) if SHEET_BASIC in xls.sheet_names else pd.DataFrame()
            if target_leave_sheet in xls.sheet_names:
                df_leave = pd.read_excel(xls, sheet_name=target_leave_sheet)
            elif SHEET_LEAVE in xls.sheet_names and selected.year == datetime.now().year and selected.month == datetime.now().month:
                df_leave = pd.read_excel(xls, sheet_name=SHEET_LEAVE)
            else:
                df_leave = pd.DataFrame()
    except Exception as e:
        print(f"讀取大假表失敗：{e}")
        return empty_attendance_frames(selected.year, selected.month)

    for col in ["姓名", "部門", "班別", "狀態", "備註"]:
        if col not in df_basic.columns:
            df_basic[col] = "啟用" if col == "狀態" else ""
    if "姓名" not in df_leave.columns:
        df_leave["姓名"] = ""

    df_basic["姓名"] = df_basic["姓名"].fillna("").astype(str).str.strip()
    df_basic["部門"] = df_basic["部門"].fillna("").astype(str).str.strip()
    df_basic["班別"] = df_basic["班別"].fillna("晚班").astype(str).str.strip().replace("", "晚班")
    df_basic["狀態"] = df_basic["狀態"].fillna("啟用").astype(str).str.strip().replace("", "啟用")
    df_basic["備註"] = df_basic["備註"].fillna("").astype(str).str.strip()
    df_leave["姓名"] = df_leave["姓名"].fillna("").astype(str).str.strip()

    df_basic = df_basic[df_basic["姓名"] != ""].drop_duplicates(subset=["姓名", "部門", "班別"], keep="first").reset_index(drop=True)
    df_leave = df_leave[df_leave["姓名"] != ""].drop_duplicates(subset=["姓名"], keep="first").reset_index(drop=True)

    modified = False
    for idx, row in df_basic.iterrows():
        if row["部門"] not in DEPARTMENTS:
            df_basic.at[idx, "部門"] = "財務"
            modified = True
        if row["班別"] not in SHIFTS:
            df_basic.at[idx, "班別"] = "晚班"
            modified = True
        if row["狀態"] not in ["啟用", "停用"]:
            df_basic.at[idx, "狀態"] = "啟用"
            modified = True

    for d in range(1, num_days + 1):
        day_col = str(d)
        if day_col not in df_leave.columns:
            df_leave[day_col] = ""
            modified = True

    valid_cols = ["姓名"] + [str(d) for d in range(1, num_days + 1)]
    other_cols = [c for c in df_leave.columns if not str(c).isdigit() and c != "姓名"]
    final_cols = [c for c in valid_cols if c in df_leave.columns] + other_cols
    if list(df_leave.columns) != final_cols:
        df_leave = df_leave[final_cols]
        modified = True

    leave_names = df_leave["姓名"].tolist()
    for _, row in df_basic.iterrows():
        name = clean_text(row["姓名"])
        if name and name not in leave_names:
            new_row = {"姓名": name}
            for d in range(1, num_days + 1):
                new_row[str(d)] = ""
            df_leave = pd.concat([df_leave, pd.DataFrame([new_row])], ignore_index=True)
            leave_names.append(name)
            modified = True

    if modified:
        save_attendance_frames(df_basic, df_leave, selected.year, selected.month)
    return df_basic, df_leave


def group_key(department, shift):
    return f"{department}__{shift}"

def build_calendar_data(year=None, month=None):
    now, num_days = now_info(year, month)
    df_basic, df_leave = init_attendance_sheets(now.year, now.month)

    leave_records = {}
    if not df_leave.empty:
        for _, l_row in df_leave.iterrows():
            name = clean_text(l_row.get("姓名", ""))
            if name:
                leave_records[name] = l_row.to_dict()

    calendar_users = []
    calendar_groups = {group_key(dept, shift): [] for dept in DEPARTMENTS for shift in SHIFTS}

    if not df_basic.empty:
        for _, row in df_basic.iterrows():
            name = clean_text(row.get("姓名", ""))
            dept = clean_text(row.get("部門", "")) or "財務"
            shift = clean_text(row.get("班別", "晚班")) or "晚班"
            if shift not in SHIFTS:
                shift = "晚班"
            status = clean_text(row.get("狀態", "啟用")) or "啟用"
            note = clean_text(row.get("備註", ""))
            if not name:
                continue
            if dept not in DEPARTMENTS:
                dept = "財務"
            user_leave_dict = leave_records.get(name, {})
            leaves = {str(d): clean_leave_value(user_leave_dict.get(str(d), "")) for d in range(1, num_days + 1)}
            leave_counts = {lt: 0 for lt in LEAVE_TYPES}
            for lv in leaves.values():
                if lv in leave_counts:
                    leave_counts[lv] += 1
            user_data = {
                "name": name, "dept": dept, "shift": shift, "status": status, "note": note,
                "leaves": leaves, "leave_counts": leave_counts,
                "rest_days": leave_counts.get("休", 0),
                "special_days": leave_counts.get("特", 0),
                "sick_days": leave_counts.get("病", 0),
            }
            calendar_users.append(user_data)
            calendar_groups[group_key(dept, shift)].append(user_data)
    return calendar_users, calendar_groups


def make_scheme(name="新方案", tasks=None, scheme_id=None):
    tasks = [clean_text(t) for t in (tasks or []) if clean_text(t)]
    return {
        "id": scheme_id or uuid.uuid4().hex[:10],
        "name": clean_text(name) or "新方案",
        "task_count": len(tasks),
        "tasks": tasks
    }


def default_work_schemes():
    return {group_key(dept, shift): {} for dept in DEPARTMENTS for shift in SHIFTS}


def normalize_work_schemes(data):
    """依部門＋班別＋人數分類：{財務__早班:{"2":[scheme]}}；相容舊版部門格式。"""
    normalized = default_work_schemes()
    if not isinstance(data, dict):
        return normalized

    def add_raw_to_key(store_key, raw):
        buckets = defaultdict(list)
        if isinstance(raw, dict):
            if all(str(k).isdigit() for k in raw.keys()):
                for count_key, schemes in raw.items():
                    if isinstance(schemes, list):
                        for item in schemes:
                            if isinstance(item, dict):
                                tasks = item.get("tasks", []) if isinstance(item.get("tasks", []), list) else []
                                scheme = make_scheme(item.get("name", "新方案"), tasks, item.get("id") or None)
                                buckets[str(scheme["task_count"])].append(scheme)
            else:
                for key, item in raw.items():
                    if isinstance(item, list):
                        scheme = make_scheme(f"方案{key}", item)
                    elif isinstance(item, dict):
                        tasks = item.get("tasks", []) if isinstance(item.get("tasks", []), list) else []
                        scheme = make_scheme(item.get("name", f"方案{key}"), tasks, item.get("id") or None)
                    else:
                        continue
                    buckets[str(scheme["task_count"])].append(scheme)
        elif isinstance(raw, list):
            for item in raw:
                if isinstance(item, dict):
                    tasks = item.get("tasks", []) if isinstance(item.get("tasks", []), list) else []
                    scheme = make_scheme(item.get("name", "新方案"), tasks, item.get("id") or None)
                    buckets[str(scheme["task_count"])].append(scheme)
        if buckets:
            normalized[store_key] = {k: v for k, v in sorted(buckets.items(), key=lambda kv: int(kv[0]) if str(kv[0]).isdigit() else 999)}

    for dept in DEPARTMENTS:
        for shift in SHIFTS:
            key = group_key(dept, shift)
            if key in data:
                add_raw_to_key(key, data.get(key, {}))
        # 舊版只有部門時，先複製到三個班別，避免原方案消失
        if dept in data:
            for shift in SHIFTS:
                key = group_key(dept, shift)
                if not normalized.get(key):
                    add_raw_to_key(key, data.get(dept, {}))
    return normalized

def load_work_schemes():
    raw = None
    if db_enabled():
        raw = kv_get("work_schemes")
    elif os.path.exists(WORK_SCHEME_FILE):
        try:
            with open(WORK_SCHEME_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except Exception as e:
            print(f"讀取工作方案失敗：{e}")
    if raw is not None:
        return normalize_work_schemes(raw)
    data = default_work_schemes()
    save_work_schemes(data)
    return data


def save_work_schemes(data):
    try:
        data = normalize_work_schemes(data)
        if db_enabled():
            kv_put("work_schemes", data)
        else:
            with open(WORK_SCHEME_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"儲存工作方案失敗：{e}")
        return False


def get_working_users_by_department_day(department, day, year=None, month=None, shift="晚班"):
    day = str(day)
    df_basic, df_leave = init_attendance_sheets(year, month)
    working = []
    leave_records = {}
    if not df_leave.empty:
        for _, row in df_leave.iterrows():
            name = clean_text(row.get("姓名", ""))
            if name:
                leave_records[name] = row.to_dict()

    for _, row in df_basic.iterrows():
        name = clean_text(row.get("姓名", ""))
        dept = clean_text(row.get("部門", ""))
        row_shift = clean_text(row.get("班別", "晚班")) or "晚班"
        if not name or dept != department or row_shift != shift:
            continue
        lv = clean_leave_value(leave_records.get(name, {}).get(day, ""))
        if lv == "":
            working.append(name)
    return working


def load_history_df():
    cols = ["日期", "年份", "月份", "日", "部門", "班別", "人員", "工作", "方案", "分配方式", "產生時間", "批次ID"]
    if db_enabled():
        try:
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("SELECT date, year, month, day, dept, shift, person, task, scheme, method, created_at, batch_id FROM work_history ORDER BY id")
                rows = cur.fetchall()
            df = pd.DataFrame(rows, columns=cols)
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
            return df[cols]
        except Exception as e:
            print(f"讀取工作分配歷史失敗（DB）：{e}")
            return pd.DataFrame(columns=cols)
    if os.path.exists(HISTORY_FILE):
        try:
            df = pd.read_excel(HISTORY_FILE, sheet_name=SHEET_HISTORY)
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
            return df[cols]
        except Exception as e:
            print(f"讀取工作分配歷史失敗：{e}")
    return pd.DataFrame(columns=cols)


def _save_history_df_to_db(df):
    import psycopg2.extras as extras

    def cv(v, numeric=False):
        s = clean_text(v)
        if numeric:
            try:
                return int(float(s)) if s != "" else None
            except Exception:
                return None
        return s

    rows = []
    if df is not None and not df.empty:
        for _, r in df.iterrows():
            rows.append((
                cv(r.get("日期")), cv(r.get("年份"), True), cv(r.get("月份"), True), cv(r.get("日"), True),
                cv(r.get("部門")), cv(r.get("班別")), cv(r.get("人員")), cv(r.get("工作")), cv(r.get("方案")),
                cv(r.get("分配方式")), cv(r.get("產生時間")), cv(r.get("批次ID")),
            ))
    with pg_conn() as conn, conn.cursor() as cur:
        cur.execute("TRUNCATE work_history")
        if rows:
            extras.execute_values(
                cur,
                "INSERT INTO work_history (date, year, month, day, dept, shift, person, task, scheme, method, created_at, batch_id) VALUES %s",
                rows,
            )
        conn.commit()


def save_history_df(df):
    if db_enabled():
        try:
            _save_history_df_to_db(df)
            return True, ""
        except Exception as e:
            return False, f"寫入資料庫失敗：{e}"
    try:
        with pd.ExcelWriter(HISTORY_FILE, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=SHEET_HISTORY, index=False)
        return True, ""
    except Exception as e:
        return False, ensure_excel_not_open_error(e)


def monthly_person_counts(department=None, year=None, month=None, shift=None):
    now, _ = now_info()
    year = year or now.year
    month = month or now.month
    df = load_history_df()
    if df.empty:
        return {}
    filtered = df[(df["年份"].astype(str) == str(year)) & (df["月份"].astype(str) == str(month))]
    if department:
        filtered = filtered[filtered["部門"].astype(str) == department]
    if shift:
        filtered = filtered[filtered["班別"].astype(str) == shift]
    counts = filtered.groupby("人員").size().to_dict()
    return {clean_text(k): int(v) for k, v in counts.items() if clean_text(k)}


def select_scheme_for_count(department, worker_count, shift="晚班", schemes=None):
    if schemes is None:
        schemes = load_work_schemes()
    bucket = schemes.get(group_key(department, shift), {}).get(str(worker_count), [])
    if not bucket:
        return None
    # 同人數若有多個方案，先用第一個；使用者可調整順序，或之後再做指定方案
    return bucket[0]


def balanced_random_order(users, department, year, month, local_counts=None, shift="晚班", base_counts=None):
    """隨機但偏向次數少的人；若可能，讓月累積差距保持在 ±2。"""
    counts = dict(base_counts) if base_counts is not None else monthly_person_counts(department, year, month, shift)
    if local_counts:
        for k, v in local_counts.items():
            counts[k] = counts.get(k, 0) + int(v)
    for u in users:
        counts.setdefault(u, 0)

    remaining = users[:]
    ordered = []
    while remaining:
        min_count = min(counts.get(u, 0) for u in remaining)
        # 優先選最少者；允許差距最多 2，超過時更強制選最少者
        candidates = [u for u in remaining if counts.get(u, 0) == min_count]
        chosen = random.choice(candidates)
        ordered.append(chosen)
        remaining.remove(chosen)
        counts[chosen] += 1
    return ordered


def task_tags(task_str):
    """把一個工作位字串拆成標籤，例如 "HS 買卡" -> ["HS", "買卡"]。"""
    return [t for t in str(task_str or "").split() if t.strip()]


def monthly_task_counts(department, shift, year=None, month=None, exclude_day=None):
    """本月該部門班別、已存歷史中每人各工作標籤的累積：{人員: {標籤: 次數}}。
    exclude_day 可排除某一天（單日重排時避免把即將被覆蓋的舊資料算進去）。"""
    selected, _ = now_info(year, month)
    y, m = selected.year, selected.month
    df = load_history_df()
    if df.empty:
        return {}
    try:
        sub = df[
            (pd.to_numeric(df["年份"], errors="coerce") == y)
            & (pd.to_numeric(df["月份"], errors="coerce") == m)
            & (df["部門"].astype(str) == str(department))
            & (df["班別"].astype(str) == str(shift))
        ]
        if exclude_day is not None:
            sub = sub[pd.to_numeric(sub["日"], errors="coerce") != int(exclude_day)]
    except Exception:
        return {}
    counts = {}
    for _, row in sub.iterrows():
        person = clean_text(row.get("人員", ""))
        if not person:
            continue
        bucket = counts.setdefault(person, {})
        for tag in task_tags(row.get("工作", "")):
            bucket[tag] = bucket.get(tag, 0) + 1
    return counts


def assign_positions_balanced(workers, positions, task_counts):
    """把當天的工作位公平分給 workers，讓每人各「工作標籤」的累積盡量平均，
    避免同一項工作一直落在同一人。task_counts 唯讀不修改。
    回傳 [{"user": 人員, "task": 工作位}, ...]，每位 worker 一筆。
    人多於工作位時，多出來的人標「未設定工作」；工作位多於人時，多餘的工作位捨棄。
    """
    workers = list(workers)
    positions = list(positions)
    sim = {w: dict(task_counts.get(w, {})) for w in workers}
    n_assign = min(len(workers), len(positions))
    pos_order = list(range(len(positions)))
    random.shuffle(pos_order)
    # 標籤較多（較難安排）的工作位先處理，平手者維持隨機順序
    pos_order.sort(key=lambda i: -len(task_tags(positions[i])))
    pos_order = pos_order[:n_assign]
    remaining = list(workers)
    chosen = {}
    for pi in pos_order:
        tags = task_tags(positions[pi])
        random.shuffle(remaining)          # 平手隨機
        best_w, best_cost = None, None
        for w in remaining:
            cost = sum(sim[w].get(t, 0) for t in tags)
            if best_cost is None or cost < best_cost:
                best_cost, best_w = cost, w
        chosen[best_w] = positions[pi]
        for t in tags:
            sim[best_w][t] = sim[best_w].get(t, 0) + 1
        remaining.remove(best_w)
    return [{"user": w, "task": chosen.get(w, "未設定工作")} for w in workers]


class MonthlyAssignmentContext:
    """整月／批次分配時預先載入假表、方案與歷史，避免每天重複讀檔。"""
    def __init__(self, department, shift, year=None, month=None, seed_task_counts=False):
        selected, _ = now_info(year, month)
        self.year, self.month = selected.year, selected.month
        self.department = department
        self.shift = shift
        df_basic, df_leave = init_attendance_sheets(self.year, self.month)
        # 預先篩出此部門＋班別的人員
        self.members = []
        for _, row in df_basic.iterrows():
            name = clean_text(row.get("姓名", ""))
            dept = clean_text(row.get("部門", ""))
            row_shift = clean_text(row.get("班別", "晚班")) or "晚班"
            if name and dept == department and row_shift == shift:
                self.members.append(name)
        # 預建請假紀錄字典（全假表，依姓名查）
        self.leave_records = {}
        if not df_leave.empty:
            for _, row in df_leave.iterrows():
                name = clean_text(row.get("姓名", ""))
                if name:
                    self.leave_records[name] = row.to_dict()
        # 各讀一次：方案與已存歷史的本月每人分配次數
        self.schemes = load_work_schemes()
        self.base_counts = monthly_person_counts(department, self.year, self.month, shift)
        # 每人×每工作標籤本月累積，平均隨機分配工作時用來公平輪替；
        # 整月重產時從 0 起算（seed_task_counts=False），逐日累積；到月底重排才從已固定的前段起算。
        self.base_task_counts = monthly_task_counts(department, shift, self.year, self.month) if seed_task_counts else {}

    def working_users(self, day):
        day = str(day)
        return [m for m in self.members
                if clean_leave_value(self.leave_records.get(m, {}).get(day, "")) == ""]


def create_assignment_for_day(department, day, mode="balanced_random", save=True, batch_id=None, local_counts=None, year=None, month=None, shift="晚班", ctx=None):
    selected, _ = now_info(year, month)
    year, month = selected.year, selected.month
    if ctx is not None:
        workers = ctx.working_users(day)
        base_counts = ctx.base_counts
        schemes = ctx.schemes
    else:
        workers = get_working_users_by_department_day(department, day, year, month, shift)
        base_counts = None
        schemes = None
    worker_count = len(workers)
    if worker_count < 1:
        return {"status": "skip", "message": f"{department} {day} 號沒有上班人員", "day": day, "department": department}

    scheme = select_scheme_for_count(department, worker_count, shift, schemes=schemes)
    if not scheme:
        return {"status": "error", "message": f"{department} {day} 號上班 {worker_count} 人，但沒有 {worker_count} 人方案", "day": day, "department": department, "worker_count": worker_count}

    tasks = scheme.get("tasks", [])
    date_obj = datetime(year, month, int(day))
    batch_id = batch_id or uuid.uuid4().hex[:12]
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if mode == "fixed":
        # 固定順序：人員照原順序、工作照方案順序對位
        assignments = [{"user": u, "task": (tasks[i] if i < len(tasks) else "未設定工作")}
                       for i, u in enumerate(workers)]
    elif mode == "random":
        # 純隨機：人員與工作各自打散後對位（不參考歷史）
        ordered_workers = balanced_random_order(workers, department, year, month, local_counts=local_counts, shift=shift, base_counts=base_counts)
        rtasks = tasks[:]
        random.shuffle(rtasks)
        assignments = [{"user": u, "task": (rtasks[i] if i < len(rtasks) else "未設定工作")}
                       for i, u in enumerate(ordered_workers)]
    else:
        # 平均隨機：依「每人 × 每個工作標籤」本月累積做公平輪替，
        # 讓同一項工作盡量平均落在不同人，不再集中在同一人。
        if ctx is not None:
            base_task_counts = ctx.base_task_counts
        else:
            base_task_counts = monthly_task_counts(department, shift, year, month, exclude_day=int(day))
        assignments = assign_positions_balanced(workers, tasks, base_task_counts)
        if ctx is not None:
            # 逐日累積回 ctx，讓後面的日子知道前面排過什麼，整月持續維持每項工作公平
            for a in assignments:
                bucket = ctx.base_task_counts.setdefault(a["user"], {})
                for t in task_tags(a["task"]):
                    bucket[t] = bucket.get(t, 0) + 1

    result = {
        "status": "success",
        "date": date_obj.strftime("%Y-%m-%d"),
        "day": int(day),
        "department": department,
        "shift": shift,
        "worker_count": worker_count,
        "workers": workers,
        "scheme": scheme.get("name", "未命名方案"),
        "scheme_id": scheme.get("id", ""),
        "mode": mode,
        "assignments": assignments,
        "created_at": created_at,
        "batch_id": batch_id,
    }

    if save:
        append_history([result])
    return result


def append_history(results):
    df = load_history_df()
    rows = []
    now, _ = now_info()
    for result in results:
        if result.get("status") != "success":
            continue
        dt = datetime.strptime(result["date"], "%Y-%m-%d")
        for item in result["assignments"]:
            rows.append({
                "日期": result["date"],
                "年份": dt.year,
                "月份": dt.month,
                "日": dt.day,
                "部門": result["department"],
                "班別": result.get("shift", "晚班"),
                "人員": item["user"],
                "工作": item["task"],
                "方案": result["scheme"],
                "分配方式": "平均隨機" if result.get("mode") in ["balanced_random", "random"] else "固定順序",
                "產生時間": result["created_at"],
                "批次ID": result["batch_id"],
            })
    if rows:
        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
        return save_history_df(df)
    return True, ""


def filter_history_by_range(start_date=None, end_date=None, department=None, person=None, shift=None):
    """依日期區間、部門、姓名查詢歷史紀錄。首頁不預先載入整個月份，避免一次顯示太多資料。"""
    df = load_history_df()
    if df.empty:
        return [], []

    filtered = df.copy()
    filtered["日期_dt"] = pd.to_datetime(filtered["日期"], errors="coerce")
    filtered = filtered[filtered["日期_dt"].notna()]

    if start_date:
        start_dt = pd.to_datetime(start_date, errors="coerce")
        if pd.notna(start_dt):
            filtered = filtered[filtered["日期_dt"] >= start_dt]

    if end_date:
        end_dt = pd.to_datetime(end_date, errors="coerce")
        if pd.notna(end_dt):
            filtered = filtered[filtered["日期_dt"] <= end_dt]

    if department and department in DEPARTMENTS:
        filtered = filtered[filtered["部門"].astype(str) == department]
    if shift and shift in SHIFTS:
        filtered = filtered[filtered["班別"].astype(str) == shift]

    if person:
        filtered = filtered[filtered["人員"].astype(str).str.strip() == person]

    if filtered.empty:
        return [], []

    filtered = filtered.sort_values(["日期_dt", "部門", "班別", "人員"])
    filtered = filtered.drop(columns=["日期_dt"], errors="ignore")

    records = filtered.to_dict("records")
    stats = (
        filtered.groupby(["部門", "班別", "人員"])
        .size()
        .reset_index(name="區間分配次數")
        .sort_values(["部門", "區間分配次數", "人員"], ascending=[True, False, True])
        .to_dict("records")
    )
    return records, stats


def duty_history_columns():
    return ["日期", "年份", "月份", "日", "星期", "班別", "值日生", "產生時間", "批次ID"]


def load_duty_history_df():
    cols = duty_history_columns()
    if db_enabled():
        try:
            with pg_conn() as conn, conn.cursor() as cur:
                cur.execute("SELECT date, year, month, day, weekday, shift, duty_person, created_at, batch_id FROM duty_history ORDER BY id")
                rows = cur.fetchall()
            df = pd.DataFrame(rows, columns=cols)
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
            return df[cols]
        except Exception as e:
            print(f"讀取值日生歷史失敗（DB）：{e}")
            return pd.DataFrame(columns=cols)
    if os.path.exists(DUTY_HISTORY_FILE):
        try:
            df = pd.read_excel(DUTY_HISTORY_FILE)
        except Exception:
            df = pd.DataFrame(columns=cols)
    else:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


def _save_duty_history_to_db(df):
    import psycopg2.extras as extras

    def cv(v, numeric=False):
        s = clean_text(v)
        if numeric:
            try:
                return int(float(s)) if s != "" else None
            except Exception:
                return None
        return s

    rows = []
    if df is not None and not df.empty:
        for _, r in df.iterrows():
            rows.append((
                cv(r.get("日期")), cv(r.get("年份"), True), cv(r.get("月份"), True), cv(r.get("日"), True),
                cv(r.get("星期")), cv(r.get("班別")), cv(r.get("值日生")), cv(r.get("產生時間")), cv(r.get("批次ID")),
            ))
    with pg_conn() as conn, conn.cursor() as cur:
        cur.execute("TRUNCATE duty_history")
        if rows:
            extras.execute_values(cur, "INSERT INTO duty_history (date, year, month, day, weekday, shift, duty_person, created_at, batch_id) VALUES %s", rows)
        conn.commit()


def save_duty_history_df(df):
    if db_enabled():
        try:
            _save_duty_history_to_db(df)
            return True, ""
        except Exception as e:
            return False, f"寫入資料庫失敗：{e}"
    try:
        df.to_excel(DUTY_HISTORY_FILE, index=False)
        return True, ""
    except Exception as e:
        return False, ensure_excel_not_open_error(e)


def append_duty_history(daily_schedule, shift, batch_id=None):
    df = load_duty_history_df()
    batch_id = batch_id or uuid.uuid4().hex[:12]
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for item in daily_schedule:
        try:
            dt = datetime.strptime(str(item.get("date", "")), "%Y-%m-%d")
            date_str = dt.strftime("%Y-%m-%d")
        except Exception:
            try:
                dt = datetime.strptime(str(item.get("date", "")), "%Y年%m月%d日")
                date_str = dt.strftime("%Y-%m-%d")
            except Exception:
                continue
        rows.append({
            "日期": date_str,
            "年份": dt.year,
            "月份": dt.month,
            "日": dt.day,
            "星期": item.get("week_day", ""),
            "班別": shift,
            "值日生": item.get("user", ""),
            "產生時間": created_at,
            "批次ID": batch_id,
        })
    if rows:
        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
        return save_duty_history_df(df)
    return True, ""


def weekly_clean_key(year, month, shift):
    selected, _ = now_info(year, month)
    return f"{selected.year}_{selected.month:02d}__{shift}"


def load_weekly_clean():
    if db_enabled():
        d = kv_get("weekly_clean")
        return d if isinstance(d, dict) else {}
    if os.path.exists(WEEKLY_CLEAN_FILE):
        try:
            with open(WEEKLY_CLEAN_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}
    return {}


def save_weekly_clean(data):
    try:
        if db_enabled():
            kv_put("weekly_clean", data)
        else:
            with open(WEEKLY_CLEAN_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"儲存每週清潔失敗：{e}")
        return False


def get_weekly_clean(year, month, shift):
    return load_weekly_clean().get(weekly_clean_key(year, month, shift), [])


def set_weekly_clean(year, month, shift, schedule):
    data = load_weekly_clean()
    data[weekly_clean_key(year, month, shift)] = schedule
    return save_weekly_clean(data)


def clear_weekly_clean(year, month, shift):
    data = load_weekly_clean()
    data.pop(weekly_clean_key(year, month, shift), None)
    return save_weekly_clean(data)


def duty_person_counts(shift=None):
    """從值日生歷史即時計算每人累積值日次數（可依班別），作為公平基礎。"""
    df = load_duty_history_df()
    if df.empty:
        return {}
    if shift and "班別" in df.columns:
        df = df[df["班別"].astype(str).str.strip() == shift]
    counts = df.groupby("值日生").size().to_dict()
    return {clean_text(k): int(v) for k, v in counts.items() if clean_text(k)}


def filter_duty_history_by_range(start_date=None, end_date=None, person=None, shift=None):
    df = load_duty_history_df()
    if df.empty:
        return [], []
    filtered = df.copy()
    filtered["日期_dt"] = pd.to_datetime(filtered["日期"], errors="coerce")
    filtered = filtered[filtered["日期_dt"].notna()]
    if start_date:
        start_dt = pd.to_datetime(start_date, errors="coerce")
        if pd.notna(start_dt):
            filtered = filtered[filtered["日期_dt"] >= start_dt]
    if end_date:
        end_dt = pd.to_datetime(end_date, errors="coerce")
        if pd.notna(end_dt):
            filtered = filtered[filtered["日期_dt"] <= end_dt]
    if shift and shift in SHIFTS and "班別" in filtered.columns:
        filtered = filtered[filtered["班別"].astype(str).str.strip() == shift]
    if person:
        filtered = filtered[filtered["值日生"].astype(str).str.strip() == person]
    if filtered.empty:
        return [], []
    filtered = filtered.sort_values(["日期_dt", "值日生"]).drop(columns=["日期_dt"], errors="ignore")
    records = filtered.to_dict("records")
    stats = (
        filtered.groupby(["值日生"])
        .size()
        .reset_index(name="區間值日次數")
        .sort_values(["區間值日次數", "值日生"], ascending=[False, True])
        .to_dict("records")
    )
    return records, stats



def latest_assignment_records_for_day(date_str, department=None, shift=None):
    """取得指定日期最新一次產生的工作分配；避免重複產生後同一天顯示多批資料。"""
    records, _ = filter_history_by_range(date_str, date_str, department, "", shift)
    if not records:
        return []
    try:
        df = pd.DataFrame(records)
        if "產生時間" in df.columns:
            df["產生時間_dt"] = pd.to_datetime(df["產生時間"], errors="coerce")
            df = df.sort_values(["產生時間_dt", "批次ID"], na_position="first")
        latest_batch = clean_text(df.iloc[-1].get("批次ID", ""))
        if latest_batch and "批次ID" in df.columns:
            df = df[df["批次ID"].astype(str).str.strip() == latest_batch]
        df = df.drop(columns=["產生時間_dt"], errors="ignore")
        return df.to_dict("records")
    except Exception:
        return records


def latest_duty_record_for_day(date_str, shift=None):
    """取得指定日期最新一次產生的值日生（可指定班別）；多位時值日生欄合併為「A、B」。"""
    records, _ = filter_duty_history_by_range(date_str, date_str, "", shift)
    if not records:
        return None
    try:
        df = pd.DataFrame(records)
        if "產生時間" in df.columns:
            df["產生時間_dt"] = pd.to_datetime(df["產生時間"], errors="coerce")
            df = df.sort_values(["產生時間_dt", "批次ID"], na_position="first")
        latest_batch = clean_text(df.iloc[-1].get("批次ID", ""))
        if latest_batch and "批次ID" in df.columns:
            df = df[df["批次ID"].astype(str).str.strip() == latest_batch]
        names = [clean_text(x) for x in df["值日生"].tolist() if clean_text(x)]
        row = df.iloc[-1].drop(labels=["產生時間_dt"], errors="ignore").to_dict()
        row["值日生"] = "、".join(names)
        return row
    except Exception:
        return records[-1]



def latest_duty_records_for_month(year, month, shift=None):
    """取得指定年月（可指定班別）最新一次產生的每日值日生整月結果。"""
    df = load_duty_history_df()
    if df.empty:
        return []
    try:
        filtered = df[(df["年份"].astype(str) == str(year)) & (df["月份"].astype(str) == str(month))].copy()
        if shift and shift in SHIFTS and "班別" in filtered.columns:
            filtered = filtered[filtered["班別"].astype(str).str.strip() == shift]
        if filtered.empty:
            return []
        filtered["產生時間_dt"] = pd.to_datetime(filtered["產生時間"], errors="coerce")
        filtered = filtered.sort_values(["產生時間_dt", "批次ID", "日"], na_position="first")
        latest_batch = clean_text(filtered.iloc[-1].get("批次ID", ""))
        if latest_batch:
            filtered = filtered[filtered["批次ID"].astype(str).str.strip() == latest_batch]
        filtered = filtered.sort_values(["日"]).drop(columns=["產生時間_dt"], errors="ignore")
        return filtered.to_dict("records")
    except Exception:
        return []


def build_month_work_matrix(year, month, department, shift, num_days, group_users):
    """整月『人×日』工作矩陣（某部門×班別）。
    回傳 {people:[姓名...], rows:{姓名:{日:工作字串}}, leaves:{姓名:{日:假別}}}。
    每天取『最新批次』，與單日顯示邏輯一致。"""
    start = f"{year}-{month:02d}-01"
    end = f"{year}-{month:02d}-{num_days:02d}"
    try:
        records, _ = filter_history_by_range(start, end, department, "", shift)
    except Exception:
        records = []
    rows = {}
    if records:
        try:
            df = pd.DataFrame(records)
            if "日" in df.columns:
                df["__day"] = pd.to_numeric(df["日"], errors="coerce")
            else:
                df["__day"] = pd.to_datetime(df.get("日期"), errors="coerce").dt.day
            df["__dt"] = pd.to_datetime(df["產生時間"], errors="coerce") if "產生時間" in df.columns else pd.NaT
            for day, sub in df.groupby("__day"):
                if pd.isna(day):
                    continue
                day = int(day)
                sort_cols = ["__dt", "批次ID"] if "批次ID" in sub.columns else ["__dt"]
                sub = sub.sort_values(sort_cols, na_position="first")
                if "批次ID" in sub.columns:
                    latest_batch = clean_text(sub.iloc[-1].get("批次ID", ""))
                    if latest_batch:
                        sub = sub[sub["批次ID"].astype(str).str.strip() == latest_batch]
                for _, r in sub.iterrows():
                    nm = clean_text(r.get("人員", ""))
                    wk = clean_text(r.get("工作", ""))
                    if nm:
                        rows.setdefault(nm, {})[day] = wk
        except Exception as e:
            print(f"建立整月工作矩陣失敗：{e}")
    people = []
    leaves = {}
    for u in (group_users or []):
        nm = clean_text(u.get("name", ""))
        if nm and nm not in people:
            people.append(nm)
            lv = {}
            for k, v in (u.get("leaves", {}) or {}).items():
                vv = clean_leave_value(v)
                if vv:
                    try:
                        lv[int(k)] = vv
                    except Exception:
                        pass
            leaves[nm] = lv
    for nm in rows:
        if nm not in people:
            people.append(nm)
            leaves.setdefault(nm, {})
    return {"people": people, "rows": rows, "leaves": leaves}


def build_month_position_view(department, shift, year, month, num_days, schemes=None, duty_set_by_day=None):
    """日×工作位 月表（仿手繪班表）。
    回傳:
      reference: [{count:int, tasks:[工作位字串...]}]  依人數高到低（工作方案參考）
      max_slots: int  最多工作位數
      has_any: bool   本月是否已有任何排班
      day_rows: [{day, wd, weekend, has, slots:[人名...], duty:"A、B"}]
    """
    if schemes is None:
        schemes = load_work_schemes()
    gk = group_key(department, shift)
    bucket = schemes.get(gk, {}) or {}
    reference = []
    for ck in sorted((int(k) for k in bucket.keys() if str(k).isdigit()), reverse=True):
        lst = bucket.get(str(ck)) or []
        if lst:
            reference.append({"count": ck, "tasks": list(lst[0].get("tasks", []))})

    start = f"{year}-{month:02d}-01"
    end = f"{year}-{month:02d}-{num_days:02d}"
    try:
        records, _ = filter_history_by_range(start, end, department, "", shift)
    except Exception:
        records = []
    by_day = {}
    if records:
        try:
            df = pd.DataFrame(records)
            if "日" in df.columns:
                df["__day"] = pd.to_numeric(df["日"], errors="coerce")
            else:
                df["__day"] = pd.to_datetime(df.get("日期"), errors="coerce").dt.day
            df["__dt"] = pd.to_datetime(df["產生時間"], errors="coerce") if "產生時間" in df.columns else pd.NaT
            for day, sub in df.groupby("__day"):
                if pd.isna(day):
                    continue
                day = int(day)
                if "批次ID" in sub.columns:
                    sub = sub.sort_values(["__dt", "批次ID"], na_position="first")
                    latest_batch = clean_text(sub.iloc[-1].get("批次ID", ""))
                    if latest_batch:
                        sub = sub[sub["批次ID"].astype(str).str.strip() == latest_batch]
                items = []
                for _, r in sub.iterrows():
                    nm = clean_text(r.get("人員", ""))
                    tk = clean_text(r.get("工作", ""))
                    if nm:
                        items.append((nm, tk))
                if items:
                    by_day[day] = items
        except Exception as e:
            print(f"建立日×工作位失敗：{e}")

    _wd = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}
    max_slots = 0
    for ref in reference:
        max_slots = max(max_slots, len(ref["tasks"]))
    day_rows = []
    has_any = False
    for day in range(1, num_days + 1):
        items = by_day.get(day, [])
        slots = []
        if items:
            has_any = True
            sch = select_scheme_for_count(department, len(items), shift, schemes)
            order_tasks = list(sch.get("tasks", [])) if sch else []
            pool = list(items)
            # 先照方案工作位順序對到的人，靠左排
            for t in order_tasks:
                for i, (nm, tk) in enumerate(pool):
                    if tk == t:
                        slots.append(nm)
                        pool.pop(i)
                        break
            # 對不到方案（例如方案被改過）的人，按原順序緊接補上，不留前導空格
            for nm, tk in pool:
                slots.append(nm)
        wi = datetime(year, month, day).weekday()
        max_slots = max(max_slots, len(slots))
        day_rows.append({
            "day": day, "wd": _wd[wi], "weekend": wi >= 5,
            "has": bool(items), "slots": slots,
            "duty": "、".join((duty_set_by_day or {}).get(day, [])),
        })
    return {"reference": reference, "max_slots": max_slots, "has_any": has_any, "day_rows": day_rows}


def history_for_current_month():
    now, _ = now_info()
    start_date = f"{now.year}-{now.month:02d}-01"
    end_date = f"{now.year}-{now.month:02d}-{now.day:02d}"
    return filter_history_by_range(start_date, end_date)



def build_attendance_summary(calendar_groups, group_keys_to_show, num_days):
    summary = {}
    for dept in group_keys_to_show:
        users = calendar_groups.get(dept, [])
        day_rows = []
        for d in range(1, num_days + 1):
            working = 0
            leave = 0
            for u in users:
                lv = clean_leave_value(u.get("leaves", {}).get(str(d), ""))
                if lv:
                    leave += 1
                else:
                    working += 1
            day_rows.append({"day": d, "working": working, "leave": leave, "total": len(users)})
        summary[dept] = day_rows
    return summary


def build_anomaly_report(year, month, department, shift, ctx=None):
    """檢查本月可能要人工確認的異常。"""
    selected, num_days = now_info(year, month)
    year, month = selected.year, selected.month
    anomalies = []
    if ctx is None:
        ctx = MonthlyAssignmentContext(department, shift, year, month)
    scheme_buckets = ctx.schemes.get(group_key(department, shift), {})

    for day in range(1, num_days + 1):
        workers = ctx.working_users(day)
        if len(workers) == 0:
            anomalies.append({"level": "danger", "date": f"{month}/{day}", "message": f"{department} {shift} 無上班人員，當天無法分配工作。"})
        elif str(len(workers)) not in scheme_buckets:
            anomalies.append({"level": "warn", "date": f"{month}/{day}", "message": f"{department} {shift} 上班 {len(workers)} 人，但沒有 {len(workers)} 人方案。"})

    # 值日生若在假表上有休假，也提醒；依該班別的值日生檢查。
    duty_records = latest_duty_records_for_month(year, month, shift)
    if duty_records:
        for rec in duty_records:
            nm = clean_text(rec.get("值日生", ""))
            day = clean_text(rec.get("日", ""))
            lv = clean_leave_value(ctx.leave_records.get(nm, {}).get(str(day), ""))
            if lv:
                anomalies.append({"level": "warn", "date": f"{month}/{day}", "message": f"值日生 {nm} 當天是假別「{lv}」，請確認是否需要重排值日生。"})
    return anomalies


def copy_previous_month_leave(year, month, department=None, shift=None):
    """把上個月同一批人員的假表標記複製到目前月份；天數不足時只複製可對應日期。"""
    selected, num_days = now_info(year, month)
    y, m = selected.year, selected.month
    prev_y, prev_m = (y - 1, 12) if m == 1 else (y, m - 1)
    _, prev_days = now_info(prev_y, prev_m)
    df_basic, df_leave = init_attendance_sheets(y, m)
    _, prev_leave = init_attendance_sheets(prev_y, prev_m)
    if prev_leave.empty:
        return False, "上個月沒有可複製的假表資料"

    target_names = set()
    for _, row in df_basic.iterrows():
        nm = clean_text(row.get("姓名", ""))
        dept = clean_text(row.get("部門", ""))
        sh = clean_text(row.get("班別", "")) or "晚班"
        if not nm:
            continue
        if department and dept != department:
            continue
        if shift and sh != shift:
            continue
        target_names.add(nm)

    prev_map = {}
    for _, row in prev_leave.iterrows():
        nm = clean_text(row.get("姓名", ""))
        if nm:
            prev_map[nm] = row.to_dict()

    copied = 0
    for idx, row in df_leave.iterrows():
        nm = clean_text(row.get("姓名", ""))
        if nm not in target_names or nm not in prev_map:
            continue
        for d in range(1, min(num_days, prev_days) + 1):
            df_leave.at[idx, str(d)] = clean_leave_value(prev_map[nm].get(str(d), ""))
        copied += 1

    ok, msg = save_attendance_frames(df_basic, df_leave, y, m)
    if not ok:
        return False, msg
    return True, f"已從 {prev_y}/{prev_m} 複製 {copied} 位人員的假表標記到 {y}/{m}。"


def make_backup_zip():
    os.makedirs(BASE_DIR, exist_ok=True)
    backup_name = f"finance_schedule_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    backup_path = os.path.join(BASE_DIR, backup_name)
    files = [
        FILE_DUTY, FILE_LEAVE, WORK_SCHEME_FILE, HISTORY_FILE, DUTY_HISTORY_FILE, PLAN_STATUS_FILE, OPERATION_LOG_FILE, SYSTEM_SETTINGS_FILE, WEEKLY_CLEAN_FILE
    ]
    with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            if os.path.exists(f):
                zf.write(f, arcname=os.path.basename(f))
        zf.writestr("backup_info.txt", f"財務部門智能排班系統備份\n建立時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    return backup_path


def restore_backup_zip(file_storage):
    if not file_storage or not file_storage.filename.lower().endswith(".zip"):
        return False, "請上傳 zip 備份檔"
    allowed = {os.path.basename(p) for p in [FILE_DUTY, FILE_LEAVE, WORK_SCHEME_FILE, HISTORY_FILE, DUTY_HISTORY_FILE, PLAN_STATUS_FILE, OPERATION_LOG_FILE, SYSTEM_SETTINGS_FILE, WEEKLY_CLEAN_FILE]}
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "upload.zip")
        file_storage.save(src)
        try:
            with zipfile.ZipFile(src, "r") as zf:
                names = [n for n in zf.namelist() if os.path.basename(n) in allowed]
                if not names:
                    return False, "備份檔內找不到可還原的系統資料"
                # 還原前自動留一份備份
                make_backup_zip()
                for n in names:
                    base = os.path.basename(n)
                    target = os.path.join(BASE_DIR, base)
                    with zf.open(n) as rf, open(target, "wb") as wf:
                        shutil.copyfileobj(rf, wf)
            return True, "備份已還原。若畫面仍是舊資料，請重新整理頁面。"
        except zipfile.BadZipFile:
            return False, "備份檔格式錯誤，無法解壓縮"
        except Exception as e:
            return False, f"還原失敗：{e}"


def export_month_report_excel(year, month, department, shift):
    selected, num_days = now_info(year, month)
    y, m = selected.year, selected.month
    start = f"{y}-{m:02d}-01"
    end = f"{y}-{m:02d}-{num_days:02d}"
    records, stats = filter_history_by_range(start, end, department, "", shift)
    duty_records = latest_duty_records_for_month(y, m)
    calendar_users, calendar_groups = build_calendar_data(y, m)
    selected_key = group_key(department, shift)
    leave_rows = []
    for u in calendar_groups.get(selected_key, []):
        row = {"姓名": u["name"], "部門": u["dept"], "班別": u["shift"], "休合計": u["leave_counts"].get("休", 0), "特合計": u["leave_counts"].get("特", 0), "病合計": u["leave_counts"].get("病", 0)}
        for d in range(1, num_days + 1):
            row[str(d)] = u["leaves"].get(str(d), "")
        leave_rows.append(row)
    anomalies = build_anomaly_report(y, m, department, shift)
    out = os.path.join(BASE_DIR, f"{department}_{shift}_{y}_{m:02d}_月報表.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(records).to_excel(writer, sheet_name="工作分配", index=False)
        pd.DataFrame(stats).to_excel(writer, sheet_name="分配統計", index=False)
        pd.DataFrame(duty_records).to_excel(writer, sheet_name="每日值日生", index=False)
        pd.DataFrame(leave_rows).to_excel(writer, sheet_name="假表統計", index=False)
        pd.DataFrame(anomalies).to_excel(writer, sheet_name="異常提醒", index=False)
    return out


def export_month_report_pdf(year, month, department, shift):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        font_name = 'STSong-Light'
    except Exception:
        font_name = 'Helvetica'
    selected, num_days = now_info(year, month)
    y, m = selected.year, selected.month
    start = f"{y}-{m:02d}-01"
    end = f"{y}-{m:02d}-{num_days:02d}"
    records, stats = filter_history_by_range(start, end, department, "", shift)
    anomalies = build_anomaly_report(y, m, department, shift)
    duty_records = latest_duty_records_for_month(y, m)
    out = os.path.join(BASE_DIR, f"{department}_{shift}_{y}_{m:02d}_月報表.pdf")
    doc = SimpleDocTemplate(out, pagesize=landscape(A4), rightMargin=12*mm, leftMargin=12*mm, topMargin=10*mm, bottomMargin=10*mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CJKTitle', fontName=font_name, fontSize=18, leading=24, spaceAfter=8))
    styles.add(ParagraphStyle(name='CJK', fontName=font_name, fontSize=9, leading=12))
    story = [Paragraph(f"財務部門智能排班系統 - {department} {shift} {y}/{m:02d} 月報表", styles['CJKTitle'])]
    story.append(Paragraph(f"匯出時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['CJK']))
    story.append(Spacer(1, 6))
    if stats:
        data = [["部門", "班別", "人員", "分配次數"]] + [[s.get("部門", ""), s.get("班別", ""), s.get("人員", ""), s.get("區間分配次數", "")] for s in stats]
        story.append(Paragraph("區間每人分配次數統計", styles['CJKTitle']))
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font_name),('BACKGROUND',(0,0),(-1,0),colors.orange),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),8)]))
        story.append(t); story.append(Spacer(1, 10))
    if records:
        story.append(Paragraph("工作分配明細", styles['CJKTitle']))
        data = [["日期", "人員", "工作", "方案", "產生時間"]]
        for r in records[:300]:
            data.append([r.get("日期", ""), r.get("人員", ""), r.get("工作", ""), r.get("方案", ""), r.get("產生時間", "")])
        t = Table(data, repeatRows=1, colWidths=[24*mm, 24*mm, 55*mm, 45*mm, 42*mm])
        t.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font_name),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f97316')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),7),('VALIGN',(0,0),(-1,-1),'TOP')]))
        story.append(t)
    story.append(PageBreak())
    story.append(Paragraph("每日值日生", styles['CJKTitle']))
    if duty_records:
        data = [["日期", "星期", "值日生"]] + [[r.get("日期", ""), r.get("星期", ""), r.get("值日生", "")] for r in duty_records]
    else:
        data = [["提醒"], ["尚未產生每日值日生"]]
    t = Table(data, repeatRows=1)
    t.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font_name),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0ea5e9')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),8)]))
    story.append(t); story.append(Spacer(1, 10))
    story.append(Paragraph("異常提醒", styles['CJKTitle']))
    if anomalies:
        data = [["日期", "等級", "內容"]] + [[a.get("date", ""), a.get("level", ""), a.get("message", "")] for a in anomalies]
    else:
        data = [["狀態"], ["目前未發現異常"]]
    t = Table(data, repeatRows=1, colWidths=[25*mm, 25*mm, 190*mm] if anomalies else None)
    t.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font_name),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#dc2626')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),8),('VALIGN',(0,0),(-1,-1),'TOP')]))
    story.append(t)
    doc.build(story)
    return out



@app.route("/login", methods=["POST"])
def login():
    pwd = clean_text(request.form.get("password", ""))
    next_url = request.form.get("next") or url_for("index")
    if verify_admin_password(pwd):
        session["admin_logged_in"] = True
        log_operation("管理員登入", content="管理員已登入")
        return redirect(next_url)
    return redirect(url_for("index", login_msg="密碼錯誤，請重新輸入"))


@app.route("/logout")
def logout():
    log_operation("管理員登出", content="管理員已登出")
    session.pop("admin_logged_in", None)
    return redirect(url_for("index"))


@app.before_request
def automatic_backup_guard():
    # 管理員執行重要 POST 動作前，自動留一份備份。避免每次點擊都備份，60 秒內只留一次。
    mutating_prefixes = ("/api/", "/restore_backup", "/add_calendar_user", "/add_duty_user", "/delete_duty_user")
    if request.method == "POST" and request.path.startswith(mutating_prefixes):
        if is_admin() and request.endpoint not in ["login"]:
            now_ts = datetime.now().timestamp()
            last_ts = session.get("last_auto_backup_ts", 0)
            if now_ts - float(last_ts or 0) > 60:
                make_auto_backup("before_change")
                session["last_auto_backup_ts"] = now_ts

@app.route("/")
def index():
    real_now = datetime.now()
    selected_year = int(request.args.get("year", real_now.year) or real_now.year)
    selected_month = int(request.args.get("month", real_now.month) or real_now.month)
    selected_department = clean_text(request.args.get("department", "財務")) or "財務"
    selected_shift = clean_text(request.args.get("shift", "晚班")) or "晚班"
    if selected_department not in DEPARTMENTS:
        selected_department = "財務"
    if selected_shift not in SHIFTS:
        selected_shift = "晚班"

    now, num_days = now_info(selected_year, selected_month)
    duty_users = load_duty_members()
    calendar_users, calendar_groups = build_calendar_data(now.year, now.month)
    history_people = sorted({u.get("name", "") for u in calendar_users if u.get("name", "")})
    duty_history_people = load_duty_members()
    work_schemes = load_work_schemes()
    history_records, history_stats = [], []
    selected_group_key = group_key(selected_department, selected_shift)
    display_departments = [selected_department]
    display_group_keys = [selected_group_key]
    attendance_summary = build_attendance_summary(calendar_groups, display_group_keys, num_days)
    selected_day_default = real_now.day if now.year == real_now.year and now.month == real_now.month and real_now.day <= num_days else 1
    try:
        sel_day = int(request.args.get("day", selected_day_default) or selected_day_default)
    except Exception:
        sel_day = selected_day_default
    if sel_day < 1 or sel_day > num_days:
        sel_day = selected_day_default
    today_date = f"{now.year}-{now.month:02d}-{sel_day:02d}"
    today_assignment_records = latest_assignment_records_for_day(today_date, selected_department, selected_shift)
    today_duty_record = latest_duty_record_for_day(today_date, selected_shift)
    month_duty_records = latest_duty_records_for_month(now.year, now.month, selected_shift)

    # 第五分頁：整月工作表（日×工作位＋方案參考＋值日生），三部門各一份
    month_duty_set_by_day = {}
    for _rec in (month_duty_records or []):
        try:
            _dd = int(_rec.get("日"))
        except Exception:
            continue
        _nm = clean_text(_rec.get("值日生", ""))
        if not _nm:
            continue
        month_duty_set_by_day.setdefault(_dd, [])
        if _nm not in month_duty_set_by_day[_dd]:
            month_duty_set_by_day[_dd].append(_nm)
    month_pos_by_dept = {}
    for _dept in DEPARTMENTS:
        month_pos_by_dept[_dept] = build_month_position_view(
            _dept, selected_shift, now.year, now.month, num_days, work_schemes, month_duty_set_by_day)
    plan_status = get_plan_status(now.year, now.month, selected_department, selected_shift)
    anomaly_ctx = MonthlyAssignmentContext(selected_department, selected_shift, now.year, now.month)
    anomalies = build_anomaly_report(now.year, now.month, selected_department, selected_shift, ctx=anomaly_ctx)
    combined_stats = build_combined_stats(selected_department, selected_shift, now.year, now.month)

    # 每週清潔卡片：替每位清潔人員補上「部門」「該週日期範圍」「是否本週」
    weekly_raw = get_weekly_clean(now.year, now.month, selected_shift)
    name_dept = {}
    for u in calendar_users:
        if u.get("shift") == selected_shift and u.get("name"):
            name_dept.setdefault(u["name"], u.get("dept", ""))
    weekly_clean_view = []
    n_weeks = len(weekly_raw)
    for i, wk in enumerate(weekly_raw):
        start = i * 7 + 1
        end = num_days if i == n_weeks - 1 else min(start + 6, num_days)
        users = [{"name": nm, "dept": name_dept.get(nm, "")} for nm in wk.get("users", [])]
        is_current = (now.year == real_now.year and now.month == real_now.month
                      and start <= real_now.day <= end)
        weekly_clean_view.append({
            "label": wk.get("week", f"第 {i+1} 週"),
            "date_range": f"{now.month}/{start} – {now.month}/{end}",
            "users": users,
            "current": is_current,
        })
    weekly_clean_participants = len({u["name"] for wk in weekly_clean_view for u in wk["users"]})

    return render_template(
        "index.html",
        weekly_clean_view=weekly_clean_view,
        weekly_clean_participants=weekly_clean_participants,
        current_year=now.year,
        current_month=now.month,
        current_day=sel_day,
        combined_stats=combined_stats,
        selected_year=now.year,
        selected_month=now.month,
        selected_department=selected_department,
        selected_shift=selected_shift,
        selected_group_key=selected_group_key,
        shifts=SHIFTS,
        total_days=num_days,
        days=list(range(1, num_days + 1)),
        departments=DEPARTMENTS,
        display_departments=display_departments,
        leave_types=LEAVE_TYPES,
        duty_users=duty_users,
        calendar_users=calendar_users,
        calendar_groups=calendar_groups,
        history_people=history_people,
        duty_history_people=duty_history_people,
        attendance_summary=attendance_summary,
        work_schemes=work_schemes,
        work_items_map=load_work_items_map(),
        history_records=history_records,
        history_stats=history_stats,
        daily_schedule=[],
        weekly_schedule=get_weekly_clean(now.year, now.month, selected_shift),
        duty_members_by_shift=duty_members_by_shift(),
        today_date=today_date,
        today_assignment_records=today_assignment_records,
        today_duty_record=today_duty_record,
        month_duty_records=month_duty_records,
        month_pos_by_dept=month_pos_by_dept,
        plan_status=plan_status,
        anomalies=anomalies,
        is_admin=is_admin(),
        current_role=current_role(),
        operation_logs=load_operation_log_df().tail(80).iloc[::-1].to_dict("records"),
        system_settings=load_system_settings(),
        auto_backup=auto_backup_summary(),
        shift_time_label=shift_time_label(selected_shift),
    )


@app.route("/add_calendar_user", methods=["POST"])
def add_calendar_user():

    redir = require_admin_redirect()
    if redir:
        return redir
    """只新增到假表：新增不同姓名會往下新增，不會取代舊人員。"""
    name = clean_text(request.form.get("username", ""))
    dept = clean_text(request.form.get("department", "")) or "財務"
    shift = clean_text(request.form.get("shift", "晚班")) or "晚班"
    if shift not in SHIFTS:
        shift = "晚班"
    if dept not in DEPARTMENTS:
        dept = "財務"
    year = request.form.get("year") or datetime.now().year
    month = request.form.get("month") or datetime.now().month
    selected_department = dept
    if not name:
        return redirect(url_for("index", year=year, month=month, department=selected_department, shift=shift))

    df_basic, df_leave = init_attendance_sheets(year, month)

    existing_keys = {(clean_text(r.get("姓名", "")), clean_text(r.get("部門", "")), clean_text(r.get("班別", "晚班"))) for _, r in df_basic.iterrows()}
    if (name, dept, shift) not in existing_keys:
        df_basic = pd.concat([
            df_basic,
            pd.DataFrame([{"姓名": name, "部門": dept, "班別": shift, "狀態": "啟用", "備註": ""}])
        ], ignore_index=True)

    leave_names = [clean_text(n) for n in df_leave.get("姓名", pd.Series(dtype=str)).tolist()]
    if name not in leave_names:
        _, num_days = now_info(year, month)
        row = {"姓名": name}
        for d in range(1, num_days + 1):
            row[str(d)] = ""
        df_leave = pd.concat([df_leave, pd.DataFrame([row])], ignore_index=True)

    save_attendance_frames(df_basic, df_leave, year, month)
    log_operation("新增假表人員", dept, shift, year, month, f"新增 {name} 到假表")
    return redirect(url_for("index", year=year, month=month, department=selected_department, shift=shift))


@app.route("/api/update_member", methods=["POST"])
def update_member():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    old_name = clean_text(req.get("old_name", ""))
    new_name = clean_text(req.get("new_name", ""))
    dept = clean_text(req.get("department", ""))
    status = clean_text(req.get("status", "啟用")) or "啟用"
    note = clean_text(req.get("note", ""))
    if not old_name or not new_name:
        return jsonify({"status": "error", "message": "姓名不可空白"})
    if dept not in DEPARTMENTS:
        return jsonify({"status": "error", "message": "部門錯誤"})
    if status not in ["啟用", "停用"]:
        status = "啟用"
    df_basic, df_leave = init_attendance_sheets()
    if old_name not in df_basic["姓名"].tolist():
        return jsonify({"status": "error", "message": "找不到人員"})
    if old_name != new_name and new_name in df_basic["姓名"].tolist():
        return jsonify({"status": "error", "message": "新姓名已存在"})
    df_basic.loc[df_basic["姓名"] == old_name, ["姓名", "部門", "狀態", "備註"]] = [new_name, dept, status, note]
    df_leave.loc[df_leave["姓名"] == old_name, "姓名"] = new_name
    ok, msg = save_attendance_frames(df_basic, df_leave)
    if not ok:
        return jsonify({"status": "error", "message": msg})
    return jsonify({"status": "success", "message": "人員資料已更新"})


@app.route("/api/delete_member", methods=["POST"])
def delete_member():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    name = clean_text(req.get("name", ""))
    if not name:
        return jsonify({"status": "error", "message": "缺少姓名"})
    df_basic, df_leave = init_attendance_sheets(req.get("year"), req.get("month"))
    dept = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", ""))
    mask = df_basic["姓名"] == name
    if dept:
        mask = mask & (df_basic["部門"].astype(str).str.strip() == dept)
    if shift:
        mask = mask & (df_basic["班別"].astype(str).str.strip() == shift)
    df_basic = df_basic[~mask].reset_index(drop=True)

    # 刪除人員時，所有月份假表都一起移除此人，避免之後仍被分配。
    if db_enabled():
        try:
            with pg_conn() as conn, conn.cursor() as cur:
                if dept and shift:
                    cur.execute("DELETE FROM members WHERE name=%s AND dept=%s AND shift=%s", (name, dept, shift))
                elif dept:
                    cur.execute("DELETE FROM members WHERE name=%s AND dept=%s", (name, dept))
                else:
                    cur.execute("DELETE FROM members WHERE name=%s", (name,))
                cur.execute("SELECT COUNT(*) FROM members WHERE name=%s", (name,))
                if int(cur.fetchone()[0]) == 0:
                    cur.execute("DELETE FROM leaves WHERE name=%s", (name,))
                conn.commit()
            return jsonify({"status": "success", "message": "人員已刪除"})
        except Exception as e:
            return jsonify({"status": "error", "message": f"刪除失敗：{e}"})
    if os.path.exists(FILE_LEAVE):
        try:
            sheets = {}
            with pd.ExcelFile(FILE_LEAVE, engine="openpyxl") as xls:
                for sheet in xls.sheet_names:
                    if sheet == SHEET_BASIC:
                        continue
                    df = pd.read_excel(xls, sheet_name=sheet)
                    if "姓名" in df.columns:
                        df = df[df["姓名"].fillna("").astype(str).str.strip() != name].reset_index(drop=True)
                    sheets[sheet] = df
            with pd.ExcelWriter(FILE_LEAVE, engine="openpyxl", mode="w") as writer:
                df_basic.to_excel(writer, sheet_name=SHEET_BASIC, index=False)
                for sheet, df in sheets.items():
                    df.to_excel(writer, sheet_name=str(sheet)[:31], index=False)
        except Exception as e:
            return jsonify({"status": "error", "message": ensure_excel_not_open_error(e)})
    return jsonify({"status": "success", "message": "人員已刪除"})


@app.route("/api/toggle_leave", methods=["POST"])
def toggle_leave():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    name = clean_text(req.get("name", ""))
    day = clean_text(req.get("day", ""))
    status = clean_text(req.get("status", "休"))
    if not name or not day:
        return jsonify({"status": "error", "message": "缺少姓名或日期"})
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    if status not in LEAVE_TYPES:
        status = "休"
    _, df_leave = init_attendance_sheets(year, month)
    if day not in df_leave.columns:
        df_leave[day] = ""
    idx = df_leave[df_leave["姓名"] == name].index
    if len(idx) == 0:
        return jsonify({"status": "error", "message": "找不到該同仁"})
    old_value = clean_leave_value(df_leave.at[idx[0], day])
    new_value = "" if old_value == status else status
    df_leave.at[idx[0], day] = new_value
    df_basic, _ = init_attendance_sheets(year, month)
    ok, msg = save_attendance_frames(df_basic, df_leave, year, month)
    if not ok:
        return jsonify({"status": "error", "message": msg})
    log_operation("修改假表", content=f"{name} {month}/{day} 由「{old_value or '上班'}」改為「{new_value or '上班'}」", year=year, month=month)
    # 路二：自動偵測，重排該人所屬部門班別的「當天到月底」，把平均補回來
    reassign_summaries = auto_reassign_after_leave_change(name, int(day), year, month)
    if reassign_summaries:
        log_operation("自動補位（請假連動）", year=year, month=month,
                      content="；".join(reassign_summaries),
                      reason=f"{name} {month}/{day} 假別變更為「{new_value or '上班'}」")
    # 值日生連動：若該人在值日生名單，重排其班別的「當天到月底」值日生
    duty_summaries = auto_reassign_duty_after_leave_change(name, int(day), year, month)
    if duty_summaries:
        log_operation("自動補位（值日生連動）", year=year, month=month,
                      content="；".join(duty_summaries),
                      reason=f"{name} {month}/{day} 假別變更為「{new_value or '上班'}」")
    return jsonify({"status": "success", "value": new_value, "reassign": reassign_summaries + duty_summaries})


@app.route("/api/add_work_scheme", methods=["POST"])
def add_work_scheme():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    scheme_name = clean_text(req.get("scheme_name", "")) or "新方案"
    tasks = parse_tasks(req.get("tasks", ""))
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status": "error", "message": "部門或班別錯誤"})
    if len(tasks) < 1:
        return jsonify({"status": "error", "message": "請至少輸入 1 個工作內容"})
    schemes = load_work_schemes()
    scheme = make_scheme(scheme_name, tasks)
    bucket = str(scheme["task_count"])
    schemes.setdefault(group_key(department, shift), {}).setdefault(bucket, []).append(scheme)
    if not save_work_schemes(schemes):
        return jsonify({"status": "error", "message": "新增方案失敗"})
    log_operation("新增工作方案", department, shift, content=f"{scheme_name}：{bucket} 人方案")
    return jsonify({"status": "success", "message": f"已新增 {scheme_name}，自動分類為 {bucket} 人方案", "scheme": scheme})


@app.route("/api/add_work_item", methods=["POST"])
def add_work_item():
    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", "")) or "晚班"
    name = clean_text(req.get("name", ""))
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status": "error", "message": "部門或班別錯誤"})
    if not name:
        return jsonify({"status": "error", "message": "請輸入項目名稱"})
    if " " in name or "\u3000" in name:
        return jsonify({"status": "error", "message": "項目名稱請勿包含空格"})
    m = load_work_items_map()
    key = group_key(department, shift)
    items = m.get(key, [])
    if any(it["name"] == name for it in items):
        return jsonify({"status": "error", "message": "這個項目已經存在了"})
    items.append({"name": name})
    m[key] = items
    ok, err = save_work_items_map(m)
    if not ok:
        return jsonify({"status": "error", "message": f"儲存失敗：{err}"})
    log_operation("新增工作項目", department, shift, content=name)
    return jsonify({"status": "success", "message": f"已新增項目「{name}」", "department": department, "shift": shift, "items": items})


@app.route("/api/delete_work_item", methods=["POST"])
def delete_work_item():
    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", "")) or "晚班"
    name = clean_text(req.get("name", ""))
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status": "error", "message": "部門或班別錯誤"})
    m = load_work_items_map()
    key = group_key(department, shift)
    items = m.get(key, [])
    new_items = [it for it in items if it["name"] != name]
    if len(new_items) == len(items):
        return jsonify({"status": "error", "message": "找不到這個項目"})
    m[key] = new_items
    ok, err = save_work_items_map(m)
    if not ok:
        return jsonify({"status": "error", "message": f"儲存失敗：{err}"})
    log_operation("刪除工作項目", department, shift, content=name)
    return jsonify({"status": "success", "message": f"已刪除項目「{name}」", "department": department, "shift": shift, "items": new_items})


@app.route("/api/save_work_scheme", methods=["POST"])
def save_work_scheme():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    scheme_id = clean_text(req.get("scheme_id", ""))
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    scheme_name = clean_text(req.get("scheme_name", "")) or "未命名方案"
    tasks = parse_tasks(req.get("tasks", ""))
    if department not in DEPARTMENTS or shift not in SHIFTS or not scheme_id:
        return jsonify({"status": "error", "message": "部門、班別或方案 ID 錯誤"})
    if len(tasks) < 1:
        return jsonify({"status": "error", "message": "至少要有 1 個工作"})

    schemes = load_work_schemes()
    found = None
    # 先從所有人數桶移除舊方案，再依新工作數量放回正確分類
    skey = group_key(department, shift)
    for bucket in list(schemes.get(skey, {}).keys()):
        new_list = []
        for item in schemes[skey][bucket]:
            if item.get("id") == scheme_id:
                found = item
            else:
                new_list.append(item)
        schemes[skey][bucket] = new_list
        if not schemes[skey][bucket]:
            del schemes[skey][bucket]

    scheme = make_scheme(scheme_name, tasks, scheme_id=scheme_id)
    bucket = str(scheme["task_count"])
    schemes.setdefault(skey, {}).setdefault(bucket, []).append(scheme)
    if not save_work_schemes(schemes):
        return jsonify({"status": "error", "message": "儲存方案失敗"})
    log_operation("儲存工作方案", department, shift, content=f"{scheme_name}：{bucket} 人方案")
    return jsonify({"status": "success", "message": f"已儲存，並分類到 {bucket} 人方案", "scheme": scheme})


@app.route("/api/delete_work_scheme", methods=["POST"])
def delete_work_scheme():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    scheme_id = clean_text(req.get("scheme_id", ""))
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    if department not in DEPARTMENTS or shift not in SHIFTS or not scheme_id:
        return jsonify({"status": "error", "message": "資料錯誤"})
    schemes = load_work_schemes()
    deleted = False
    skey = group_key(department, shift)
    for bucket in list(schemes.get(skey, {}).keys()):
        before = len(schemes[skey][bucket])
        schemes[skey][bucket] = [s for s in schemes[skey][bucket] if s.get("id") != scheme_id]
        if len(schemes[skey][bucket]) != before:
            deleted = True
        if not schemes[skey][bucket]:
            del schemes[skey][bucket]
    if not deleted:
        return jsonify({"status": "error", "message": "找不到方案"})
    save_work_schemes(schemes)
    log_operation("刪除工作方案", department, shift, content=f"刪除方案 ID：{scheme_id}")
    return jsonify({"status": "success", "message": "方案已刪除"})


@app.route("/api/generate_work_assignment", methods=["POST"])
def generate_work_assignment():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    day = clean_text(req.get("day", ""))
    mode = clean_text(req.get("mode", "balanced_random")) or "balanced_random"
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status": "error", "message": "部門或班別錯誤"})
    if not day or not day.isdigit():
        return jsonify({"status": "error", "message": "請選擇日期"})
    if mode not in ["fixed", "balanced_random", "random"]:
        mode = "balanced_random"
    result = create_assignment_for_day(department, int(day), mode=mode, save=True, year=year, month=month, shift=shift)
    if result.get("status") != "success":
        return jsonify({"status": "error", "message": result.get("message", "產生失敗"), "result": result})
    return jsonify({"status": "success", "result": result})



def replace_month_department_history(year, month, department, shift, results):
    """整月分配採用覆蓋制：同部門同年月先清掉舊結果，再寫入本次整月結果。"""
    df = load_history_df()
    if not df.empty:
        mask = (
            (df["年份"].astype(str) == str(year)) &
            (df["月份"].astype(str) == str(month)) &
            (df["部門"].astype(str) == str(department)) &
            (df["班別"].astype(str) == str(shift))
        )
        df = df[~mask].reset_index(drop=True)
    rows = []
    for result in results:
        if result.get("status") != "success":
            continue
        dt = datetime.strptime(result["date"], "%Y-%m-%d")
        for item in result.get("assignments", []):
            rows.append({
                "日期": result["date"],
                "年份": dt.year,
                "月份": dt.month,
                "日": dt.day,
                "部門": result["department"],
                "班別": result.get("shift", "晚班"),
                "人員": item["user"],
                "工作": item["task"],
                "方案": result["scheme"],
                "分配方式": "平均隨機" if result.get("mode") in ["balanced_random", "random"] else "固定順序",
                "產生時間": result["created_at"],
                "批次ID": result["batch_id"],
            })
    if rows:
        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
    return save_history_df(df)


def produce_duty_for_month(shift, year, month):
    """產生（或重產）某班別本月每日值日生與每週清潔。會先清掉本月該班別舊值日生避免疊加。
    回傳 dict：status / message / shift / daily_schedule / weekly_schedule。"""
    if shift not in SHIFTS:
        shift = default_duty_shift()
    duty_users = load_duty_members(shift)
    if len(duty_users) < 1:
        return {"status": "error", "message": f"「{shift}」值日生名單內沒有同仁，請先在該班別新增值日生"}
    now, num_days = now_info(year, month)
    year, month = now.year, now.month
    _, df_leave = init_attendance_sheets(year, month)
    week_day_map = {0: "星期一", 1: "星期二", 2: "星期三", 3: "星期四", 4: "星期五", 5: "星期六", 6: "星期日"}
    leave_records = {}
    if not df_leave.empty:
        for _, row in df_leave.iterrows():
            nm = clean_text(row.get("姓名", ""))
            if nm:
                leave_records[nm] = row.to_dict()

    # 先刪本月該班別舊值日生（避免疊加），再以剩餘歷史算公平基礎
    dfh = load_duty_history_df()
    if not dfh.empty:
        old_mask = (
            (dfh["年份"].astype(str) == str(year)) &
            (dfh["月份"].astype(str) == str(month)) &
            (dfh["班別"].astype(str).str.strip() == str(shift))
        )
        if old_mask.any():
            dfh = dfh[~old_mask].reset_index(drop=True)
            save_duty_history_df(dfh)

    # 只看當月：名單上的人每人從 0 起算，整月平均分配（不累積跨月歷史）
    daily_counts = {user: 0 for user in duty_users}

    duty_count = daily_duty_count_for_shift(shift)
    daily_schedule = []
    for day in range(1, num_days + 1):
        date_obj = datetime(year, month, day)
        available_today = [u for u in duty_users
                           if clean_leave_value(leave_records.get(u, {}).get(str(day), "")) == ""]
        if not available_today:
            available_today = duty_users.copy()
        pick = min(duty_count, len(available_today))
        chosen_list = []
        for _ in range(pick):
            pool = [u for u in available_today if u not in chosen_list]
            if not pool:
                break
            min_count = min(daily_counts[u] for u in pool)
            candidates = [u for u in pool if daily_counts[u] == min_count]
            chosen = random.choice(candidates)
            chosen_list.append(chosen)
            daily_counts[chosen] += 1
        ds = date_obj.strftime("%Y-%m-%d")
        wd = week_day_map[date_obj.weekday()]
        for chosen in chosen_list:
            daily_schedule.append({"date": ds, "week_day": wd, "user": chosen})

    ok, msg = append_duty_history(daily_schedule, shift)
    if not ok:
        return {"status": "error", "message": msg}

    _settings = load_system_settings()
    cleaner_count = max(1, int(_settings.get("weekly_cleaner_count", 3)))
    for _s in _settings.get("shifts", []):
        if _s.get("name") == shift:
            try:
                cleaner_count = max(1, int(_s.get("cleaner_count", cleaner_count)))
            except Exception:
                pass
            break
    weekly_counts = {u: 0 for u in duty_users}
    weekly_schedule = []
    for week_num in range(1, 5):
        week_users = []
        for _ in range(min(cleaner_count, len(duty_users))):
            pool = [u for u in duty_users if u not in week_users]
            min_count = min(weekly_counts[u] for u in pool)
            candidates = [u for u in pool if weekly_counts[u] == min_count]
            chosen = random.choice(candidates)
            week_users.append(chosen)
            weekly_counts[chosen] += 1
        weekly_schedule.append({"week": f"第 {week_num} 週", "users": week_users})
    set_weekly_clean(year, month, shift, weekly_schedule)

    log_operation("產生每日值日生與清潔", shift=shift, year=year, month=month, content=f"{shift}｜每日 {duty_count} 人，共 {len(daily_schedule)} 筆，每週清潔 {len(weekly_schedule)} 週")
    return {"status": "success", "message": f"已產生 {year}/{month}「{shift}」每日值日生，並存入歷史", "shift": shift, "daily_schedule": daily_schedule, "weekly_schedule": weekly_schedule}


def duty_exists_for_month(shift, year, month):
    """該班別本月是否已產生過值日生。"""
    now, _ = now_info(year, month)
    df = load_duty_history_df()
    if df.empty:
        return False
    mask = (
        (df["年份"].astype(str) == str(now.year)) &
        (df["月份"].astype(str) == str(now.month)) &
        (df["班別"].astype(str).str.strip() == str(shift))
    )
    return bool(mask.any())


@app.route("/api/generate_month_assignments", methods=["POST"])
def generate_month_assignments():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    mode = clean_text(req.get("mode", "balanced_random")) or "balanced_random"
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status": "error", "message": "部門或班別錯誤"})
    if is_plan_locked(year, month, department, shift):
        return jsonify({"status": "error", "message": "本月班表已鎖定，請先解鎖後再重新產生整月工作表。"})
    selected, num_days = now_info(year, month)
    year, month = selected.year, selected.month
    ctx = MonthlyAssignmentContext(department, shift, year, month)
    batch_id = uuid.uuid4().hex[:12]
    results = []
    errors = []
    # 為了整月內也公平，使用本次本地計數暫時修正
    local_counts = {}
    for day in range(1, num_days + 1):
        result = create_assignment_for_day(department, day, mode=mode, save=False, batch_id=batch_id, local_counts=local_counts, year=year, month=month, shift=shift, ctx=ctx)
        if result.get("status") == "success":
            results.append(result)
            for item in result.get("assignments", []):
                local_counts[item["user"]] = local_counts.get(item["user"], 0) + 1
        elif result.get("status") == "error":
            errors.append(result.get("message"))
    ok, msg = replace_month_department_history(year, month, department, shift, results)
    if not ok:
        return jsonify({"status": "error", "message": msg})
    # 同步：該班別本月若還沒值日生，就自動產一份（含每週清潔），達成「工作表＋值日生一次產生」
    duty_note = ""
    if not duty_exists_for_month(shift, year, month):
        dres = produce_duty_for_month(shift, year, month)
        if dres.get("status") == "success":
            duty_note = f"，並已自動產生「{shift}」本月每日值日生"
    log_operation("產生／更新整月工作表", department, shift, year, month, f"共 {len(results)} 天，錯誤 {len(errors)} 筆{duty_note}")
    return jsonify({
        "status": "success",
        "message": f"已重新產生並覆蓋 {department} {shift} {year}/{month} 整月工作表，共 {len(results)} 天{duty_note}。當日分配結果會直接從這份整月結果讀取。",
        "results": results,
        "errors": errors[:20]
    })


def replace_day_department_history(year, month, day, department, shift, result):
    """臨時請假補位：只覆蓋指定日期＋部門＋班別，不動整個月份其他天。"""
    selected, _ = now_info(year, month)
    year, month = selected.year, selected.month
    date_str = f"{year}-{month:02d}-{int(day):02d}"
    df = load_history_df()
    if not df.empty:
        mask = (
            (df["日期"].astype(str) == date_str) &
            (df["部門"].astype(str) == str(department)) &
            (df["班別"].astype(str) == str(shift))
        )
        df = df[~mask].reset_index(drop=True)
    rows = []
    if result.get("status") == "success":
        dt = datetime.strptime(result["date"], "%Y-%m-%d")
        for item in result.get("assignments", []):
            rows.append({
                "日期": result["date"],
                "年份": dt.year,
                "月份": dt.month,
                "日": dt.day,
                "部門": result["department"],
                "班別": result.get("shift", "晚班"),
                "人員": item["user"],
                "工作": item["task"],
                "方案": result["scheme"],
                "分配方式": "臨時補位／平均隨機" if result.get("mode") in ["balanced_random", "random"] else "臨時補位／固定順序",
                "產生時間": result["created_at"],
                "批次ID": result["batch_id"],
            })
    if rows:
        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
    return save_history_df(df)


def reassign_from_day(department, shift, start_day, year, month):
    """路二：重排 [start_day..月底] 的工作分配，讓本月剩餘天數補回平均。
    僅在該部門班別本月已排過整月工作時才動作；start_day 之前已執行的天不動。"""
    selected, num_days = now_info(year, month)
    year, month = selected.year, selected.month
    start_day = max(1, int(start_day))

    df = load_history_df()
    if df.empty:
        return {"reassigned": False, "days": 0}
    month_mask = (
        (df["年份"].astype(str) == str(year)) &
        (df["月份"].astype(str) == str(month)) &
        (df["部門"].astype(str) == str(department)) &
        (df["班別"].astype(str) == str(shift))
    )
    if not month_mask.any():
        # 該部門班別本月尚未排過整月工作 → 不自動補位（避免無中生有只排幾天）
        return {"reassigned": False, "days": 0}

    # 刪除 start_day 起的舊分配，保留 1~(start_day-1)
    drop_mask = month_mask & (pd.to_numeric(df["日"], errors="coerce") >= start_day)
    df = df[~drop_mask].reset_index(drop=True)
    ok, msg = save_history_df(df)
    if not ok:
        return {"reassigned": False, "days": 0, "error": msg}

    # 此時歷史只剩 1~(start_day-1)，ctx.base_counts 即為已固定部分；逐天往後補回平均
    ctx = MonthlyAssignmentContext(department, shift, year, month, seed_task_counts=True)
    batch_id = uuid.uuid4().hex[:12]
    local_counts = {}
    results = []
    for day in range(start_day, num_days + 1):
        r = create_assignment_for_day(department, day, mode="balanced_random", save=False,
                                      batch_id=batch_id, local_counts=local_counts,
                                      year=year, month=month, shift=shift, ctx=ctx)
        if r.get("status") == "success":
            results.append(r)
            for item in r.get("assignments", []):
                local_counts[item["user"]] = local_counts.get(item["user"], 0) + 1
    if results:
        append_history(results)
    return {"reassigned": True, "days": len(results), "start_day": start_day}


def auto_reassign_after_leave_change(name, day, year, month):
    """某人劃假／銷假後，自動重排其所屬部門班別的 [day..月底]（路二）。回傳重排摘要。"""
    selected, _ = now_info(year, month)
    year, month = selected.year, selected.month
    df_basic, _ = init_attendance_sheets(year, month)
    combos = set()
    for _, row in df_basic.iterrows():
        nm = clean_text(row.get("姓名", ""))
        if nm != name:
            continue
        dept = clean_text(row.get("部門", "")) or "財務"
        sh = clean_text(row.get("班別", "晚班")) or "晚班"
        if dept in DEPARTMENTS and sh in SHIFTS:
            combos.add((dept, sh))
    summaries = []
    for dept, sh in combos:
        res = reassign_from_day(dept, sh, day, year, month)
        if res.get("reassigned"):
            summaries.append(f"{dept}｜{sh}：已自動重排 {res.get('start_day')} 號起共 {res.get('days')} 天")
    return summaries


def reassign_duty_from_day(shift, start_day, year, month):
    """值日生版『到月底重排』：保留 [1..start_day-1]，重排 [start_day..月底]，公平拉平。
    僅在該班別本月已產生過值日生時才動作。挑人邏輯與 generate_schedule 完全一致。"""
    selected, num_days = now_info(year, month)
    year, month = selected.year, selected.month
    start_day = max(1, int(start_day))

    duty_users = load_duty_members(shift)
    if len(duty_users) < 1:
        return {"reassigned": False, "days": 0}

    df = load_duty_history_df()
    if df.empty:
        return {"reassigned": False, "days": 0}
    month_mask = (
        (df["年份"].astype(str) == str(year)) &
        (df["月份"].astype(str) == str(month)) &
        (df["班別"].astype(str).str.strip() == str(shift))
    )
    if not month_mask.any():
        return {"reassigned": False, "days": 0}

    # 刪除 start_day 起的舊值日生，保留 1~(start_day-1)
    drop_mask = month_mask & (pd.to_numeric(df["日"], errors="coerce") >= start_day)
    df = df[~drop_mask].reset_index(drop=True)
    ok, msg = save_duty_history_df(df)
    if not ok:
        return {"reassigned": False, "days": 0, "error": msg}

    # 公平基礎：只看本月已保留的 1~(start_day-1)，讓整月平均（不累積跨月歷史）
    daily_counts = {u: 0 for u in duty_users}
    if not df.empty:
        kept = df[
            (df["年份"].astype(str) == str(year)) &
            (df["月份"].astype(str) == str(month)) &
            (df["班別"].astype(str).str.strip() == str(shift))
        ]
        for _, _r in kept.iterrows():
            _nm = clean_text(_r.get("值日生", ""))
            if _nm in daily_counts:
                daily_counts[_nm] += 1

    _, df_leave = init_attendance_sheets(year, month)
    leave_records = {}
    if not df_leave.empty:
        for _, row in df_leave.iterrows():
            nm = clean_text(row.get("姓名", ""))
            if nm:
                leave_records[nm] = row.to_dict()

    duty_count = daily_duty_count_for_shift(shift)
    if duty_count < 1:
        return {"reassigned": False, "days": 0}
    week_day_map = {0: "星期一", 1: "星期二", 2: "星期三", 3: "星期四", 4: "星期五", 5: "星期六", 6: "星期日"}
    now, _ = now_info(year, month)
    daily_schedule = []
    for day in range(start_day, num_days + 1):
        date_obj = datetime(now.year, now.month, day)
        available_today = [u for u in duty_users
                           if clean_leave_value(leave_records.get(u, {}).get(str(day), "")) == ""]
        if not available_today:
            available_today = duty_users.copy()
        pick = min(duty_count, len(available_today))
        chosen_list = []
        for _ in range(pick):
            pool = [u for u in available_today if u not in chosen_list]
            if not pool:
                break
            min_count = min(daily_counts[u] for u in pool)
            candidates = [u for u in pool if daily_counts[u] == min_count]
            chosen = random.choice(candidates)
            chosen_list.append(chosen)
            daily_counts[chosen] += 1
        ds = date_obj.strftime("%Y-%m-%d")
        wd = week_day_map[date_obj.weekday()]
        for chosen in chosen_list:
            daily_schedule.append({"date": ds, "week_day": wd, "user": chosen})

    if daily_schedule:
        ok, msg = append_duty_history(daily_schedule, shift)
        if not ok:
            return {"reassigned": False, "days": 0, "error": msg}
    days_count = len({d["date"] for d in daily_schedule})
    return {"reassigned": True, "days": days_count, "start_day": start_day}


def auto_reassign_duty_after_leave_change(name, day, year, month):
    """某人劃假／銷假後，自動重排其所在值日生班別的 [day..月底]。回傳摘要。"""
    selected, _ = now_info(year, month)
    year, month = selected.year, selected.month
    summaries = []
    for sh in SHIFTS:
        if name in load_duty_members(sh):
            res = reassign_duty_from_day(sh, day, year, month)
            if res.get("reassigned"):
                summaries.append(f"{sh} 值日生：已自動重排 {res.get('start_day')} 號起共 {res.get('days')} 天")
    return summaries


@app.route("/api/reassign_day", methods=["POST"])
def reassign_day():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    day = clean_text(req.get("day", ""))
    mode = clean_text(req.get("mode", "balanced_random")) or "balanced_random"
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status": "error", "message": "部門或班別錯誤"})
    locked = is_plan_locked(year, month, department, shift)
    reason = clean_text(req.get("reason", ""))
    if locked and not reason:
        return jsonify({"status": "error", "message": "本月班表已鎖定，臨時補位必須填寫原因。"})
    if not day or not str(day).isdigit():
        return jsonify({"status": "error", "message": "缺少日期"})
    if mode not in ["fixed", "balanced_random", "random"]:
        mode = "balanced_random"
    batch_id = uuid.uuid4().hex[:12]
    result = create_assignment_for_day(department, int(day), mode=mode, save=False, batch_id=batch_id, year=year, month=month, shift=shift)
    if result.get("status") != "success":
        return jsonify({"status": "error", "message": result.get("message", "補位失敗"), "result": result})
    ok, msg = replace_day_department_history(year, month, int(day), department, shift, result)
    if not ok:
        return jsonify({"status": "error", "message": msg})
    log_operation("臨時補位當天", department, shift, year, month, f"{result['date']} 重新補位", reason)
    return jsonify({
        "status": "success",
        "message": f"已重新補位 {result['date']} {department} {shift}，只更新這一天，不影響整個月份其他日期。",
        "result": result
    })


@app.route("/api/history", methods=["GET"])
def api_history():
    start_date = clean_text(request.args.get("start_date", ""))
    end_date = clean_text(request.args.get("end_date", ""))
    department = clean_text(request.args.get("department", ""))
    person = clean_text(request.args.get("person", ""))
    shift = clean_text(request.args.get("shift", ""))

    if not start_date or not end_date:
        return jsonify({"status": "error", "message": "請先選擇開始日期與結束日期"})

    records, stats = filter_history_by_range(start_date, end_date, department, person, shift)
    return jsonify({"status": "success", "records": records, "stats": stats})


@app.route("/api/duty_history", methods=["GET"])
def api_duty_history():
    start_date = clean_text(request.args.get("start_date", ""))
    end_date = clean_text(request.args.get("end_date", ""))
    person = clean_text(request.args.get("person", ""))
    shift = clean_text(request.args.get("shift", ""))
    if not start_date or not end_date:
        return jsonify({"status": "error", "message": "請先選擇開始日期與結束日期"})
    records, stats = filter_duty_history_by_range(start_date, end_date, person, shift)
    return jsonify({"status": "success", "records": records, "stats": stats})



@app.route("/api/set_plan_status", methods=["POST"])
def api_set_plan_status():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    status = clean_text(req.get("status", "草稿中")) or "草稿中"
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status": "error", "message": "部門或班別錯誤"})
    ok, msg = set_plan_status(year, month, department, shift, status)
    if not ok:
        return jsonify({"status": "error", "message": msg})
    log_operation("班表狀態變更", department, shift, year, month, f"狀態改為：{status}")
    return jsonify({"status": "success", "message": f"{department} {shift} {year}/{month} 班表狀態已更新為：{status}"})


@app.route("/api/copy_previous_month", methods=["POST"])
def api_copy_previous_month():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", ""))
    if department and department not in DEPARTMENTS:
        return jsonify({"status": "error", "message": "部門錯誤"})
    if shift and shift not in SHIFTS:
        return jsonify({"status": "error", "message": "班別錯誤"})
    ok, msg = copy_previous_month_leave(year, month, department or None, shift or None)
    if ok:
        log_operation("複製上個月假表", department, shift, year, month, msg)
    return jsonify({"status": "success" if ok else "error", "message": msg})



@app.route("/save_settings", methods=["POST"])
def save_settings_route():
    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    current = load_system_settings()
    title = clean_text(request.form.get("system_title", current.get("system_title", "財務部門智能排班系統"))) or "財務部門智能排班系統"
    pwd_input = clean_text(request.form.get("admin_password", ""))
    if pwd_input:
        admin_password = generate_password_hash(pwd_input)
    else:
        admin_password = current.get("admin_password", "9527")
    departments_raw = request.form.get("departments", "")
    departments = [clean_text(x) for x in departments_raw.replace("，", ",").replace("\n", ",").split(",") if clean_text(x)]
    if not departments:
        departments = current.get("departments", ["財務", "BBP", "四方"])
    shift_names = request.form.getlist("shift_name")
    shift_starts = request.form.getlist("shift_start")
    shift_ends = request.form.getlist("shift_end")
    shift_duty_counts = request.form.getlist("shift_duty_count")
    shift_cleaner_counts = request.form.getlist("shift_cleaner_count")
    shifts = []
    for i, name in enumerate(shift_names):
        nm = clean_text(name)
        if not nm:
            continue
        try:
            dc = max(0, int(shift_duty_counts[i])) if i < len(shift_duty_counts) and clean_text(shift_duty_counts[i]) != "" else 1
        except Exception:
            dc = 1
        try:
            cc = max(1, int(shift_cleaner_counts[i])) if i < len(shift_cleaner_counts) and clean_text(shift_cleaner_counts[i]) != "" else 3
        except Exception:
            cc = 3
        shifts.append({"name": nm, "start": clean_text(shift_starts[i] if i < len(shift_starts) else ""), "end": clean_text(shift_ends[i] if i < len(shift_ends) else ""), "duty_count": dc, "cleaner_count": cc})
    if not shifts:
        shifts = current.get("shifts", default_system_settings()["shifts"])
    try:
        weekly_cleaner_count = max(1, int(request.form.get("weekly_cleaner_count", current.get("weekly_cleaner_count", 3))))
    except Exception:
        weekly_cleaner_count = 3
    try:
        auto_backup_keep = max(3, int(request.form.get("auto_backup_keep", current.get("auto_backup_keep", 10))))
    except Exception:
        auto_backup_keep = 10
    data = {
        "system_title": title,
        "admin_password": admin_password,
        "departments": departments,
        "shifts": shifts,
        "weekly_cleaner_count": weekly_cleaner_count,
        "auto_backup_enabled": request.form.get("auto_backup_enabled") == "on",
        "auto_backup_keep": auto_backup_keep,
    }
    ok, msg = save_system_settings(data)
    if ok:
        log_operation("修改系統設定", content="系統設定已更新")
    return redirect(url_for("index", restore_msg=msg))


@app.route("/today")
def today_dashboard():
    real_now = datetime.now()
    selected_year = int(request.args.get("year", real_now.year) or real_now.year)
    selected_month = int(request.args.get("month", real_now.month) or real_now.month)
    selected_department = clean_text(request.args.get("department", DEPARTMENTS[0])) or DEPARTMENTS[0]
    selected_shift = clean_text(request.args.get("shift", SHIFTS[-1])) or SHIFTS[-1]
    if selected_department not in DEPARTMENTS: selected_department = DEPARTMENTS[0]
    if selected_shift not in SHIFTS: selected_shift = SHIFTS[-1]
    now, num_days = now_info(selected_year, selected_month)
    default_day = real_now.day if now.year == real_now.year and now.month == real_now.month and real_now.day <= num_days else 1
    try:
        day = int(request.args.get("day", default_day) or default_day)
    except Exception:
        day = default_day
    if day < 1 or day > num_days:
        day = default_day
    date_str = f"{now.year}-{now.month:02d}-{day:02d}"
    _, calendar_groups = build_calendar_data(now.year, now.month)
    selected_key = group_key(selected_department, selected_shift)
    users = calendar_groups.get(selected_key, [])
    working = [u for u in users if not u.get("leaves", {}).get(str(day), "")]
    leaves = [u for u in users if u.get("leaves", {}).get(str(day), "")]
    records = latest_assignment_records_for_day(date_str, selected_department, selected_shift)
    duty_record = latest_duty_record_for_day(date_str, selected_shift)
    return render_template("today.html", system_settings=load_system_settings(), departments=DEPARTMENTS, shifts=SHIFTS, selected_department=selected_department, selected_shift=selected_shift, selected_year=now.year, selected_month=now.month, current_day=day, total_days=num_days, today_date=date_str, working=working, leaves=leaves, records=records, duty_record=duty_record, shift_time_label=shift_time_label(selected_shift))


@app.route("/download_latest_auto_backup")
def download_latest_auto_backup():
    backups = sorted(Path(AUTO_BACKUP_DIR).glob("auto_backup_*.zip"), key=lambda x: x.stat().st_mtime, reverse=True) if os.path.exists(AUTO_BACKUP_DIR) else []
    if not backups:
        path = make_auto_backup("manual_latest")
    else:
        path = str(backups[0])
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

@app.route("/download_backup")
def download_backup():
    path = make_backup_zip()
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/restore_backup", methods=["POST"])
def restore_backup():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    ok, msg = restore_backup_zip(request.files.get("backup_file"))
    return redirect(url_for("index", restore_msg=msg))


@app.route("/export_month_excel")
def export_month_excel():
    year = request.args.get("year") or datetime.now().year
    month = request.args.get("month") or datetime.now().month
    department = clean_text(request.args.get("department", "財務")) or "財務"
    shift = clean_text(request.args.get("shift", "晚班")) or "晚班"
    if department not in DEPARTMENTS:
        department = "財務"
    if shift not in SHIFTS:
        shift = "晚班"
    path = export_month_report_excel(year, month, department, shift)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/export_month_pdf")
def export_month_pdf():
    year = request.args.get("year") or datetime.now().year
    month = request.args.get("month") or datetime.now().month
    department = clean_text(request.args.get("department", "財務")) or "財務"
    shift = clean_text(request.args.get("shift", "晚班")) or "晚班"
    if department not in DEPARTMENTS:
        department = "財務"
    if shift not in SHIFTS:
        shift = "晚班"
    path = export_month_report_pdf(year, month, department, shift)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/download_duty_history")
def download_duty_history():
    if not os.path.exists(DUTY_HISTORY_FILE):
        df = load_duty_history_df()
        save_duty_history_df(df)
    return send_file(DUTY_HISTORY_FILE, as_attachment=True, download_name="duty_schedule_history.xlsx")


@app.route("/download_history")
def download_history():
    if not os.path.exists(HISTORY_FILE):
        df = load_history_df()
        save_history_df(df)
    return send_file(HISTORY_FILE, as_attachment=True, download_name="work_assignment_history.xlsx")


@app.route("/add_duty_user", methods=["POST"])
def add_duty_user():

    redir = require_admin_redirect()
    if redir:
        return redir
    name = clean_text(request.form.get("username", ""))
    shift = clean_text(request.form.get("shift", "")) or default_duty_shift()
    if shift not in SHIFTS:
        shift = default_duty_shift()
    if name:
        df = load_duty_members_df()
        exists = ((df["姓名"].astype(str).str.strip() == name) & (df["班別"].astype(str).str.strip() == shift)).any()
        if not exists:
            df = pd.concat([df, pd.DataFrame([{"姓名": name, "班別": shift}])], ignore_index=True)
            if not save_duty_members_df(df):
                return redirect(url_for("index", login_msg="值日生名單儲存失敗，請確認資料檔案沒有被佔用後再試。"))
            log_operation("新增值日生", shift=shift, content=f"新增值日生 {name}（{shift}）")
    return redirect(url_for("index",
        department=clean_text(request.form.get("selected_department", "")),
        shift=clean_text(request.form.get("selected_shift", "")) or shift,
        year=request.form.get("year", ""),
        month=request.form.get("month", "")))


@app.route("/delete_duty_user/<shift>/<name>")
def delete_duty_user(shift, name):

    redir = require_admin_redirect()
    if redir:
        return redir
    name = clean_text(name)
    shift = clean_text(shift)
    df = load_duty_members_df()
    mask = (df["姓名"].astype(str).str.strip() == name) & (df["班別"].astype(str).str.strip() == shift)
    if mask.any():
        df = df[~mask].reset_index(drop=True)
        save_duty_members_df(df)
        log_operation("刪除值日生", shift=shift, content=f"刪除值日生 {name}（{shift}）")
    return redirect(url_for("index",
        department=clean_text(request.args.get("department", "")),
        shift=shift,
        year=request.args.get("year", ""),
        month=request.args.get("month", "")))


@app.route("/api/generate_schedule", methods=["POST"])
def generate_schedule():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    shift = clean_text(req.get("shift", "")) or default_duty_shift()
    res = produce_duty_for_month(shift, year, month)
    return jsonify(res)


@app.route("/api/clear_schedule", methods=["POST"])
def clear_schedule():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    shift = clean_text(req.get("shift", "")) or default_duty_shift()
    if shift not in SHIFTS:
        shift = default_duty_shift()
    clear_weekly_clean(year, month, shift)
    return jsonify({"status": "success", "message": f"已清除「{shift}」本月每週清潔顯示結果（每日值日生仍保留在歷史中）"})


@app.route("/api/reset_counts", methods=["POST"])
def reset_counts():

    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    return jsonify({"status": "success", "message": "公平計數現在會依「值日生歷史」即時計算，不需要手動重置。若要真正歸零，請清空值日生歷史。"})


@app.route("/api/swap_leave", methods=["POST"])
def api_swap_leave():
    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    person_a = clean_text(req.get("person_a", ""))
    person_b = clean_text(req.get("person_b", ""))
    day_a = clean_text(req.get("day_a", ""))
    day_b = clean_text(req.get("day_b", ""))
    reassign = bool(req.get("reassign", True))
    reason = clean_text(req.get("reason", ""))
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status":"error","message":"部門或班別錯誤"})
    if not person_a or not person_b or not str(day_a).isdigit() or not str(day_b).isdigit():
        return jsonify({"status":"error","message":"請完整選擇兩位人員與兩個日期"})
    locked = is_plan_locked(year, month, department, shift)
    if locked and not reason:
        return jsonify({"status":"error","message":"班表已鎖定，換假／換班必須填寫原因"})
    selected, num_days = now_info(year, month)
    da, db = int(day_a), int(day_b)
    if da < 1 or da > num_days or db < 1 or db > num_days:
        return jsonify({"status":"error","message":"日期超出本月範圍"})
    df_basic, df_leave = init_attendance_sheets(selected.year, selected.month)
    idx_a = df_leave[df_leave["姓名"].astype(str).str.strip() == person_a].index
    idx_b = df_leave[df_leave["姓名"].astype(str).str.strip() == person_b].index
    if len(idx_a) == 0 or len(idx_b) == 0:
        return jsonify({"status":"error","message":"找不到人員，請確認兩位都已新增到假表"})
    old_a = clean_leave_value(df_leave.at[idx_a[0], str(da)])
    old_b = clean_leave_value(df_leave.at[idx_b[0], str(db)])
    df_leave.at[idx_a[0], str(da)] = old_b
    df_leave.at[idx_b[0], str(db)] = old_a
    ok, msg = save_attendance_frames(df_basic, df_leave, selected.year, selected.month)
    if not ok:
        return jsonify({"status":"error","message":msg})
    reassign_msgs = []
    if reassign:
        for d in sorted({da, db}):
            batch_id = uuid.uuid4().hex[:12]
            result = create_assignment_for_day(department, d, mode="balanced_random", save=False, batch_id=batch_id, year=selected.year, month=selected.month, shift=shift)
            if result.get("status") == "success":
                ok2, msg2 = replace_day_department_history(selected.year, selected.month, d, department, shift, result)
                reassign_msgs.append(f"{d} 號：" + ("已重新補位" if ok2 else msg2))
            else:
                reassign_msgs.append(f"{d} 號：{result.get('message','無法補位')}")
    content = f"{person_a} {selected.month}/{da}（原：{old_a or '上班'}）與 {person_b} {selected.month}/{db}（原：{old_b or '上班'}）互換"
    if reassign_msgs:
        content += "；" + "、".join(reassign_msgs)
    log_operation("換假／換班", department, shift, selected.year, selected.month, content, reason)
    return jsonify({"status":"success","message":"換假／換班完成。" + ("\n" + "\n".join(reassign_msgs) if reassign_msgs else ""), "reassign_messages": reassign_msgs})


@app.route("/api/swap_work", methods=["POST"])
def api_swap_work():
    admin_error = require_admin_json()
    if admin_error:
        return admin_error
    req = request.json or {}
    year = req.get("year") or datetime.now().year
    month = req.get("month") or datetime.now().month
    department = clean_text(req.get("department", ""))
    shift = clean_text(req.get("shift", "晚班")) or "晚班"
    day = clean_text(req.get("day", ""))
    person_a = clean_text(req.get("person_a", ""))
    person_b = clean_text(req.get("person_b", ""))
    reason = clean_text(req.get("reason", ""))
    if department not in DEPARTMENTS or shift not in SHIFTS:
        return jsonify({"status":"error","message":"部門或班別錯誤"})
    if not day or not str(day).isdigit() or not person_a or not person_b:
        return jsonify({"status":"error","message":"請選擇日期與兩位人員"})
    locked = is_plan_locked(year, month, department, shift)
    if locked and not reason:
        return jsonify({"status":"error","message":"班表已鎖定，換工作必須填寫原因"})
    selected, _ = now_info(year, month)
    date_str = f"{selected.year}-{selected.month:02d}-{int(day):02d}"
    df = load_history_df()
    if df.empty:
        return jsonify({"status":"error","message":"目前沒有工作分配歷史，無法換工作"})
    mask = ((df["日期"].astype(str) == date_str) & (df["部門"].astype(str) == department) & (df["班別"].astype(str) == shift))
    day_df = df[mask].copy()
    if day_df.empty:
        return jsonify({"status":"error","message":"這一天尚未產生工作分配"})
    day_df["產生時間_dt"] = pd.to_datetime(day_df["產生時間"], errors="coerce")
    day_df = day_df.sort_values(["產生時間_dt", "批次ID"], na_position="first")
    latest_batch = clean_text(day_df.iloc[-1].get("批次ID", ""))
    latest_mask = mask & (df["批次ID"].astype(str).str.strip() == latest_batch)
    idx_a = df[latest_mask & (df["人員"].astype(str).str.strip() == person_a)].index
    idx_b = df[latest_mask & (df["人員"].astype(str).str.strip() == person_b)].index
    if len(idx_a) == 0 or len(idx_b) == 0:
        return jsonify({"status":"error","message":"最新當日分配內找不到其中一位人員"})
    ia, ib = idx_a[0], idx_b[0]
    work_a, work_b = clean_text(df.at[ia, "工作"]), clean_text(df.at[ib, "工作"])
    df.at[ia, "工作"] = work_b
    df.at[ib, "工作"] = work_a
    df.at[ia, "分配方式"] = clean_text(df.at[ia, "分配方式"]) + "／換工作"
    df.at[ib, "分配方式"] = clean_text(df.at[ib, "分配方式"]) + "／換工作"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df.at[ia, "產生時間"] = now_str
    df.at[ib, "產生時間"] = now_str
    ok, msg = save_history_df(df)
    if not ok:
        return jsonify({"status":"error","message":msg})
    content = f"{date_str} {person_a}（{work_a}）與 {person_b}（{work_b}）互換工作"
    log_operation("換工作", department, shift, selected.year, selected.month, content, reason)
    return jsonify({"status":"success","message":"換工作完成：" + content})


@app.route("/api/operation_logs", methods=["GET"])
def api_operation_logs():
    start_date = clean_text(request.args.get("start_date", ""))
    end_date = clean_text(request.args.get("end_date", ""))
    action = clean_text(request.args.get("action", ""))
    department = clean_text(request.args.get("department", ""))
    shift = clean_text(request.args.get("shift", ""))
    df = load_operation_log_df()
    if df.empty:
        return jsonify({"status":"success","records":[]})
    filtered = df.copy()
    filtered["時間_dt"] = pd.to_datetime(filtered["時間"], errors="coerce")
    filtered = filtered[filtered["時間_dt"].notna()]
    if start_date:
        sd = pd.to_datetime(start_date, errors="coerce")
        if pd.notna(sd): filtered = filtered[filtered["時間_dt"] >= sd]
    if end_date:
        ed = pd.to_datetime(end_date, errors="coerce")
        if pd.notna(ed): filtered = filtered[filtered["時間_dt"] <= ed + pd.Timedelta(days=1)]
    if action:
        filtered = filtered[filtered["操作類型"].astype(str).str.contains(action, na=False)]
    if department and department in DEPARTMENTS:
        filtered = filtered[filtered["部門"].astype(str) == department]
    if shift and shift in SHIFTS:
        filtered = filtered[filtered["班別"].astype(str) == shift]
    filtered = filtered.sort_values("時間_dt", ascending=False).drop(columns=["時間_dt"], errors="ignore")
    return jsonify({"status":"success","records":filtered.head(300).to_dict("records")})


@app.route("/download_operation_log")
def download_operation_log():
    if not os.path.exists(OPERATION_LOG_FILE):
        save_operation_log_df(load_operation_log_df())
    return send_file(OPERATION_LOG_FILE, as_attachment=True, download_name="operation_log.xlsx")

# 啟動時初始化資料庫（建表 + 首次遷移皆為冪等，gunicorn 多 worker 下每個 worker 都安全執行）
init_db()

if __name__ == "__main__":
    make_auto_backup("startup")
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
