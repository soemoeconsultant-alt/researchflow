import os
import re
import time
import json
import html
import datetime
import hashlib
import shutil
import subprocess
from zoneinfo import ZoneInfo
from html.parser import HTMLParser
import pandas as pd
import streamlit as st
from pathlib import Path
from google import genai
from pptx import Presentation
from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import RGBColor as DocxRGBColor

st.set_page_config(
    page_title="ResearchFlow | AI Market Research",
    page_icon="🔮",
    layout="wide"
)

# App version — shown in the sidebar and written into backup manifests, so the
# running version can always be verified from the code itself (v3.15 had none).
APP_VERSION = "3.24.4"

# --- v3.18 language switch (EN / MM) ---
# T("English", "မြန်မာ") returns the text for the language chosen in the sidebar.
# If no Myanmar text is given it falls back to English. Page ids stay English.
if "lang" not in st.session_state:
    st.session_state["lang"] = "EN"


def T(en, mm=None):
    if mm is not None and st.session_state.get("lang") == "MM":
        return mm
    return en

APP_DIR = Path(__file__).resolve().parent
# Localhost layout: app.py lives in <BASE>/website/, data in <BASE>/.
# Online layout: app.py at the repo root, data in the repo root.
if (APP_DIR.parent / "audio").is_dir() or (APP_DIR.parent / "output").is_dir():
    BASE_DIR = APP_DIR.parent
else:
    BASE_DIR = APP_DIR
AUDIO_DIR = BASE_DIR / "audio"
OUTPUT_DIR = BASE_DIR / "output"
MERGED_DIR = OUTPUT_DIR / "merged"
ANALYSIS_DIR = BASE_DIR / "analysis"
REPORTS_DIR = BASE_DIR / "reports"
NOTES_DIR = BASE_DIR / "notes"

AUDIO_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
MERGED_DIR.mkdir(exist_ok=True)
ANALYSIS_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)
NOTES_DIR.mkdir(exist_ok=True)


# --- v3.24 UI preferences: theme + language remembered between sessions.
# Stored in .ui_prefs.json next to the app (a runtime file, not in the repo;
# no secrets, no personal data — just "light"/"dark" and "EN"/"MM").
# On Streamlit Cloud the file lives in the deployed app folder, so it resets
# on a redeploy/reboot; the in-session choice always works.
UI_PREFS_PATH = BASE_DIR / ".ui_prefs.json"


def _load_ui_prefs():
    try:
        data = json.loads(UI_PREFS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_ui_prefs(prefs):
    try:
        _atomic_write_bytes(UI_PREFS_PATH, json.dumps(prefs, ensure_ascii=False).encode("utf-8"))
    except Exception:
        pass


# Seed theme + language once per session from the saved prefs (a fresh
# browser session starts with empty session state, so this is the only
# moment the file can be read). Later reruns keep the in-session choice.
if "_ui_prefs_loaded" not in st.session_state:
    _prefs = _load_ui_prefs()
    if _prefs.get("theme") in ("light", "dark"):
        st.session_state["theme"] = _prefs["theme"]
    if _prefs.get("lang") in ("EN", "MM"):
        st.session_state["lang"] = _prefs["lang"]
    st.session_state["_ui_prefs_loaded"] = True
if "theme" not in st.session_state:
    st.session_state["theme"] = "light"


def get_transcript_files():
    files = [f for f in OUTPUT_DIR.glob("*.txt")
             if f.name != "failed_files.txt" and "_conflict_" not in f.name]
    return sorted(files, key=lambda f: f.name)


# --- Merge chunks into full transcripts ---
def chunk_group_key(name):
    stem = name[:-4] if name.endswith(".txt") else name
    stem = re.sub(r"\(\d+\)$", "", stem)
    return stem


def chunk_order(name):
    m = re.search(r"\((\d+)\)\.txt$", name)
    return int(m.group(1)) if m else 1


def get_chunk_groups():
    groups = {}
    for f in get_transcript_files():
        if f.name.endswith("_full.txt"):
            continue
        groups.setdefault(chunk_group_key(f.name), []).append(f)
    for key in groups:
        groups[key] = sorted(groups[key], key=lambda f: chunk_order(f.name))
    return groups


def get_full_transcripts():
    """Merged full files + single-chunk files.
    Returns (full_files, pending_groups)."""
    groups = get_chunk_groups()
    fulls = []
    pending = []
    for key, files in sorted(groups.items()):
        merged = MERGED_DIR / f"{key}_full.txt"
        if merged.exists():
            fulls.append(merged)
        elif len(files) == 1:
            fulls.append(files[0])
        else:
            pending.append((key, len(files)))
    return fulls, pending


def short_ids_for(names):
    ids = []
    for name in names:
        stem = name[:-4] if name.endswith(".txt") else name
        stem = stem.replace("_full", "").replace("interview_", "")
        ids.append(stem)
    return ids


# --- Storage: local folders, or Google Drive when credentials exist ---
# Drive is the source of truth when credentials are configured:
#   ResearchFlow / <Project> / {audio, transcripts, notes, reports}
# ("Default" keeps the legacy flat folders on the local PC.)
# Auth order (v3): OAuth user credentials first — the consultant account
# itself, so uploaded files belong to him and count against his storage.
# A service account owns the files it uploads and has no storage quota of
# its own, so its uploads are rejected by Google; it stays only as fallback.
# Local folders are always the working cache, so every page works the same
# with or without Drive. Without credentials the app runs on localhost
# exactly as before.
DRIVE_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
]
DRIVE_ROOT_NAME = "ResearchFlow"
# The ResearchFlow folder in the consultant's Drive (created 2026-11-14, shared
# with the service account as Editor). A service account cannot find a shared
# folder via "'root' in parents" — shared items are not in its own root — so
# use this folder id directly when it resolves; otherwise fall back to lookup.
DRIVE_ROOT_ID = "1eiun1ZiNj-9rf3LWfdQEWwJczWjPi3jJ"

# dir_key -> Drive folder name (analysis txt + plan txt live in "reports",
# merged full transcripts live in "transcripts" alongside chunk txt).
DIR_TO_DRIVE = {
    "audio": "audio",
    "transcripts": "transcripts",
    "merged": "transcripts",
    "analysis": "reports",
    "reports": "reports",
    "notes": "notes",
}
LOCAL_DIR = {
    "audio": AUDIO_DIR,
    "transcripts": OUTPUT_DIR,
    "merged": MERGED_DIR,
    "analysis": ANALYSIS_DIR,
    "reports": REPORTS_DIR,
    "notes": NOTES_DIR,
}


# --- Projects (v3.4 shell): one Project = one container. "Default" keeps the
# legacy flat folders; any other project keeps its own folders under
# projects/<name>/ locally and ResearchFlow/<name>/ on Drive.
def current_project():
    """Working project name from session state (default 'Default')."""
    return st.session_state.get("active_project", "Default")


def refresh_project_dirs(project):
    """Point the local working folders at the project (rebinds the module
    globals; every page reads them at runtime, so the switch applies
    immediately). Creates the folders if missing."""
    global AUDIO_DIR, OUTPUT_DIR, MERGED_DIR, ANALYSIS_DIR, REPORTS_DIR, NOTES_DIR, LOCAL_DIR
    base = BASE_DIR if project == "Default" else BASE_DIR / "projects" / project
    AUDIO_DIR = base / "audio"
    OUTPUT_DIR = base / "output"
    MERGED_DIR = OUTPUT_DIR / "merged"
    ANALYSIS_DIR = base / "analysis"
    REPORTS_DIR = base / "reports"
    NOTES_DIR = base / "notes"
    for d in (AUDIO_DIR, OUTPUT_DIR, MERGED_DIR, ANALYSIS_DIR, REPORTS_DIR, NOTES_DIR):
        d.mkdir(parents=True, exist_ok=True)
    LOCAL_DIR = {
        "audio": AUDIO_DIR,
        "transcripts": OUTPUT_DIR,
        "merged": MERGED_DIR,
        "analysis": ANALYSIS_DIR,
        "reports": REPORTS_DIR,
        "notes": NOTES_DIR,
    }


def _drive_root_id():
    """The ResearchFlow root folder id (pinned DRIVE_ROOT_ID, else lookup)."""
    if DRIVE_ROOT_ID:
        try:
            info = _drive_service().files().get(fileId=DRIVE_ROOT_ID, fields="id,mimeType").execute()
            if info.get("mimeType") == "application/vnd.google-apps.folder":
                return DRIVE_ROOT_ID
        except Exception:
            pass
    return _drive_ensure_folder(DRIVE_ROOT_NAME, "root")


def list_drive_projects():
    """Project folder names under the ResearchFlow root (Drive)."""
    if not drive_available():
        return []
    try:
        svc = _drive_service()
        root = _drive_root_id()
        names, token = [], None
        while True:
            res = svc.files().list(
                q=f"'{root}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
                fields="files(name),nextPageToken", pageSize=100, pageToken=token,
            ).execute()
            names += [f["name"] for f in res.get("files", [])]
            token = res.get("nextPageToken")
            if not token:
                break
        return sorted(names)
    except Exception:
        return []


def list_projects():
    """All known projects: 'Default' first, then Drive folders + local project dirs."""
    names = ["Default"]
    for n in list_drive_projects():
        if n not in names:
            names.append(n)
    local_root = BASE_DIR / "projects"
    if local_root.exists():
        for d in sorted(local_root.iterdir()):
            if d.is_dir() and d.name not in names:
                names.append(d.name)
    return names


def ensure_project(name):
    """Create a project's folders (Drive + local). Returns the cleaned name, or None."""
    clean = re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9 _\-\u1000-\u109F]+", "", (name or "").strip())).strip()
    if not clean or clean == "Default":
        return None
    if drive_available():
        try:
            root = _drive_root_id()
            proj = _drive_ensure_folder(clean, root)
            for sub in ("audio", "transcripts", "notes", "reports"):
                _drive_ensure_folder(sub, proj)
        except Exception:
            pass
    base = BASE_DIR / "projects" / clean
    for sub in ("audio", "output", "output/merged", "analysis", "reports", "notes"):
        (base / sub).mkdir(parents=True, exist_ok=True)
    return clean

# --- Client Brief (v3.21, Stage 1): one brief.json per project, in notes/
# (Drive notes folder + local). Fields follow the 5-stage table exactly:
# Client info, project info, objectives, target audience, scope, timeline,
# deliverables, approval (Draft → Approved). Discussion Guide lives here too
# (moved from the Market Analysis planner — no duplicate input there).
BRIEF_APPROVALS = ["Draft", "Pending", "Approved", "Revision"]
BRIEF_DELIVERABLES = ["PPTX Report", "DOCX Report", "PDF Report", "Presentation"]


def load_brief():
    """Saved brief for the current project, or {} when none exists."""
    try:
        data = json.loads((NOTES_DIR / "brief.json").read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_brief(data):
    """Save the brief (Drive + local, via the notes folder)."""
    return store_write_text("notes", "brief.json", json.dumps(data, ensure_ascii=False, indent=2))


_SA_INFO = None
_SA_TRIED = False
_OAUTH_INFO = None
_OAUTH_TRIED = False
_DRIVE_SERVICE = None
_DOCS_SERVICE = None
_FOLDER_IDS = {}


def _load_sa_info():
    """Service-account credentials, from (in order):
    st.secrets["gcp_service_account"] (dict), env GOOGLE_SERVICE_ACCOUNT_JSON
    (JSON string), or drive-service-account.json next to app.py."""
    try:
        sa = st.secrets.get("gcp_service_account")
        if sa:
            return dict(sa)
    except Exception:
        pass
    try:
        raw_secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        if isinstance(raw_secret, dict):
            return dict(raw_secret)
        if raw_secret:
            return json.loads(raw_secret)
    except Exception:
        pass
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw:
        try:
            return json.loads(raw)
        except Exception:
            pass
    p = APP_DIR / "drive-service-account.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return None


def _load_oauth_info():
    """OAuth user credentials (the consultant account itself), from (in order):
    st.secrets["google_oauth"] (dict with client_id, client_secret,
    refresh_token, optional token_uri) or env GOOGLE_OAUTH_JSON (JSON string
    with the same keys). The refresh token is obtained once on the user's PC
    (get_refresh_token.py) and pasted into Streamlit secrets — the app never
    asks the user to sign in. Returns the dict, or None."""
    try:
        o = st.secrets.get("google_oauth")
        # NOTE (v3.1 fix): st.secrets sections are AttrDict, NOT plain dict,
        # so isinstance(o, dict) is False and OAuth was silently skipped.
        # Duck-type instead: any mapping with a refresh_token counts.
        if o is not None and o.get("refresh_token"):
            return dict(o)
    except Exception:
        pass
    raw = os.environ.get("GOOGLE_OAUTH_JSON")
    if raw:
        try:
            o = json.loads(raw)
            if isinstance(o, dict) and o.get("refresh_token"):
                return dict(o)
        except Exception:
            pass
    return None


def _load_all_auth():
    """Load both credential sources once (OAuth first, service account second)."""
    global _SA_INFO, _SA_TRIED, _OAUTH_INFO, _OAUTH_TRIED
    if not _SA_TRIED:
        _SA_TRIED = True
        _SA_INFO = _load_sa_info()
    if not _OAUTH_TRIED:
        _OAUTH_TRIED = True
        _OAUTH_INFO = _load_oauth_info()


def drive_available():
    """True when Drive credentials exist AND the Google API packages are installed."""
    _load_all_auth()
    if not (_OAUTH_INFO or _SA_INFO):
        return False
    try:
        import googleapiclient.discovery  # noqa: F401
        from google.oauth2 import service_account  # noqa: F401
        from google.oauth2 import credentials  # noqa: F401
        return True
    except Exception:
        return False


def drive_auth_mode():
    """Which credentials Drive uses: "oauth", "service", or None."""
    _load_all_auth()
    if _OAUTH_INFO:
        return "oauth"
    if _SA_INFO:
        return "service"
    return None


def docs_available():
    return drive_available()


# --- Login (v3.3): Google sign-in + allowlist ---
# Config lives in Streamlit secrets (never in the repo):
#   [auth]
#   redirect_uri = "https://researchflowmaster.streamlit.app/oauth2callback"
#   cookie_secret = "<random string>"
#   client_id = "<Web OAuth client id>"
#   client_secret = "<Web OAuth client secret>"
#   server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
#   [access]
#   owners = ["soemoe.consultant@gmail.com"]
#   allowed = ["soemoekyawww@gmail.com"]
# Production fails closed if [auth] is missing. Localhost can be opened explicitly
# with RESEARCHFLOW_ALLOW_LOCAL_OPEN_ACCESS=1.
# NOTE: the consent screen is in Testing mode, so Google only lets test users
# sign in — add every allowed email as a test user (or switch to Production).

def _auth_configured():
    try:
        return bool(st.secrets.get("auth"))
    except Exception:
        return False


def _allowed_emails():
    """(owners, allowed) email lists from secrets [access]."""
    owners, allowed = [], []
    try:
        acc = st.secrets.get("access")
        if acc is not None:
            owners = [str(e).strip().lower() for e in (acc.get("owners") or [])]
            allowed = [str(e).strip().lower() for e in (acc.get("allowed") or [])]
    except Exception:
        pass
    return owners, allowed


def current_user_email():
    try:
        return (st.user.email or "").strip().lower()
    except Exception:
        return ""


def current_user_is_owner():
    owners, _ = _allowed_emails()
    return current_user_email() in owners


def _is_local_request():
    """Return True only for loopback hostnames; unknown hosts are not local."""
    try:
        host = str(st.context.headers.get("host", "")).split(":", 1)[0].strip().lower()
        return host in {"localhost", "127.0.0.1", "::1"}
    except Exception:
        return False


# --- Landing page (v3.21): signed-out visitors see ResearchFlow itself, not
# a bare login form. Presentation only — the gate, [auth] requirement,
# redirect_uri and allowlist below are unchanged. This style block is separate
# from CUSTOM_CSS because login_gate() runs before CUSTOM_CSS is injected.
LOGIN_CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=Noto+Sans+Myanmar:wght@400;500;600;700;800&display=swap');
.rf-land { font-family: 'IBM Plex Sans', 'Noto Sans Myanmar', 'Myanmar Text', -apple-system, 'Segoe UI', sans-serif; background: linear-gradient(115deg, #EAF3FF 0%, #DDF3F8 55%, #D9F5EF 100%); border: 1px solid #D5E5F7; border-radius: 18px; padding: 34px 30px 26px; text-align: center; box-shadow: 0 10px 28px rgba(16,38,74,.08); }
.rf-land h1 { color: #10264A; font-size: 34px; font-weight: 800; margin: 0 0 2px; line-height: 1.6; }
.rf-land h1 span.b { color: #2563EB; }
.rf-land h1 span.t { color: #0D9488; }
.rf-land .sub { color: #475569; font-size: 14px; font-weight: 600; margin-bottom: 12px; line-height: 1.6; }
.rf-land p { color: #475569; font-size: 15px; line-height: 1.7; margin: 0 auto 14px; max-width: 560px; }
.rf-land .stages { display: flex; flex-wrap: wrap; gap: 6px; justify-content: center; margin: 0 auto 6px; max-width: 560px; }
.rf-land .stg { background: #FFFFFF; border: 1px solid #D5E5F7; border-radius: 999px; padding: 5px 12px; font-size: 12.5px; font-weight: 600; color: #10264A; line-height: 1.6; }
.rf-land .stg b { color: #0E8F7E; }
.rf-land-note { text-align: center; color: #5E6B78; font-size: 13px; margin-top: 12px; line-height: 1.6; }
</style>
"""


def _landing_header():
    """Branded landing card shown before login (logo, name, what it does)."""
    st.markdown(LOGIN_CSS, unsafe_allow_html=True)
    _l, _c, _r = st.columns([1, 2.2, 1])
    with _c:
        st.markdown(
            '<div class="rf-land">'
            '<svg width="44" height="44" viewBox="0 0 32 32" fill="none">'
            '<rect x="2" y="19" width="6" height="10" rx="2" fill="#2563EB"/>'
            '<rect x="11" y="12" width="6" height="17" rx="2" fill="#0891B2"/>'
            '<rect x="20" y="5" width="6" height="24" rx="2" fill="#0D9488"/>'
            '</svg>'
            '<h1><span class="b">Research</span><span class="t">Flow</span></h1>'
            '<div class="sub">' + T("Market Research AI", "Market Research AI") + '</div><p>'
            + T("Turn interview and survey data into clear findings and client-ready "
                "reports — from client brief to final presentation, one project, one place.",
                "Interview နဲ့ survey data တွေကို ရှင်းလင်းတဲ့ finding နဲ့ client-ready report "
                "အဖြစ် ပြောင်းပေးတယ် — client brief ကနေ နောက်ဆုံး တင်ပြမှုအထိ၊ project "
                "တစ်ခုတည်း၊ တစ်နေရာတည်း။")
            + '</p><div class="stages">'
            + "".join(
                f'<span class="stg"><b>{i}</b> {n}</span>'
                for i, n in [
                    (1, T("Brief", "Brief")),
                    (2, T("Research Design", "သုတေသန ဒီဇိုင်း")),
                    (3, T("Fieldworks", "ကွင်းဆင်းလုပ်ငန်း")),
                    (4, T("Data Cleansing", "ဒေတာ သန့်စင်ခြင်း")),
                    (5, T("Analysis & Reports", "ခွဲခြမ်းစိတ်ဖြာမှုနဲ့ Report")),
                ]
            )
            + '</div></div>',
            unsafe_allow_html=True,
        )


def login_gate():
    """Require Google sign-in + allowlist. Open access is opt-in on localhost only.
    Auth behaviour is unchanged in v3.21 — only the signed-out view is a
    branded landing page instead of a bare login form."""
    if not _auth_configured():
        local_open = os.environ.get("RESEARCHFLOW_ALLOW_LOCAL_OPEN_ACCESS", "").strip() == "1"
        if local_open and _is_local_request():
            return ""
        _landing_header()
        st.error(T("[auth] configuration is missing. Online deployment is closed for security.", "[auth] configuration မရှိပါ။ Online deployment ကို လုံခြုံရေးအတွက် ပိတ်ထားပါတယ်။"))
        st.info(T("For local development, set RESEARCHFLOW_ALLOW_LOCAL_OPEN_ACCESS=1 on localhost. For online use, add [auth] and [access] to the Streamlit secrets.", "Local development အတွက် localhost မှာ RESEARCHFLOW_ALLOW_LOCAL_OPEN_ACCESS=1 ကို သတ်မှတ်ပါ။ Online အသုံးပြုရန် Streamlit secrets ထဲတွင် [auth] နှင့် [access] ကို ထည့်ပါ။"))
        st.stop()
    if not st.user.is_logged_in:
        _landing_header()
        _l, _c, _r = st.columns([1, 2.2, 1])
        with _c:
            st.login()
            st.markdown(
                '<div class="rf-land-note">'
                + T("Sign in with Google — only accounts approved by the owner can enter.",
                    "Google နဲ့ sign in ဝင်ပါ — owner က ခွင့်ပြုထားတဲ့ account ပဲ ဝင်ရတယ်။")
                + '</div>',
                unsafe_allow_html=True,
            )
        st.stop()
    email = current_user_email()
    owners, allowed = _allowed_emails()
    if email not in owners + allowed:
        _landing_header()
        st.error(T("This account cannot sign in — only accounts approved by the owner can.", "ဒီ account နဲ့ ဝင်လို့ မရဘူး — owner က ခွင့်ပြုထားတဲ့ account ပဲ ဝင်ရတယ်။"))
        st.write(T(f"Signed in as: {email}", f"ဝင်ထားတဲ့ account: {email}"))
        if st.button(T("Log out", "Log out ထွက်မယ်")):
            st.logout()
        st.stop()
    return email


def _sa_credentials():
    _load_all_auth()
    if not _SA_INFO:
        return None
    try:
        from google.oauth2 import service_account
        return service_account.Credentials.from_service_account_info(_SA_INFO, scopes=DRIVE_SCOPES)
    except Exception:
        return None


def _drive_credentials():
    """OAuth user credentials first (files belong to the user, whose storage
    quota applies); service account only as fallback. The google-auth
    library refreshes the OAuth access token automatically on first use."""
    _load_all_auth()
    if _OAUTH_INFO:
        try:
            from google.oauth2.credentials import Credentials
            return Credentials(
                token=None,
                refresh_token=_OAUTH_INFO.get("refresh_token"),
                client_id=_OAUTH_INFO.get("client_id"),
                client_secret=_OAUTH_INFO.get("client_secret"),
                token_uri=_OAUTH_INFO.get("token_uri") or "https://oauth2.googleapis.com/token",
                scopes=DRIVE_SCOPES,
            )
        except Exception:
            pass
    return _sa_credentials()


def _drive_service():
    global _DRIVE_SERVICE
    if _DRIVE_SERVICE is None:
        try:
            from googleapiclient.discovery import build
            _DRIVE_SERVICE = build("drive", "v3", credentials=_drive_credentials())
        except Exception:
            _DRIVE_SERVICE = False
    return _DRIVE_SERVICE or None


def _docs_service():
    global _DOCS_SERVICE
    if _DOCS_SERVICE is None:
        try:
            from googleapiclient.discovery import build
            _DOCS_SERVICE = build("docs", "v1", credentials=_drive_credentials())
        except Exception:
            _DOCS_SERVICE = False
    return _DOCS_SERVICE or None


def _drive_find_folder(name, parent_id):
    svc = _drive_service()
    q = (
        f"name = '{name}' and mimeType = 'application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents and trashed = false"
    )
    res = svc.files().list(q=q, fields="files(id,name)", pageSize=5).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def _drive_ensure_folder(name, parent_id):
    fid = _drive_find_folder(name, parent_id)
    if fid:
        return fid
    svc = _drive_service()
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    return svc.files().create(body=meta, fields="id").execute()["id"]


def _drive_folder_id(folder_name):
    """Folder id for ResearchFlow / <project> / <folder_name> (created if missing).
    Cache is keyed by (project, folder) so switching projects never returns
    another project's folder."""
    project = current_project()
    key = (project, folder_name)
    if key in _FOLDER_IDS:
        return _FOLDER_IDS[key]
    root = _drive_root_id()
    proj = _drive_ensure_folder(project, root)
    fid = _drive_ensure_folder(folder_name, proj)
    _FOLDER_IDS[key] = fid
    return fid


def drive_list_files(folder_name):
    """List non-trashed files directly inside a Drive folder."""
    svc = _drive_service()
    fid = _drive_folder_id(folder_name)
    items, token = [], None
    while True:
        res = svc.files().list(
            q=f"'{fid}' in parents and trashed = false",
            fields="files(id,name,mimeType,modifiedTime),nextPageToken",
            pageSize=100, pageToken=token,
        ).execute()
        items += res.get("files", [])
        token = res.get("nextPageToken")
        if not token:
            break
    return items


def _drive_find_file(folder_name, name):
    for f in drive_list_files(folder_name):
        if f["name"] == name and f.get("mimeType") != "application/vnd.google-apps.folder":
            return f
    return None


def drive_download(file_id, dest_path):
    from googleapiclient.http import MediaIoBaseDownload
    svc = _drive_service()
    req = svc.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def drive_upload(local_path, folder_name, name=None):
    """Upload a local file to a Drive folder (updates it if the name exists). Returns file id."""
    from googleapiclient.http import MediaFileUpload
    svc = _drive_service()
    name = name or Path(local_path).name
    media = MediaFileUpload(str(local_path), resumable=True)
    existing = _drive_find_file(folder_name, name)
    if existing:
        return svc.files().update(fileId=existing["id"], media_body=media, fields="id").execute()["id"]
    meta = {"name": name, "parents": [_drive_folder_id(folder_name)]}
    return svc.files().create(body=meta, media_body=media, fields="id").execute()["id"]


def drive_move_to_originals(name):
    """Move an audio file in Drive to audio/originals (kept safe, out of the Batch queue)."""
    if not drive_available():
        return
    try:
        svc = _drive_service()
        audio_id = _drive_folder_id("audio")
        orig_id = _drive_ensure_folder("originals", audio_id)
        f = _drive_find_file("audio", name)
        if f:
            svc.files().update(fileId=f["id"], addParents=orig_id, removeParents=audio_id, fields="id").execute()
    except Exception:
        pass


def _drive_name_belongs(dir_key, name):
    """Which local dir a file from a Drive folder belongs to."""
    if dir_key == "merged":
        return name.endswith("_full.txt")
    if dir_key == "transcripts":
        return name.endswith(".txt") and not name.endswith("_full.txt")
    if dir_key == "analysis":
        return name.endswith(".txt")
    if dir_key == "reports":
        return name.endswith((".pptx", ".docx"))
    return True


# --- v3.16 data safety: atomic save, versioned backup, retry queue, backup/restore ---
SYNC_QUEUE_PATH = BASE_DIR / "sync_queue.json"
VERSIONS_KEEP = 5
BACKUP_DIRS = ("audio", "transcripts", "merged", "analysis", "reports", "notes")


def _file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _atomic_write_bytes(path, data):
    """Write via a temp file in the same folder + os.replace, so a crash or a
    closed app mid-write can never leave a half-written file behind."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    with open(tmp, "wb") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, path)
    return path


def _atomic_write_text(path, text):
    return _atomic_write_bytes(path, text.encode("utf-8"))


def _backup_version(path):
    """Before overwriting a text file, keep a timestamped copy in a `.versions`
    subfolder (newest 5 kept). Audio/binary files are NOT versioned (too big)
    — atomic write only. Returns the backup path or None."""
    path = Path(path)
    if not path.exists() or path.suffix.lower() not in (".txt", ".md"):
        return None
    vdir = path.parent / ".versions"
    vdir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    bak = vdir / f"{path.stem}.{stamp}.bak"
    n = 1
    while bak.exists():  # two saves in the same microsecond must not collide
        n += 1
        bak = vdir / f"{path.stem}.{stamp}-{n}.bak"
    try:
        shutil.copy2(path, bak)
    except Exception:
        return None
    olds = sorted(vdir.glob(f"{path.stem}.*.bak"), key=lambda f: f.name)
    for old in olds[:-VERSIONS_KEEP]:
        try:
            old.unlink()
        except Exception:
            pass
    return bak


def _load_sync_queue():
    try:
        items = json.loads(SYNC_QUEUE_PATH.read_text(encoding="utf-8"))
        return items if isinstance(items, list) else []
    except Exception:
        return []


def _save_sync_queue(items):
    try:
        _atomic_write_text(SYNC_QUEUE_PATH, json.dumps(items, ensure_ascii=False, indent=1))
    except Exception:
        pass


def queue_upload(dir_key, name, error=""):
    """Remember a failed Drive upload so it is retried automatically.
    De-duplicated by (project, dir_key, name). Returns True if newly queued."""
    items = _load_sync_queue()
    project = st.session_state.get("active_project", "Default")
    path = str(LOCAL_DIR[dir_key] / name)
    if any(i.get("project") == project and i.get("dir_key") == dir_key
           and i.get("name") == name for i in items):
        return False
    items.append({
        "project": project, "dir_key": dir_key, "name": name, "path": path,
        "queued_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "attempts": 0, "error": str(error)[:200],
    })
    _save_sync_queue(items)
    return True


def retry_sync_queue():
    """Re-upload every queued file for the current project.
    Returns (uploaded, still_pending). Entries whose local file is gone are dropped."""
    if not drive_available():
        return 0, len(_load_sync_queue())
    items = _load_sync_queue()
    project = st.session_state.get("active_project", "Default")
    remaining, done = [], 0
    for it in items:
        if it.get("project", "Default") != project:
            remaining.append(it)
            continue
        path = Path(it.get("path") or (LOCAL_DIR[it["dir_key"]] / it["name"]))
        try:
            if path.exists():
                drive_upload(path, DIR_TO_DRIVE[it["dir_key"]], it["name"])
                done += 1
            # local file gone: drop the entry silently
        except Exception as e:
            it["error"] = str(e)[:200]
            it["attempts"] = it.get("attempts", 0) + 1
            remaining.append(it)
    _save_sync_queue(remaining)
    return done, len(remaining)


def build_backup_zip():
    """Zip the current project's local folders + a manifest (version, project,
    time, file list with sha256). Skips .versions history and temp files.
    Returns (filename, bytes)."""
    import io
    import zipfile
    project = current_project()
    buf = io.BytesIO()
    manifest_files = []
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for dir_key in BACKUP_DIRS:
            d = LOCAL_DIR[dir_key]
            if not d.exists():
                continue
            for p in sorted(d.rglob("*")):
                if not p.is_file():
                    continue
                if ".versions" in p.parts or p.name.startswith("."):
                    continue
                arc = f"{dir_key}/{p.relative_to(d).as_posix()}"
                zf.write(p, arc)
                manifest_files.append({"path": arc, "sha256": _file_sha256(p),
                                       "size": p.stat().st_size})
    manifest = {
        "app": "ResearchFlow", "app_version": APP_VERSION, "project": project,
        "created_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "files": manifest_files,
    }
    with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=1))
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"researchflow-backup-{project}-{stamp}.zip", buf.getvalue()


# --- v3.17 restore safety limits (zip-bomb + traversal protection) ---
RESTORE_MAX_COMPRESSED = 1024 * 1024 * 1024      # 1 GB uploaded ZIP
RESTORE_MAX_TOTAL = 4 * 1024 * 1024 * 1024        # 4 GB uncompressed total
RESTORE_MAX_FILE = 1024 * 1024 * 1024             # 1 GB single file
RESTORE_MAX_ENTRIES = 10000


def read_backup_manifest(data):
    """Read just the manifest from a backup ZIP. Returns (manifest, error)."""
    import io
    import zipfile
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except Exception as e:
        return None, T(f"Cannot open ZIP: {e}", f"ZIP ဖွင့်လို့ မရဘူး: {e}")
    try:
        return json.loads(zf.read("manifest.json").decode("utf-8")), None
    except Exception:
        return None, T("manifest.json not found — this may not be a ResearchFlow backup.", "manifest.json မတွေ့ဘူး — ResearchFlow backup မဟုတ်နိုင်ဘူး။")


def _restore_report(rep):
    if rep.get("restored"):
        st.success(T(f"Restored {len(rep['restored'])} file(s) ✓", f"Restored {len(rep['restored'])} ဖိုင် ✓"))
    if rep.get("skipped"):
        st.warning(T("Skipped: ", "ကျော်သွားတာ: ") + "; ".join(rep["skipped"][:5]))
    if rep.get("errors"):
        st.error("; ".join(rep["errors"]))


def restore_backup_zip(data, force=False):
    """Restore a validated ResearchFlow backup into the active project."""
    import io
    import zipfile
    report = {"restored": [], "skipped": [], "errors": []}
    if len(data) > RESTORE_MAX_COMPRESSED:
        report["errors"].append(T("Backup ZIP is too large (over 1 GB).", "Backup ZIP ကြီးလွန်းတယ် (1 GB ထက် ကြီးတယ်)။"))
        return report
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
        infos = zf.infolist()
    except Exception as e:
        report["errors"].append(T(f"Cannot open ZIP: {e}", f"ZIP ဖွင့်လို့ မရဘူး: {e}"))
        return report
    if len(infos) > RESTORE_MAX_ENTRIES:
        report["errors"].append(T("Too many files in the backup.", "Backup ထဲမှာ ဖိုင်များလွန်းတယ်။"))
        return report
    if sum(i.file_size for i in infos) > RESTORE_MAX_TOTAL:
        report["errors"].append(T("Backup would expand to over 4 GB — not accepted.", "Backup ဖြည်ရင် 4 GB ထက် ကြီးမယ် — လက်မခံဘူး။"))
        return report
    names = [i.filename for i in infos]
    if len(names) != len(set(names)):
        report["errors"].append(T("Duplicate file name in the ZIP.", "ZIP ထဲတွင် ဖိုင်အမည် ထပ်နေသည်။"))
        return report
    manifest_info = next((i for i in infos if i.filename == "manifest.json"), None)
    if not manifest_info or manifest_info.file_size > 5 * 1024 * 1024:
        report["errors"].append(T("manifest.json is missing or has an invalid size.", "manifest.json မတွေ့ဘူး သို့မဟုတ် အရွယ်အစား မမှန်ပါ။"))
        return report
    try:
        manifest = json.loads(zf.read(manifest_info).decode("utf-8"))
    except Exception:
        report["errors"].append(T("Cannot read manifest.json.", "manifest.json ကို ဖတ်မရပါ။"))
        return report
    if not isinstance(manifest, dict) or manifest.get("app") != "ResearchFlow":
        report["errors"].append(T("Not a ResearchFlow backup manifest.", "ResearchFlow backup manifest မဟုတ်ပါ။"))
        return report
    if not isinstance(manifest.get("project"), str) or not isinstance(manifest.get("files"), list):
        report["errors"].append(T("Backup manifest structure is invalid.", "Backup manifest structure မမှန်ပါ။"))
        return report
    backup_project = manifest["project"]
    if backup_project != current_project() and not force:
        report["project_mismatch"] = (backup_project, current_project())
        return report
    manifest_files = {}
    for entry in manifest["files"]:
        if not isinstance(entry, dict):
            report["errors"].append(T("Invalid manifest file entry.", "Manifest file entry မမှန်ပါ။"))
            return report
        path, digest, size = entry.get("path"), entry.get("sha256"), entry.get("size")
        if (not isinstance(path, str) or not isinstance(digest, str)
                or not re.fullmatch(r"[0-9a-fA-F]{64}", digest)
                or not isinstance(size, int) or size < 0 or size > RESTORE_MAX_FILE
                or path in manifest_files):
            report["errors"].append(T("Invalid file path/hash/size in the manifest.", "Manifest ထဲက file path/hash/size မမှန်ပါ။"))
            return report
        manifest_files[path] = entry
    zip_files = {n for n in names if n != "manifest.json" and not n.endswith("/")}
    if zip_files != set(manifest_files):
        report["errors"].append(T("ZIP file list does not match the manifest file list.", "ZIP ဖိုင်စာရင်းနှင့် manifest ဖိုင်စာရင်း မကိုက်ညီပါ။"))
        return report
    for info in infos:
        name = info.filename
        if name == "manifest.json" or name.endswith("/"):
            continue
        entry = manifest_files[name]
        # ZIP paths must be canonical POSIX relative paths; reject Windows and POSIX traversal.
        if "\\" in name or name.startswith("/") or re.match(r"^[A-Za-z]:", name):
            report["skipped"].append(f"{name} (unsafe path)")
            continue
        parts = name.split("/")
        if len(parts) < 2 or parts[0] not in BACKUP_DIRS or any(x in ("", ".", "..") for x in parts):
            report["skipped"].append(f"{name} (unsafe or unknown folder)")
            continue
        dir_key, rel = parts[0], "/".join(parts[1:])
        if info.file_size > RESTORE_MAX_FILE or info.file_size != entry["size"]:
            report["skipped"].append(T(f"{name} (file size mismatch)", f"{name} (file size မကိုက်)"))
            continue
        dest = LOCAL_DIR[dir_key] / rel
        try:
            if not dest.resolve().is_relative_to(LOCAL_DIR[dir_key].resolve()):
                report["skipped"].append(f"{name} (unsafe path)")
                continue
            raw = zf.read(info)
        except Exception:
            report["skipped"].append(T(f"{name} (cannot read)", f"{name} (ဖတ်မရပါ)"))
            continue
        if hashlib.sha256(raw).hexdigest().lower() != entry["sha256"].lower():
            report["skipped"].append(T(f"{name} (checksum mismatch)", f"{name} (checksum မကိုက်)"))
            continue
        dest.parent.mkdir(parents=True, exist_ok=True)
        _backup_version(dest)
        _atomic_write_bytes(dest, raw)
        report["restored"].append(name)
        if drive_available():
            try:
                drive_upload(dest, DIR_TO_DRIVE[dir_key], dest.name)
            except Exception as e:
                queue_upload(dir_key, dest.name, e)
    return report


def sync_from_drive():
    """Drive -> local sync that NEVER silently overwrites a newer local file.
    - Same content: skip.
    - Local newer: keep local, queue it for upload, save Drive's copy aside as
      <stem>_conflict_<timestamp><suffix> so nothing is lost.
    - Drive newer: back up the local file first, then replace it.
    - Audio: download only when missing or Drive is newer (big files)."""
    if not drive_available():
        return
    import datetime
    for dir_key, drive_folder in DIR_TO_DRIVE.items():
        local_dir = LOCAL_DIR[dir_key]
        try:
            items = drive_list_files(drive_folder)
        except Exception:
            continue
        for f in items:
            name = f["name"]
            if f.get("mimeType") == "application/vnd.google-apps.folder":
                continue
            if not _drive_name_belongs(dir_key, name):
                continue
            dest = local_dir / name
            try:
                drv_ts = 0
                dt = f.get("modifiedTime", "")
                if dt:
                    drv_ts = datetime.datetime.fromisoformat(dt.replace("Z", "+00:00")).timestamp()
                is_audio = name.endswith((".mp3", ".wav", ".m4a"))
                if dest.exists() and is_audio and drv_ts and drv_ts <= dest.stat().st_mtime:
                    continue  # local audio is newer — keep it, no download
                tmp = dest.with_name(f".{dest.name}.{os.getpid()}.dl")
                try:
                    drive_download(f["id"], tmp)
                except Exception:
                    try:
                        tmp.unlink()
                    except Exception:
                        pass
                    continue
                if dest.exists():
                    if _file_sha256(tmp) == _file_sha256(dest):
                        tmp.unlink()
                        continue
                    if is_audio:
                        os.replace(tmp, dest)  # Drive audio newer; no versioning (big)
                        continue
                    if dest.stat().st_mtime >= drv_ts:
                        # Local is newer: keep local, queue upload, keep Drive copy aside.
                        # If an identical conflict copy already exists, don't pile up another.
                        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
                        conflict = dest.with_name(f"{dest.stem}_conflict_{stamp}{dest.suffix}")
                        _existing = [c for c in dest.parent.glob(f"{dest.stem}_conflict_*{dest.suffix}")]
                        if any(_file_sha256(c) == _file_sha256(tmp) for c in _existing):
                            tmp.unlink()
                        else:
                            os.replace(tmp, conflict)
                        queue_upload(dir_key, name, "local newer than Drive — kept local")
                    else:
                        # Drive is newer: back up local first, then replace.
                        _backup_version(dest)
                        os.replace(tmp, dest)
                else:
                    os.replace(tmp, dest)
            except Exception:
                continue


def _safe_storage_path(dir_key, name):
    """Resolve a single safe filename inside an approved project storage folder."""
    if dir_key not in LOCAL_DIR:
        raise ValueError("Unknown storage folder")
    if not isinstance(name, str) or not name or name in {".", ".."}:
        raise ValueError("Invalid filename")
    if Path(name).name != name or "/" in name or "\\" in name or "\x00" in name:
        raise ValueError("Filename must not contain path components")
    if any(ord(ch) < 32 for ch in name):
        raise ValueError("Filename contains control characters")
    base = LOCAL_DIR[dir_key].resolve()
    path = (base / name).resolve()
    if not path.is_relative_to(base):
        raise ValueError("Storage path escapes project folder")
    return path


def store_write_text(dir_key, name, text):
    """Write a text file locally (atomic + versioned backup), then upload it to
    Drive when Drive mode is on. A failed upload goes into the retry queue.
    Unchanged content is skipped entirely (stops rerun duplicate uploads)."""
    path = _safe_storage_path(dir_key, name)
    data = text.encode("utf-8")
    if path.exists() and _file_sha256(path) == hashlib.sha256(data).hexdigest():
        return path  # same content: nothing to write or upload
    _backup_version(path)
    _atomic_write_bytes(path, data)
    if drive_available():
        try:
            drive_upload(path, DIR_TO_DRIVE[dir_key], name)
        except Exception as e:
            queue_upload(dir_key, name, e)
            st.warning(f"Drive upload မအောင်မြင်ဘူး: {name} — {str(e)[:200]} (local မှာ သိမ်းပြီးပြီ။ Queue ထဲထည့်ထားတယ်၊ internet ပြန်ရရင် အလိုအလျောက် ပြန်တင်မယ်။)")
    return path


def store_write_bytes(dir_key, name, data):
    """Write a binary file locally (atomic), then upload it to Drive when Drive
    mode is on. A failed upload goes into the retry queue.
    Unchanged content is skipped entirely (stops rerun duplicate uploads)."""
    path = _safe_storage_path(dir_key, name)
    if path.exists() and _file_sha256(path) == hashlib.sha256(data).hexdigest():
        return path  # same content: nothing to write or upload
    _backup_version(path)
    _atomic_write_bytes(path, data)
    if drive_available():
        try:
            drive_upload(path, DIR_TO_DRIVE[dir_key], name)
        except Exception as e:
            queue_upload(dir_key, name, e)
            st.warning(f"Drive upload မအောင်မြင်ဘူး: {name} — {str(e)[:200]} (local မှာ သိမ်းပြီးပြီ။ Queue ထဲထည့်ထားတယ်၊ internet ပြန်ရရင် အလိုအလျောက် ပြန်တင်မယ်။)")
    return path


def store_upload_existing(dir_key, name):
    """Upload a local file created by another tool (ffmpeg split, PPTX/DOCX export).
    A failed upload goes into the retry queue."""
    if not drive_available():
        return
    path = _safe_storage_path(dir_key, name)
    if path.exists():
        try:
            drive_upload(path, DIR_TO_DRIVE[dir_key], name)
        except Exception as e:
            queue_upload(dir_key, name, e)
            st.warning(f"Drive upload မအောင်မြင်ဘူး: {name} — {str(e)[:200]} (local မှာ သိမ်းပြီးပြီ။ Queue ထဲထည့်ထားတယ်၊ internet ပြန်ရရင် အလိုအလျောက် ပြန်တင်မယ်။)")

def next_version_number(stem):
    existing = sorted(ANALYSIS_DIR.glob(f"{stem}_v*_analysis.txt"), key=lambda f: f.name)
    nums = []
    for f in existing:
        m = re.search(r"_v(\d+)_analysis\.txt$", f.name)
        if m:
            nums.append(int(m.group(1)))
    return (max(nums) + 1) if nums else 1


# --- Google Docs (Report Editor) ---
def _doc_id_from_input(value):
    value = (value or "").strip()
    if not value:
        return None
    m = re.search(r"/d/([A-Za-z0-9_-]+)", value)
    return m.group(1) if m else value


def drive_create_doc(title, text):
    """Create a Google Doc in the Drive reports folder from analysis text.
    '## ' lines become headings. Returns (doc_id, url)."""
    svc = _drive_service()
    docs = _docs_service()
    meta = {
        "name": title,
        "mimeType": "application/vnd.google-apps.document",
        "parents": [_drive_folder_id("reports")],
    }
    doc_id = svc.files().create(body=meta, fields="id").execute()["id"]
    lines = text.splitlines() or [""]
    plain_parts = []
    headings = []
    idx = 1
    for line in lines:
        if line.startswith("## "):
            clean = line[3:].strip()
            start = idx
            plain_parts.append(clean)
            idx += len(clean)
            if clean:
                headings.append((start, idx))
        else:
            plain_parts.append(line)
            idx += len(line)
        plain_parts.append("\n")
        idx += 1
    plain = "".join(plain_parts) or " "
    requests = [{"insertText": {"location": {"index": 1}, "text": plain}}]
    for start, end in headings:
        requests.append({
            "updateParagraphStyle": {
                "range": {"startIndex": start, "endIndex": end},
                "paragraphStyle": {"namedStyleType": "HEADING_1"},
                "fields": "namedStyleType",
            }
        })
    docs.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
    return doc_id, f"https://docs.google.com/document/d/{doc_id}/edit"


def docs_read_text(doc_id):
    """Read a Google Doc's text. Heading paragraphs come back as '## ' lines,
    so the analysis structure survives the round trip."""
    docs = _docs_service()
    doc = docs.documents().get(documentId=doc_id).execute()
    out = []
    for el in doc.get("body", {}).get("content", []):
        para = el.get("paragraph")
        if not para:
            continue
        style = para.get("paragraphStyle", {}).get("namedStyleType", "")
        text = "".join(e.get("textRun", {}).get("content", "") for e in para.get("elements", []))
        text = text.rstrip("\n")
        if not text.strip():
            continue
        if style in ("HEADING_1", "HEADING_2", "TITLE"):
            out.append("## " + text.strip())
        else:
            out.append(text)
    return "\n".join(out).strip() + "\n"


DRIVE_MODE = drive_available()
DRIVE_AUTH = drive_auth_mode()  # "oauth" | "service" | None


# --- Gemini setup (Phase 2) ---
MODELS = ["gemini-3.5-flash", "gemini-3.8-flash"]


def get_gemini_client():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        try:
            api_key = st.secrets.get("GEMINI_API_KEY")
        except Exception:
            api_key = None
    if not api_key:
        return None
    return genai.Client(api_key=api_key)


def ask_gemini(prompt):
    """Try each model in order; return (text, model_used)."""
    client = get_gemini_client()
    if client is None:
        return None, "GEMINI_API_KEY not set"
    last_error = None
    for model in MODELS:
        try:
            response = client.models.generate_content(model=model, contents=prompt)
            return response.text, model
        except Exception as e:
            last_error = e
    return None, f"All models failed: {last_error}"


def check_gemini_models():
    """List the models this API key can use. Listing is not a generation
    request, so it does not use the daily request quota.
    Returns (sorted model names, error)."""
    client = get_gemini_client()
    if client is None:
        return None, "GEMINI_API_KEY not set"
    try:
        names = []
        for m in client.models.list():
            actions = getattr(m, "supported_actions", None) or []
            if actions and "generateContent" not in actions:
                continue
            n = str(getattr(m, "name", ""))
            names.append(n[len("models/"):] if n.startswith("models/") else n)
        return sorted(set(names)), None
    except Exception as e:
        return None, str(e)[:300]


def run_system_check():
    """Setup checks only (no Gemini generation). Returns a list of
    (ok, label, detail) where ok is True / False / None (= warning)."""
    import importlib.util
    rows = []
    key_ok = get_gemini_client() is not None
    rows.append((key_ok, "Gemini API key", "set" if key_ok else "GEMINI_API_KEY not found (env or secrets)"))
    names, err = check_gemini_models()
    if err:
        rows.append((False, "Gemini model list", err))
    else:
        for m in MODELS:
            rows.append((m in names, f"Model {m}",
                         "available" if m in names else "NOT in this key's model list — change MODELS in the app"))
        flash = [n for n in names if "flash" in n][:15]
        rows.append((None, "Flash models this key can use", ", ".join(flash) if flash else "none found"))
    ff = shutil.which("ffmpeg")
    rows.append((bool(ff), "ffmpeg", "found" if ff else "missing — mp3/m4a splitting will not work (WAV still works)"))
    fp = shutil.which("ffprobe")
    rows.append((bool(fp), "ffprobe", "found" if fp else "missing — audio duration cannot be read"))
    ox = importlib.util.find_spec("openpyxl") is not None
    rows.append((ox, "openpyxl", "installed" if ox else "missing — Quantitative Analysis cannot read XLSX"))
    rows.append((True if DRIVE_MODE else None, "Storage",
                 f"Google Drive ({DRIVE_AUTH})" if DRIVE_MODE else "Local folders only"))
    rows.append((_auth_configured(), "Login ([auth] secrets)", "configured" if _auth_configured() else "not configured"))
    return rows


def is_rate_limit_error(e):
    error_text = str(e).lower()
    status = getattr(getattr(e, "response", None), "status_code", None)
    if status is None:
        status = getattr(e, "code", None)
    if status is None:
        status = getattr(e, "status_code", None)
    return (
        status == 429
        or "429" in error_text
        or "resource_exhausted" in error_text
        or "too many requests" in error_text
    )


def transcribe_audio(audio_path):
    """Upload audio to Gemini and transcribe with generate_content
    (same call as the working analysis). Returns (text, model_used)."""
    client = get_gemini_client()
    if client is None:
        return None, "GEMINI_API_KEY not set"
    try:
        audio_file = client.files.upload(file=str(audio_path))
    except Exception as e:
        return None, f"Upload failed: {e}"
    prompt = PROMPT_LIBRARY["transcription"]["template"]
    last_error = None
    for model in MODELS:
        try:
            result = client.models.generate_content(
                model=model,
                contents=[prompt, audio_file],
            )
            return result.text, model
        except Exception as e:
            last_error = e
            if is_rate_limit_error(e):
                return None, "429 — quota ပြည့်နေတယ်, နောက်မှ ပြန်လုပ်"
    return None, f"All models failed: {last_error}"


def audio_duration(path):
    """Duration in seconds via ffprobe, or None if unavailable."""
    ffprobe = shutil.which("ffprobe")
    if not ffprobe:
        return None
    try:
        r = subprocess.run(
            [ffprobe, "-v", "error", "-show_entries", "format=duration",
             "-of", "csv=p=0", str(path)],
            capture_output=True, text=True, timeout=60,
        )
        return float(r.stdout.strip())
    except Exception:
        return None


def split_audio(src_path, chunk_min):
    """Split a long audio file into chunk_min-minute pieces named base(N).ext.
    Returns (chunk_paths, error). Uses ffmpeg; WAV fallback via the wave module."""
    import contextlib
    import wave
    base = src_path.stem
    ext = src_path.suffix.lower().lstrip(".")
    chunk_sec = chunk_min * 60
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        out_pattern = str(AUDIO_DIR / f"{base}(%d).{ext}")
        r = subprocess.run(
            [ffmpeg, "-y", "-i", str(src_path), "-f", "segment",
             "-segment_time", str(chunk_sec), "-c", "copy",
             "-segment_start_number", "1", out_pattern],
            capture_output=True, text=True, timeout=600,
        )
        if r.returncode != 0:
            return None, r.stderr[-500:] or "ffmpeg failed"
        chunks = sorted(
            AUDIO_DIR.glob(f"{base}(*).{ext}"),
            key=lambda f: chunk_order(f.name),
        )
        return chunks, None
    if ext == "wav":
        with contextlib.closing(wave.open(str(src_path), "rb")) as w:
            params = w.getparams()
            framerate = w.getframerate()
            n_frames = w.getnframes()
            chunk_frames = chunk_sec * framerate
            chunks = []
            idx = 1
            while (idx - 1) * chunk_frames < n_frames:
                w.setpos((idx - 1) * chunk_frames)
                frames = w.readframes(min(chunk_frames, n_frames - (idx - 1) * chunk_frames))
                out = AUDIO_DIR / f"{base}({idx}).wav"
                with contextlib.closing(wave.open(str(out), "wb")) as o:
                    o.setparams(params)
                    o.writeframes(frames)
                chunks.append(out)
                idx += 1
        return chunks, None
    return None, "ffmpeg မရှိသေးဘူး — mp3/m4a ခွဲဖို့ ffmpeg လိုတယ်"


# --- ResearchFlow AI Prompt Library (v3.11) ---
# One source of truth for every prompt Gemini is given. Pages call the
# library, never inline prompt text. The 7 engines:
#   1 Transcription, 2 Research Planning, 3 Qualitative Analysis,
#   4 Quantitative Analysis, 5 Cross-Interview Intelligence,
#   6 Report Generation, 7 Executive Summary.
# Fix 1 (evidence): every finding needs n= count + a verbatim quote with
# speaker + source file; no evidence -> "No evidence in transcripts".
# Fix 2 (plan structure): plan_based_prompt() uses the plan's sections exactly.
# Fix 3 (report format): all analysis output uses "## " headings + "- " bullets
# so parse_sections() can turn it into PPTX/DOCX automatically.

EVIDENCE_RULES = (
    "Evidence rules (apply to every finding):\n"
    "- State how many respondents said it: n=X of Y (single interview: n=1 of 1).\n"
    "- Support every finding with at least one verbatim quote, written exactly as:\n"
    '  > "..." — Speaker N, <source file>\n'
    "- A finding with no evidence in the transcripts must NOT be written. "
    "If a planned section has no evidence, write exactly this under it:\n"
    "  No evidence in transcripts\n"
    "- Mark anything you could not hear or read clearly as [unclear]. "
    "Never guess a name, number, or brand.\n"
)

OUTPUT_FORMAT = (
    "Output format (so the report export can read it automatically):\n"
    "- Section headings exactly as: ## <Section name> — on their own line, no other heading levels.\n"
    "- Findings as short paragraphs or \"- \" bullets (one finding per bullet).\n"
    "- Counts exactly as n=X/Y. Quotes exactly as: > \"...\" — Speaker N, <source file>\n"
    "- No tables, no bold headings, no extra sections.\n"
)

CROSS_INTERVIEW_RULES = (
    "Cross-respondent rules (more than one interview together):\n"
    "- For each theme, state how many respondents mentioned it: n=X of Y.\n"
    "- Give one supporting verbatim quote per theme: > \"...\" — Speaker N, <source file>\n"
    "- Note divergent or contradictory views where they exist — do not hide disagreement.\n"
    "- A pattern needs at least 2 respondents; one respondent's view is a single view, not a pattern.\n"
    "- If interviews disagree on a fact (price, brand, place), say so and name the sources.\n"
)

TRANSCRIPTION_TEMPLATE = (
    "Transcribe this audio in Burmese. Separate speakers clearly: "
    "Speaker 1 = Respondent, Speaker 2 = Moderator. "
    "Do not summarize. Preserve the original spoken meaning. "
    "Medicine, brand, and disease names in English. "
    "Mark unclear parts as [unclear]."
)

PLANNING_TEMPLATE = (
    "You are a senior market research analyst. Propose an analysis plan for this project. "
    "Read the client brief, discussion guide, questionnaire, and interview transcripts together. "
    "Do NOT write findings yet — only the plan.\n"
    "Rules: Base everything ONLY on the brief, guide, questionnaire, and transcripts. Do not invent objectives or questions. "
    "Use the questionnaire to know what was asked and how the data was collected, and match analysis sections to both the guide topics and the survey questions. "
    "Write in English.\n\n"
    "Structure your answer exactly like this:\n"
    "## Project Type\n"
    "(e.g. IDI consumer, FGD persona, survey, marketing)\n"
    "## Objective\n"
    "(restate what the client wants to know, from the brief)\n"
    "## Recommended Framework\n"
    "(framework name + WHY it fits this brief and guide)\n"
    "## Questionnaire\n"
    "(what survey questions were asked, which tool collected the data, and how the survey data links to the analysis)\n"
    "## Analysis Sections\n"
    "(each section: title + what question it answers)\n"
    "## Report Outline\n"
    "(what goes in the PPTX/DOCX report, in order)\n"
    "## Gaps & Notes\n"
    "(guide or questionnaire topics with no or little evidence in the transcripts; anything missing)\n\n"
)

PLAN_APPLY_TEMPLATE = (
    "You are a senior market research analyst. Follow this approved analysis plan exactly. "
    "Use ONLY information from the transcripts. Do not invent findings, numbers, or quotes. "
    "Write in English. Medicine, brand, and disease names in English. Findings in short paragraphs.\n\n"
    "Structure control (follow strictly):\n"
    "- Use the plan's \"## Analysis Sections\" headings EXACTLY, in the plan's order, word for word. "
    "Do not add, rename, merge, or drop a section.\n"
    "- The plan's \"## Report Outline\" is a report plan, not an analysis section — do not write findings under it.\n"
    "- If a section has no evidence in the transcripts, write exactly this under its heading:\n"
    "  No evidence in transcripts\n\n"
    + EVIDENCE_RULES + "\n" + OUTPUT_FORMAT + "\n"
    "Approved analysis plan:\n{plan}\n\n"
)

QUALITATIVE_SINGLE_TEMPLATE = (
    "You are a senior market research analyst. Analyze the interview transcript below in depth, at a professional level.\n"
    "Rules: Use ONLY information from the transcript. Do not invent findings, "
    "numbers, or quotes. Write in English. Medicine, brand, and disease names in English. "
    "Findings in short paragraphs. Go beyond surface summary: explain WHY behind behaviours, "
    "note drivers and barriers, and distinguish what is said by one respondent from what is a pattern.\n\n"
    + EVIDENCE_RULES + "\n" + OUTPUT_FORMAT + "\n"
    "Structure your answer exactly like this:\n"
    "## Key Findings\n"
    "## Customer Needs\n"
    "## Market Needs\n"
    "## Pain Points\n"
    "## Behaviour & Journey\n"
    "(the purchase/usage journey step by step, with what triggers each step)\n"
    "## Brand & Market View\n"
    "(brand perception and positioning, purchase drivers and barriers including price and the pharmacy's role, channel and shopper behaviour)\n"
    "## Persona\n"
    "(1-2 personas built from behaviour and motivation, each with who they are, what drives them, what blocks them)\n"
    "## Key Quotes\n"
    "(for each quote: > \"...\" — Speaker N, <source file>)\n"
    "## Recommendations\n"
    "(each with priority High, Medium or Low, and which finding it follows from)\n\n"
)

QUALITATIVE_MULTI_TEMPLATE = (
    "You are a senior market research analyst. Analyze these {count} interview transcripts together in depth, at a professional level.\n"
    "Rules: Use ONLY information from the transcripts. Do not invent findings, "
    "numbers, or quotes. Write in English. Medicine, brand, and disease names in English. "
    "Findings in short paragraphs. Go beyond surface summary: explain WHY behind behaviours, note drivers and barriers.\n\n"
    + CROSS_INTERVIEW_RULES + "\n" + EVIDENCE_RULES + "\n" + OUTPUT_FORMAT + "\n"
    "Structure your answer exactly like this:\n"
    "## Key Findings\n"
    "## Cross-Respondent Themes\n"
    "(for each theme: the theme, then n=X of Y, one supporting quote, and any divergent view)\n"
    "## Customer Needs\n"
    "## Market Needs\n"
    "## Pain Points\n"
    "## Behaviour & Journey\n"
    "(the purchase/usage journey step by step, with what triggers each step)\n"
    "## Brand & Market View\n"
    "(brand perception and positioning, purchase drivers and barriers including price and the pharmacy's role, channel and shopper behaviour)\n"
    "## Persona\n"
    "(1-2 personas built from behaviour and motivation, each with who they are, what drives them, what blocks them)\n"
    "## Key Quotes\n"
    "(for each quote: > \"...\" — Speaker N, <source file>)\n"
    "## Recommendations\n"
    "(each with priority High, Medium or Low, and which finding it follows from)\n\n"
)

QUANT_INTERPRET_TEMPLATE = (
    "You are a senior market research analyst. Interpret the survey results below. "
    "Use ONLY the numbers given. Do not invent counts, percentages, or findings. "
    "Write in English.\n\n"
    "Rules: quote every number exactly as given (count, %, n=). A pattern needs a clear "
    "difference in the numbers — do not over-claim from a small sample. "
    "If the survey cannot answer something, say so under \"What Needs a Follow-up\" — do not guess.\n\n"
    "Structure your answer exactly like this:\n"
    "## Key Patterns\n"
    "(the 3-5 clearest patterns, each with its exact count and % from the data)\n"
    "## Notable Differences\n"
    "(where groups or answers differ clearly, with exact numbers)\n"
    "## What Needs a Follow-up\n"
    "(questions the survey cannot answer — suggest what to ask in interviews)\n\n"
)

REPORT_FORMAT_TEMPLATE = (
    "Report export contract (how Gemini output becomes PPTX/DOCX automatically):\n"
    + OUTPUT_FORMAT +
    "- The export reads \"## \" headings as slide/section titles and \"- \" bullets as points.\n"
    "- Anything outside that format (tables, bold headings, other heading levels) is dropped by the export.\n"
    "- Templates: Business Clean, Topline, Dark Navy.\n"
)

EXEC_SUMMARY_TEMPLATE = (
    "You are a senior market research analyst. Write a one-page executive summary (topline) "
    "of the analysis below, for a client who has 2 minutes. "
    "Use ONLY information from the analysis. Do not invent findings, numbers, or quotes. "
    "Write in English. Medicine, brand, and disease names in English.\n\n"
    "Structure your answer exactly like this:\n"
    "## Objective\n"
    "(one paragraph — what the client wanted to know)\n"
    "## Key Findings\n"
    "(3-5 findings only, each with n=X/Y and one short quote)\n"
    "## Recommendations\n"
    "(3 recommendations maximum, each with priority High, Medium or Low)\n\n"
    + EVIDENCE_RULES + "\n" + OUTPUT_FORMAT + "\n"
    "Analysis:\n"
)

PROMPT_LIBRARY = {
    "transcription": {
        "name": "1. Transcription Engine",
        "icon": "🎙️",
        "purpose": "Turns audio into a Burmese transcript with speakers separated (Respondent / Moderator). No summary, no invention — unclear parts marked [unclear].",
        "used_by": "Upload & Transcribe page (transcribe_audio)",
        "template": TRANSCRIPTION_TEMPLATE,
    },
    "planning": {
        "name": "2. Research Planning Engine",
        "icon": "🗺️",
        "purpose": "Reads brief + discussion guide + questionnaire + transcripts and proposes an analysis plan (sections + report outline). Findings are NOT written here — only the plan. The apply variant runs an approved plan with strict structure control.",
        "used_by": "Market Analysis page (Analysis Planner, plan_based_prompt)",
        "template": PLANNING_TEMPLATE,
        "variants": {"Apply an approved plan": PLAN_APPLY_TEMPLATE},
    },
    "qualitative": {
        "name": "3. Qualitative Analysis Engine",
        "icon": "🔍",
        "purpose": "Deep analysis of interviews — findings, needs, pain points, journey, brand view, persona, quotes, recommendations. Evidence rules (n= + verbatim quote + speaker) and the fixed report format apply to every finding.",
        "used_by": "Market Analysis page (Analyze / Analyze multi)",
        "template": QUALITATIVE_SINGLE_TEMPLATE,
        "variants": {"Multi-interview": QUALITATIVE_MULTI_TEMPLATE},
    },
    "quantitative": {
        "name": "4. Quantitative Analysis Engine",
        "icon": "📈",
        "purpose": "Survey numbers (counts, %, means, cross-tabs) are computed from the data in Python — never generated by AI. This prompt only interprets the computed numbers, quoting them exactly.",
        "used_by": "Quantitative Analysis page (Interpret with AI)",
        "template": QUANT_INTERPRET_TEMPLATE,
    },
    "cross_interview": {
        "name": "5. Cross-Interview Intelligence",
        "icon": "🔀",
        "purpose": "Compares interviews: themes with n=X of Y, one quote per theme, divergent views named, patterns need 2+ respondents. Used inside the multi-interview analysis prompt.",
        "used_by": "Market Analysis page (Analyze multi)",
        "template": CROSS_INTERVIEW_RULES,
    },
    "report_generation": {
        "name": "6. Report Generation Engine",
        "icon": "📑",
        "purpose": "The format contract between Gemini output and the PPTX/DOCX export: \"## \" headings become slides/sections, \"- \" bullets become points. Anything outside the format is dropped by the export.",
        "used_by": "Reports page (Export PPTX / DOCX, parse_sections)",
        "template": REPORT_FORMAT_TEMPLATE,
    },
    "executive_summary": {
        "name": "7. Executive Summary Engine",
        "icon": "📝",
        "purpose": "Turns a saved analysis into a one-page client topline: objective, 3-5 key findings with n= and quotes, max 3 recommendations. Same evidence rules and format as the analysis engines.",
        "used_by": "Reports page (Create executive summary)",
        "template": EXEC_SUMMARY_TEMPLATE,
    },
}


def plan_based_prompt(plan_text, combined_text, count):
    base = PLAN_APPLY_TEMPLATE.format(plan=plan_text)
    if count > 1:
        base += (
            f"Analyze these {count} interview transcripts together, following the plan.\n\n"
            "Transcripts:\n" + combined_text
        )
    else:
        base += "Transcript:\n" + combined_text
    return base


# --- PPTX report (Phase 3) ---
TEMPLATES = {
    "Business Clean": {"bg": "FFFFFF", "title": "0070C0", "body": "262626"},
    "Topline": {"bg": "FFFFFF", "title": "10283B", "body": "262626"},
    "Dark Navy": {"bg": "10283B", "title": "FFFFFF", "body": "FFFFFF"},
}


def parse_sections(text):
    """Parse Markdown H2 sections while retaining any preamble or plain text."""
    sections = []
    current = None
    preamble = []
    for raw_line in str(text or "").splitlines():
        if raw_line.startswith("## "):
            if current is None and any(x.strip() for x in preamble):
                sections.append({"title": "Overview", "lines": [x.strip() for x in preamble if x.strip()]})
            current = {"title": raw_line[3:].strip() or "Section", "lines": []}
            sections.append(current)
        else:
            line = raw_line.strip()
            if not line:
                continue
            if line.startswith(("- ", "* ")):
                line = line[2:]
            if current is None:
                preamble.append(line)
            else:
                current["lines"].append(line)
    if current is None and any(x.strip() for x in preamble):
        sections.append({"title": "Analysis", "lines": [x.strip() for x in preamble if x.strip()]})
    return sections


class _HtmlToText(HTMLParser):
    """Convert HTML back to plain text with ## headings (for versions saved earlier as HTML)."""
    def __init__(self):
        super().__init__()
        self.out = []

    def handle_starttag(self, tag, attrs):
        if tag in ("h1", "h2"):
            self.out.append("\n## ")
        elif tag == "li":
            self.out.append("\n- ")
        elif tag == "p":
            self.out.append("\n")
        elif tag == "br":
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in ("p", "h1", "h2"):
            self.out.append("\n")

    def handle_data(self, data):
        if data.strip():
            self.out.append(data)


def html_to_text(html_text):
    parser = _HtmlToText()
    parser.feed(html_text)
    text = "".join(parser.out)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def looks_like_html(text):
    return "<p" in text or "<h1" in text or "<h2" in text or "<li" in text


def _set_bg(slide, hex_color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(hex_color)


def _color_runs(paragraphs, hex_color):
    for para in paragraphs:
        for run in para.runs:
            run.font.color.rgb = RGBColor.from_string(hex_color)


def analysis_to_pptx(analysis_text, title_text, out_path, template="Business Clean"):
    t = TEMPLATES.get(template, TEMPLATES["Business Clean"])
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    _set_bg(slide, t["bg"])
    slide.shapes.title.text = title_text
    _color_runs(slide.shapes.title.text_frame.paragraphs, t["title"])
    slide.placeholders[1].text = "Market Research AI — ResearchFlow"
    _color_runs(slide.placeholders[1].text_frame.paragraphs, t["body"])
    for section in parse_sections(analysis_text):
        # Keep slides readable: split long sections instead of overflowing one slide.
        chunks, current, char_count = [], [], 0
        for line in section["lines"]:
            if current and (len(current) >= 14 or char_count + len(line) > 1500):
                chunks.append(current)
                current, char_count = [], 0
            current.append(line)
            char_count += len(line)
        if current or not chunks:
            chunks.append(current)
        for part_index, lines in enumerate(chunks, start=1):
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            _set_bg(slide, t["bg"])
            slide_title = section["title"] if part_index == 1 else f"{section['title']} (continued {part_index})"
            slide.shapes.title.text = slide_title
            _color_runs(slide.shapes.title.text_frame.paragraphs, t["title"])
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for i, line in enumerate(lines):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = line
                p.level = 0
                for run in p.runs:
                    run.font.color.rgb = RGBColor.from_string(t["body"])
    prs.save(out_path)
    return out_path


def analysis_to_docx(analysis_text, title_text, out_path):
    doc = Document()
    doc.add_heading(title_text, level=0)
    doc.add_paragraph("Market Research AI — ResearchFlow")
    for section in parse_sections(analysis_text):
        heading = doc.add_heading(section["title"], level=1)
        for run in heading.runs:
            run.font.color.rgb = DocxRGBColor(0x00, 0x70, 0xC0)
        for line in section["lines"]:
            doc.add_paragraph(line)
    doc.save(out_path)
    return out_path


# Login gate: authentication is mandatory unless explicit local-only opt-in is set.
_MY_EMAIL = login_gate()

# --- v3.5 look: dark navy sidebar, white cards, blue accent (design reference) ---
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=Noto+Sans+Myanmar:wght@400;500;600;700;800&display=swap');
/* v3.24.1 Myanmar rendering (#1): the font stack must reach EVERY text-bearing
   element, not just .stApp descendants. Root cause of the audit breakage:
   (a) Streamlit/BaseWeb render several widgets with their own explicit fonts
   (widget labels, radio options, select controls, uploader, alerts), and
   dropdown menus / popovers / tooltips live in body-level portals OUTSIDE
   .stApp, so inheritance from .stApp never reached them; (b) tight
   line-heights (1.1-1.25) crush Myanmar stacked marks together, which reads
   as overlapping/piled glyphs; (c) letter-spacing on table/stage headers
   disturbs Myanmar shaping. Fixed here: enumerated selectors (no "*" font
   selector anywhere), open line-height everywhere, letter-spacing removed
   where Myanmar can appear. Streamlit's own icons are Material Symbols
   ligatures (e.g. keyboard_double_arrow_right for the collapsed-sidebar
   button) and render as literal text if their font is overridden — v3.24.1
   broke them this way: the .stApp span rule below (specificity 0,1,1) beat
   the icon font, and the span[class*="material-symbols"] guard never
   matched anything (Streamlit 1.67 sets the icon font via an emotion
   hashed class, not a material-symbols-* class). v3.24.2 restores the icon
   font on the real identifiers (see the stIconMaterial rule below). */
.stApp { font-family: 'IBM Plex Sans', 'Noto Sans Myanmar', 'Myanmar Text', 'Pyidaungsu', -apple-system, 'Segoe UI', sans-serif; }
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp p,
.stApp div, .stApp span, .stApp label, .stApp button, .stApp input,
.stApp textarea, .stApp select, .stApp li, .stApp td, .stApp th,
.stApp a, .stApp small, .stApp summary
{ font-family: 'IBM Plex Sans', 'Noto Sans Myanmar', 'Myanmar Text', 'Pyidaungsu', -apple-system, 'Segoe UI', sans-serif; }
/* Dropdown menus (Streamlit 1.67 has no BaseWeb — the old data-baseweb
   selectors matched nothing, so the v3.24.1 menu font/line-height rules
   were dead). The 1.67 select/multiselect menus render inside the widget
   element, so they inherit .stApp's font; this explicit rule keeps Myanmar
   option text on the same stack and line height. */
[data-testid="stSelectboxVirtualDropdown"], [data-testid="stMultiSelectDropdown"],
[data-testid="stSelectboxVirtualDropdown"] li, [data-testid="stMultiSelectDropdown"] li,
[data-testid="stSelectboxVirtualDropdown"] div, [data-testid="stMultiSelectDropdown"] div
{ font-family: 'IBM Plex Sans', 'Noto Sans Myanmar', 'Myanmar Text', 'Pyidaungsu', -apple-system, 'Segoe UI', sans-serif; }
/* Myanmar script needs open line height; tight tracking breaks stacked marks. */
[data-testid="stMarkdownContainer"] p { line-height: 1.7; }
[data-testid="stMarkdownContainer"] li { line-height: 1.7; }
[data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] label { line-height: 1.7; }
[data-testid="stCaptionContainer"] { line-height: 1.6; }
.stApp h1, .stApp h2, .stApp h3, .stApp h4 { line-height: 1.7; }
/* v3.24.3 (#7): Myanmar stacking diacritics render tight in buttons and
   labels at 1.6 — open them to 1.8 (headings 1.7 above). */
.stApp button, .stApp label { line-height: 1.8; }
[data-testid="stSelectboxVirtualDropdown"] li, [data-testid="stMultiSelectDropdown"] li { line-height: 1.6; }
input::placeholder, textarea::placeholder { font-family: 'IBM Plex Sans', 'Noto Sans Myanmar', 'Myanmar Text', 'Pyidaungsu', -apple-system, 'Segoe UI', sans-serif; color: #64748B; opacity: 1; line-height: 1.6; }
[data-testid="stSelectbox"] input::placeholder, [data-testid="stMultiSelect"] input::placeholder { color: #64748B; opacity: 1; }
/* v3.24.2 (#5): restore the Material Symbols icon font. Streamlit 1.67
   renders every Material icon as
   <span data-testid="stIconMaterial" aria-hidden="true" translate="no">
   with font-family "Material Symbols Rounded" applied by an emotion hashed
   class (specificity 0,1,0) — the .stApp span rule above (0,1,1) overrode
   it, so icon names rendered as literal text ("uploadUpload",
   "_arrow_right", "keyboard_double"). These rules target the real
   identifiers and win with !important; the class*="material-symbols" line
   is kept only as a fallback (it never matched in 1.67). Never put the app
   font stack on an icon span. */
span[data-testid="stIconMaterial"],
span[aria-hidden="true"][translate="no"],
span[role="img"][translate="no"] { font-family: 'Material Symbols Rounded', sans-serif !important; }
span[class*="material-symbols"] { font-family: 'Material Symbols Outlined', 'Material Symbols Rounded', sans-serif !important; }
:root {
  --rf-navy: #071B3A;
  --rf-navy-2: #0B2A52;
  --rf-blue: #2563EB;
  --rf-cyan: #0891B2;
  --rf-teal: #0D9488;
  --rf-text: #16222E;
  --rf-muted: #5E6B78;
  --rf-border: #D9DEE4;
  --rf-surface: #FFFFFF;
  /* v3.19 tokens */
  --rf-ground: #F3F5F7;
  --rf-ink: #16222E;
  --rf-line: #D9DEE4;
  --rf-accent: #0E8F7E;
  --rf-accent-soft: #E4F2F0;
  --rf-attn: #C26A00;
}

/* v3.24.1: flex columns must be allowed to shrink below their content width,
   or one wide table/label forces its whole row to overflow its neighbours (#9). */
[data-testid="stColumn"] { min-width: 0; }
[data-testid="stVerticalBlock"] { min-width: 0; }
[data-testid="stVerticalBlockBorderWrapper"] { min-width: 0; }

/* App canvas */
.stApp {
  background: var(--rf-ground);
  color: var(--rf-text);
}
.block-container {
  width: 100%;
  max-width: 1540px;
  padding: 4.5rem clamp(1.2rem, 2.2vw, 2.25rem) 6rem;
  /* v3.24.1: 6rem bottom padding so page content is never underneath
     Streamlit's floating "Manage app" button (#14, #15); 4.5rem top padding
     clears the fixed header. */
}

/* Compact sidebar: reserve more screen space for the workspace.
   Width rules apply ONLY while expanded — an unqualified min-width: 250px
   !important stops the sidebar collapsing (the page stays pushed right,
   "left behind"). When collapsed, Streamlit's own rules must apply untouched.
   v3.24.4 (#5): the `> div` width is scoped to [data-testid=
   "stSidebarContent"] — the bare `> div` also matched the resize handle
   (data-testid="stSidebarResizeHandle", a direct child div of the section),
   which must stay 8px wide. Narrow widths: below Streamlit's md breakpoint
   (768px viewport) <main> becomes position:absolute and the expanded sidebar
   floats OVER the content by design (mobile pattern: tap outside to close,
   auto-collapses on resize-down, starts collapsed on a fresh narrow load).
   That is why a ~1013px window with a wide browser side panel (viewport
   <= 768px) showed the sidebar covering the main content — it is not an
   app rule breaking at ~1000px. At viewport >= 768px the sidebar pushes
   content (flex row); the 900px rule below only caps its width. Do NOT
   force push-layout below 768px — that would break phone use. */
section[data-testid="stSidebar"][aria-expanded="true"] {
  width: 250px !important;
  min-width: 250px !important;
  max-width: 250px !important;
  background: var(--rf-surface);
  border-right: 1px solid var(--rf-line);
}
section[data-testid="stSidebar"][aria-expanded="true"] > [data-testid="stSidebarContent"] {
  width: 250px !important;
}
@media (min-width: 901px) {
  section[data-testid="stSidebar"][aria-expanded="true"] {
    width: 250px !important;
    min-width: 250px !important;
    max-width: 250px !important;
  }
}

/* v3.24.1: the Streamlit header/toolbar is platform chrome, not app content.
   Give it a solid theme background so scrolled content disappears cleanly
   behind it instead of showing through / over a floating light bar (#7, #13).
   Dark theme overrides this below with its own solid colour.
   v3.24.3 (#5): the header background MUST be fully opaque with !important
   (the 1.67 header is transparent when it has no logo/nav/toolbar content,
   so headings sliding under it stayed half-visible); the bottom border marks
   where content disappears, and the high z-index keeps it above page content. */
[data-testid="stHeader"] { background: #FFFFFF !important; border-bottom: 1px solid #E2E8F0; z-index: 1000000 !important; }
[data-testid="stToolbar"] { color: #172B4D; background: #FFFFFF !important; }
/* v3.24.3 (#4, #6): collapse Streamlit 1.67's sidebar header. The header is
   3.75rem tall with a 2rem logo spacer (stLogoSpacer — rendered because the
   app sets no st.logo) plus a bottom margin, which together with the 3.5rem
   scroll-container padding put the logo ~185px from the top. The spacer
   serves no purpose here, so hide it; shrink the header to the collapse
   button alone and drop its margin. The collapse button (stSidebarCollapseButton)
   is visibility:hidden until the sidebar is hovered (verified in the 1.67
   bundle: the jx wrapper toggles visibility on showSidebarCollapse) — force
   it visible on load so it is always reachable. */
[data-testid="stSidebarHeader"] { height: auto !important; min-height: 2.5rem !important; margin-bottom: 0 !important; justify-content: flex-end !important; }
[data-testid="stLogoSpacer"] { display: none !important; }
[data-testid="stSidebarCollapseButton"] { visibility: visible !important; opacity: 1 !important; }

/* v3.24.4 (#2): the sidebar top padding is cut from 3.5rem to 0.5rem.
   The 3.5rem was added in v3.24.1 on the belief that the fixed Streamlit
   header overlays the top of the sidebar — verified against the 1.67 bundle
   it does NOT (at viewport >= 768px the header lives inside the main column
   only; below 768px the sidebar paints ABOVE the header, z-index header+1).
   The padding was pure dead space: with the header row already collapsed to
   2.5rem it put the logo ~160px from the top. 0.5rem keeps a small clearance
   so the collapse button is not jammed at y=0. */
[data-testid="stSidebar"] > div {
  padding-top: 0.5rem !important;
  /* v3.24.2 (#6): anchor/focus scrolling in the sidebar stops below the
     pinned logo instead of sliding a nav item under it.
     v3.24.4: 5rem — the stuck logo now pins at top:0 (see below), so its
     bottom edge sits ~62px down; anchor scrolling must clear 80px. */
  scroll-padding-top: 5rem !important;
}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
  padding-top: 0 !important;
}
/* v3.24.3 (#4): Streamlit 1.67's stSidebarUserContent adds padding-top
   (spacing.twoXL when page nav sits above) — remove it; the logo must start
   right below the collapsed sidebar header. */
[data-testid="stSidebarUserContent"] { padding-top: 0 !important; }

/* v3.19 light sidebar: hairline border, no gradient. The active page is marked
   by an accent bar on its left edge. Stage entries carry the stage number;
   the other pages of a stage are indented under it. */
[data-testid="stSidebar"] {
  background: var(--rf-surface);
  border-right: 1px solid var(--rf-line);
}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: .3rem; }
[data-testid="stSidebar"] h1 {
  font-size: 1.35rem;
  line-height: 1.6;
  margin-bottom: .15rem;
}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
  font-size: .76rem; color: var(--rf-muted);
}
[data-testid="stSidebar"] .stButton > button {
  min-height: 2.6rem; padding: .35rem .75rem; border-radius: 6px;
  font-size: .9rem; font-weight: 500; justify-content: flex-start; gap: .6rem;
  background: transparent; color: var(--rf-ink); border: none; box-shadow: none;
  transform: none; line-height: 1.6; white-space: normal;
}
/* Streamlit centers nav content in an inner div INSIDE the button, so
   justify-content on the button alone cannot move it (v3.13 lesson). */
[data-testid="stSidebar"] .stButton button div { justify-content: flex-start !important; }
[data-testid="stSidebar"] .stButton > button p { color: var(--rf-ink); text-align: left; line-height: 1.8; white-space: normal; }
[data-testid="stSidebar"] .stButton > button:hover {
  background: #EEF1F4; border: none; color: var(--rf-ink); transform: none;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"],
[data-testid="stSidebar"] .stButton > button[data-testid="stBaseButton-primary"] {
  background: var(--rf-accent-soft); border: none; color: var(--rf-ink);
  font-weight: 600; box-shadow: inset 3px 0 0 var(--rf-accent);
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] p,
[data-testid="stSidebar"] .stButton > button[data-testid="stBaseButton-primary"] p {
  color: var(--rf-ink); font-weight: 600;
}
[data-testid="stSidebar"] [class*="st-key-navsub"] button {
  padding-left: 2.3rem; font-size: .84rem;
}
[data-testid="stSidebar"] [class*="st-key-navsub"] button p { color: var(--rf-muted); }
[data-testid="stSidebar"] [class*="st-key-navsub"] button[kind="primary"] p,
[data-testid="stSidebar"] [class*="st-key-navsub"] button[data-testid="stBaseButton-primary"] p {
  color: var(--rf-ink);
}
[data-testid="stSidebar"] [data-testid="stExpander"] {
  background: transparent; border: none; border-radius: 6px;
}
[data-testid="stSidebar"] [data-testid="stExpander"] details summary {
  padding-top: .45rem; padding-bottom: .45rem;
}
[data-testid="stSidebar"] hr { border-color: var(--rf-line); }

/* Sidebar top block stays visible (v3.21 logo pin; v3.24: only the logo is
   pinned — the old current-project label it used to pin below is gone).
   DECISION (b), user-chosen: the logo STAYS pinned, with a SOLID OPAQUE
   backing so scrolling content can neither show through it nor cover it.
   v3.24.4 — REGRESSION in v3.24.3: the logo was NOT pinned at all (it
   scrolled off the top entirely). Root cause, verified against the 1.67 DOM:
   the v3.24.3 selector used `stSidebarUserContent > stVerticalBlock`, but
   the sidebar's children are wrapped in a plain <div> (no testid) BETWEEN
   them — real DOM: stSidebarUserContent > div > stVerticalBlock >
   div[data-testid="stElementContainer"] (logo). The `>` combinator matched
   NOTHING, so the rule was dead. The fix uses a descendant combinator
   through that wrapper, and targets the logo's element container directly
   with :has(.rf-logo) — no reliance on child order, no way to match the
   wrong element. Sticky top is 0 (not 3.75rem): the main header does not
   cover the sidebar (it lives in the main column at viewport >= 768px, and
   below 768px the sidebar paints ABOVE it, z-index header+1), so the logo
   pins at the very top — the user's "keep pinned at top". The wrapper gets
   an opaque sidebar-coloured background (!important) plus z-index 20, and
   the .rf-logo element itself gets the same solid background as a second
   layer (dark: #0F172A via DARK_CSS). At rest nothing is hidden; while
   scrolling, nav items pass behind the opaque logo — the intended pinned
   behaviour. The nested-block rule after this one still un-sticks anything
   inside columns, so the "English" / "🌙 Dark" buttons can never become
   sticky and paint over the logo again (v3.24.2 root cause). */
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stVerticalBlock"] > div[data-testid="stElementContainer"]:has(.rf-logo) {
  position: sticky !important; top: 0; z-index: 20; background: var(--rf-surface) !important;
  border-bottom: 1px solid var(--rf-line); padding-bottom: .15rem;
}
/* Never let a button inside a sidebar column become sticky (see above). */
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] [data-testid="stVerticalBlock"] > div {
  position: static !important;
}

/* v3.24: the current-project label is gone (project shows in the stage
   header and the top bar instead) — its styles were removed with it. */

/* Log out was cramped against the user card — give it room to breathe. */
[data-testid="stSidebar"] [class*="st-key-logout_btn"] { margin-top: .8rem; }
[data-testid="stSidebar"] [class*="st-key-logout_btn"] button {
  padding-top: .55rem; padding-bottom: .55rem; margin-bottom: .4rem;
}

/* Sidebar logo + user card */
/* v3.24.3 (logo decision b): the logo element itself carries the same solid
   opaque sidebar background as its sticky wrapper, so the logo text can
   never be covered by scrolling content even if the wrapper rule is missed. */
.rf-logo { display: flex; align-items: center; gap: 10px; padding: 0 4px 14px; background: var(--rf-surface); }
.rf-logo-tx { color: var(--rf-ink); font-size: 19px; font-weight: 700; line-height: 1.5; }
.rf-logo-sub { color: var(--rf-muted); font-size: 11px; margin-top: 2px; }
.rf-user { display: flex; align-items: center; gap: 10px; padding: 12px 10px 4px; margin-top: 8px; border-top: 1px solid var(--rf-line); }
.rf-user-av { width: 34px; height: 34px; border-radius: 50%; background: var(--rf-accent); color: #FFFFFF; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 15px; flex-shrink: 0; }
.rf-user-nm { color: var(--rf-ink); font-weight: 600; font-size: 14px; line-height: 1.6; }
.rf-user-rl { color: var(--rf-muted); font-size: 12px; line-height: 1.6; }

/* Typography and section headings */
h1, h2, h3 { color: var(--rf-ink); letter-spacing: 0; }
/* v3.24.1: headings must never slide under the fixed Streamlit header (#13).
   v3.24.3 (#5): 6rem — the header is 3.75rem tall, so anchor-scrolled
   headings stop 36px clear of it.
   v3.24.4 (#3): scroll-padding belongs on the REAL scroll container.
   The main pane (<main>, data-testid="stMain") is the element with
   overflow:auto and height:100dvh — the document itself never scrolls, so
   the old `html { scroll-padding-top }` alone did nothing for anchor jumps.
   The scroll-to-top-on-page-change itself is done in Python (a gated
   st.html script just before the page dispatch) — the v3.24.3 note claiming
   st.html strips <script> was WRONG for 1.67: the Python API exposes
   st.html(body, unsafe_allow_javascript=True) (html.py:52) and the
   frontend Html chunk re-creates <script> elements so they run in the main
   document (not iframed). No native-navigation refactor was needed. */
h1, h2, h3, .rf-sh { scroll-margin-top: 6rem; }
html { scroll-padding-top: 6rem; }
[data-testid="stMain"] { scroll-padding-top: 6rem; }
h1 { font-weight: 600; }
h2, h3 { font-weight: 600; }
[data-testid="stCaptionContainer"] { color: var(--rf-muted); }

/* Hero banner — light blue-to-cyan gradient */
.rf-hero {
  position: relative; overflow: hidden;
  background: linear-gradient(115deg, #EAF3FF 0%, #DDF3F8 55%, #D9F5EF 100%);
  border: 1px solid #D5E5F7; border-radius: 18px;
  padding: 20px 26px; margin: 4px 0 16px;
  box-shadow: 0 10px 28px rgba(16,38,74,.08);
}
.rf-hero:after {
  content: ""; position: absolute; width: 300px; height: 300px; right: -70px; top: -170px;
  border-radius: 50%; background: radial-gradient(circle, rgba(24,198,217,.22), rgba(24,198,217,0) 68%);
  pointer-events: none;
}
.rf-hero h1 { color: #10264A; font-size: clamp(24px, 2.6vw, 34px); margin: 0 0 6px; font-weight: 800; line-height: 1.6; overflow-wrap: anywhere; }
.rf-hero h1 span.b { color: #2563EB; }
.rf-hero h1 span.t { color: #0D9488; }
.rf-hero p { color: #475569; font-size: 16px; margin: 0; max-width: 760px; line-height: 1.7; }

/* Modern metric cards */
.rf-metric { position: relative; overflow: hidden; background: #FFFFFF; border: 1px solid var(--rf-border); border-radius: 16px; padding: 18px; box-shadow: 0 5px 18px rgba(16,38,74,.06); }
.rf-metric:before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, #2563EB, #0891B2, #0D9488); }
.rf-metric-ico { width: 40px; height: 40px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 20px; margin-bottom: 10px; }
.rf-metric-label { color: #64748B; font-size: 13px; font-weight: 600; line-height: 1.6; overflow-wrap: anywhere; }
.rf-metric-value { color: #10264A; font-size: 30px; font-weight: 800; line-height: 1.4; }

/* Workflow cards and progress ring */
.rf-step {
  background: linear-gradient(160deg, #FFFFFF 0%, #F8FAFF 100%);
  border: 1px solid #E5E9F7; border-radius: 15px; padding: 14px 10px;
  text-align: center;
  box-shadow: 0 5px 16px rgba(16,38,74,.05);
  transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease;
}
/* v3.21: per-stage sub-items under each workflow card, so stages 3 and 4
   no longer look alike. */
.rf-step .subs { text-align: left; margin-top: 8px; border-top: 1px solid #EEF2F8; padding-top: 8px; }
.rf-step .sub { color: #475569; font-size: 11px; line-height: 1.6; white-space: normal; overflow-wrap: anywhere; }
.rf-step .sub .ok { color: #0E7A5F; font-weight: 700; }
.rf-step .sub .no { color: #94A3B8; }
.rf-step .sub .cm { color: #94A3B8; }
.rf-step .sub i { color: #94A3B8; font-style: normal; }
.rf-step:hover { transform: translateY(-3px); border-color: #BFD9F2; box-shadow: 0 10px 24px rgba(18,100,197,.12); }
.rf-step .ico { width: 52px; height: 52px; border-radius: 50%; margin: 0 auto 8px; display: flex; align-items: center; justify-content: center; font-size: 24px; }
.rf-step .nm { font-weight: 700; color: #10264A; font-size: 13px; line-height: 1.6; overflow-wrap: anywhere; }
.rf-step .ds { color: #64748B; font-size: 11px; margin-top: 5px; line-height: 1.6; overflow-wrap: anywhere; }
.rf-step .st { color: #2563EB; font-size: 12px; font-weight: 700; margin-top: 8px; }
.rf-ring {
  width: 104px; height: 104px; border-radius: 50%; margin: 12px auto;
  background: conic-gradient(from 220deg, #2563EB calc(var(--p)*1%), #0891B2 0, #E9ECF7 0);
  display: flex; align-items: center; justify-content: center; position: relative;
  box-shadow: 0 8px 22px rgba(18,100,197,.12);
}
.rf-ring:before { content: ""; position: absolute; width: 78px; height: 78px; border-radius: 50%; background: #FFFFFF; }
.rf-ring span { position: relative; z-index: 1; font-size: 22px; font-weight: 800; color: #10264A; }

/* Recent activity table with status badges */
/* v3.24.2 (#7): table-layout fixed + overflow-wrap anywhere squeezed the
   table at 1341px and broke headers/cells mid-word ("Dat e & Tim e").
   Cells are now nowrap and the table scrolls horizontally inside its own
   wrapper instead of squeezing — the wrapper's overflow-x already handles it.
   v3.24.3 (#3): the nowrap fix kept the Status column cut off at the right
   edge (only a green sliver of the "Completed" pill showed, behind the
   table's own horizontal scrollbar). The table's natural width (~690px) is
   far wider than its ~440px column, so scrolling was unavoidable. Now the
   table is fixed-layout at 100% of its container: Date/Activity/Status wrap
   at WORD boundaries only (never mid-word), the long File name truncates
   with an ellipsis (full name on hover via the title attribute), and the
   Status column gets a guaranteed 20% so the whole pill always fits — no
   horizontal scrolling at 1341px. */
.rf-table-wrap { overflow-x: auto; }
.rf-table { width: 100%; border-collapse: collapse; font-size: 14px; table-layout: fixed; }
.rf-table th:nth-child(1), .rf-table td:nth-child(1) { width: 26%; }
.rf-table th:nth-child(2), .rf-table td:nth-child(2) { width: 30%; }
.rf-table th:nth-child(3), .rf-table td:nth-child(3) { width: 24%; }
.rf-table th:nth-child(4), .rf-table td:nth-child(4) { width: 20%; }
.rf-table th { text-align: left; color: #64748B; font-size: 12px; font-weight: 700; letter-spacing: 0; padding: 10px 8px; border-bottom: 1px solid var(--rf-border); line-height: 1.6; white-space: normal; overflow-wrap: normal; word-break: normal; overflow: hidden; }
.rf-table td { padding: 8px; border-bottom: 1px solid #EEF2F8; color: #10264A; vertical-align: middle; line-height: 1.6; white-space: normal; overflow-wrap: normal; word-break: normal; overflow: hidden; }
.rf-table td:nth-child(2) { white-space: nowrap; text-overflow: ellipsis; }
.rf-table tr:last-child td { border-bottom: none; }
.rf-badge { display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 700; background: #E2F5EE; color: #0E7A5F; white-space: nowrap; }
/* v3.24.2 (#7) safety net: on narrow screens the Dashboard bottom row
   (Recent Activity + Quick Actions + System Status) stacks vertically
   instead of squeezing its three cards side by side. */
@media (max-width: 1200px) {
  [data-testid="stHorizontalBlock"]:has(.rf-table-wrap) { flex-wrap: wrap !important; }
  [data-testid="stHorizontalBlock"]:has(.rf-table-wrap) > [data-testid="stColumn"] { flex: 1 1 100% !important; min-width: 0 !important; }
}

/* Guide + system architecture */
.rf-guide { background: #FFFFFF; border: 1px solid var(--rf-border); border-radius: 16px; padding: 20px 22px; box-shadow: 0 5px 18px rgba(16,38,74,.05); margin: 0 0 16px; }
.rf-arch { background: #F4F8FF; border: 1px solid var(--rf-border); border-radius: 12px; padding: 14px 16px; text-align: center; font-weight: 700; color: #10264A; }
.rf-arch small { display: block; font-weight: 500; color: #64748B; font-size: 12px; margin-top: 2px; }
.rf-arch-arrow { text-align: center; color: #94A3B8; font-size: 18px; line-height: 1.6; padding: 2px 0; }

/* Buttons and form controls */
.stButton > button, .stDownloadButton > button, [data-testid="stLinkButton"] a {
  border-radius: 6px; border: 1px solid var(--rf-line); font-weight: 600; font-size: 14px;
  transition: background .15s ease, border-color .15s ease; box-shadow: none;
  line-height: 1.6; white-space: normal; height: auto; min-height: 2.5rem;
}
/* v3.24.1: Streamlit truncates long button labels with an ellipsis — let
   labels (incl. Myanmar) wrap instead of "▶ Start Ne…" (#10).
   v3.24.2 (#7): overflow-wrap anywhere broke labels MID-WORD in narrow
   columns ("Uplo ad Files", "Creat e") — wrap at word boundaries only. */
.stButton > button p, .stDownloadButton > button p { white-space: normal !important; overflow: visible !important; text-overflow: clip !important; line-height: 1.6; overflow-wrap: normal; word-break: normal; }
/* v3.24.1: horizontal radio option bars (Stage 4/5 section + step bars) must
   wrap instead of crushing long Myanmar labels (#3). */
[data-testid="stRadio"] div[role="radiogroup"] { flex-wrap: wrap; }
.stButton > button[kind="primary"] {
  color: #FFFFFF; border: 0; background: var(--rf-accent); box-shadow: none;
}
.stButton > button[kind="primary"] p { color: #FFFFFF; }
.stButton > button[kind="primary"]:hover { background: #1D4ED8; }
.stButton > button:hover, .stDownloadButton > button:hover { border-color: var(--rf-accent); }
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea,
[data-testid="stSelectbox"] div:has(> input), [data-testid="stMultiSelect"] div:has(> input),
[data-testid="stNumberInput"] input {
  border-radius: 6px; border-color: var(--rf-line); background: #FFFFFF;
}
[data-testid="stFileUploader"] section {
  border: 1.5px dashed #B6C2CC; border-radius: 8px; background: #FFFFFF;
}

/* Tables, tabs, expanders, alerts */
[data-testid="stDataFrame"] { border: 1px solid var(--rf-line); border-radius: 8px; overflow: hidden; }
[data-testid="stTabs"] button[role="tab"] { font-weight: 600; }
[data-testid="stTabs"] button[aria-selected="true"] { color: var(--rf-accent); }
[data-testid="stExpander"] { border: 1px solid var(--rf-line); border-radius: 8px; background: #FFFFFF; }
[data-testid="stAlertContainer"] { border-radius: 8px; }
hr { border-color: var(--rf-line); }

/* v3.23 stage header: stage number + name + project + status badge. */
.rf-sh { display: flex; align-items: center; justify-content: space-between; gap: 16px;
  background: #FFFFFF; border: 1px solid var(--rf-line); border-left: 4px solid var(--rf-accent);
  border-radius: 10px; padding: 12px 18px; margin: 0 0 10px; }
.rf-sh-k { color: var(--rf-muted); font-size: 12px; font-weight: 600; letter-spacing: 0; line-height: 1.6; }
.rf-sh-t { color: var(--rf-ink); font-size: 20px; font-weight: 700; line-height: 1.7; overflow-wrap: anywhere; }
.rf-sh-p { color: var(--rf-muted); font-size: 13px; line-height: 1.6; }
.rf-sh-b { padding: 4px 12px; border-radius: 999px; font-size: 12px; font-weight: 700; white-space: nowrap; }
.rf-sh-b.done { background: #E2F5EE; color: #0E7A5F; }
.rf-sh-b.wip { background: #FFF1DD; color: #C26A00; }
.rf-sh-b.todo { background: #EEF1F4; color: #5E6B78; }

/* v3.24 top bar chips (project + user) + KPI/workflow flex rows + status pills. */
.rf-chip { background: var(--rf-surface); border: 1px solid var(--rf-line); border-radius: 999px; padding: 9px 14px; font-size: 13px; font-weight: 600; color: var(--rf-ink); white-space: normal; text-align: center; margin-top: 2px; overflow-wrap: anywhere; line-height: 1.6; }
.rf-kpirow { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
.rf-kpirow .rf-metric { flex: 1 1 200px; min-width: 0; }
.rf-flow { display: flex; gap: 6px; align-items: stretch; flex-wrap: wrap; margin-bottom: 8px; min-width: 0; }
.rf-flow .rf-step { flex: 1 1 150px; min-width: 140px; }
.rf-flow-arrow { flex: 0 0 10px; align-self: center; text-align: center; color: #94A3B8; font-size: 18px; }
.rf-pill { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 700; white-space: nowrap; }
.pill-done { background: #E2F5EE; color: #0E7A5F; }
.pill-wip { background: #DBEAFE; color: #1D4ED8; }
.pill-todo { background: #EEF1F4; color: #5E6B78; }
.rf-current { display: inline-block; background: #DBEAFE; color: #1D4ED8; border-radius: 999px; padding: 4px 12px; font-size: 12px; font-weight: 700; margin-bottom: 10px; }

@media (max-width: 1100px) {
  /* v3.24.1 (#11): under 1100px the arrows eat the width — drop them and let
     the workflow cards wrap two-per-row instead of clipping. */
  .rf-flow-arrow { display: none; }
  .rf-flow .rf-step { flex: 1 1 220px; }
}
@media (max-width: 900px) {
  section[data-testid="stSidebar"][aria-expanded="true"],
  section[data-testid="stSidebar"][aria-expanded="true"] > [data-testid="stSidebarContent"] {
    width: min(82vw, 280px) !important;
    min-width: min(82vw, 280px) !important;
    max-width: min(82vw, 280px) !important;
  }
  .block-container { padding: 4.5rem 1rem 6rem; }
  .rf-hero { padding: 23px 20px; }
  .rf-step { min-height: 132px; }
}
</style>
"""
# --- v3.24 theme system: centralized config, CSS variables, one UI.
# Light is the default. The sidebar buttons change st.session_state["theme"]
# and rerun; this block re-injects the matching variables before any page
# content renders, so the whole app switches at once. The sidebar colour is
# decided by the theme (Light = white, Dark = #0F172A) — the redesign brief's
# "dark navy sidebar" is the Dark-theme look, not a fixed colour.
THEMES = {
    "light": {
        "ground": "#F4F7FC", "surface": "#FFFFFF", "text": "#172B4D",
        "ink": "#172B4D", "muted": "#64748B", "line": "#E2E8F0",
        "border": "#E2E8F0", "accent": "#2563EB", "accent-soft": "#E8EEFB",
        "blue": "#2563EB", "cyan": "#0891B2", "teal": "#0D9488",
        "attn": "#C26A00",
    },
    "dark": {
        "ground": "#0B1220", "surface": "#111827", "text": "#F8FAFC",
        "ink": "#F8FAFC", "muted": "#CBD5E1", "line": "#334155",
        "border": "#334155", "accent": "#3B82F6", "accent-soft": "#1E293B",
        "blue": "#60A5FA", "cyan": "#22D3EE", "teal": "#2DD4BF",
        "attn": "#F59E0B",
    },
}
_THEME_NAME = st.session_state.get("theme", "light")
_THEME = THEMES.get(_THEME_NAME, THEMES["light"])

# Dark Mode overrides — applied after CUSTOM_CSS when theme == dark. The
# :root variables are redefined by the theme block below; these rules cover
# the component colours that CUSTOM_CSS hardcodes instead of variables.
# v3.24.1: Streamlit itself still runs its LIGHT theme (this app fakes dark
# mode with CSS), so every native widget defaults to dark text on a white
# control. Each widget family gets an explicit dark override below (#2-#8,
# #12). Dataframes and charts keep Streamlit's styling where CSS cannot
# reach them — noted as a limitation, not faked.
DARK_CSS = """<style>
.stApp { background: #0B1220; color: #F8FAFC; }
section[data-testid="stSidebar"][aria-expanded="true"] { background: #0F172A; border-right-color: #334155; }
[data-testid="stSidebar"] { background: #0F172A; border-right-color: #334155; }
[data-testid="stSidebar"] .stButton > button:hover { background: #1E293B; }
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stVerticalBlock"] > div[data-testid="stElementContainer"]:has(.rf-logo) { background: #0F172A !important; }
.rf-logo { background: #0F172A !important; }
.rf-hero { background: linear-gradient(115deg, #0F1B33 0%, #12233F 55%, #0F2A33 100%); border-color: #1E293B; }
.rf-hero h1 { color: #F8FAFC; }
.rf-hero p { color: #CBD5E1; }
.rf-metric { background: #1E293B; border-color: #334155; }
.rf-metric-label { color: #CBD5E1; }
.rf-metric-value { color: #F8FAFC; }
.rf-step { background: linear-gradient(160deg, #1E293B 0%, #16213A 100%); border-color: #334155; }
.rf-step .nm { color: #F8FAFC; }
.rf-step .ds { color: #CBD5E1; }
.rf-step .subs { border-top-color: #334155; }
.rf-step .sub { color: #CBD5E1; }
.rf-step .sub .no, .rf-step .sub .cm, .rf-step .sub i { color: #94A3B8; }
.rf-step .sub .ok { color: #34D399; }
.rf-step:hover { border-color: #3B82F6; }
.rf-ring:before { background: #111827; }
.rf-ring span { color: #F8FAFC; }
.rf-sh { background: #1E293B; border-color: #334155; }
.rf-sh-t { color: #F8FAFC; }
.rf-sh-p, .rf-sh-k { color: #CBD5E1; }
.rf-sh-b.done { background: #064E3B; color: #6EE7B7; }
.rf-sh-b.wip { background: #451A03; color: #FDBA74; }
.rf-sh-b.todo { background: #1E293B; color: #94A3B8; }
.pill-done { background: #064E3B; color: #6EE7B7; }
.pill-wip { background: #1E3A8A; color: #93C5FD; }
.pill-todo { background: #1E293B; color: #94A3B8; }
.rf-current { background: #1E3A8A; color: #93C5FD; }
.rf-table th { color: #94A3B8; border-bottom-color: #334155; }
.rf-table td { color: #F8FAFC; border-bottom-color: #1E293B; }
.rf-badge { background: #064E3B; color: #6EE7B7; }
.rf-guide { background: #1E293B; border-color: #334155; }
.rf-arch { background: #0F172A; border-color: #334155; color: #F8FAFC; }
.rf-arch small { color: #94A3B8; }
.stButton > button, .stDownloadButton > button, [data-testid="stLinkButton"] a { background: #1E293B; color: #F8FAFC; border-color: #334155; }
.stButton > button[kind="primary"] { background: #2563EB; color: #FFFFFF; }
.stButton > button[kind="primary"]:hover { background: #1D4ED8; }
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea, [data-testid="stSelectbox"] div:has(> input), [data-testid="stMultiSelect"] div:has(> input), [data-testid="stNumberInput"] input { background: #0F172A; color: #F8FAFC; border-color: #334155; }
/* v3.24.2 (#1): dropdown menus in Dark — Streamlit 1.67 has no BaseWeb, so
   the old div[data-baseweb="menu"] rules matched nothing and menus stayed
   white. Real testids: stSelectboxVirtualDropdown / stMultiSelectDropdown. */
[data-testid="stSelectboxVirtualDropdown"], [data-testid="stMultiSelectDropdown"] { background: #1E293B; border: 1px solid #334155; }
[data-testid="stSelectboxVirtualDropdown"] li, [data-testid="stMultiSelectDropdown"] li { color: #F8FAFC; }
[data-testid="stFileUploader"] section { background: #0F172A; border-color: #475569; }
[data-testid="stExpander"] { background: #1E293B; border-color: #334155; }
[data-testid="stDataFrame"] { border-color: #334155; }
[data-testid="stCodeBlock"] { background: #0F172A; }
[data-testid="stCodeBlock"] code { color: #E2E8F0; }
.rf-chip { background: #1E293B; border-color: #334155; color: #F8FAFC; }
.rf-land { background: linear-gradient(115deg, #0F1B33 0%, #12233F 55%, #0F2A33 100%); border-color: #1E293B; }
.rf-land h1 { color: #F8FAFC; }
.rf-land .sub, .rf-land p { color: #CBD5E1; }
.rf-land .stg { background: #1E293B; border-color: #334155; color: #F8FAFC; }
.rf-land-note { color: #94A3B8; }
/* v3.24.1 — Dark theme widget re-theming (#2-#8, #12). */
/* Streamlit chrome: solid themed header so no white bar floats over dark
   content or over the sidebar header when the sidebar scrolls (#7, #13).
   v3.24.3 (#5): bottom border so the header's edge is visible against dark
   content; the collapse button (forced visible in CUSTOM_CSS) gets a light
   colour here — the 1.67 bundle gives its icon fadedText60 from the LIGHT
   theme, which is faint dark-grey on the dark sidebar. */
[data-testid="stHeader"] { background: #0F172A !important; border-bottom: 1px solid #334155; }
[data-testid="stHeader"] button { color: #F8FAFC !important; }
[data-testid="stHeader"] a { color: #F8FAFC !important; }
[data-testid="stToolbar"] { color: #F8FAFC !important; background: #0F172A !important; }
[data-testid="stSidebarCollapseButton"] { color: #CBD5E1 !important; }
[data-testid="stSidebarCollapseButton"] span, [data-testid="stSidebarCollapseButton"] svg { color: #CBD5E1 !important; fill: #CBD5E1 !important; }
/* Widget labels, captions and radio/checkbox option text (#2, #3).
   v3.24.3 (#7): radio/checkbox option labels (e.g. "ခွဲခြမ်းစိတ်ဖြာမှု") get
   line-height 1.8 so Myanmar stacking diacritics are not cramped. */
[data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] label,
[data-testid="stWidgetLabel"] div { color: #F8FAFC !important; line-height: 1.7; }
[data-testid="stCaptionContainer"] { color: #CBD5E1 !important; }
[data-testid="stRadio"] label p, [data-testid="stRadio"] label div,
[data-testid="stRadio"] label span,
[data-testid="stCheckbox"] label p, [data-testid="stCheckbox"] label div,
[data-testid="stCheckbox"] label span { color: #F8FAFC !important; line-height: 1.8; }
[data-testid="stRadio"] label:has(input:checked) p,
[data-testid="stRadio"] label:has(input:checked) div { color: #FFFFFF !important; font-weight: 700; }
[data-testid="stRadio"] input[type="radio"], [data-testid="stCheckbox"] input[type="checkbox"] { accent-color: #3B82F6; }
/* Select / dropdown controls + their menus (#4). Streamlit 1.67 removed
   BaseWeb entirely — every div[data-baseweb="..."] selector in v3.24.1
   matched NOTHING, which is why fields and menus stayed white in Dark.
   Verified against the 1.67 frontend bundle: a selectbox field is the div
   that directly holds the <input> inside [data-testid="stSelectbox"]; a
   multiselect field is the same inside [data-testid="stMultiSelect"]; the
   opened menus are [data-testid="stSelectboxVirtualDropdown"] and
   [data-testid="stMultiSelectDropdown"] (position: fixed, bg from the
   theme). The old BaseWeb popover/tooltip/calendar/tag rules are dropped —
   they matched nothing and the app uses no popovers. */
[data-testid="stSelectbox"] div:has(> input),
[data-testid="stMultiSelect"] div:has(> input) { background-color: #0F172A !important; color: #F8FAFC !important; border-color: #334155 !important; }
[data-testid="stSelectbox"] input, [data-testid="stMultiSelect"] input { color: #F8FAFC !important; }
[data-testid="stSelectbox"] input::placeholder, [data-testid="stMultiSelect"] input::placeholder { color: #94A3B8 !important; opacity: 1 !important; }
[data-testid="stSelectbox"] button, [data-testid="stMultiSelect"] button { color: #F8FAFC !important; }
[data-testid="stSelectbox"] svg, [data-testid="stMultiSelect"] svg { fill: #CBD5E1 !important; }
/* v3.24.3 (#1): the multiselect's WHITE CHEVRON PATCH. The v3.24.2 rule
   themed only div:has(> input) — inside a multiselect that is the TAGS
   container, not the field itself. Verified against the 1.67 bundle
   (Multiselect chunk): the field container is the div directly holding
   [data-testid="stMultiSelectTagsContainer"], with backgroundColor
   secondaryBg (white in the light theme Streamlit runs) and the "Open"
   button (aria-label="Open", transparent background) sitting on it at the
   right end — so the tags area went dark while the strip behind the chevron
   stayed white. Theme the field container itself dark, and keep the Open
   button transparent with a light chevron so the patch disappears. */
[data-testid="stMultiSelect"] div:has(> [data-testid="stMultiSelectTagsContainer"]) { background-color: #0F172A !important; border-color: #334155 !important; }
[data-testid="stMultiSelect"] button[aria-label="Open"] { background-color: transparent !important; color: #F8FAFC !important; }
[data-testid="stMultiSelect"] button[aria-label="Open"] span { color: #F8FAFC !important; }
[data-testid="stSelectboxVirtualDropdown"],
[data-testid="stMultiSelectDropdown"] { background-color: #1E293B !important; border: 1px solid #334155 !important; }
[data-testid="stSelectboxVirtualDropdown"] ul, [data-testid="stMultiSelectDropdown"] ul,
[data-testid="stSelectboxVirtualDropdown"] li, [data-testid="stMultiSelectDropdown"] li,
[data-testid="stSelectboxVirtualDropdown"] div, [data-testid="stMultiSelectDropdown"] div { color: #F8FAFC !important; background-color: #1E293B !important; }
[data-testid="stSelectboxVirtualDropdown"] li:hover, [data-testid="stSelectboxVirtualDropdown"] li[aria-selected="true"],
[data-testid="stMultiSelectDropdown"] li:hover, [data-testid="stMultiSelectDropdown"] li[aria-selected="true"] { background-color: #334155 !important; }
[data-testid="stSelectboxVirtualDropdown"] [data-hovered], [data-testid="stMultiSelectDropdown"] [data-hovered],
[data-testid="stSelectboxVirtualDropdown"] [data-focused], [data-testid="stMultiSelectDropdown"] [data-focused] { background-color: #334155 !important; }
/* Search + all text placeholders (#8). */
input::placeholder, textarea::placeholder { color: #94A3B8 !important; opacity: 1 !important; }
/* File uploader: Browse button + dropzone text + size caption (#5). */
[data-testid="stFileUploader"] section { background: #0F172A !important; border-color: #475569 !important; color: #F8FAFC !important; }
[data-testid="stFileUploader"] section p, [data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section small, [data-testid="stFileUploader"] section div { color: #F8FAFC !important; }
[data-testid="stFileUploader"] section small { color: #CBD5E1 !important; }
[data-testid="stFileUploader"] button { background: #1E293B !important; color: #F8FAFC !important; border: 1px solid #475569 !important; }
[data-testid="stFileUploader"] button:hover { background: #334155 !important; }
[data-testid="stFileUploader"] button p, [data-testid="stFileUploader"] button span { color: #F8FAFC !important; }
/* Info / alert boxes (#6). Streamlit 1.67: the outer box is
   [data-testid="stAlertContainer"] — the old [data-testid="stAlert"]
   matched nothing, so v3.24.1 flipped the text white while the box stayed
   white (invisible). Background AND text are themed together here. */
[data-testid="stAlertContainer"] { background-color: #1E293B !important; border: 1px solid #334155 !important; }
[data-testid="stAlertContainer"] > div { background-color: #1E293B !important; }
[data-testid="stAlertContainer"] p, [data-testid="stAlertContainer"] div, [data-testid="stAlertContainer"] span, [data-testid="stAlertContainer"] li { color: #F8FAFC !important; }
[data-testid="stAlertContainer"] a { color: #93C5FD !important; }
[data-testid="stAlertContentInfo"], [data-testid="stAlertContentSuccess"],
[data-testid="stAlertContentWarning"], [data-testid="stAlertContentError"] { color: #F8FAFC !important; }
/* Audio recorder widget (#12). Streamlit 1.67: [data-testid="stAudioInput"]
   is the OUTER wrapper — the white box is the inner container div, which
   the old rule never reached. Theme the inner divs + the timer text.
   NOT fixable via CSS: the waveform itself is a canvas drawn by WaveSurfer
   with colours passed as JS props from the (light) Streamlit theme
   (waveColor = primary, playback colour = fadedText40 on secondaryBg), so
   its dots stay light-grey; the timer/buttons/container are themed here. */
[data-testid="stAudioInput"] { border: 1px solid #334155 !important; border-radius: 8px; }
[data-testid="stAudioInput"] div { background-color: #1E293B !important; }
[data-testid="stAudioInput"] p, [data-testid="stAudioInput"] span { color: #F8FAFC !important; }
/* v3.24.3 (#2): the TIMER WHITE RECTANGLE. The v3.24.2 rule set only the
   timer's text colour — but the timer (stAudioInputWaveformTimeCode) is a
   <span> with its OWN backgroundColor: secondaryBg (verified in the 1.67
   AudioInput chunk: the ge component is i(`span`, ...) with
   backgroundColor: e.colors.secondaryBg), and the broad div rule never
   matched a span — so the digits went white on a white box. Theme the
   span's background dark as well as its text. */
[data-testid="stAudioInputWaveformTimeCode"] { color: #F8FAFC !important; background-color: #1E293B !important; }
[data-testid="stAudioInput"] button { background-color: #334155 !important; color: #F8FAFC !important; border-color: #475569 !important; }
[data-testid="stAudioInput"] button svg { fill: #F8FAFC !important; }
/* Tabs, expander headers, number-input steppers, date/time inputs. */
[data-testid="stTabs"] button[role="tab"] { color: #CBD5E1 !important; }
[data-testid="stTabs"] button[aria-selected="true"] { color: #60A5FA !important; }
[data-testid="stExpander"] summary p, [data-testid="stExpander"] summary div { color: #F8FAFC !important; }
[data-testid="stNumberInput"] button { background: #1E293B !important; color: #F8FAFC !important; }
[data-testid="stDateInput"] input, [data-testid="stTimeInput"] input { background: #0F172A !important; color: #F8FAFC !important; border-color: #334155 !important; }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
# Theme variables override CUSTOM_CSS's :root (same names, later in the
# cascade) — one place decides every surface/text/accent colour.
_theme_root = "<style>\n:root {\n" + "\n".join(
    f"  --rf-{_k}: {_v};" for _k, _v in _THEME.items()
) + "\n}\n</style>"
st.markdown(_theme_root, unsafe_allow_html=True)
if _THEME_NAME == "dark":
    st.markdown(DARK_CSS, unsafe_allow_html=True)

# v3.24.4 (#4): make sure Noto Sans Myanmar actually LOADS. The font stack
# and the @import in CUSTOM_CSS look right on paper, yet bold Myanmar text
# (the "မြန်မာ" button, "တင်ပြမှု" heading, "ခွဲခြမ်းစိတ်ဖြာမှု" radio
# labels — all weight 600/700) still renders with stacked/medial marks
# piled up, while regular-weight Myanmar is clean, and raising line-height
# to 1.8 had no visible effect. That pattern points at the FONT, not the
# spacing: if Noto Sans Myanmar never loads, the text falls back to the
# system Myanmar font, whose bold face positions stacked marks tightly.
# The @import lives inside a <style> block injected through st.markdown,
# which is sanitised — the import may never reach the browser. So load the
# same Google Fonts stylesheet a second way: a tiny script (st.html with
# unsafe_allow_javascript=True — the Python API DOES expose it in 1.67,
# html.py:52; the frontend re-creates <script> elements so they run in the
# main document, not an iframe) that appends a <link> to document.head.
# A script-created <link> is not subject to HTML sanitising, so it loads
# even if the @import is stripped. The URL is identical to the @import, so
# the browser de-duplicates it if the @import did load. Runs once per
# session (the <link> stays in <head> after this element unmounts).
# If the marks are STILL tight after this, check in the live browser whether
# fonts.googleapis.com is reachable at all (computed font-family /
# document.fonts) — a blocked font CDN would need a bundled font instead.
if "_font_link_loaded" not in st.session_state:
    st.session_state["_font_link_loaded"] = True
    st.html(
        """<script>(function(){if(!document.querySelector('link[data-rf-font]')){var l=document.createElement('link');l.rel='stylesheet';l.setAttribute('data-rf-font','1');l.href='https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=Noto+Sans+Myanmar:wght@400;500;600;700;800&display=swap';document.head.appendChild(l);}})();</script>""",
        unsafe_allow_javascript=True,
    )

# --- Shell (v3.4): 5-stage nav, one Project = one container. Modern, professional,
# minimalist: white, one blue accent, generous spacing, no extra decoration.
STAGES = [
    ("1 · Client Brief", "Client Brief"),
    ("2 · Research Design", "Research Design"),
    ("3 · Fieldworks", "Fieldworks"),
    ("4 · Data Cleansing", "Data Cleansing"),
    ("5 · Analysis & Reports", "Market Analysis"),
]
PAGE_STAGE = {
    "Client Brief": 1,
    "Research Design": 2,
    "Fieldworks": 3,
    "Upload & Transcribe": 4, "Batch Processing": 4, "Upload & Split": 4,
    "Transcripts": 4, "Data Cleansing": 4,
    "Market Analysis": 5, "Quantitative Analysis": 5, "Reports": 5, "Report Editor": 5,
    "Presentation": 5, "Insights": 5,
}

if "page" not in st.session_state:
    st.session_state["page"] = "Dashboard"

st.sidebar.markdown(
    '<div class="rf-logo">'
    '<svg width="32" height="32" viewBox="0 0 32 32" fill="none">'
    '<rect x="2" y="19" width="6" height="10" rx="2" fill="#2563EB"/>'
    '<rect x="11" y="12" width="6" height="17" rx="2" fill="#0891B2"/>'
    '<rect x="20" y="5" width="6" height="24" rx="2" fill="#0D9488"/>'
    '</svg>'
    '<div><div class="rf-logo-tx">ResearchFlow</div>'
    '<div class="rf-logo-sub">Market Research AI</div></div></div>',
    unsafe_allow_html=True,
)

# v3.18 language switch. The value is kept in session_state["lang"] (a plain key,
# not a widget key) so it survives the login screen and page changes.
# v3.24: language is chosen with two segmented buttons below (the old radio
# rendered the "မြန်မာ" label with overlapping/stacked glyphs in the sidebar).
_LANG_CODES = ["EN", "MM"]
_LANG_NAMES = {"EN": "English", "MM": "မြန်မာ"}

# Project selector: list once per session (Drive call), refresh after creating one.
if "project_list" not in st.session_state:
    st.session_state["project_list"] = list_projects()
projects = st.session_state["project_list"]
current = st.session_state.get("active_project", "Default")
if current not in projects:
    projects = [current] + projects
if "project_select_target" in st.session_state:
    st.session_state["project_select"] = st.session_state.pop("project_select_target")
# v3.24: the "📁 Default" label is gone — the sidebar shows ONLY the
# "+ New project" button (the current project already shows in the stage
# header and in the top bar). The switch dropdown still appears only when
# more than one project exists.
selected = current
if len(projects) > 1:
    selected = st.sidebar.selectbox(
        T("Switch project", "Project ပြောင်းမယ်"), projects,
        index=projects.index(current),
        key="project_select",
    )
if selected != current:
    st.session_state["active_project"] = selected
    st.session_state["drive_synced"] = False
    st.rerun()
refresh_project_dirs(selected)

# v3.21: a single visible button — it opens the New Project page, where the
# details are entered. No collapsed expander form in the sidebar any more.
if st.sidebar.button(T("＋ New project", "＋ Project အသစ်"), key="new_project_btn", use_container_width=True):
    st.session_state["page"] = "New Project"
    st.rerun()

# v3.24: language + theme as segmented buttons, below the Project area,
# above the nav. Plain session-state keys + rerun (no widget keys that could
# go stale); both choices are saved to .ui_prefs.json so they survive a reload.
st.sidebar.caption(T("Language", "ဘာသာစကား"))
_lc1, _lc2 = st.sidebar.columns(2)
with _lc1:
    if st.button("English", key="lang_en", type="primary" if st.session_state.get("lang", "EN") == "EN" else "secondary", use_container_width=True):
        st.session_state["lang"] = "EN"
        _save_ui_prefs({"theme": st.session_state.get("theme", "light"), "lang": "EN"})
        st.rerun()
with _lc2:
    if st.button("မြန်မာ", key="lang_mm", type="primary" if st.session_state.get("lang", "EN") == "MM" else "secondary", use_container_width=True):
        st.session_state["lang"] = "MM"
        _save_ui_prefs({"theme": st.session_state.get("theme", "light"), "lang": "MM"})
        st.rerun()
st.sidebar.caption(T("Theme", "အပြင်အဆင်"))
_tc1, _tc2 = st.sidebar.columns(2)
with _tc1:
    if st.button("☀️ Light", key="theme_light", type="primary" if st.session_state.get("theme", "light") == "light" else "secondary", use_container_width=True):
        st.session_state["theme"] = "light"
        _save_ui_prefs({"theme": "light", "lang": st.session_state.get("lang", "EN")})
        st.rerun()
with _tc2:
    if st.button("🌙 Dark", key="theme_dark", type="primary" if st.session_state.get("theme", "light") == "dark" else "secondary", use_container_width=True):
        st.session_state["theme"] = "dark"
        _save_ui_prefs({"theme": "dark", "lang": st.session_state.get("lang", "EN")})
        st.rerun()

_PAGE_ICONS = {
    "Dashboard": ":material/home:", "Client Brief": ":material/description:",
    "Research Design": ":material/design_services:",
    "Upload & Transcribe": ":material/database:", "Batch Processing": ":material/layers:",
    "Upload & Split": ":material/content_cut:",
    "Fieldworks": ":material/groups:", "Transcripts": ":material/article:",
    "Data Cleansing": ":material/cleaning_services:",
    "Market Analysis": ":material/bar_chart:", "Quantitative Analysis": ":material/pie_chart:",
    "Reports": ":material/assignment:",
    "Report Editor": ":material/edit:", "Presentation": ":material/slideshow:",
    "Insights": ":material/lightbulb:",
    "Prompt Library": ":material/library_books:",
}
_PAGE_ORDER = [
    "Dashboard", "Client Brief", "Research Design",
    "Fieldworks",
    "Upload & Transcribe", "Batch Processing", "Upload & Split",
    "Transcripts", "Data Cleansing",
    "Market Analysis", "Quantitative Analysis", "Reports", "Report Editor",
    "Presentation", "Insights", "Prompt Library",
]

_PAGE_MM = {
    "Dashboard": "ပင်မစာမျက်နှာ", "Client Brief": "Client Brief",
    "Research Design": "သုတေသန ဒီဇိုင်း",
    "Upload & Transcribe": "Upload နဲ့ Transcript လုပ်ခြင်း", "Batch Processing": "Batch လုပ်ဆောင်ခြင်း",
    "Upload & Split": "Upload နဲ့ ခွဲခြင်း", "Fieldworks": "ကွင်းဆင်းလုပ်ငန်း",
    "Transcripts": "Transcript များ", "Data Cleansing": "ဒေတာ သန့်စင်ခြင်း",
    "Market Analysis": "ဈေးကွက် ခွဲခြမ်းစိတ်ဖြာမှု",
    "Quantitative Analysis": "ပမာဏ ခွဲခြမ်းစိတ်ဖြာမှု",
    "Reports": "Report များ", "Report Editor": "Report ပြင်ဆင်ရန်",
    "Presentation": "တင်ပြမှု", "Prompt Library": "Prompt Library",
    "Insights": "Insights",
}
_STAGE_MM = {
    "Client Brief": "Client Brief", "Research Design": "သုတေသန ဒီဇိုင်း",
    "Fieldworks": "ကွင်းဆင်းလုပ်ငန်း", "Data Cleansing": "ဒေတာ သန့်စင်ခြင်း",
    "Analysis & Reports": "ခွဲခြမ်းစိတ်ဖြာမှုနဲ့ Report", "Presentation": "တင်ပြမှု",
}


def page_label(page_id):
    return T(page_id, _PAGE_MM.get(page_id))


def _sl(label):
    """Translate a stage label like '1 · Client Brief' (number kept)."""
    num, nm = label.split(" · ", 1)
    return f"{num} · {T(nm, _STAGE_MM.get(nm))}"


# v3.23 nav: the sidebar is Dashboard + 5 stage entries + Prompt Library.
# The old pages are kept (same internal page ids, same code); each stage groups
# them into sections / steps shown at the top of the stage page.
_STAGE_NAME_EN = {1: "Client Brief", 2: "Research Design", 3: "Fieldworks",
                  4: "Data Cleansing", 5: "Analysis & Reports"}
_STAGE_HOME = {1: "Client Brief", 2: "Research Design", 3: "Fieldworks",
               4: "Upload & Transcribe", 5: "Market Analysis"}
# stage -> [(section EN, section MM, [(page id, step EN, step MM), ...]), ...]
# A stage that is not listed here has one page and no section bar.
_STAGE_SECTIONS = {
    4: [
        ("Data Preparation", "ဒေတာ ပြင်ဆင်ခြင်း", [
            ("Upload & Transcribe", "Upload & Transcribe", "Upload နဲ့ Transcript"),
            ("Batch Processing", "Batch", "Batch"),
            ("Upload & Split", "Split", "ခွဲခြင်း"),
            ("Transcripts", "Transcripts & Merge", "Transcript နဲ့ Merge"),
        ]),
        ("Quality Control", "Quality Control", [
            ("Data Cleansing", "Quality Control", "Quality Control"),
        ]),
    ],
    5: [
        ("Analysis", "ခွဲခြမ်းစိတ်ဖြာမှု", [
            ("Market Analysis", "Qualitative", "Qualitative"),
            ("Quantitative Analysis", "Quantitative", "Quantitative"),
        ]),
        ("Insights", "Insights", [
            ("Insights", "Insights", "Insights"),
        ]),
        ("Report & Presentation", "Report နဲ့ တင်ပြမှု", [
            ("Reports", "Reports", "Report များ"),
            ("Report Editor", "Editor", "ပြင်ဆင်ရန်"),
            ("Presentation", "Presentation", "တင်ပြမှု"),
        ]),
    ],
}
assert all(p in PAGE_STAGE or p in ("Dashboard", "Prompt Library") for p in _PAGE_ORDER)
assert all(PAGE_STAGE[p] == _n for _n, _secs in _STAGE_SECTIONS.items()
           for _sec in _secs for p, _a, _b in _sec[2])


def _section_index(stage_no, page_id):
    for _i, _sec in enumerate(_STAGE_SECTIONS.get(stage_no, [])):
        if any(p == page_id for p, _a, _b in _sec[2]):
            return _i
    return 0


def _go_section(stage_no):
    """Section bar clicked: open the last page used in that section (or its first)."""
    _i = st.session_state.get(f"sec_radio_{stage_no}", 0)
    _ids = [p for p, _a, _b in _STAGE_SECTIONS[stage_no][_i][2]]
    _last = st.session_state.get(f"sec_last_{stage_no}_{_i}")
    st.session_state["page"] = _last if _last in _ids else _ids[0]


def _go_step(stage_no, sec_i):
    """Step bar clicked: open that step's page."""
    _ids = [p for p, _a, _b in _STAGE_SECTIONS[stage_no][sec_i][2]]
    st.session_state["page"] = _ids[st.session_state.get(f"step_radio_{stage_no}_{sec_i}", 0)]


def _nav_btn(label, key, active, target):
    if st.sidebar.button(
        label, key=key, use_container_width=True,
        type="primary" if active else "secondary",
    ):
        st.session_state["page"] = target
        st.rerun()


_cur_stage = PAGE_STAGE.get(st.session_state["page"])
_nav_btn(page_label("Dashboard"), "nav_Dashboard", st.session_state["page"] == "Dashboard", "Dashboard")
for _n in range(1, 6):
    _nm = _STAGE_NAME_EN[_n]
    _nav_btn(
        f"{_n}  {T(_nm, _STAGE_MM.get(_nm))}", f"nav_stage{_n}", _cur_stage == _n,
        st.session_state.get(f"stage_last_{_n}", _STAGE_HOME[_n]),
    )
st.sidebar.divider()
_nav_btn(page_label("Prompt Library"), "nav_Prompt Library", st.session_state["page"] == "Prompt Library", "Prompt Library")

_storage_label = (
    T("💾 Storage: Google Drive (OAuth)", "💾 သိမ်းဆည်းမှု: Google Drive (OAuth)") if DRIVE_AUTH == "oauth"
    else T("💾 Storage: Google Drive", "💾 သိမ်းဆည်းမှု: Google Drive") if DRIVE_AUTH == "service"
    else T("💾 Storage: Local (PC)", "💾 သိမ်းဆည်းမှု: Local (PC)")
)
st.sidebar.caption(f"{_storage_label} · v{APP_VERSION}")

_pending_q = [i for i in _load_sync_queue()
              if i.get("project", "Default") == st.session_state.get("active_project", "Default")]
if _pending_q:
    st.sidebar.warning(T(f"⏳ Sync queue: {len(_pending_q)} file(s) waiting to upload to Drive", f"⏳ Sync queue: {len(_pending_q)} ဖိုင် Drive ကို တင်ဖို့ စောင့်နေတယ်"))
    if st.sidebar.button(T("↻ Retry upload", "↻ ပြန်တင်မယ်"), key="retry_queue"):
        _done, _left = retry_sync_queue()
        if _left:
            st.sidebar.warning(T(f"{_done} file(s) uploaded, {_left} left", f"{_done} ဖိုင် တင်ပြီးပြီ၊ {_left} ဖိုင် ကျန်သေးတယ်"))
        else:
            st.sidebar.success(T(f"{_done} file(s) uploaded ✓", f"{_done} ဖိုင် တင်ပြီးပြီ ✓"))
        st.rerun()

st.sidebar.markdown(
    '<div class="rf-user"><div class="rf-user-av">S</div>'
    '<div><div class="rf-user-nm">Soe Moe</div>'
    '<div class="rf-user-rl">Research Team</div></div></div>',
    unsafe_allow_html=True,
)

if _MY_EMAIL:
    if st.sidebar.button(T("Log out", "Log out ထွက်မယ်"), key="logout_btn"):
        st.logout()

page = st.session_state["page"]

# v3.24 top bar: global search + current project + user. No notification
# bell — there is no notification system, so nothing fake is shown.
_tb_search, _tb_proj, _tb_user = st.columns([5.6, 1.8, 1.6])  # v3.24.1 (#10): wider chip columns so project/user pills are never truncated
with _tb_search:
    _search_q = st.text_input(
        "Search", key="global_search", label_visibility="collapsed",
        placeholder=T("🔍 Search projects, files, or features...", "🔍 Project, ဖိုင်၊ feature ရှာပါ..."),
    )
with _tb_proj:
    st.markdown(f'<div class="rf-chip">📁 {html.escape(current)}</div>', unsafe_allow_html=True)
with _tb_user:
    st.markdown('<div class="rf-chip">👤 Soe Moe</div>', unsafe_allow_html=True)
if _search_q:
    _q = _search_q.strip().lower()
    _hits = []
    for _p in st.session_state.get("project_list", ["Default"]):
        if _q in _p.lower():
            _hits.append((f"📁 Project: {_p}", _p, "project"))
    _scan = ((AUDIO_DIR, "Upload & Transcribe"), (OUTPUT_DIR, "Transcripts"),
             (ANALYSIS_DIR, "Market Analysis"), (REPORTS_DIR, "Reports"))
    for _folder, _target in _scan:
        for _f in _folder.glob("*"):
            if _f.is_file() and _q in _f.name.lower():
                _hits.append((f"📄 {_f.name}", _target, "file"))
    if _hits:
        for _i, (_label, _target, _kind) in enumerate(_hits[:6]):
            if st.button(_label, key=f"search_hit_{_i}"):
                if _kind == "project":
                    st.session_state["project_select_target"] = _target
                else:
                    st.session_state["page"] = _target
                st.rerun()
    else:
        st.caption(T("No matches in this project.", "ဒီ project ထဲမှာ မတွေ့ဘူး။"))

# Navigation lives in the sidebar only (the top stepper row was removed in v3.9).

def _data_prep_counts():
    """Real counts behind Stage 4 (Data Preparation): audio, transcripts,
    merged full transcripts, split chunks, still-untranscribed audio."""
    audio_files = [f for f in AUDIO_DIR.glob("*") if f.is_file()]
    audio_n = len(audio_files)
    tr_n = len(get_transcript_files())
    merged_n = len(list(MERGED_DIR.glob("*_full.txt")))
    chunk_n = len([f for f in AUDIO_DIR.glob("*(*).wav")]
                  + [f for f in AUDIO_DIR.glob("*(*).mp3")]
                  + [f for f in AUDIO_DIR.glob("*(*).m4a")])
    pending_n = len([f for f in audio_files
                     if f.suffix.lower() in {".mp3", ".wav", ".m4a"}
                     and not (OUTPUT_DIR / (f.stem + ".txt")).exists()])
    return audio_n, tr_n, merged_n, chunk_n, pending_n


def _stage_item(en, mm, state):
    """One workflow sub-item: (English, Myanmar, state) with state 'done',
    'todo' (feature exists, no data yet) or 'coming' (not built yet)."""
    return (en, mm, state)


def stage_status():
    """(label, description, done, sub_items) per stage, from real files only.
    v3.21: each stage lists its own sub-items with per-item status, so
    Fieldworks (Stage 3) and Data Cleansing (Stage 4) no longer look alike —
    Stage 3 is collecting data, Stage 4 is preparing it. A stage counts as
    done when its output files exist in this project."""
    audio_n, tr_n, merged_n, chunk_n, pending_n = _data_prep_counts()
    an_n = len(list(ANALYSIS_DIR.glob("*_analysis.txt")))
    rep_n = len(list(REPORTS_DIR.glob("*.pptx"))) + len(list(REPORTS_DIR.glob("*.docx")))
    pptx_n = len(list(REPORTS_DIR.glob("*.pptx")))
    quant_n = len(list(ANALYSIS_DIR.glob("quant_*.txt")))
    multi_n = len(list(ANALYSIS_DIR.glob("multi_*_analysis.txt")))
    note_n = len([f for f in NOTES_DIR.glob("*")
                  if f.is_file() and f.name not in {"brief.json", "project.json"}])
    any_files = audio_n + tr_n + an_n + rep_n > 0
    brief = load_brief()
    return [
        (_sl("1 · Client Brief"),
         T("Define objectives & scope", "ရည်ရွယ်ချက်နဲ့ scope သတ်မှတ်ခြင်း"),
         brief.get("approval") == "Approved",
         [
             _stage_item("Client info", "Client info", "done" if brief.get("client_name") else "todo"),
             _stage_item("Project info", "Project info", "done" if brief.get("project_code") or brief.get("background") else "todo"),
             _stage_item("Objectives", "ရည်ရွယ်ချက်များ", "done" if brief.get("objectives") else "todo"),
             _stage_item("Target audience", "ပစ်မှတ်အုပ်စု", "done" if brief.get("target_audience") else "todo"),
             _stage_item("Scope", "Scope", "done" if brief.get("scope") else "todo"),
             _stage_item("Timeline", "အချိန်ဇယား", "done" if brief.get("timeline_start") and brief.get("timeline_end") else "todo"),
             _stage_item("Deliverables", "ထုတ်ပေးမည့်အရာများ", "done" if brief.get("deliverables") else "todo"),
             _stage_item("Approval", "အတည်ပြုချက်", "done" if brief.get("approval") == "Approved" else "todo"),
         ]),
        (_sl("2 · Research Design"),
         T("Design methodology, sampling & instruments", "နည်းလမ်း၊ sampling နဲ့ tool တွေ ဒီဇိုင်းဆွဲခြင်း"),
         False,
         [
             _stage_item("Approach (Qual / Quant / Mixed)", "Approach (Qual / Quant / Mixed)", "coming"),
             _stage_item("IDI / FGD / survey", "IDI / FGD / survey", "coming"),
             _stage_item("Sampling", "Sampling", "coming"),
             _stage_item("DG / questionnaire / screener", "DG / questionnaire / screener", "coming"),
             _stage_item("Fieldwork plan", "ကွင်းဆင်းအစီအစဉ်", "coming"),
         ]),
        (_sl("3 · Fieldworks"),
         T("Collect data in the field", "ကွင်းဆင်းပြီး ဒေတာ ကောက်ယူခြင်း"),
         audio_n > 0,
         [
             _stage_item("Form", "Form", "coming"),
             _stage_item("F2F discussion", "F2F discussion", "done" if audio_n > 0 else "todo"),
             _stage_item("Submit data", "ဒေတာ တင်ခြင်း", "coming"),
             _stage_item("Record", "အသံသွင်းခြင်း", "done" if audio_n > 0 else "todo"),
             _stage_item("Interview management", "Interview စီမံခန့်ခွဲခြင်း", "coming"),
             _stage_item("Survey data", "Survey ဒေတာ", "done" if quant_n > 0 else "todo"),
             _stage_item("Field notes", "ကွင်းဆင်းမှတ်စု", "done" if note_n > 0 else "todo"),
             _stage_item("Tracking", "ခြေရာခံခြင်း", "done" if any_files else "todo"),
         ]),
        (_sl("4 · Data Cleansing"),
         T("Prepare data: upload, transcribe, merge + quality control", "ဒေတာ ပြင်ဆင်ခြင်း: upload, transcribe, merge + အရည်အသွေး စစ်ဆေးခြင်း"),
         tr_n > 0,
         [
             _stage_item("Upload", "Upload", "done" if audio_n > 0 else "todo"),
             _stage_item("Batch", "Batch", "done" if audio_n > 0 and pending_n == 0 else "todo"),
             _stage_item("Split", "ခွဲခြင်း", "done" if chunk_n > 0 else "todo"),
             _stage_item("Transcribe", "Transcript လုပ်ခြင်း", "done" if tr_n > 0 else "todo"),
             _stage_item("Merge", "ပေါင်းခြင်း", "done" if merged_n > 0 else "todo"),
             _stage_item("QC: check", "QC: စစ်ဆေးခြင်း", "coming"),
             _stage_item("QC: correct", "QC: ပြင်ဆင်ခြင်း", "coming"),
             _stage_item("QC: status", "QC: အခြေအနေ", "coming"),
             _stage_item("QC: audit trail", "QC: audit trail", "coming"),
         ]),
        (_sl("5 · Analysis & Reports"),
         T("Generate insights & reports", "Insight နဲ့ report ထုတ်လုပ်ခြင်း"),
         an_n > 0 or rep_n > 0,
         [
             _stage_item("Qual analysis", "Qual ခွဲခြမ်းစိတ်ဖြာမှု", "done" if an_n > 0 else "todo"),
             _stage_item("Quant analysis", "Quant ခွဲခြမ်းစိတ်ဖြာမှု", "done" if quant_n > 0 else "todo"),
             _stage_item("Insights (Qual + Quant)", "Insights (Qual + Quant)", "done" if multi_n > 0 else "todo"),
             _stage_item("Report (PPTX / DOCX / PDF)", "Report (PPTX / DOCX / PDF)", "done" if rep_n > 0 else "todo"),
             _stage_item("Presentation (PPTX export)", "တင်ပြမှု (PPTX export)", "done" if pptx_n > 0 else "todo"),
         ]),
    ]


def _stage_badge(stage):
    """Honest badge from a stage_status() tuple: Done ONLY when every
    sub-item is done (no 'todo' and no 'coming'); In progress when at least
    one sub-item is done; otherwise Not started. (v3.24 fix: stages 3/4/5
    used to show Done from file counts while sub-items were still coming.)"""
    _label, _desc, _done_flag, _subs = stage
    if _subs and all(_s == "done" for _en, _mm, _s in _subs):
        return (T("✅ Done", "✅ ပြီးပြီ"), "done")
    if any(_s == "done" for _en, _mm, _s in _subs):
        return (T("● In progress", "● လုပ်နေဆဲ"), "wip")
    return (T("○ Not started", "○ မစရသေးဘူး"), "todo")


def recent_activity(limit=5):
    """Most recent files across the project folders: (time, file, activity, status)."""
    rows = []
    for folder, kind in ((AUDIO_DIR, "audio"), (OUTPUT_DIR, "transcript"),
                         (ANALYSIS_DIR, "analysis"), (REPORTS_DIR, "report")):
        for f in folder.glob("*"):
            if not f.is_file() or f.name == "failed_files.txt":
                continue
            rows.append((f.stat().st_mtime, f.name, kind))
    rows.sort(key=lambda r: r[0], reverse=True)
    out = []
    for mtime, name, kind in rows[:limit]:
        dt = datetime.datetime.fromtimestamp(mtime, tz=ZoneInfo("Asia/Rangoon"))
        label = {
            "audio": T("Audio file added", "Audio ဖိုင် ထည့်ပြီး"),
            "transcript": T("Transcription completed", "Transcription ပြီးပြီ"),
            "analysis": T("Analysis saved", "Analysis သိမ်းပြီး"),
            "report": T("Report exported", "Report ထုတ်ပြီး"),
        }[kind]
        out.append((dt.strftime("%b %d, %Y %I:%M %p"), name, label, T("Completed", "ပြီးပြီ")))
    return out


# Drive is the source of truth: sync once per session, then pages work on local files.
if DRIVE_MODE and not st.session_state.get("drive_synced"):
    with st.spinner(T("Syncing with Google Drive...", "Google Drive နဲ့ sync လုပ်နေပါတယ်...")):
        sync_from_drive()
        retry_sync_queue()
    st.session_state["drive_synced"] = True

# v3.23 stage header + section / step bars (shown on stage pages only).
_stage_no = PAGE_STAGE.get(page)
if _stage_no:
    st.session_state[f"stage_last_{_stage_no}"] = page
    _stg = stage_status()[_stage_no - 1]
    _sh_txt, _sh_cls = _stage_badge(_stg)
    _sh_nm = _STAGE_NAME_EN[_stage_no]
    st.markdown(
        '<div class="rf-sh"><div><div class="rf-sh-k">' + T("Stage", "အဆင့်") + f" {_stage_no} / 5</div>"
        '<div class="rf-sh-t">' + html.escape(T(_sh_nm, _STAGE_MM.get(_sh_nm))) + "</div>"
        '<div class="rf-sh-p">📁 ' + html.escape(current_project()) + "</div></div>"
        '<span class="rf-sh-b ' + _sh_cls + '">' + _sh_txt + "</span></div>",
        unsafe_allow_html=True,
    )
    _secs = _STAGE_SECTIONS.get(_stage_no)
    if _secs:
        _si = _section_index(_stage_no, page)
        if len(_secs) > 1:
            _sec_key = f"sec_radio_{_stage_no}"
            st.session_state[_sec_key] = _si
            st.radio(
                T("Section", "အပိုင်း"), list(range(len(_secs))),
                format_func=lambda _i, _secs=_secs: T(_secs[_i][0], _secs[_i][1]),
                horizontal=True, key=_sec_key, label_visibility="collapsed",
                on_change=_go_section, args=(_stage_no,),
            )
        _steps = _secs[_si][2]
        if len(_steps) > 1:
            _step_key = f"step_radio_{_stage_no}_{_si}"
            st.session_state[_step_key] = [p for p, _a, _b in _steps].index(page)
            st.radio(
                T("Step", "အဆင့်ခွဲ"), list(range(len(_steps))),
                format_func=lambda _i, _steps=_steps: T(_steps[_i][1], _steps[_i][2]),
                horizontal=True, key=_step_key, label_visibility="collapsed",
                on_change=_go_step, args=(_stage_no, _si),
            )
        st.session_state[f"sec_last_{_stage_no}_{_si}"] = page

# v3.24.4 (#3): scroll the main pane to the top when the page CHANGES.
# Streamlit preserves main's scroll position across reruns, so a page opened
# from a scrolled position used to open half-scrolled (Data Cleansing's
# "Upload & Transcribe" heading below the fold; the Dashboard search box and
# hero heading half-cut under the toolbar on navigation). CSS cannot force a
# scroll-to-top. The v3.24.3 note claimed st.html strips <script> — that is
# WRONG for 1.67: the Python API exposes st.html(body,
# unsafe_allow_javascript=True) (html.py:52) and the frontend Html chunk
# re-creates <script> elements so they execute in the main document (not
# iframed). The script is emitted ONLY on a real page change (tracked in
# session state), never on widget reruns, so interacting with a widget on a
# long page never yanks the user to the top. The 60 ms delay lets Streamlit
# finish committing the new page (and any scroll restoration) first.
# LIMIT: free manual scrolling under the fixed header is normal behaviour —
# content slides behind the opaque toolbar and the user scrolls it back into
# view; that cannot (and should not) be prevented with CSS.
if st.session_state.get("_last_scrolled_page") != page:
    st.session_state["_last_scrolled_page"] = page
    st.html(
        """<script>setTimeout(function(){var m=document.querySelector('[data-testid="stMain"]');if(m){m.scrollTop=0;}window.scrollTo(0,0);},60);</script>""",
        unsafe_allow_javascript=True,
    )

if page == "Dashboard":

    st.markdown(
        '<div class="rf-hero"><h1>' + T("Welcome to", "ကြိုဆိုပါတယ်") + ' <span class="b">Research</span>'
        '<span class="t">Flow</span></h1><p>'
        + T("Transform your interview data into actionable insights.",
            "Interview data တွေကို အသုံးဝင်တဲ့ insight အဖြစ် ပြောင်းလဲပေးပါတယ်။")
        + "</p></div>",
        unsafe_allow_html=True,
    )
    _hc1, _hc2, _hc3 = st.columns([2.2, 2.2, 4.6])  # v3.24.1 (#10): wider hero buttons so labels are never truncated
    with _hc1:
        if st.button(T("▶ Start New Project", "▶ Project အသစ် စမယ်"), key="hero_new_project", type="primary", use_container_width=True):
            st.session_state["page"] = "New Project"
            st.rerun()
    with _hc2:
        if st.button(T("📖 View Guide", "📖 လမ်းညွှန်ကြည့်မယ်"), key="hero_view_guide", use_container_width=True):
            st.session_state["show_guide"] = not st.session_state.get("show_guide", False)
            st.rerun()
    if st.session_state.get("show_guide"):
        st.markdown(
            '<div class="rf-guide"><h3>' + T("📖 How ResearchFlow works", "📖 ResearchFlow အလုပ်လုပ်ပုံ") + '</h3><p>'
            + T("One project flows through 5 stages — brief, design, fieldwork, "
                "cleanse, analyse & report. Files live in Google Drive "
                "(ResearchFlow / project / audio, transcripts, notes, reports) and "
                "this app works on those same files.",
                "Project တစ်ခုက အဆင့် ၅ ဆင့် ဖြတ်သန်းတယ် — brief, ဒီဇိုင်း, ကွင်းဆင်း, "
                "သန့်စင်, ခွဲခြမ်းစိတ်ဖြာပြီး report ထုတ်လုပ်ခြင်း။ ဖိုင်တွေက Google Drive "
                "(ResearchFlow / project / audio, transcripts, notes, reports) ထဲမှာ ရှိပြီး "
                "ဒီ app က အဲဒီဖိုင်တွေပေါ်မှာပဲ အလုပ်လုပ်တယ်။")
            + "</p></div>",
            unsafe_allow_html=True,
        )
        st.markdown(T("**System Architecture**", "**စနစ်ဖွဲ့စည်းပုံ**"))
        st.markdown(
            '<div class="rf-arch">' + T("Research Project Management", "Research Project စီမံခန့်ခွဲမှု")
            + '<small>' + T("Client Brief • Project Setup • Objectives", "Client Brief • Project သတ်မှတ်ခြင်း • ရည်ရွယ်ချက်များ") + '</small></div>'
            '<div class="rf-arch-arrow">↓</div>',
            unsafe_allow_html=True,
        )
        _ac1, _ac2 = st.columns(2)
        with _ac1:
            st.markdown('<div class="rf-arch">' + T("Fieldworks", "ကွင်းဆင်းလုပ်ငန်း") + '<small>' + T("Interviews • Surveys • FGDs", "Interview • Survey • FGD") + '</small></div>', unsafe_allow_html=True)
        with _ac2:
            st.markdown('<div class="rf-arch">' + T("Batch Processing", "Batch လုပ်ဆောင်ခြင်း") + '<small>' + T("Audio • Documents", "Audio • စာရွက်စာတမ်းများ") + '</small></div>', unsafe_allow_html=True)
        st.markdown('<div class="rf-arch-arrow">↓</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="rf-arch">' + T("AI Processing Engine", "AI လုပ်ဆောင်စနစ်")
            + '<small>' + T("Transcription • Data Cleansing", "Transcription • ဒေတာ သန့်စင်ခြင်း") + '</small></div>'
            '<div class="rf-arch-arrow">↓</div>'
            '<div class="rf-arch">' + T("Research Intelligence", "သုတေသန Intelligence")
            + '<small>' + T("Qualitative • Quantitative • Market Analysis", "Qualitative • Quantitative • ဈေးကွက် ခွဲခြမ်းစိတ်ဖြာမှု") + '</small></div>'
            '<div class="rf-arch-arrow">↓</div>'
            '<div class="rf-arch">' + T("AI Reports &amp; Presentation", "AI Report နဲ့ တင်ပြမှု")
            + '<small>' + T("Insights • Recommendations • Client Reports", "Insight • အကြံပြုချက်များ • Client Report များ") + '</small></div>',
            unsafe_allow_html=True,
        )
        # v3.21: the old caption was wrong — surveys ARE in Stage 3
        # ("Survey data"), and Coding / Video are not in the 5-stage table.

    _projects = st.session_state.get("project_list", ["Default"])
    _transcripts = get_transcript_files()
    _reports = list(REPORTS_DIR.glob("*.pptx")) + list(REPORTS_DIR.glob("*.docx"))
    _audio = [f for f in AUDIO_DIR.glob("*") if f.is_file()]
    _cards = [
        ("📁", T("Total Projects", "Project စုစုပေါင်း"), len(_projects), "#2563EB"),
        ("📝", T("Transcripts", "Transcript များ"), len(_transcripts), "#0D9488"),
        ("📊", T("Research Reports", "သုတေသန Report များ"), len(_reports), "#0891B2"),
        ("🎙️", T("Audio Files", "Audio ဖိုင်များ"), len(_audio), "#7C3AED"),
    ]
    _kpi_html = "".join(
        f'<div class="rf-metric"><div class="rf-metric-ico" style="background:{_c}1A">{_ico}</div>'
        f'<div class="rf-metric-label">{_lbl}</div>'
        f'<div class="rf-metric-value">{_val}</div></div>'
        for _ico, _lbl, _val, _c in _cards
    )
    st.markdown(f'<div class="rf-kpirow">{_kpi_html}</div>', unsafe_allow_html=True)

    _stages = stage_status()
    _done = sum(1 for _s in _stages if _stage_badge(_s)[1] == "done")
    _pct = round(_done / len(_stages) * 100)
    _icons = ["📋", "📐", "👥", "🧹", "📊"]
    _step_colors = ["#2563EB", "#0891B2", "#0D9488", "#2563EB", "#0891B2"]

    st.subheader(T("🚀 Research Workflow", "🚀 သုတေသန လုပ်ငန်းစဉ်"))
    st.caption(T("Follow the complete research process from brief to presentation", "Brief ကနေ တင်ပြမှုအထိ သုတေသနလုပ်ငန်းစဉ် အပြည့်အစုံကို လိုက်ကြည့်ပါ"))
    # Current step = first stage that is not Done (computed from real status).
    _cur_step = next((i + 1 for i, _s in enumerate(_stages) if _stage_badge(_s)[1] != "done"), None)
    if _cur_step:
        st.markdown('<div class="rf-current">' + T(f"Current: Step {_cur_step}", f"အခု အဆင့်: {_cur_step}") + '</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="rf-current">' + T("✅ All stages complete", "✅ အဆင့်အားလုံး ပြီးပြီ") + '</div>', unsafe_allow_html=True)
    _flow_cards = []
    for _i, (_label, _desc, _d, _subs) in enumerate(_stages):
        _nm = _label.split(" · ", 1)[1]
        _b_txt, _b_cls = _stage_badge((_label, _desc, _d, _subs))
        _pill_cls = {"done": "pill-done", "wip": "pill-wip", "todo": "pill-todo"}[_b_cls]
        _subs_html = "".join(
            '<div class="sub"><span class="'
            + ("ok" if _s == "done" else "no" if _s == "todo" else "cm")
            + '">'
            + ("✓" if _s == "done" else "○" if _s == "todo" else "–")
            + "</span> " + html.escape(T(_en, _mm))
            + (" <i>(" + T("coming", "မလာသေးဘူး") + ")</i>" if _s == "coming" else "")
            + "</div>"
            for _en, _mm, _s in _subs
        )
        _flow_cards.append(
            f'<div class="rf-step"><div class="ico" style="background:{_step_colors[_i]}1A">{_icons[_i]}</div>'
            f'<div class="nm">{_i + 1}. {html.escape(_nm)}</div>'
            f'<div class="ds">{html.escape(_desc)}</div>'
            f'<div class="st"><span class="rf-pill {_pill_cls}">{_b_txt}</span></div>'
            f'<div class="subs">{_subs_html}</div></div>'
        )
    st.markdown('<div class="rf-flow">' + '<div class="rf-flow-arrow">→</div>'.join(_flow_cards) + '</div>', unsafe_allow_html=True)

    st.subheader(T("💡 Project Progress", "💡 Project တိုးတက်မှု"))
    _pc1, _pc2 = st.columns([1, 2])
    with _pc1:
        st.markdown(
            f'<div class="rf-ring" style="--p:{_pct}"><span>{_pct}%</span></div>',
            unsafe_allow_html=True,
        )
        st.caption(T(f"{_done} of {len(_stages)} stages complete", f"အဆင့် {len(_stages)} ခုအနက် {_done} ခု ပြီးပြီ"))
    with _pc2:
        _s1, _s2 = st.columns(2)
        for _i, (_label, _desc, _d, _subs) in enumerate(_stages):
            with (_s1 if _i < 3 else _s2):
                _mk = {"done": "✅ ", "wip": "● ", "todo": "○ "}[_stage_badge((_label, _desc, _d, _subs))[1]]
                st.write(_mk + _label)

    # v3.24: Recent Activities + Quick Actions + System Status, side by side.
    # Every row is real data only — no fake activities, no fake deltas.
    _ra, _qa, _ss = st.columns([1.6, 0.85, 1])  # v3.24.3 (#3): Recent Activity needs more room so the Status column fits; Quick Actions buttons are short and System Status wraps fine in less width
    with _ra:
        with st.container(border=True):
            st.subheader(T("🕐 Recent Activity", "🕐 မကြာသေးခင်က လုပ်ဆောင်ချက်များ"))
            _act = recent_activity()
            if _act:
                _rows = "".join(
                    f"<tr><td>{html.escape(_t)}</td><td title=\"{html.escape(_f, quote=True)}\">{html.escape(_f)}</td><td>{html.escape(_a)}</td>"
                    f'<td><span class="rf-badge">{html.escape(_s)}</span></td></tr>'
                    for _t, _f, _a, _s in _act
                )  # v3.24.3 (#3): the File cell truncates with an ellipsis in CSS — the title attribute keeps the full name on hover.
                st.markdown(
                    '<div class="rf-table-wrap"><table class="rf-table"><tr><th>' + T("Date &amp; Time", "ရက်စွဲနဲ့ အချိန်") + '</th><th>'
                    + T("File / Project", "ဖိုင် / Project") + "</th><th>" + T("Activity", "လုပ်ဆောင်ချက်")
                    + "</th><th>" + T("Status", "အခြေအနေ") + "</th></tr>" + _rows + "</table></div>",
                    unsafe_allow_html=True,
                )
            else:
                st.info(T("No activity yet — record or upload audio to begin.", "လုပ်ဆောင်ချက် မရှိသေးဘူး — စဖို့ audio record လုပ်ပါ၊ ဒါမှမဟုတ် upload လုပ်ပါ။"))
    with _qa:
        with st.container(border=True):
            st.subheader(T("⚡ Quick Actions", "⚡ အမြန်လုပ်ဆောင်ချက်များ"))
            # v3.24.2 (#7): the 2 sub-columns squeezed labels to ~180px and
            # broke them mid-word ("Uplo ad Files") — stack the 4 buttons
            # vertically in the full column width instead.
            if st.button(T("🎙️ Start Interview", "🎙️ Interview စမယ်"), key="qa_record", use_container_width=True):
                st.session_state["page"] = "Fieldworks"
                st.rerun()
            if st.button(T("📊 Run Analysis", "📊 Analysis လုပ်မယ်"), key="qa_analyze", use_container_width=True):
                st.session_state["page"] = "Market Analysis"
                st.rerun()
            if st.button(T("📤 Upload Files", "📤 ဖိုင်တင်မယ်"), key="qa_upload", use_container_width=True):
                st.session_state["page"] = "Upload & Transcribe"
                st.rerun()
            if st.button(T("📑 Create Report", "📑 Report လုပ်မယ်"), key="qa_report", use_container_width=True):
                st.session_state["page"] = "Reports"
                st.rerun()
    with _ss:
        with st.container(border=True):
            st.subheader(T("💚 System Status", "💚 စနစ် အခြေအနေ"))
            _drive_st = (T("✅ Connected (OAuth)", "✅ ချိတ်ထားတယ် (OAuth)") if DRIVE_AUTH == "oauth"
                         else T("✅ Connected", "✅ ချိတ်ထားတယ်") if DRIVE_AUTH == "service"
                         else T("○ Local only", "○ Local ပဲ"))
            _gem_st = (T("✅ Ready", "✅ အဆင်သင့်") if get_gemini_client() is not None
                       else T("⚠️ No API key", "⚠️ API key မရှိဘူး"))
            st.write(f"☁️ **{T('Google Drive', 'Google Drive')}** — {_drive_st}")
            st.write(f"✦ **{T('Gemini AI', 'Gemini AI')}** — {_gem_st}")
            st.write(f"💾 **{T('Local Storage', 'Local Storage')}** — {T('✅ Healthy', '✅ ကောင်းတယ်')}")
            st.write(f"📦 **{T('Backup', 'Backup')}** — {T('○ On demand', '○ လိုအပ်မှ လုပ်မယ်')}")
            st.caption(T("Create a backup ZIP in Backup & Restore below.", "Backup & Restore မှာ ZIP လုပ်လို့ရတယ်။"))

    st.subheader(T("💾 Backup & Restore", "💾 Backup နဲ့ Restore"))
    _b1, _b2 = st.columns(2)
    with _b1:
        st.write(T("You can save the whole project to your PC as a ZIP.", "Project တခုလုံး ZIP အဖြစ် PC မှာ သိမ်းထားလို့ရတယ်။"))
        if st.button(T("📦 Create backup ZIP", "📦 Backup ZIP လုပ်မယ်"), key="create_backup"):
            with st.spinner(T("Creating backup ZIP...", "Backup ZIP လုပ်နေပါတယ်...")):
                _bname, _bdata = build_backup_zip()
                st.session_state["backup_zip"] = (_bname, _bdata)
        if "backup_zip" in st.session_state:
            _bname, _bdata = st.session_state["backup_zip"]
            st.download_button(T("⬇️ Download backup", "⬇️ Backup ဒေါင်းလုတ်"), data=_bdata, file_name=_bname,
                               mime="application/zip", key="download_backup")
            st.caption(f"{_bname} — {len(_bdata) / 1024 / 1024:.1f} MB")
    with _b2:
        st.write(T("You can restore from a backup ZIP. Existing files are backed up first, then replaced.", "Backup ZIP ကနေ ပြန်ထည့်လို့ရတယ်။ ရှိပြီးသားဖိုင်တွေကို အရင် backup ယူပြီးမှ အစားထိုးမယ်။"))
        _up = st.file_uploader(T("Choose a backup ZIP", "Backup ZIP ရွေးပါ"), type=["zip"], key="restore_upload")
        if _up is not None:
            _rdata = _up.getvalue()
            _rman, _rerr = read_backup_manifest(_rdata)
            if _rerr:
                st.error(_rerr)
            elif _rman.get("project", "Default") != current_project():
                _bproj = _rman.get("project", "Default")
                st.warning(T(f"⚠️ This backup is from project **{_bproj}**. Restore it into **{current_project()}** now? If this is a mistake, Project A's data will end up in Project B.", f"⚠️ ဒီ Backup က **{_bproj}** Project ကပါ။ အခု **{current_project()}** Project မှာ restore လုပ်မလား? မှားရင် Project B မှာ Project A data ရောက်သွားမယ်။"))
                if st.button(T(f"↩️ Restore anyway ({_bproj} → {current_project()})", f"↩️ ဒါပေမယ့် Restore လုပ်မယ် ({_bproj} → {current_project()})"), key="confirm_restore"):
                    with st.spinner(T("Restoring...", "Restore လုပ်နေပါတယ်...")):
                        _restore_report(restore_backup_zip(_rdata, force=True))
            elif st.button(T("↩️ Restore backup", "↩️ Backup ပြန်ထည့်မယ်"), key="do_restore"):
                with st.spinner(T("Restoring...", "Restore လုပ်နေပါတယ်...")):
                    _restore_report(restore_backup_zip(_rdata))

    with st.expander(T("🩺 System check", "🩺 စနစ် စစ်ဆေးမှု")):
        st.caption(T("Checks setup only — no Gemini request quota is used.",
                     "Setup ကိုပဲ စစ်တာ — Gemini request quota မကုန်ဘူး။"))
        if st.button(T("Run check", "စစ်မယ်"), key="run_system_check"):
            with st.spinner(T("Checking...", "စစ်နေပါတယ်...")):
                st.session_state["sys_check"] = run_system_check()
        for _ok, _lab, _det in st.session_state.get("sys_check", []):
            _icon = "✅" if _ok else ("⚠️" if _ok is None else "❌")
            st.write(f"{_icon} **{_lab}** — {_det}")

elif page == "New Project":
    # v3.21: the ONE place to create a project. Reached only from the sidebar
    # "+ New project" button (or the Dashboard hero button) — no sidebar
    # expander, no duplicate form on the Client Brief page.
    st.title(T("＋ New Project", "＋ Project အသစ်"))
    st.caption(T("One project = one container: brief, design, fieldwork, "
                 "cleansing, analysis and reports, saved per project to Drive + local.",
                 "Project တစ်ခုက container တစ်ခု: brief, ဒီဇိုင်း, ကွင်းဆင်း, သန့်စင်, "
                 "ခွဲခြမ်းစိတ်ဖြာမှုနဲ့ report — Drive + local မှာ project အလိုက် သိမ်းတယ်။"))
    _np_name = st.text_input(T("Project name", "Project နာမည်"), key="np_name")
    _np_desc = st.text_area(T("Description", "ဖော်ပြချက်"), height=80, key="np_desc")
    _np_members = st.text_input(T("Add members (names, comma-separated)", "အဖွဲ့ဝင် ထည့်ပါ (နာမည်များ, ကော်မာနဲ့ ခြားပါ)"), key="np_members")
    _np_c1, _np_c2 = st.columns(2)
    if _np_c1.button(T("Create project", "Project ဖွင့်"), key="np_create", type="primary"):
        _created = ensure_project(_np_name)
        if _created:
            st.session_state["active_project"] = _created
            if _created not in st.session_state["project_list"]:
                st.session_state["project_list"].append(_created)
            refresh_project_dirs(_created)
            if _np_desc.strip() or _np_members.strip():
                store_write_text("notes", "project.json", json.dumps({
                    "name": _created,
                    "description": _np_desc.strip(),
                    "members": [m.strip() for m in _np_members.split(",") if m.strip()],
                }, ensure_ascii=False, indent=2))
            st.session_state["project_select_target"] = _created
            st.session_state["page"] = "Client Brief"
            st.rerun()
        else:
            st.warning(T("Please type a valid project name.", "Project နာမည် မှန်မှန်ရိုက်ပါ။"))
    if _np_c2.button(T("← Back to Dashboard", "← Dashboard ကို ပြန်သွားမယ်"), key="np_back"):
        st.session_state["page"] = "Dashboard"
        st.rerun()

elif page == "Client Brief":
    # Stage 1: the brief form. Fields follow the 5-stage table exactly —
    # Client info, project info, objectives, target audience, scope, timeline,
    # deliverables, approval (Draft → Approved) + Discussion Guide (moved here
    # from the Market Analysis planner). Saved per project as brief.json
    # (Drive notes folder + local). The old placeholder text is gone.
    st.title(T("📋 Client Brief", "📋 Client Brief"))
    _brief = load_brief()
    _appr = _brief.get("approval", "Draft")
    st.caption(T(f"Approval: {_appr}", f"အတည်ပြုချက်: {_appr}"))
    with st.form("brief_form"):
        st.subheader(T("Client info", "Client info"))
        _bc1, _bc2 = st.columns(2)
        _client_name = _bc1.text_input(T("Client name", "Client နာမည်"), value=_brief.get("client_name", ""), key="brief_client_name")
        _client_contact = _bc2.text_input(T("Contact person", "ဆက်သွယ်ရမည့်သူ"), value=_brief.get("client_contact", ""), key="brief_client_contact")
        _client_email = st.text_input(T("Contact email", "ဆက်သွယ်ရန် email"), value=_brief.get("client_email", ""), key="brief_client_email")
        st.subheader(T("Project info", "Project info"))
        _bp1, _bp2 = st.columns(2)
        _project_code = _bp1.text_input(T("Project code", "Project code"), value=_brief.get("project_code", ""), key="brief_project_code")
        _background = st.text_area(T("Background", "နောက်ခံ"), value=_brief.get("background", ""), height=100, key="brief_background")
        st.subheader(T("Objectives", "ရည်ရွယ်ချက်များ"))
        _objectives = st.text_area(T("What should this research answer?", "ဒီသုတေသနက ဘာကို ဖြေပေးရမလဲ။"), value=_brief.get("objectives", ""), height=100, key="brief_objectives")
        st.subheader(T("Target audience", "ပစ်မှတ်အုပ်စု"))
        _audience = st.text_area(T("Who are the respondents?", "ဘယ်သူတွေကို မေးမလဲ။"), value=_brief.get("target_audience", ""), height=80, key="brief_target_audience")
        st.subheader(T("Scope", "Scope"))
        _scope = st.text_area(T("What is in scope — and out of scope?", "ဘာတွေ ပါမလဲ — ဘာတွေ မပါဘူး။"), value=_brief.get("scope", ""), height=80, key="brief_scope")
        st.subheader(T("Timeline", "အချိန်ဇယား"))
        _bt1, _bt2 = st.columns(2)
        _start = _bt1.text_input(T("Start date (YYYY-MM-DD)", "စမည့်ရက် (YYYY-MM-DD)"), value=_brief.get("timeline_start", ""), key="brief_start")
        _end = _bt2.text_input(T("End date (YYYY-MM-DD)", "ပြီးမည့်ရက် (YYYY-MM-DD)"), value=_brief.get("timeline_end", ""), key="brief_end")
        st.subheader(T("Deliverables", "ထုတ်ပေးမည့်အရာများ"))
        _delivs = st.multiselect(T("Deliverables", "ထုတ်ပေးမည့်အရာများ"), BRIEF_DELIVERABLES, default=[d for d in _brief.get("deliverables", []) if d in BRIEF_DELIVERABLES], key="brief_deliverables")
        st.subheader(T("Discussion Guide", "Discussion Guide"))
        _dg = st.text_area(T("Discussion Guide (sections / questions)", "Discussion Guide (အပိုင်း / မေးခွန်းများ)"), value=_brief.get("discussion_guide", ""), height=120, key="brief_dg")
        st.subheader(T("Approval", "အတည်ပြုချက်"))
        _approval = st.selectbox(T("Approval status", "အတည်ပြုမှု အခြေအနေ"), BRIEF_APPROVALS, index=BRIEF_APPROVALS.index(_appr) if _appr in BRIEF_APPROVALS else 0, key="brief_approval")
        if st.form_submit_button(T("💾 Save brief", "💾 Brief သိမ်းမယ်"), type="primary"):
            _data = {
                "client_name": _client_name.strip(), "client_contact": _client_contact.strip(),
                "client_email": _client_email.strip(), "project_code": _project_code.strip(),
                "background": _background.strip(), "objectives": _objectives.strip(),
                "target_audience": _audience.strip(), "scope": _scope.strip(),
                "timeline_start": _start.strip(), "timeline_end": _end.strip(),
                "deliverables": list(_delivs), "discussion_guide": _dg.strip(),
                "approval": _approval,
            }
            save_brief(_data)
            st.success(T("Brief saved.", "Brief သိမ်းပြီးပြီ။"))
            st.rerun()

elif page == "Research Design":
    # Stage 2 shell page: the full builder (approach, sampling, instruments,
    # fieldwork plan) comes in the Stage 2 build — the list below already
    # follows the 5-stage table exactly.
    st.title(T("📐 Research Design", "📐 သုတေသန ဒီဇိုင်း"))
    st.info(
        T("Approach (Qual / Quant / Mixed), IDI / FGD / survey, sampling, "
          "instruments (Discussion Guide / questionnaire / screener — AI draft) "
          "+ fieldwork plan — coming in the Stage 2 build.",
          "Approach (Qual / Quant / Mixed), IDI / FGD / survey, sampling, "
          "instruments (Discussion Guide / questionnaire / screener — AI draft) "
          "+ fieldwork plan — Stage 2 build မှာ လာမယ်။")
    )

elif page == "Data Cleansing":
    # Stage 4 shell page: file counts now; per-file check/clean/unclear/issue
    # status comes in the Stage 4 build.
    st.title(T("🧹 Data Cleansing", "🧹 ဒေတာ သန့်စင်ခြင်း"))
    st.info(
        T("Quality Control — check each raw input (checked / clean / unclear / "
          "issue), status + audit trail — coming in the Stage 4 build.",
          "Quality Control — raw input တစ်ခုချင်းကို စစ်တာ (checked / clean / "
          "unclear / issue), status + audit trail — Stage 4 build မှာ လာမယ်။")
    )

    st.subheader(T("Data Preparation", "ဒေတာ ပြင်ဆင်ခြင်း"))
    _pa, _pt, _pm, _pc, _pp = _data_prep_counts()
    for _lbl, _ok in [
        (T("Upload", "Upload"), _pa > 0),
        (T("Batch", "Batch"), _pa > 0 and _pp == 0),
        (T("Split", "ခွဲခြင်း"), _pc > 0),
        (T("Transcribe", "Transcript လုပ်ခြင်း"), _pt > 0),
        (T("Merge", "ပေါင်းခြင်း"), _pm > 0),
    ]:
        st.write(("✅ " if _ok else "○ ") + _lbl)

elif page == "Insights":
    # Stage 5: Qual + Quant combined. The builder comes in the Stage 5 build;
    # for now it shows which inputs already exist (real counts only).
    st.title(T("💡 Insights", "💡 Insights"))
    _in_qual = len(list(ANALYSIS_DIR.glob("*_analysis.txt")))
    _in_quant = len(list(ANALYSIS_DIR.glob("quant_*.txt")))
    st.info(T("Insights (Qual + Quant combined) — coming in the Stage 5 build. Inputs it will use:",
              "Insights (Qual + Quant ပေါင်း) — Stage 5 build မှာ လာမယ်။ သုံးမယ့် input တွေ:"))
    st.write(("✅ " if _in_qual else "○ ") + T(f"Qualitative analyses saved: {_in_qual}", f"သိမ်းထားတဲ့ Qualitative analysis: {_in_qual}"))
    st.write(("✅ " if _in_quant else "○ ") + T(f"Quantitative summaries saved: {_in_quant}", f"သိမ်းထားတဲ့ Quantitative summary: {_in_quant}"))

elif page == "Presentation":
    # Stage 5 shell page: presentation is part of Stage 5 — the PPTX comes
    # from the report on the Reports page.
    st.title(T("🎤 Presentation", "🎤 တင်ပြမှု"))
    st.info(
        T("Presentation is part of Stage 5 — export the PPTX from the report "
          "on the Reports page.",
          "Presentation က Stage 5 ထဲမှာ ပါတယ် — Reports စာမျက်နှာမှာ report ကနေ "
          "PPTX ထုတ်လို့ရတယ်။")
    )

elif page == "Upload & Transcribe":
    st.title(T("🎙️ Upload & Transcribe", "🎙️ Upload နဲ့ Transcript လုပ်ခြင်း"))

    uploaded = st.file_uploader(T("Upload an audio file (mp3 / wav / m4a)", "Audio file တင်ပါ (mp3 / wav / m4a)"), type=["mp3", "wav", "m4a"], key="dc_audio_upload")
    audio_path = None
    if uploaded is not None:
        safe_name = Path(uploaded.name).name
        if safe_name != uploaded.name or safe_name in {"", ".", ".."}:
            st.error(T("Invalid file name.", "ဖိုင်အမည် မမှန်ပါ။"))
            st.stop()
        audio_path = _safe_storage_path("audio", safe_name)
        _uhash = hashlib.sha256(uploaded.getbuffer()).hexdigest()
        upload_key = (current_project(), safe_name, _uhash)
        if st.session_state.get("last_audio_upload") != upload_key:
            store_write_bytes("audio", safe_name, uploaded.getbuffer())
            st.session_state["last_audio_upload"] = upload_key
        st.caption(T(f"Saved: {uploaded.name}", f"သိမ်းပြီး: {uploaded.name}"))
    else:
        existing = sorted(AUDIO_DIR.glob("*.mp3")) + sorted(AUDIO_DIR.glob("*.wav")) + sorted(AUDIO_DIR.glob("*.m4a"))
        if existing:
            audio_names = [f.name for f in existing]
            audio_selected = st.selectbox(T("Choose an existing audio file", "ရှိပြီးသား audio ရွေးပါ"), audio_names, key="dc_audio_existing")
            audio_path = AUDIO_DIR / audio_selected
        else:
            st.info(T("No audio file yet — upload one first.", "Audio file မရှိသေးပါ — အရင် upload လုပ်ပါ။"))

    if audio_path is not None and st.button("Transcribe"):
        with st.spinner(T("Transcribing... (long audio can take a few minutes)", "Transcribing... (audio ကြာရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)")):
            text, model = transcribe_audio(audio_path)
        if text:
            out_file = store_write_text("transcripts", audio_path.stem + ".txt", text)
            st.success(f"Done — model: {model}")
            st.caption(T(f"Saved: {out_file.name} — you can reopen it on the Transcripts page.", f"Saved: {out_file.name} — Transcripts page မှာ ပြန်ဖွင့်လို့ရတယ်။"))
        else:
            st.error(f"Failed: {model}")

elif page == "Batch Processing":
    st.title(T("📁 Batch Processing", "📁 Batch လုပ်ဆောင်ခြင်း"))

    audio_files = sorted(AUDIO_DIR.glob("*.mp3")) + sorted(AUDIO_DIR.glob("*.wav")) + sorted(AUDIO_DIR.glob("*.m4a"))
    if not audio_files:
        st.info(T("No audio file yet — upload one on the Upload & Transcribe page.", "Audio file မရှိသေးပါ — Upload & Transcribe page မှာ upload လုပ်ပါ။"))
    else:
        pending = []
        done = []
        for f in audio_files:
            if (OUTPUT_DIR / (f.stem + ".txt")).exists():
                done.append(f.name)
            else:
                pending.append(f.name)
        st.write(T(f"Total: {len(audio_files)} | Done: {len(done)} | Remaining: {len(pending)}", f"စုစုပေါင်း: {len(audio_files)} | ပြီးပြီ: {len(done)} | ကျန်သေးတယ်: {len(pending)}"))
        if pending:
            st.write(T("Remaining files:", "ကျန်တဲ့ file တွေ:"))
            for name in pending:
                st.write(f"- {name}")
        if pending and st.button(T("Transcribe next 5", "နောက် 5 ခု transcribe လုပ်မယ်")):
            progress = st.progress(0)
            batch = pending[:5]
            for i, name in enumerate(batch):
                audio_path = AUDIO_DIR / name
                with st.spinner(f"Transcribing {name}... ({i+1}/{len(batch)})"):
                    text, model = transcribe_audio(audio_path)
                if text:
                    store_write_text("transcripts", audio_path.stem + ".txt", text)
                    st.success(f"Done — {name} ({model})")
                else:
                    st.error(f"Failed — {name}: {model}")
                    if "429" in model:
                        st.warning(T("429 — quota is full, batch stopped. Try again later.", "429 — quota ပြည့်နေတယ်, batch ရပ်လိုက်ပြီ။ နောက်မှ ပြန်လုပ်။"))
                        break
                progress.progress((i + 1) / len(batch))
                if i < len(batch) - 1:
                    time.sleep(60)
            st.info(T("Batch done — to do the next 5, wait 15 minutes, then press the button again.", "Batch ပြီးပြီ — နောက် 5 ခု လုပ်ချင်ရင် 15 မိနစ် နားပြီးမှ ခလုတ် ပြန်နှိပ်ပါ။"))
        elif not pending:
            # v3.21: a done batch used to leave a near-empty page — show the
            # processed files instead.
            st.success(T("All audio transcribed.", "Audio အားလုံး transcript ပြီးပြီ။"))
            st.write(T("Transcribed files:", "Transcript ပြီးတဲ့ file တွေ:"))
            for name in done:
                st.write(f"- {name} ✅")

elif page == "Fieldworks":
    st.title(T("🎙️ Fieldworks", "🎙️ ကွင်းဆင်းလုပ်ငန်း"))
    st.caption(T("Record one chunk at a time with the browser mic → auto-transcribe. 15 minutes per chunk. 1 chunk = 1 request (a 2h FGD ≈ 8 chunks).", "Browser mic နဲ့ chunk တစ်ခုချင်း record လုပ် → auto-transcribe. 15 မိနစ် တစ်ပိုင်း. 1 chunk = 1 request (2h FGD ≈ 8 chunks)."))
    st.info(T("If the browser asks for the mic, press Allow. If many people talk at once, speaker separation gets less accurate — check anything marked [unclear].", "Browser က mic ခွင့်တောင်းရင် Allow နှိပ်ပါ. လူအများ တစ်ပြိုင်နက် ပြောရင် speaker ခွဲတာ တိကျမှု လျော့နိုင်တယ် — [unclear] လို့ မှတ်ထားတာကို ပြန်စစ်ပါ."))

    session_name = st.text_input(T("Session name", "Session နာမည်"), value="fgd_01", key="fgd_session")
    session_name = re.sub(r"[^\w\-]+", "_", session_name.strip()) or "fgd_01"

    def session_chunk_files():
        files = []
        for f in AUDIO_DIR.glob(f"{session_name}(*).wav"):
            m = re.search(r"\((\d+)\)\.wav$", f.name)
            if m:
                files.append((int(m.group(1)), f))
        return sorted(files)

    chunks = session_chunk_files()
    if chunks:
        st.write(f"Chunks: {len(chunks)}")
        for n, f in chunks:
            txt = OUTPUT_DIR / (f.stem + ".txt")
            status = "✅ transcribed" if txt.exists() else "⏳ pending"
            st.write(f"- Chunk {n}: {f.name} — {status}")
    else:
        st.info(T("No chunks yet — record one below.", "Chunk မရှိသေးဘူး — အောက်မှာ record လုပ်ပါ။"))

    audio_data = st.audio_input(T("🎙️ Record a chunk (about 15 min)", "🎙️ Record a chunk (15 မိနစ် လောက်)"))
    if audio_data is not None:
        audio_bytes = audio_data.read()
        audio_hash = hashlib.sha256(audio_bytes).hexdigest()
        record_key = (current_project(), session_name, audio_hash)
        if st.session_state.get("last_recorded_key") != record_key:
            next_n = (chunks[-1][0] + 1) if chunks else 1
            chunk_path = AUDIO_DIR / f"{session_name}({next_n}).wav"
            store_write_bytes("audio", chunk_path.name, audio_bytes)
            st.session_state["last_recorded_key"] = record_key
            st.session_state["last_chunk"] = (current_project(), session_name, chunk_path.name)
            st.success(f"Saved: {chunk_path.name}")
            txt_path = OUTPUT_DIR / (chunk_path.stem + ".txt")
            with st.spinner(T(f"Transcribing chunk {next_n}... (long audio can take a few minutes)", f"Chunk {next_n} ကို transcribing... (audio ကြာရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)")):
                text, model = transcribe_audio(chunk_path)
            if text:
                store_write_text("transcripts", txt_path.name, text)
                st.success(f"Done — model: {model}")
            else:
                st.error(T(f"Failed: {model} — retry with Transcribe pending below.", f"Failed: {model} — အောက်က Transcribe pending နဲ့ ပြန်လုပ်လို့ရတယ်။"))
            st.rerun()
        else:
            st.caption(T(f"This chunk is already saved: {st.session_state.get('last_chunk')}", f"ဒီ chunk က သိမ်းပြီးသား: {st.session_state.get('last_chunk')}"))
            last_info = st.session_state.get("last_chunk", ())
            last_name = last_info[2] if len(last_info) == 3 and last_info[:2] == (current_project(), session_name) else ""
            txt_path = OUTPUT_DIR / (Path(last_name).stem + ".txt") if last_name else None
            if last_name and txt_path.exists():
                with st.expander(T("View transcript", "Transcript ကြည့်ပါ")):
                    st.text(txt_path.read_text(encoding="utf-8"))

    pending_chunks = [(n, f) for n, f in session_chunk_files() if not (OUTPUT_DIR / (f.stem + ".txt")).exists()]
    if pending_chunks and st.button(T("Transcribe pending chunks", "Transcribe pending chunks")):
        for n, f in pending_chunks:
            with st.spinner(f"Transcribing chunk {n}..."):
                text, model = transcribe_audio(f)
            if text:
                store_write_text("transcripts", f.stem + ".txt", text)
                st.success(f"Done — chunk {n} ({model})")
            else:
                st.error(f"Failed — chunk {n}: {model}")
                if "429" in model:
                    st.warning(T("429 — quota is full, try again later.", "429 — quota ပြည့်နေတယ်, နောက်မှ ပြန်လုပ်။"))
                    break

elif page == "Upload & Split":
    st.title(T("✂️ Upload & Split", "✂️ Upload နဲ့ ခွဲခြင်း"))
    st.caption(T("Split a long audio (recorded on a device) into 15-minute chunks → then transcribe with Batch Processing.", "အရှည်ကြီး audio (device နဲ့ record လုပ်ထားတာ) ကို 15 မိနစ် အပိုင်းတွေ ခွဲ → ပြီးမှ Batch Processing နဲ့ transcribe လုပ်."))
    if DRIVE_MODE:
        st.info(T("Splitting needs ffmpeg (for mp3/m4a). WAV works without ffmpeg. An uploaded file is saved straight to Google Drive (ResearchFlow / project / audio).", "Split လုပ်ဖို့ ffmpeg လိုတယ် (mp3/m4a အတွက်). WAV ဆိုရင် ffmpeg မရှိလည်း ရတယ်. Upload လုပ်တဲ့ file က Google Drive (ResearchFlow / project / audio) ထဲ တိုက်ရိုက် သိမ်းမယ်."))
    else:
        st.info(T("Splitting needs ffmpeg (for mp3/m4a). WAV works without ffmpeg. Upload straight to this website for now.", "Split လုပ်ဖို့ ffmpeg လိုတယ် (mp3/m4a အတွက်). WAV ဆိုရင် ffmpeg မရှိလည်း ရတယ်. အခု website ကို တိုက်ရိုက် တင်ပါ."))

    src_mode = st.radio(T("Where is the audio from?", "Audio ဘယ်ကယူမလဲ"), ["upload", "existing"], format_func=lambda v: T("Upload", "Upload လုပ်မယ်") if v == "upload" else T("Choose an existing file", "ရှိပြီးသား file ရွေးမယ်"), horizontal=True)
    src_path = None
    if src_mode == "upload":
        up = st.file_uploader(T("Upload an audio file (mp3 / wav / m4a)", "Audio file တင်ပါ (mp3 / wav / m4a)"), type=["mp3", "wav", "m4a"], key="split_upload")
        if up is not None:
            src_path = AUDIO_DIR / up.name
            if not src_path.exists():
                store_write_bytes("audio", up.name, up.getbuffer())
            st.caption(f"Saved: {up.name}")
    else:
        existing = sorted(AUDIO_DIR.glob("*.mp3")) + sorted(AUDIO_DIR.glob("*.wav")) + sorted(AUDIO_DIR.glob("*.m4a"))
        if existing:
            sel = st.selectbox(T("Choose audio", "Audio ရွေးပါ"), [f.name for f in existing], key="split_existing")
            src_path = AUDIO_DIR / sel
        else:
            st.info(T("No audio file yet — upload one first.", "Audio file မရှိသေးပါ — အရင် upload လုပ်ပါ။"))

    if src_path is not None:
        dur = audio_duration(src_path)
        if dur:
            st.caption(T(f"Duration: {dur / 60:.1f} min", f"Duration: {dur / 60:.1f} မိနစ်"))
        chunk_min = st.selectbox(T("Minutes per chunk", "တစ်ပိုင်း ဘယ်လောက် မိနစ်"), [10, 15], index=1, key="split_len")
        if st.button(T("Split into chunks", "အပိုင်းတွေ ခွဲမယ်")):
            with st.spinner("Splitting..."):
                chunks, err = split_audio(src_path, chunk_min)
            if err:
                st.error(f"Failed: {err}")
                if "ffmpeg" in err:
                    st.info("PowerShell မှာ ဒါ run လုပ်ပြီး ffmpeg install လုပ်ပါ (ပြီးရင် PowerShell အသစ်ဖွင့်): winget install ffmpeg")
            elif chunks:
                st.success(f"Done — {len(chunks)} chunks:")
                for c in chunks:
                    st.write(f"- {c.name}")
                    store_upload_existing("audio", c.name)
                orig_dir = AUDIO_DIR / "originals"
                orig_dir.mkdir(exist_ok=True)
                shutil.move(str(src_path), str(orig_dir / src_path.name))
                drive_move_to_originals(src_path.name)
                st.caption(T("Original file moved to audio/originals/ — the chunks are now in Batch Processing.", "မူရင်း file ကို audio/originals/ ထဲ ရွှေ့ပြီးပြီ — Batch Processing မှာ chunks တွေ ပေါ်နေပြီ။"))
            else:
                st.warning(T("No chunks — the file may be shorter than one chunk.", "Chunk မထွက်ဘူး — file က chunk တစ်ပိုင်းထက် တိုနေလို့ဖြစ်နိုင်တယ်။"))

elif page == "Transcripts":
    st.title(T("📝 Transcripts", "📝 Transcript များ"))

    st.subheader(T("🔗 Merge chunks — full transcript", "🔗 Chunk တွေ ပေါင်း — full transcript"))
    groups = get_chunk_groups()
    multi_groups = {k: v for k, v in groups.items() if len(v) > 1}
    if multi_groups:
        for key, files_in_group in sorted(multi_groups.items()):
            merged = MERGED_DIR / f"{key}_full.txt"
            status = "✅ merged" if merged.exists() else f"{len(files_in_group)} chunks"
            with st.expander(f"{key}: {status}"):
                for i, f in enumerate(files_in_group, start=1):
                    st.write(f"{i}. {f.name}")
        if st.button(T("Merge all into full transcripts", "Full transcript ဖြစ်အောင် အကုန် Merge လုပ်မယ်")):
            for key, files_in_group in sorted(multi_groups.items()):
                parts = []
                for i, f in enumerate(files_in_group, start=1):
                    parts.append(f"--- chunk {i} ({f.name}) ---\n")
                    parts.append(f.read_text(encoding="utf-8"))
                out = store_write_text("merged", f"{key}_full.txt", "\n\n".join(parts))
            st.success(T("Merged — full transcripts ready in output/merged/", "Merge ပြီးပြီ — full transcript တွေ output/merged/ မှာ အဆင်သင့်ဖြစ်ပြီ။"))
    else:
        st.info(T("No interview with many chunks.", "Chunk အများကြီး ရှိတဲ့ interview မရှိဘူး။"))

    st.divider()

    st.subheader(T("📄 Full transcripts (merged)", "📄 Full transcript များ (merge ပြီးသား)"))
    merged_files = sorted(MERGED_DIR.glob("*_full.txt"), key=lambda f: f.name)
    if merged_files:
        for file in merged_files:
            with st.expander(file.name):
                st.text(file.read_text(encoding="utf-8"))
    else:
        # v3.21: the old text told the user to press a Merge button that does
        # not exist when there are no multi-chunk groups — say which case it is.
        if multi_groups:
            st.info(T("No merged transcript yet — press the Merge button above.", "Merge မလုပ်ရသေးဘူး — အပေါ်က Merge ခလုတ် နှိပ်ပါ။"))
        else:
            st.info(T("Nothing to merge — each transcript here is already a full transcript.", "Merge လုပ်စရာ မရှိဘူး — ဒီမှာရှိတဲ့ transcript တစ်ခုချင်းစီက full transcript ဖြစ်ပြီးသား။"))

    st.divider()

    files = get_transcript_files()

    if files:
        for file in files:
            with st.expander(file.name):
                st.text(file.read_text(encoding="utf-8"))
    else:
        st.warning(T("No transcript file found yet.", "Transcript ဖိုင် မတွေ့ရှိသေးပါ။"))

elif page == "Market Analysis":
    st.title(T("📊 Market Analysis", "📊 ဈေးကွက် ခွဲခြမ်းစိတ်ဖြာမှု"))

    # Analysis uses FULL transcripts only — never chunks.
    full_files, pending = get_full_transcripts()
    if not full_files:
        st.warning(T("No full transcript yet — merge on the Transcripts page first.", "Full transcript မရှိသေးပါ — အရင် Transcripts page မှာ Merge လုပ်ပါ။"))
    else:
        names = [f.name for f in full_files]
        file_map = {f.name: f for f in full_files}
        if pending:
            st.caption(T("Not merged yet: " + ", ".join(f"{k} ({n} chunks)" for k, n in pending) + " — merge on the Transcripts page.", "Merge မလုပ်ရသေးဘူး: " + ", ".join(f"{k} ({n} chunks)" for k, n in pending) + " — Transcripts page မှာ Merge လုပ်ပါ။"))

        st.subheader(T("📋 Analysis Planner", "📋 Analysis Planner"))
        st.caption(T("Before analysis: look at the Brief + Discussion Guide + Questionnaire + transcripts together and plan first.", "Analysis မလုပ်ခင်: Brief + Discussion Guide + Questionnaire + transcript ပေါင်းကြည့်ပြီး plan အရင် ဆွဲမယ်။"))
        # v3.21: Client Brief / Objective + Discussion Guide live on the
        # Client Brief page now (Stage 1) — the planner reads them from the
        # saved brief instead of asking again. No duplicate inputs here.
        _brief = load_brief()
        brief_text = _brief.get("objectives", "")
        dg_text = _brief.get("discussion_guide", "")
        if brief_text or dg_text:
            with st.expander(T("Brief & Discussion Guide (from the Client Brief page)", "Brief နဲ့ Discussion Guide (Client Brief page ကနေ)")):
                if brief_text:
                    st.markdown(f"**{T('Objective', 'ရည်ရွယ်ချက်')}**")
                    st.text(brief_text)
                if dg_text:
                    st.markdown(f"**{T('Discussion Guide', 'Discussion Guide')}**")
                    st.text(dg_text)
                st.caption(T("Edit these on the Client Brief page.", "ဒါတွေကို Client Brief page မှာ ပြင်ပါ။"))
        else:
            st.info(T("No brief saved yet — fill in the Client Brief page first.", "Brief မသိမ်းရသေးဘူး — အရင် Client Brief page မှာ ဖြည့်ပါ။"))
        questionnaire_text = st.text_area(T("Questionnaire (survey questions — paste here)", "Questionnaire (survey မေးခွန်းတွေ — paste လုပ်)"), key="questionnaire_text", height=150)
        collection_tool = st.selectbox(
            T("Data collection tool", "Data ကောက်တဲ့ tool"),
            ["Google Forms", "SurveyMonkey", "Kobo Collect", T("Other", "အခြား")],
            key="collection_tool",
        )
        plan_transcripts = st.multiselect(
            T("Pick transcripts for the plan", "Plan အတွက် transcript တွေ ရွေးပါ"),
            names,
            key="plan_transcripts",
        )
        st.caption(T("The transcripts you pick for the plan must match the ones you pick for Analyze below.", "Plan အတွက် ရွေးတဲ့ transcript တွေနဲ့ အောက်က Analyze အတွက် ရွေးတာ တူရမယ်။"))
        if st.button("Plan analysis"):
            if not brief_text.strip() and not dg_text.strip() and not questionnaire_text.strip():
                st.warning(T("Add at least one of Brief / Discussion Guide / Questionnaire.", "Brief / Discussion Guide / Questionnaire အနည်းဆုံး တစ်ခု ထည့်ပါ။"))
            elif not plan_transcripts:
                st.warning(T("Pick at least 1 transcript for the plan.", "Plan အတွက် transcript အနည်းဆုံး ၁ ခု ရွေးပါ။"))
            else:
                combined_plan = ""
                for name in plan_transcripts:
                    combined_plan += f"\n\n=== {name} ===\n"
                    combined_plan += file_map[name].read_text(encoding="utf-8")
                plan_prompt = (
                    PROMPT_LIBRARY["planning"]["template"]
                    + "\n\nClient Brief / Objective:\n" + brief_text + "\n\nDiscussion Guide:\n" + dg_text + "\n\nQuestionnaire:\n" + questionnaire_text + "\n\nData collection tool: " + collection_tool + "\n\nTranscripts:\n" + combined_plan
                )
                with st.spinner(T("Planning... (many transcripts can take a few minutes)", "Planning... (transcript များရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)")):
                    plan_result, plan_model = ask_gemini(plan_prompt)
                if plan_result:
                    st.success(f"Done — model: {plan_model}")
                    plan_ids = short_ids_for(plan_transcripts)
                    plan_file = store_write_text("analysis", f"plan_{'-'.join(plan_ids)}.txt", plan_result)
                    st.session_state["current_plan"] = plan_result
                    st.session_state["current_plan_name"] = plan_file.name
                    st.caption(f"Saved: {plan_file.name}")
                else:
                    st.error(f"Failed: {plan_model}")

        plan_files = sorted(ANALYSIS_DIR.glob("plan_*.txt"), key=lambda f: f.name)
        if plan_files:
            plan_names = [f.name for f in plan_files]
            plan_selected = st.selectbox(T("Pick a saved plan", "သိမ်းထားတဲ့ plan ရွေးပါ"), plan_names)
            if st.button("Load plan"):
                st.session_state["current_plan"] = (ANALYSIS_DIR / plan_selected).read_text(encoding="utf-8")
                st.session_state["current_plan_name"] = plan_selected
                st.success(f"Loaded: {plan_selected}")

        if st.session_state.get("current_plan"):
            st.caption(f"Current plan: {st.session_state.get('current_plan_name')}")
            with st.expander(T("View plan", "Plan ကြည့်ပါ")):
                st.markdown(st.session_state["current_plan"])

        st.divider()
        has_plan = st.session_state.get("current_plan") is not None
        use_plan = st.checkbox(
            T("Analyse following the approved plan", "Approved plan အတိုင်း analyse လုပ်မယ်"),
            value=has_plan,
            disabled=not has_plan,
        )
        mode = st.radio(T("Analysis mode", "Analysis ပုံစံ"), ["single", "multi"], format_func=lambda v: T("One transcript", "တစ်ခုတည်း") if v == "single" else T("Many (multi)", "အများကြီး (multi)"), horizontal=True)
        if mode == "single":
            default_idx = names.index("interview_01_full.txt") if "interview_01_full.txt" in names else 0
            selected = st.selectbox(T("Pick a transcript", "Transcript ရွေးပါ"), names, index=default_idx)
            if st.button("Analyze"):
                text = file_map[selected].read_text(encoding="utf-8")
                if use_plan and st.session_state.get("current_plan"):
                    prompt = plan_based_prompt(st.session_state["current_plan"], text, 1)
                else:
                    prompt = PROMPT_LIBRARY["qualitative"]["template"] + "\n\nTranscript:\n" + text
                with st.spinner(T("Analyzing... (please wait)", "Analyzing... (ခဏစောင့်ပါ)")):
                    result, model = ask_gemini(prompt)
                if result:
                    st.success(f"Done — model: {model}")
                    analysis_file = store_write_text("analysis", selected.replace(".txt", "_analysis.txt"), result)
                    st.session_state["last_analysis"] = analysis_file.name
                    st.caption(T(f"Saved: {analysis_file.name} — reopen it under Saved analyses below.", f"Saved: {analysis_file.name} — အောက်က Saved analyses မှာ ပြန်ဖွင့်လို့ရတယ်။"))
                else:
                    st.error(f"Failed: {model}")
        else:
            selected_multi = st.multiselect(
                T("Pick transcripts (about 6 at most — more takes longer)", "Transcript တွေ ရွေးပါ (အများဆုံး 6 ခု လောက်ရွေး — များလွန်းရင် ကြာမယ်)"),
                names,
            )
            if selected_multi and st.button("Analyze multi"):
                combined = ""
                for name in selected_multi:
                    combined += f"\n\n=== {name} ===\n"
                    combined += file_map[name].read_text(encoding="utf-8")
                if use_plan and st.session_state.get("current_plan"):
                    prompt = plan_based_prompt(st.session_state["current_plan"], combined, len(selected_multi))
                else:
                    prompt = (
                        PROMPT_LIBRARY["qualitative"]["variants"]["Multi-interview"].format(count=len(selected_multi))
                        + "\n\nTranscripts:\n" + combined
                    )
                with st.spinner(T("Analyzing... (many transcripts can take a few minutes)", "Analyzing... (transcript များရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)")):
                    result, model = ask_gemini(prompt)
                if result:
                    st.success(f"Done — model: {model}")
                    short_ids = short_ids_for(selected_multi)
                    analysis_file = store_write_text("analysis", f"multi_{'-'.join(short_ids)}_analysis.txt", result)
                    st.session_state["last_analysis"] = analysis_file.name
                    st.caption(T(f"Saved: {analysis_file.name} — reopen it under Saved analyses below.", f"Saved: {analysis_file.name} — အောက်က Saved analyses မှာ ပြန်ဖွင့်လို့ရတယ်။"))
                else:
                    st.error(f"Failed: {model}")

        st.divider()
        st.subheader(T("💾 Saved analyses", "💾 သိမ်းထားတဲ့ analysis များ"))
        saved = sorted(ANALYSIS_DIR.glob("*_analysis.txt"), key=lambda f: f.name)
        if saved:
            saved_names = [f.name for f in saved]
            last = st.session_state.get("last_analysis")
            default_saved = saved_names.index(last) if last in saved_names else 0
            saved_selected = st.selectbox(T("Pick a saved analysis", "သိမ်းထားတာ ရွေးပါ"), saved_names, index=default_saved)
            saved_text = (ANALYSIS_DIR / saved_selected).read_text(encoding="utf-8")
            if looks_like_html(saved_text):
                saved_text = html_to_text(saved_text)
            st.markdown(saved_text)
            st.download_button(
                "Download saved analysis (.txt)",
                data=(ANALYSIS_DIR / saved_selected).read_bytes(),
                file_name=saved_selected,
                mime="text/plain",
            )
        else:
            st.info(T("No saved analysis yet.", "သိမ်းထားတဲ့ analysis မရှိသေးပါ။"))

elif page == "Quantitative Analysis":
    st.title("📊 Quantitative Analysis")
    st.caption("Survey data (Google Forms / SurveyMonkey / Kobo Collect export — CSV / XLSX) → counts, %, charts, cross-tab.")

    uploaded = st.file_uploader(T("Upload a survey file (CSV / XLSX)", "Survey file တင်ပါ (CSV / XLSX)"), type=["csv", "xlsx"])
    df = None
    if uploaded is not None:
        try:
            if uploaded.name.endswith(".csv"):
                df = pd.read_csv(uploaded)
            else:
                df = pd.read_excel(uploaded)
        except ImportError:
            st.error(T("openpyxl is not installed — run this in PowerShell: .\\.venv\\Scripts\\python.exe -m pip install openpyxl", "openpyxl မရှိသေးဘူး — PowerShell မှာ ဒါ run လုပ်ပါ: .\\.venv\\Scripts\\python.exe -m pip install openpyxl"))
        except Exception as e:
            st.error(T(f"Could not read the file: {e}", f"File ဖတ်မရဘူး: {e}"))
    if df is not None:
        st.write(f"Rows: {len(df)} | Columns: {len(df.columns)}")
        st.dataframe(df.head(5))
        summary_lines = [f"Quant summary — {uploaded.name}", f"Rows: {len(df)}", ""]
        cat_cols = []
        for col in df.columns:
            series = df[col].dropna()
            n = len(series)
            nunique = series.nunique()
            st.subheader(f"❓ {col}")
            summary_lines.append(f"## {col} (n={n})")
            if pd.api.types.is_numeric_dtype(series):
                counts = series.value_counts().sort_index()
                table = pd.DataFrame({"Count": counts, "%": (counts / n * 100).round(1)})
                st.dataframe(table)
                st.bar_chart(counts)
                st.caption(f"Mean: {series.mean():.2f}")
                summary_lines.append(f"Mean: {series.mean():.2f}")
                for val, cnt in counts.items():
                    summary_lines.append(f"- {val}: {cnt} ({cnt / n * 100:.1f}%)")
                if nunique <= 20:
                    cat_cols.append(col)
            elif nunique <= 20:
                counts = series.value_counts()
                table = pd.DataFrame({"Count": counts, "%": (counts / n * 100).round(1)})
                st.dataframe(table)
                st.bar_chart(counts)
                for val, cnt in counts.items():
                    summary_lines.append(f"- {val}: {cnt} ({cnt / n * 100:.1f}%)")
                cat_cols.append(col)
            else:
                st.caption(T(f"Open text — {nunique} unique answers. Examples:", f"Open text — {nunique} unique answers. နမူနာ:"))
                for val in series.head(5):
                    st.write(f"- {val}")
                summary_lines.append(f"Open text — {nunique} unique answers.")
            summary_lines.append("")
        if len(cat_cols) >= 2:
            st.divider()
            st.subheader("🔀 Cross-tab")
            col_a = st.selectbox(T("First question", "ပထမ မေးခွန်း"), cat_cols, key="xtab_a")
            col_b = st.selectbox(T("Second question", "ဒုတိယ မေးခွန်း"), cat_cols, index=1 if len(cat_cols) > 1 else 0, key="xtab_b")
            if col_a != col_b:
                ctab = pd.crosstab(df[col_a], df[col_b])
                st.dataframe(ctab)
                summary_lines.append("## Cross-tab")
                summary_lines.append(f"{col_a} x {col_b}")
                summary_lines.append(ctab.to_string())
        if st.button("Save quant summary"):
            out_file = store_write_text("analysis", f"quant_{Path(uploaded.name).stem}.txt", "\n".join(summary_lines))
            st.success(f"Saved: {out_file.name}")
            st.download_button(
                "Download quant summary (.txt)",
                data=out_file.read_bytes(),
                file_name=out_file.name,
                mime="text/plain",
            )
        if st.button("📈 Interpret with AI"):
            _qprompt = PROMPT_LIBRARY["quantitative"]["template"] + "\n\nSurvey results:\n" + "\n".join(summary_lines)
            with st.spinner("Interpreting..."):
                _qres, _qmodel = ask_gemini(_qprompt)
            if _qres:
                _qfile = store_write_text("analysis", f"quant_{Path(uploaded.name).stem}_interpretation.txt", _qres)
                st.success(f"Done — model: {_qmodel} — Saved: {_qfile.name}")
            else:
                st.error(f"Failed: {_qmodel}")

elif page == "Reports":
    st.title(T("📑 Reports", "📑 Report များ"))

    saved = sorted(ANALYSIS_DIR.glob("*_analysis.txt"), key=lambda f: f.name)
    if not saved:
        st.info(T("No analysis yet — Analyze on the Market Analysis page first.", "Analysis မရှိသေးပါ — အရင် Market Analysis မှာ Analyze လုပ်ပါ။"))
    else:
        saved_names = [f.name for f in saved]
        selected = st.selectbox(T("Pick an analysis", "Analysis ရွေးပါ"), saved_names)
        template = st.selectbox(T("Pick a template", "Template ရွေးပါ"), list(TEMPLATES.keys()))
        col_pptx, col_docx = st.columns(2)
        if col_pptx.button("Export PPTX"):
            analysis_text = (ANALYSIS_DIR / selected).read_text(encoding="utf-8")
            if looks_like_html(analysis_text):
                analysis_text = html_to_text(analysis_text)
            title_text = selected.replace("_analysis.txt", "").replace("_", " ")
            out_path = REPORTS_DIR / selected.replace("_analysis.txt", ".pptx")
            with st.spinner(T("Exporting PPTX...", "PPTX ထုတ်နေပါတယ်...")):
                analysis_to_pptx(analysis_text, title_text, out_path, template)
            store_upload_existing("reports", out_path.name)
            st.success(f"Done — {out_path.name} ({template})")
            st.download_button(
                "Download PPTX",
                data=out_path.read_bytes(),
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        if col_docx.button("Export DOCX"):
            analysis_text = (ANALYSIS_DIR / selected).read_text(encoding="utf-8")
            if looks_like_html(analysis_text):
                analysis_text = html_to_text(analysis_text)
            title_text = selected.replace("_analysis.txt", "").replace("_", " ")
            out_path = REPORTS_DIR / selected.replace("_analysis.txt", ".docx")
            with st.spinner("DOCX ထုတ်နေပါတယ်..."):
                analysis_to_docx(analysis_text, title_text, out_path)
            store_upload_existing("reports", out_path.name)
            st.success(f"Done — {out_path.name}")
            st.download_button(
                "Download DOCX",
                data=out_path.read_bytes(),
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        st.divider()
        if st.button("📝 Create executive summary (topline)"):
            analysis_text = (ANALYSIS_DIR / selected).read_text(encoding="utf-8")
            if looks_like_html(analysis_text):
                analysis_text = html_to_text(analysis_text)
            _sprompt = PROMPT_LIBRARY["executive_summary"]["template"] + analysis_text
            with st.spinner("Summarizing..."):
                _sres, _smodel = ask_gemini(_sprompt)
            if _sres:
                _sfile = store_write_text("analysis", selected.replace("_analysis.txt", "_topline.txt"), _sres)
                st.success(f"Done — model: {_smodel} — Saved: {_sfile.name}")
                st.download_button(
                    "Download topline (.txt)",
                    data=_sfile.read_bytes(),
                    file_name=_sfile.name,
                    mime="text/plain",
                )
            else:
                st.error(f"Failed: {_smodel}")

        st.divider()
        st.subheader("📁 Saved reports")
        reports = sorted(REPORTS_DIR.glob("*.pptx")) + sorted(REPORTS_DIR.glob("*.docx"))
        if reports:
            report_names = [f.name for f in reports]
            report_selected = st.selectbox(T("Pick a saved report", "သိမ်းထားတဲ့ report ရွေးပါ"), report_names)
            mime = ("application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    if report_selected.endswith(".pptx")
                    else "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.download_button(
                "Download saved report",
                data=(REPORTS_DIR / report_selected).read_bytes(),
                file_name=report_selected,
                mime=mime,
            )
        else:
            st.info(T("No saved report yet.", "သိမ်းထားတဲ့ report မရှိသေးပါ။"))

elif page == "Report Editor":
    st.title("✏️ Report Editor")
    st.caption(T("Edit in Google Docs — open the Doc, edit, then press Import back. The original file is untouched.", "Google Docs မှာ ပြင် — Doc ဖွင့်, ပြင်ပြီး, Import back နှိပ်. မူရင်း file ကို မထိဘူး."))

    saved = sorted(ANALYSIS_DIR.glob("*_analysis.txt"), key=lambda f: f.name)
    if not saved:
        st.info(T("No analysis yet — Analyze on the Market Analysis page first.", "Analysis မရှိသေးပါ — အရင် Market Analysis မှာ Analyze လုပ်ပါ။"))
    else:
        saved_names = [f.name for f in saved]
        # v3.2 fix: a widget's value may only be set BEFORE the widget is
        # created in the current run. Load buttons set a plain flag and rerun;
        # it is applied here, before the selectbox exists in this run.
        _load_target = st.session_state.pop("editor_load_target", None)
        if _load_target in saved_names:
            st.session_state["editor_select"] = _load_target
        selected = st.selectbox(T("Pick an analysis", "Analysis ရွေးပါ"), saved_names, key="editor_select")
        content = (ANALYSIS_DIR / selected).read_text(encoding="utf-8")
        if looks_like_html(content):
            content = html_to_text(content)
        edited = st.text_area("Report", value=content, height=400, key=f"ed_text_{selected}")
        stem = selected[:-len("_analysis.txt")]

        if docs_available():
            col_c, col_i = st.columns(2)
            if col_c.button("📄 Create Google Doc"):
                title = stem.replace("_", " ")
                with st.spinner(T("Creating the Google Doc...", "Google Doc ဖန်တီးနေပါတယ်...")):
                    try:
                        doc_id, url = drive_create_doc(title, edited)
                        st.session_state[f"editor_doc_id_{selected}"] = doc_id
                        st.session_state[f"editor_doc_url_{selected}"] = url
                        # v3.2 fix: fill the link field too — the text input
                        # keeps its old value unless its own key is set here.
                        st.session_state["editor_doc_input"] = url
                        st.success(T("Google Doc ready — press the link below to edit.", "Google Doc ရပြီ — အောက်က link နှိပ်ပြီး ပြင်ပါ."))
                    except Exception as e:
                        st.error(T(f"Could not create the Doc: {e}", f"Doc ဖန်တီးမရဘူး: {e}"))
            # v3.2 fix: when the selected analysis changes, show that
            # analysis's stored Doc link (or empty) instead of a stale one.
            if st.session_state.get("_editor_last_selected") != selected:
                st.session_state["_editor_last_selected"] = selected
                st.session_state["editor_doc_input"] = st.session_state.get(f"editor_doc_url_{selected}", "")
            doc_url_default = st.session_state.get(f"editor_doc_url_{selected}", "")
            doc_input = st.text_input(T("Google Doc link (or ID)", "Google Doc link (သို့မဟုတ် ID)"), value=doc_url_default, key="editor_doc_input")
            if st.session_state.get(f"editor_doc_url_{selected}"):
                st.link_button("🔗 Open Google Doc", st.session_state[f"editor_doc_url_{selected}"])
            if col_i.button("⬇️ Import back from Doc"):
                doc_id = _doc_id_from_input(doc_input)
                if not doc_id:
                    st.warning(T("Add the Doc link first.", "Doc link အရင် ထည့်ပါ."))
                else:
                    with st.spinner(T("Importing back from the Doc...", "Doc ကနေ ပြန်ယူနေပါတယ်...")):
                        try:
                            new_text = docs_read_text(doc_id)
                            out_name = f"{stem}_v{next_version_number(stem)}_analysis.txt"
                            store_write_text("analysis", out_name, new_text)
                            st.success(T(f"Saved: {out_name} — the original {selected} is unchanged.", f"Saved: {out_name} — မူရင်း {selected} က မပြောင်းဘူး."))
                            st.caption(T("On the Reports page you can pick the new version and export PPTX/DOCX.", "Reports page သွားရင် version အသစ် ရွေးပြီး PPTX/DOCX export လုပ်လို့ရပြီ။"))
                        except Exception as e:
                            st.error(f"Import မရဘူး: {e}")
        else:
            st.info(T("Google is not connected yet — Docs cannot open. Edit here and save as a new version.", "Google နဲ့ မချိတ်ရသေးဘူး — Doc ဖွင့်လို့ မရသေးဘူး. ဒီမှာ ပြင်ပြီး version အသစ် အဖြစ် သိမ်းလို့ရတယ်."))
            if st.button("💾 Save as new version"):
                out_name = f"{stem}_v{next_version_number(stem)}_analysis.txt"
                store_write_text("analysis", out_name, edited or "")
                st.success(T(f"Saved: {out_name} — the original {selected} is unchanged.", f"Saved: {out_name} — မူရင်း {selected} က မပြောင်းဘူး."))
                st.caption(T("On the Reports page you can pick the new version and export PPTX/DOCX.", "Reports page သွားရင် version အသစ် ရွေးပြီး PPTX/DOCX export လုပ်လို့ရပြီ။"))

        st.divider()
        st.subheader("Versions")
        versions = sorted(ANALYSIS_DIR.glob(f"{stem}_v*_analysis.txt"), key=lambda f: f.name)
        if versions:
            for v in versions:
                col_v, col_l = st.columns([4, 1])
                with col_v:
                    st.write(f"- {v.name}")
                with col_l:
                    if st.button("Load", key=f"ed_load_{v.name}"):
                        # v3.2 fix: do NOT write the selectbox key here — the
                        # selectbox already exists in this run (crash). Set a
                        # flag; the top of this page applies it, then reruns.
                        st.session_state["editor_load_target"] = v.name
                        st.rerun()
        else:
            st.caption(T("No new version yet.", "Version အသစ် မရှိသေးဘူး။"))

elif page == "Prompt Library":
    st.title("📚 Prompt Library")
    st.caption("Every prompt Gemini is given, in one place. Prompts live in the code — to change one, ask SM.")
    for _key, _eng in PROMPT_LIBRARY.items():
        with st.expander(f"{_eng['icon']} {_eng['name']}"):
            st.write(_eng["purpose"])
            st.caption(f"Used by: {_eng['used_by']}")
            st.code(_eng["template"], language="text")
            for _vname, _vtext in _eng.get("variants", {}).items():
                st.markdown(f"**Variant: {_vname}**")
                st.code(_vtext, language="text")

# v3.23: move on to the next stage from the bottom of any stage page.
if _stage_no and _stage_no < 5:
    st.divider()
    _nx = _stage_no + 1
    _nx_nm = _STAGE_NAME_EN[_nx]
    if st.button(
        T(f"Next: Stage {_nx} · {_nx_nm} →", f"နောက်တစ်ဆင့်: အဆင့် {_nx} · {_STAGE_MM.get(_nx_nm, _nx_nm)} →"),
        key=f"next_stage_{_stage_no}",
    ):
        st.session_state["page"] = st.session_state.get(f"stage_last_{_nx}", _STAGE_HOME[_nx])
        st.rerun()
