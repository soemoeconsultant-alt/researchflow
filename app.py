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


def get_transcript_files():
    files = [f for f in OUTPUT_DIR.glob("*.txt") if f.name != "failed_files.txt"]
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
    clean = re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9 _-]+", "", (name or "").strip())).strip()
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
#   redirect_uri = "https://researchflowmaster.streamlit.app/"
#   cookie_secret = "<random string>"
#   client_id = "<Web OAuth client id>"
#   client_secret = "<Web OAuth client secret>"
#   server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
#   [access]
#   owners = ["soemoe.consultant@gmail.com"]
#   allowed = ["soemoekyawww@gmail.com"]
# No [auth] in secrets (e.g. localhost) -> no login, open access as before.
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


def login_gate():
    """Require Google sign-in + allowlist when [auth] is configured.
    Call before any page renders. Returns the logged-in email (or "")."""
    if not _auth_configured():
        return ""
    if not st.user.is_logged_in:
        st.title("🔐 ResearchFlow — Log in")
        st.write("Google account နဲ့ log in ဝင်ပါ။")
        st.login()
        st.stop()
    email = current_user_email()
    owners, allowed = _allowed_emails()
    if email not in owners + allowed:
        st.title("🔐 ResearchFlow — Log in")
        st.error("ဒီ account နဲ့ ဝင်လို့ မရဘူး — owner က ခွင့်ပြုထားတဲ့ account ပဲ ဝင်ရတယ်။")
        st.write(f"Signed in as: {email}")
        if st.button("Log out"):
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


def sync_from_drive():
    """Drive is the source of truth: download Drive files into the local working folders.
    Text files are always refreshed from Drive; audio is downloaded only when missing
    or when the Drive copy is newer (audio files are big)."""
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
                if name.endswith((".mp3", ".wav", ".m4a")) and dest.exists():
                    dt = f.get("modifiedTime", "")
                    if dt:
                        drv = datetime.datetime.fromisoformat(dt.replace("Z", "+00:00")).timestamp()
                        if drv <= dest.stat().st_mtime:
                            continue
                drive_download(f["id"], dest)
            except Exception:
                continue


def store_write_text(dir_key, name, text):
    """Write a text file locally, then upload it to Drive when Drive mode is on."""
    path = LOCAL_DIR[dir_key] / name
    path.write_text(text, encoding="utf-8")
    if drive_available():
        try:
            drive_upload(path, DIR_TO_DRIVE[dir_key], name)
        except Exception as e:
            st.warning(f"Drive upload မအောင်မြင်ဘူး: {name} — {str(e)[:200]} (local မှာ သိမ်းပြီးပြီ။)")
    return path


def store_write_bytes(dir_key, name, data):
    """Write a binary file locally, then upload it to Drive when Drive mode is on."""
    path = LOCAL_DIR[dir_key] / name
    path.write_bytes(data)
    if drive_available():
        try:
            drive_upload(path, DIR_TO_DRIVE[dir_key], name)
        except Exception as e:
            st.warning(f"Drive upload မအောင်မြင်ဘူး: {name} — {str(e)[:200]} (local မှာ သိမ်းပြီးပြီ။)")
    return path


def store_upload_existing(dir_key, name):
    """Upload a local file created by another tool (ffmpeg split, PPTX/DOCX export)."""
    if not drive_available():
        return
    path = LOCAL_DIR[dir_key] / name
    if path.exists():
        try:
            drive_upload(path, DIR_TO_DRIVE[dir_key], name)
        except Exception as e:
            st.warning(f"Drive upload မအောင်မြင်ဘူး: {name} — {str(e)[:200]} (local မှာ သိမ်းပြီးပြီ။)")

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
    "## Questionnaire & Data Collection\n"
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
        "used_by": "Data Collection page (transcribe_audio)",
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
    sections = []
    current = None
    for line in text.splitlines():
        if line.startswith("## "):
            current = {"title": line[3:].strip(), "lines": []}
            sections.append(current)
        elif current is not None:
            line = line.strip()
            if line:
                if line.startswith(("- ", "* ")):
                    line = line[2:]
                current["lines"].append(line)
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
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        _set_bg(slide, t["bg"])
        slide.shapes.title.text = section["title"]
        _color_runs(slide.shapes.title.text_frame.paragraphs, t["title"])
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for i, line in enumerate(section["lines"]):
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


# Login gate (v3.3): runs before any page renders. No [auth] in secrets
# (e.g. localhost) -> open access, returns "".
_MY_EMAIL = login_gate()

# --- v3.5 look: dark navy sidebar, white cards, blue accent (design reference) ---
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
.stApp, .stApp * { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
:root {
  --rf-navy: #071B3A;
  --rf-navy-2: #0B2A52;
  --rf-blue: #1264C5;
  --rf-cyan: #18C6D9;
  --rf-teal: #16B89A;
  --rf-text: #10264A;
  --rf-muted: #64748B;
  --rf-border: #E5ECF5;
  --rf-surface: #FFFFFF;
}

/* App canvas */
.stApp {
  background: radial-gradient(ellipse at 85% 0%, rgba(24,198,217,.07), transparent 32%),
              radial-gradient(ellipse at 5% 12%, rgba(18,100,197,.06), transparent 28%),
              #F4F8FF;
  color: var(--rf-text);
}
.block-container {
  width: 100%;
  max-width: 1540px;
  padding: 1.35rem clamp(1.2rem, 2.2vw, 2.25rem) 2.5rem;
}

/* Compact sidebar: reserve more screen space for the workspace */
section[data-testid="stSidebar"] {
  width: 250px !important;
  min-width: 250px !important;
  max-width: 250px !important;
  background: linear-gradient(180deg, #071B3A 0%, #0B2A52 55%, #0E3A6E 100%);
  border-right: 1px solid rgba(255,255,255,.08);
}
section[data-testid="stSidebar"] > div {
  width: 250px !important;
}
@media (min-width: 901px) {
  section[data-testid="stSidebar"][aria-expanded="true"] {
    width: 250px !important;
    min-width: 250px !important;
    max-width: 250px !important;
  }
}

/* Dark AI sidebar */
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #071B3A 0%, #0B2A52 55%, #0E3A6E 100%);
  border-right: 1px solid rgba(255,255,255,.08);
}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
  gap: .45rem;
}
[data-testid="stSidebar"] h1 {
  font-size: 1.35rem;
  line-height: 1.25;
  margin-bottom: .15rem;
}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
  font-size: .76rem;
}
[data-testid="stSidebar"] .stButton > button {
  min-height: 2.15rem;
  padding: .28rem .65rem;
  border-radius: 8px;
  font-size: .84rem;
}
[data-testid="stSidebar"] [data-testid="stExpander"] details summary {
  padding-top: .45rem;
  padding-bottom: .45rem;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p { color: #F5F7FF; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { color: #AAB4D2; }
[data-testid="stSidebar"] button {
  background: rgba(255,255,255,.055); color: #F5F7FF;
  border: 1px solid rgba(255,255,255,.10); border-radius: 10px;
  transition: all .18s ease;
}
[data-testid="stSidebar"] button p { color: #F5F7FF; }
[data-testid="stSidebar"] button:hover {
  background: linear-gradient(100deg, rgba(18,100,197,.32), rgba(24,198,217,.30));
  border-color: rgba(24,198,217,.55); color: #FFFFFF;
}
[data-testid="stSidebar"] [data-testid="stExpander"] {
  background: rgba(255,255,255,.055); border: 1px solid rgba(255,255,255,.12); border-radius: 12px;
}
[data-testid="stSidebar"] [data-testid="stExpander"] p { color: #F5F7FF; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.12); }

/* Typography and section headings */
h1, h2, h3 { color: #10264A; letter-spacing: -.025em; }
h1 { font-weight: 800; }
h2, h3 { font-weight: 700; }
[data-testid="stCaptionContainer"] { color: var(--rf-muted); }

/* Hero banner — light blue-to-cyan gradient */
.rf-hero {
  position: relative; overflow: hidden;
  background: linear-gradient(115deg, #EAF3FF 0%, #DDF3F8 55%, #D9F5EF 100%);
  border: 1px solid #D5E5F7; border-radius: 18px;
  padding: 28px 32px; margin: 4px 0 16px;
  box-shadow: 0 10px 28px rgba(16,38,74,.08);
}
.rf-hero:after {
  content: ""; position: absolute; width: 300px; height: 300px; right: -70px; top: -170px;
  border-radius: 50%; background: radial-gradient(circle, rgba(24,198,217,.22), rgba(24,198,217,0) 68%);
  pointer-events: none;
}
.rf-hero h1 { color: #10264A; font-size: clamp(30px, 3.4vw, 44px); margin: 0 0 8px; font-weight: 800; }
.rf-hero h1 span.b { color: #1264C5; }
.rf-hero h1 span.t { color: #16B89A; }
.rf-hero p { color: #475569; font-size: 16px; margin: 0; max-width: 760px; }

/* Modern metric cards */
.rf-metric { position: relative; overflow: hidden; background: #FFFFFF; border: 1px solid var(--rf-border); border-radius: 16px; padding: 18px; box-shadow: 0 5px 18px rgba(16,38,74,.06); }
.rf-metric:before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, #1264C5, #18C6D9, #16B89A); }
.rf-metric-ico { width: 40px; height: 40px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 20px; margin-bottom: 10px; }
.rf-metric-label { color: #64748B; font-size: 13px; font-weight: 600; }
.rf-metric-value { color: #10264A; font-size: 30px; font-weight: 800; line-height: 1.1; }

/* Workflow cards and progress ring */
.rf-step {
  background: linear-gradient(160deg, #FFFFFF 0%, #F8FAFF 100%);
  border: 1px solid #E5E9F7; border-radius: 15px; padding: 14px 8px;
  text-align: center; min-height: 142px;
  box-shadow: 0 5px 16px rgba(16,38,74,.05);
  transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease;
}
.rf-step:hover { transform: translateY(-3px); border-color: #BFD9F2; box-shadow: 0 10px 24px rgba(18,100,197,.12); }
.rf-step .ico { width: 52px; height: 52px; border-radius: 50%; margin: 0 auto 8px; display: flex; align-items: center; justify-content: center; font-size: 24px; }
.rf-step .nm { font-weight: 750; color: #10264A; font-size: 13px; }
.rf-step .ds { color: #64748B; font-size: 11px; margin-top: 5px; }
.rf-step .st { color: #1264C5; font-size: 12px; font-weight: 700; margin-top: 8px; }
.rf-ring {
  width: 126px; height: 126px; border-radius: 50%; margin: 12px auto;
  background: conic-gradient(from 220deg, #1264C5 calc(var(--p)*1%), #18C6D9 0, #E9ECF7 0);
  display: flex; align-items: center; justify-content: center; position: relative;
  box-shadow: 0 8px 22px rgba(18,100,197,.12);
}
.rf-ring:before { content: ""; position: absolute; width: 94px; height: 94px; border-radius: 50%; background: #FFFFFF; }
.rf-ring span { position: relative; z-index: 1; font-size: 25px; font-weight: 800; color: #10264A; }

/* Recent activity table with status badges */
.rf-table { width: 100%; border-collapse: collapse; font-size: 14px; }
.rf-table th { text-align: left; color: #64748B; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; padding: 10px 12px; border-bottom: 1px solid var(--rf-border); }
.rf-table td { padding: 11px 12px; border-bottom: 1px solid #EEF2F8; color: #10264A; vertical-align: middle; }
.rf-table tr:last-child td { border-bottom: none; }
.rf-badge { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; background: #E2F5EE; color: #0E7A5F; }

/* Guide + system architecture */
.rf-guide { background: #FFFFFF; border: 1px solid var(--rf-border); border-radius: 16px; padding: 20px 22px; box-shadow: 0 5px 18px rgba(16,38,74,.05); margin: 0 0 16px; }
.rf-arch { background: #F4F8FF; border: 1px solid var(--rf-border); border-radius: 12px; padding: 14px 16px; text-align: center; font-weight: 700; color: #10264A; }
.rf-arch small { display: block; font-weight: 500; color: #64748B; font-size: 12px; margin-top: 2px; }
.rf-arch-arrow { text-align: center; color: #94A3B8; font-size: 18px; line-height: 1.2; padding: 2px 0; }

/* Buttons and form controls */
.stButton > button, .stDownloadButton > button, [data-testid="stLinkButton"] a {
  border-radius: 11px; border: 1px solid #DCE3F5; font-weight: 650; font-size: 14px;
  transition: all .18s ease; box-shadow: 0 2px 7px rgba(35,48,105,.045);
}
.stButton > button[kind="primary"] {
  color: #FFFFFF; border: 0;
  background: linear-gradient(105deg, #1264C5 0%, #18C6D9 100%);
  box-shadow: 0 6px 16px rgba(18,100,197,.22);
}
.stButton > button:hover, .stDownloadButton > button:hover { border-color: #7FB8E8; transform: translateY(-1px); }
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea,
[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
[data-testid="stNumberInput"] input {
  border-radius: 10px; border-color: #DDE3F1; background: #FFFFFF;
}
[data-testid="stFileUploader"] section {
  border: 1.5px dashed #BFD9F2; border-radius: 15px; background: #FAFBFF;
}

/* Tables, tabs, expanders, alerts */
[data-testid="stDataFrame"] { border: 1px solid var(--rf-border); border-radius: 13px; overflow: hidden; }
[data-testid="stTabs"] button[role="tab"] { font-weight: 650; }
[data-testid="stTabs"] button[aria-selected="true"] { color: #1264C5; }
[data-testid="stExpander"] { border: 1px solid var(--rf-border); border-radius: 13px; background: #FFFFFF; }
[data-testid="stAlert"] { border-radius: 12px; }
hr { border-color: #E5E9F3; }

@media (max-width: 900px) {
  section[data-testid="stSidebar"],
  section[data-testid="stSidebar"] > div {
    width: min(82vw, 280px) !important;
    min-width: min(82vw, 280px) !important;
    max-width: min(82vw, 280px) !important;
  }
  .block-container { padding: 1rem 1rem 2rem; }
  .rf-hero { padding: 23px 20px; }
  .rf-step { min-height: 132px; }
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# --- Shell (v3.4): 6-stage nav, one Project = one container. Modern, professional,
# minimalist: white, one blue accent, generous spacing, no extra decoration.
STAGES = [
    ("1 · Client Brief", "Client Brief"),
    ("2 · Data Collection", "Data Collection"),
    ("3 · Fieldworks", "Fieldworks"),
    ("4 · Data Cleansing", "Data Cleansing"),
    ("5 · Analysis & Reports", "Market Analysis"),
    ("6 · Presentation", "Presentation"),
]
PAGE_STAGE = {
    "Client Brief": 1,
    "Data Collection": 2, "Batch Processing": 2, "Upload & Split": 2,
    "Fieldworks": 3, "Transcripts": 3,
    "Data Cleansing": 4,
    "Market Analysis": 5, "Quantitative Analysis": 5, "Reports": 5, "Report Editor": 5,
    "Presentation": 6,
}

if "page" not in st.session_state:
    st.session_state["page"] = "Dashboard"

st.sidebar.title("📊 ResearchFlow")
st.sidebar.caption("Market Research AI")

# Project selector: list once per session (Drive call), refresh after creating one.
if "project_list" not in st.session_state:
    st.session_state["project_list"] = list_projects()
projects = st.session_state["project_list"]
current = st.session_state.get("active_project", "Default")
if current not in projects:
    projects = [current] + projects
if "project_select_target" in st.session_state:
    st.session_state["project_select"] = st.session_state.pop("project_select_target")
selected = st.sidebar.selectbox(
    "Project", projects,
    index=projects.index(current),
    key="project_select",
)
if selected != current:
    st.session_state["active_project"] = selected
    st.session_state["drive_synced"] = False
    st.rerun()
refresh_project_dirs(selected)

with st.sidebar.expander("＋ Project အသစ်"):
    new_name = st.text_input("Project နာမည်", key="new_project_name")
    if st.button("Project ဖွင့်", key="create_project"):
        created = ensure_project(new_name)
        if created:
            if created not in st.session_state["project_list"]:
                st.session_state["project_list"].append(created)
            st.session_state["project_select_target"] = created
            st.rerun()
        else:
            st.warning("Project နာမည် မှန်မှန်ရိုက်ပါ။")

_PAGE_ICONS = {
    "Dashboard": "🏠", "Client Brief": "📋",
    "Data Collection": "🎙️", "Batch Processing": "📁", "Upload & Split": "✂️",
    "Fieldworks": "🔴", "Transcripts": "📝", "Data Cleansing": "🧹",
    "Market Analysis": "📊", "Quantitative Analysis": "📈", "Reports": "📑",
    "Report Editor": "✏️", "Presentation": "🎤", "Prompt Library": "📚",
}
_PAGE_ORDER = [
    "Dashboard", "Client Brief",
    "Data Collection", "Batch Processing", "Upload & Split",
    "Fieldworks", "Transcripts", "Data Cleansing",
    "Market Analysis", "Quantitative Analysis", "Reports", "Report Editor",
    "Presentation", "Prompt Library",
]

for item in _PAGE_ORDER:
    _active = st.session_state["page"] == item
    if st.sidebar.button(
        f"{_PAGE_ICONS.get(item, '•')} {item}", key=f"nav_{item}",
        use_container_width=True, type="primary" if _active else "secondary",
    ):
        st.session_state["page"] = item
        st.rerun()

st.sidebar.caption(
    "💾 Storage: Google Drive (OAuth)" if DRIVE_AUTH == "oauth"
    else "💾 Storage: Google Drive" if DRIVE_AUTH == "service"
    else "💾 Storage: Local (PC)"
)

if _MY_EMAIL:
    st.sidebar.caption(f"👤 {_MY_EMAIL}")
    if st.sidebar.button("Log out"):
        st.logout()

page = st.session_state["page"]

# Top bar: functional global search across this project's files + projects.
_search_q = st.text_input(
    "Search", key="global_search", label_visibility="collapsed",
    placeholder="🔍 Search projects, files, or features...",
)
if _search_q:
    _q = _search_q.strip().lower()
    _hits = []
    for _p in st.session_state.get("project_list", ["Default"]):
        if _q in _p.lower():
            _hits.append((f"📁 Project: {_p}", _p, "project"))
    _scan = ((AUDIO_DIR, "Data Collection"), (OUTPUT_DIR, "Transcripts"),
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
        st.caption("No matches in this project.")

# Navigation lives in the sidebar only (the top stepper row was removed in v3.9).

def stage_status():
    """(label, description, done) per stage, from real file counts only.
    A stage counts as done when its output files exist in this project."""
    audio_n = len([f for f in AUDIO_DIR.glob("*") if f.is_file()])
    tr_n = len(get_transcript_files())
    an_n = len(list(ANALYSIS_DIR.glob("*_analysis.txt")))
    rep_n = len(list(REPORTS_DIR.glob("*.pptx"))) + len(list(REPORTS_DIR.glob("*.docx")))
    return [
        ("1 · Client Brief", "Define objectives & scope", False),
        ("2 · Data Collection", "Gather audio & secondary data", audio_n > 0),
        ("3 · Fieldworks", "Conduct interviews & FGDs", tr_n > 0),
        ("4 · Data Cleansing", "Clean & prepare data", False),
        ("5 · Analysis & Reports", "Generate insights & reports", an_n > 0),
        ("6 · Presentation", "Present to clients", rep_n > 0),
    ]


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
            "audio": "Audio file added",
            "transcript": "Transcription completed",
            "analysis": "Analysis saved",
            "report": "Report exported",
        }[kind]
        out.append((dt.strftime("%b %d, %Y %I:%M %p"), name, label, "Completed"))
    return out


# Drive is the source of truth: sync once per session, then pages work on local files.
if DRIVE_MODE and not st.session_state.get("drive_synced"):
    with st.spinner("Google Drive နဲ့ sync လုပ်နေပါတယ်..."):
        sync_from_drive()
    st.session_state["drive_synced"] = True

if page == "Dashboard":

    st.markdown(
        '<div class="rf-hero"><h1>Welcome to<br><span class="b">Research</span>'
        '<span class="t">Flow</span></h1>'
        "<p>Transform your interview data into actionable insights with our "
        "professional market research platform.</p></div>",
        unsafe_allow_html=True,
    )
    _hc1, _hc2, _hc3 = st.columns([1, 1, 4])
    with _hc1:
        if st.button("▶ Start New Project", key="hero_new_project", type="primary"):
            st.session_state["page"] = "Client Brief"
            st.rerun()
    with _hc2:
        if st.button("📖 View Guide", key="hero_view_guide"):
            st.session_state["show_guide"] = not st.session_state.get("show_guide", False)
            st.rerun()
    if st.session_state.get("show_guide"):
        st.markdown(
            '<div class="rf-guide"><h3>📖 How ResearchFlow works</h3>'
            "<p>One project flows through 6 stages — brief, collect, fieldwork, "
            "cleanse, analyse, present. Files live in Google Drive "
            "(ResearchFlow / project / audio, transcripts, notes, reports) and "
            "this app works on those same files.</p></div>",
            unsafe_allow_html=True,
        )
        st.markdown("**System Architecture**")
        st.markdown(
            '<div class="rf-arch">Research Project Management<small>Client Brief • Project Setup • Objectives</small></div>'
            '<div class="rf-arch-arrow">↓</div>',
            unsafe_allow_html=True,
        )
        _ac1, _ac2 = st.columns(2)
        with _ac1:
            st.markdown('<div class="rf-arch">Data Collection<small>Interviews • Surveys • FGDs</small></div>', unsafe_allow_html=True)
        with _ac2:
            st.markdown('<div class="rf-arch">Batch Processing<small>Audio • Video • Documents</small></div>', unsafe_allow_html=True)
        st.markdown('<div class="rf-arch-arrow">↓</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="rf-arch">AI Processing Engine<small>Transcription • Data Cleansing • Coding</small></div>'
            '<div class="rf-arch-arrow">↓</div>'
            '<div class="rf-arch">Research Intelligence<small>Qualitative • Quantitative • Market Analysis</small></div>'
            '<div class="rf-arch-arrow">↓</div>'
            '<div class="rf-arch">AI Reports &amp; Presentation<small>Insights • Recommendations • Client Reports</small></div>',
            unsafe_allow_html=True,
        )
        st.caption("Coding, Surveys and Video are part of the target architecture — not yet built.")

    _projects = st.session_state.get("project_list", ["Default"])
    _transcripts = get_transcript_files()
    _reports = list(REPORTS_DIR.glob("*.pptx")) + list(REPORTS_DIR.glob("*.docx"))
    _audio = [f for f in AUDIO_DIR.glob("*") if f.is_file()]
    _cards = [
        ("📁", "Total Projects", len(_projects), "#1264C5"),
        ("📝", "Transcripts", len(_transcripts), "#16B89A"),
        ("📊", "Research Reports", len(_reports), "#18C6D9"),
        ("🎙️", "Audio Files", len(_audio), "#071B3A"),
    ]
    _cols = st.columns(4)
    for _i, (_ico, _lbl, _val, _c) in enumerate(_cards):
        with _cols[_i]:
            st.markdown(
                f'<div class="rf-metric"><div class="rf-metric-ico" style="background:{_c}1A">{_ico}</div>'
                f'<div class="rf-metric-label">{_lbl}</div>'
                f'<div class="rf-metric-value">{_val}</div></div>',
                unsafe_allow_html=True,
            )

    _stages = stage_status()
    _done = sum(1 for _, _, d in _stages if d)
    _pct = round(_done / len(_stages) * 100)
    _icons = ["📋", "🎙️", "👥", "🧹", "📊", "🎤"]
    _step_colors = ["#1264C5", "#18C6D9", "#16B89A", "#1264C5", "#18C6D9", "#16B89A"]

    st.subheader("🚀 Research Workflow")
    st.caption("Follow the complete research process from brief to presentation")
    _cols = st.columns([3, 1, 3, 1, 3, 1, 3, 1, 3, 1, 3])
    _ci = 0
    for _i, (_label, _desc, _d) in enumerate(_stages):
        with _cols[_ci]:
            _nm = _label.split(" · ", 1)[1]
            _st = "✅ Done" if _d else "○ Not yet"
            st.markdown(
                f'<div class="rf-step"><div class="ico" style="background:{_step_colors[_i]}1A">{_icons[_i]}</div>'
                f'<div class="nm">{_i + 1}. {_nm}</div>'
                f'<div class="ds">{_desc}</div>'
                f'<div class="st">{_st}</div></div>',
                unsafe_allow_html=True,
            )
        _ci += 1
        if _i < len(_stages) - 1:
            with _cols[_ci]:
                st.markdown(
                    '<div style="text-align:center; font-size:22px; color:#94A3B8;'
                    ' padding-top:55px;">→</div>',
                    unsafe_allow_html=True,
                )
            _ci += 1

    st.subheader("💡 Project Progress")
    _pc1, _pc2 = st.columns([1, 2])
    with _pc1:
        st.markdown(
            f'<div class="rf-ring" style="--p:{_pct}"><span>{_pct}%</span></div>',
            unsafe_allow_html=True,
        )
        st.caption(f"{_done} of {len(_stages)} stages have files so far")
    with _pc2:
        _s1, _s2 = st.columns(2)
        for _i, (_label, _desc, _d) in enumerate(_stages):
            with (_s1 if _i < 3 else _s2):
                st.write(("✅ " if _d else "○ ") + _label)

    st.subheader("🕐 Recent Activity")
    _act = recent_activity()
    if _act:
        _rows = "".join(
            f"<tr><td>{html.escape(_t)}</td><td>{html.escape(_f)}</td><td>{html.escape(_a)}</td>"
            f'<td><span class="rf-badge">{html.escape(_s)}</span></td></tr>'
            for _t, _f, _a, _s in _act
        )
        st.markdown(
            '<table class="rf-table"><tr><th>Date &amp; Time</th><th>File / Project</th>'
            "<th>Activity</th><th>Status</th></tr>" + _rows + "</table>",
            unsafe_allow_html=True,
        )
    else:
        st.info("No activity yet — record or upload audio to begin.")

elif page == "Client Brief":
    # Stage 1 shell page: project info now; the Brief form (Scope / Objective /
    # Methodology / Result Outcome) + Communication + Meetings come in the Stage 1 build.
    st.title("📋 Client Brief")
    st.write(f"Project: **{current_project()}**")

    col1, col2, col3 = st.columns(3)
    col1.metric("Audio Files", len(list(AUDIO_DIR.glob("*"))))
    col2.metric("Transcripts", len(get_transcript_files()))
    col3.metric("Research Reports", len(list(REPORTS_DIR.glob("*.pptx"))) + len(list(REPORTS_DIR.glob("*.docx"))))

    st.divider()
    st.info(
        "Brief form (Scope / Objective / Methodology / Result Outcome) + "
        "Communication + Meetings — Stage 1 build မှာ လာမယ်။"
    )

    st.subheader("＋ Project အသစ်")
    _new = st.text_input("Project နာမည်", key="brief_new_project_name")
    if st.button("Project ဖွင့်", key="brief_create_project"):
        _created = ensure_project(_new)
        if _created:
            if _created not in st.session_state["project_list"]:
                st.session_state["project_list"].append(_created)
            st.session_state["project_select_target"] = _created
            st.rerun()
        else:
            st.warning("Project နာမည် မှန်မှန်ရိုက်ပါ။")

elif page == "Data Cleansing":
    # Stage 4 shell page: file counts now; per-file check/clean/unclear/issue
    # status comes in the Stage 4 build.
    st.title("🧹 Data Cleansing")
    st.write(f"Project: **{current_project()}**")

    col1, col2, col3 = st.columns(3)
    col1.metric("Audio Files", len(list(AUDIO_DIR.glob("*"))))
    col2.metric("Transcripts", len(get_transcript_files()))
    col3.metric("Notes", len(list(NOTES_DIR.glob("*"))))

    st.divider()
    st.info(
        "Raw input တစ်ခုချင်းကို စစ်တာ (checked / clean / unclear / issue) — "
        "Stage 4 build မှာ လာမယ်။"
    )

elif page == "Presentation":
    # Stage 6 shell page: client view + per-project share link come in the Stage 6 build.
    st.title("🎤 Presentation")
    st.write(f"Project: **{current_project()}**")

    st.divider()
    st.info(
        "Client view + per-project share link — Stage 6 build မှာ လာမယ်။ "
        "အခု Reports စာမျက်နှာမှာ PPTX / DOCX ထုတ်လို့ရတယ်။"
    )

elif page == "Data Collection":
    st.title("🎙️ Data Collection")

    uploaded = st.file_uploader("Audio file တင်ပါ (mp3 / wav / m4a)", type=["mp3", "wav", "m4a"])
    audio_path = None
    if uploaded is not None:
        audio_path = AUDIO_DIR / uploaded.name
        store_write_bytes("audio", uploaded.name, uploaded.getbuffer())
        st.caption(f"Saved: {uploaded.name}")
    else:
        existing = sorted(AUDIO_DIR.glob("*.mp3")) + sorted(AUDIO_DIR.glob("*.wav")) + sorted(AUDIO_DIR.glob("*.m4a"))
        if existing:
            audio_names = [f.name for f in existing]
            audio_selected = st.selectbox("ရှိပြီးသား audio ရွေးပါ", audio_names)
            audio_path = AUDIO_DIR / audio_selected
        else:
            st.info("Audio file မရှိသေးပါ — အရင် upload လုပ်ပါ။")

    if audio_path is not None and st.button("Transcribe"):
        with st.spinner("Transcribing... (audio ကြာရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)"):
            text, model = transcribe_audio(audio_path)
        if text:
            out_file = store_write_text("transcripts", audio_path.stem + ".txt", text)
            st.success(f"Done — model: {model}")
            st.caption(f"Saved: {out_file.name} — Transcripts page မှာ ပြန်ဖွင့်လို့ရတယ်။")
        else:
            st.error(f"Failed: {model}")

elif page == "Batch Processing":
    st.title("📁 Batch Processing")

    audio_files = sorted(AUDIO_DIR.glob("*.mp3")) + sorted(AUDIO_DIR.glob("*.wav")) + sorted(AUDIO_DIR.glob("*.m4a"))
    if not audio_files:
        st.info("Audio file မရှိသေးပါ — Data Collection page မှာ upload လုပ်ပါ။")
    else:
        pending = []
        done = []
        for f in audio_files:
            if (OUTPUT_DIR / (f.stem + ".txt")).exists():
                done.append(f.name)
            else:
                pending.append(f.name)
        st.write(f"စုစုပေါင်း: {len(audio_files)} | ပြီးပြီ: {len(done)} | ကျန်သေးတယ်: {len(pending)}")
        if pending:
            st.write("ကျန်တဲ့ file တွေ:")
            for name in pending:
                st.write(f"- {name}")
        if pending and st.button("Transcribe next 5"):
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
                        st.warning("429 — quota ပြည့်နေတယ်, batch ရပ်လိုက်ပြီ။ နောက်မှ ပြန်လုပ်။")
                        break
                progress.progress((i + 1) / len(batch))
                if i < len(batch) - 1:
                    time.sleep(60)
            st.info("Batch ပြီးပြီ — နောက် 5 ခု လုပ်ချင်ရင် 15 မိနစ် နားပြီးမှ ခလုတ် ပြန်နှိပ်ပါ။")
        elif not pending:
            st.success("Audio အားလုံး transcript ပြီးပြီ။")

elif page == "Fieldworks":
    st.title("🎙️ Fieldworks")
    st.caption("Browser mic နဲ့ chunk တစ်ခုချင်း record လုပ် → auto-transcribe. 15 မိနစ် တစ်ပိုင်း. 1 chunk = 1 request (2h FGD ≈ 8 chunks).")
    st.info("Browser က mic ခွင့်တောင်းရင် Allow နှိပ်ပါ. လူအများ တစ်ပြိုင်နက် ပြောရင် speaker ခွဲတာ တိကျမှု လျော့နိုင်တယ် — [unclear] လို့ မှတ်ထားတာကို ပြန်စစ်ပါ.")

    session_name = st.text_input("Session နာမည်", value="fgd_01", key="fgd_session")
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
        st.info("Chunk မရှိသေးဘူး — အောက်မှာ record လုပ်ပါ။")

    audio_data = st.audio_input("🎙️ Record a chunk (15 မိနစ် လောက်)")
    if audio_data is not None:
        audio_bytes = audio_data.read()
        audio_hash = hashlib.md5(audio_bytes).hexdigest()
        if st.session_state.get("last_recorded_hash") != audio_hash:
            next_n = (chunks[-1][0] + 1) if chunks else 1
            chunk_path = AUDIO_DIR / f"{session_name}({next_n}).wav"
            store_write_bytes("audio", chunk_path.name, audio_bytes)
            st.session_state["last_recorded_hash"] = audio_hash
            st.session_state["last_chunk"] = chunk_path.name
            st.success(f"Saved: {chunk_path.name}")
            txt_path = OUTPUT_DIR / (chunk_path.stem + ".txt")
            with st.spinner(f"Transcribing chunk {next_n}... (audio ကြာရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)"):
                text, model = transcribe_audio(chunk_path)
            if text:
                store_write_text("transcripts", txt_path.name, text)
                st.success(f"Done — model: {model}")
            else:
                st.error(f"Failed: {model} — အောက်က Transcribe pending နဲ့ ပြန်လုပ်လို့ရတယ်။")
            st.rerun()
        else:
            st.caption(f"ဒီ chunk က သိမ်းပြီးသား: {st.session_state.get('last_chunk')}")
            last_name = st.session_state.get("last_chunk", "")
            txt_path = OUTPUT_DIR / (Path(last_name).stem + ".txt")
            if last_name and txt_path.exists():
                with st.expander("Transcript ကြည့်ပါ"):
                    st.text(txt_path.read_text(encoding="utf-8"))

    pending_chunks = [(n, f) for n, f in session_chunk_files() if not (OUTPUT_DIR / (f.stem + ".txt")).exists()]
    if pending_chunks and st.button("Transcribe pending chunks"):
        for n, f in pending_chunks:
            with st.spinner(f"Transcribing chunk {n}..."):
                text, model = transcribe_audio(f)
            if text:
                store_write_text("transcripts", f.stem + ".txt", text)
                st.success(f"Done — chunk {n} ({model})")
            else:
                st.error(f"Failed — chunk {n}: {model}")
                if "429" in model:
                    st.warning("429 — quota ပြည့်နေတယ်, နောက်မှ ပြန်လုပ်။")
                    break

elif page == "Upload & Split":
    st.title("✂️ Upload & Split")
    st.caption("အရှည်ကြီး audio (device နဲ့ record လုပ်ထားတာ) ကို 15 မိနစ် အပိုင်းတွေ ခွဲ → ပြီးမှ Batch Processing နဲ့ transcribe လုပ်.")
    if DRIVE_MODE:
        st.info("Split လုပ်ဖို့ ffmpeg လိုတယ် (mp3/m4a အတွက်). WAV ဆိုရင် ffmpeg မရှိလည်း ရတယ်. Upload လုပ်တဲ့ file က Google Drive (ResearchFlow/Default/audio) ထဲ တိုက်ရိုက် သိမ်းမယ်.")
    else:
        st.info("Split လုပ်ဖို့ ffmpeg လိုတယ် (mp3/m4a အတွက်). WAV ဆိုရင် ffmpeg မရှိလည်း ရတယ်. Drive ကနေ တင်တာက နောက်အဆင့် — အခု website ကို တိုက်ရိုက် တင်ပါ.")

    src_mode = st.radio("Audio ဘယ်ကယူမလဲ", ["Upload လုပ်မယ်", "ရှိပြီးသား file ရွေးမယ်"], horizontal=True)
    src_path = None
    if src_mode == "Upload လုပ်မယ်":
        up = st.file_uploader("Audio file တင်ပါ (mp3 / wav / m4a)", type=["mp3", "wav", "m4a"], key="split_upload")
        if up is not None:
            src_path = AUDIO_DIR / up.name
            if not src_path.exists():
                store_write_bytes("audio", up.name, up.getbuffer())
            st.caption(f"Saved: {up.name}")
    else:
        existing = sorted(AUDIO_DIR.glob("*.mp3")) + sorted(AUDIO_DIR.glob("*.wav")) + sorted(AUDIO_DIR.glob("*.m4a"))
        if existing:
            sel = st.selectbox("Audio ရွေးပါ", [f.name for f in existing], key="split_existing")
            src_path = AUDIO_DIR / sel
        else:
            st.info("Audio file မရှိသေးပါ — အရင် upload လုပ်ပါ။")

    if src_path is not None:
        dur = audio_duration(src_path)
        if dur:
            st.caption(f"Duration: {dur / 60:.1f} မိနစ်")
        chunk_min = st.selectbox("တစ်ပိုင်း ဘယ်လောက် မိနစ်", [10, 15], index=1, key="split_len")
        if st.button("Split into chunks"):
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
                st.caption(f"မူရင်း file ကို audio/originals/ ထဲ ရွှေ့ပြီးပြီ — Batch Processing မှာ chunks တွေ ပေါ်နေပြီ။")
            else:
                st.warning("Chunk မထွက်ဘူး — file က chunk တစ်ပိုင်းထက် တိုနေလို့ဖြစ်နိုင်တယ်။")

elif page == "Transcripts":
    st.title("📝 Transcript Management")

    st.subheader("🔗 Merge chunks — full transcript")
    groups = get_chunk_groups()
    multi_groups = {k: v for k, v in groups.items() if len(v) > 1}
    if multi_groups:
        for key, files_in_group in sorted(multi_groups.items()):
            merged = MERGED_DIR / f"{key}_full.txt"
            status = "✅ merged" if merged.exists() else f"{len(files_in_group)} chunks"
            with st.expander(f"{key}: {status}"):
                for i, f in enumerate(files_in_group, start=1):
                    st.write(f"{i}. {f.name}")
        if st.button("Merge all into full transcripts"):
            for key, files_in_group in sorted(multi_groups.items()):
                parts = []
                for i, f in enumerate(files_in_group, start=1):
                    parts.append(f"--- chunk {i} ({f.name}) ---\n")
                    parts.append(f.read_text(encoding="utf-8"))
                out = store_write_text("merged", f"{key}_full.txt", "\n\n".join(parts))
            st.success("Merged — full transcripts ready in output/merged/")
    else:
        st.info("Chunk အများကြီး ရှိတဲ့ interview မရှိဘူး။")

    st.divider()

    st.subheader("📄 Full transcripts (merged)")
    merged_files = sorted(MERGED_DIR.glob("*_full.txt"), key=lambda f: f.name)
    if merged_files:
        for file in merged_files:
            with st.expander(file.name):
                st.text(file.read_text(encoding="utf-8"))
    else:
        st.info("Merge မလုပ်ရသေးဘူး — အပေါ်က Merge ခလုတ် နှိပ်ပါ။")

    st.divider()

    files = get_transcript_files()

    if files:
        for file in files:
            with st.expander(file.name):
                st.text(file.read_text(encoding="utf-8"))
    else:
        st.warning("Transcript ဖိုင် မတွေ့ရှိသေးပါ။")

elif page == "Market Analysis":
    st.title("📊 Market Research Analysis")

    # Analysis uses FULL transcripts only — never chunks.
    full_files, pending = get_full_transcripts()
    if not full_files:
        st.warning("Full transcript မရှိသေးပါ — အရင် Transcripts page မှာ Merge လုပ်ပါ။")
    else:
        names = [f.name for f in full_files]
        file_map = {f.name: f for f in full_files}
        if pending:
            st.caption("Merge မလုပ်ရသေးဘူး: " + ", ".join(f"{k} ({n} chunks)" for k, n in pending) + " — Transcripts page မှာ Merge လုပ်ပါ။")

        st.subheader("📋 Analysis Planner")
        st.caption("Analysis မလုပ်ခင်: Brief + Discussion Guide + Questionnaire + transcript ပေါင်းကြည့်ပြီး plan အရင် ဆွဲမယ်။")
        brief_text = st.text_area("Client Brief / Objective", key="brief_text", height=150)
        dg_text = st.text_area("Discussion Guide (sections / questions)", key="dg_text", height=150)
        questionnaire_text = st.text_area("Questionnaire (survey မေးခွန်းတွေ — paste လုပ်)", key="questionnaire_text", height=150)
        collection_tool = st.selectbox(
            "Data ကောက်တဲ့ tool",
            ["Google Forms", "SurveyMonkey", "Kobo Collect", "အခြား"],
            key="collection_tool",
        )
        plan_transcripts = st.multiselect(
            "Plan အတွက် transcript တွေ ရွေးပါ",
            names,
            key="plan_transcripts",
        )
        st.caption("Plan အတွက် ရွေးတဲ့ transcript တွေနဲ့ အောက်က Analyze အတွက် ရွေးတာ တူရမယ်။")
        if st.button("Plan analysis"):
            if not brief_text.strip() and not dg_text.strip() and not questionnaire_text.strip():
                st.warning("Brief / Discussion Guide / Questionnaire အနည်းဆုံး တစ်ခု ထည့်ပါ။")
            elif not plan_transcripts:
                st.warning("Plan အတွက် transcript အနည်းဆုံး ၁ ခု ရွေးပါ။")
            else:
                combined_plan = ""
                for name in plan_transcripts:
                    combined_plan += f"\n\n=== {name} ===\n"
                    combined_plan += file_map[name].read_text(encoding="utf-8")
                plan_prompt = (
                    PROMPT_LIBRARY["planning"]["template"]
                    + "\n\nClient Brief / Objective:\n" + brief_text + "\n\nDiscussion Guide:\n" + dg_text + "\n\nQuestionnaire:\n" + questionnaire_text + "\n\nData collection tool: " + collection_tool + "\n\nTranscripts:\n" + combined_plan
                )
                with st.spinner("Planning... (transcript များရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)"):
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
            plan_selected = st.selectbox("သိမ်းထားတဲ့ plan ရွေးပါ", plan_names)
            if st.button("Load plan"):
                st.session_state["current_plan"] = (ANALYSIS_DIR / plan_selected).read_text(encoding="utf-8")
                st.session_state["current_plan_name"] = plan_selected
                st.success(f"Loaded: {plan_selected}")

        if st.session_state.get("current_plan"):
            st.caption(f"Current plan: {st.session_state.get('current_plan_name')}")
            with st.expander("Plan ကြည့်ပါ"):
                st.text(st.session_state["current_plan"])

        st.divider()
        has_plan = st.session_state.get("current_plan") is not None
        use_plan = st.checkbox(
            "Approved plan အတိုင်း analyse လုပ်မယ်",
            value=has_plan,
            disabled=not has_plan,
        )
        mode = st.radio("Analysis ပုံစံ", ["တစ်ခုတည်း", "အများကြီး (multi)"], horizontal=True)
        if mode == "တစ်ခုတည်း":
            default_idx = names.index("interview_01_full.txt") if "interview_01_full.txt" in names else 0
            selected = st.selectbox("Transcript ရွေးပါ", names, index=default_idx)
            if st.button("Analyze"):
                text = file_map[selected].read_text(encoding="utf-8")
                if use_plan and st.session_state.get("current_plan"):
                    prompt = plan_based_prompt(st.session_state["current_plan"], text, 1)
                else:
                    prompt = PROMPT_LIBRARY["qualitative"]["template"] + "\n\nTranscript:\n" + text
                with st.spinner("Analyzing... (ခဏစောင့်ပါ)"):
                    result, model = ask_gemini(prompt)
                if result:
                    st.success(f"Done — model: {model}")
                    analysis_file = store_write_text("analysis", selected.replace(".txt", "_analysis.txt"), result)
                    st.session_state["last_analysis"] = analysis_file.name
                    st.caption(f"Saved: {analysis_file.name} — အောက်က Saved analyses မှာ ပြန်ဖွင့်လို့ရတယ်။")
                else:
                    st.error(f"Failed: {model}")
        else:
            selected_multi = st.multiselect(
                "Transcript တွေ ရွေးပါ (အများဆုံး 6 ခု လောက်ရွေး — များလွန်းရင် ကြာမယ်)",
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
                with st.spinner("Analyzing... (transcript များရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)"):
                    result, model = ask_gemini(prompt)
                if result:
                    st.success(f"Done — model: {model}")
                    short_ids = short_ids_for(selected_multi)
                    analysis_file = store_write_text("analysis", f"multi_{'-'.join(short_ids)}_analysis.txt", result)
                    st.session_state["last_analysis"] = analysis_file.name
                    st.caption(f"Saved: {analysis_file.name} — အောက်က Saved analyses မှာ ပြန်ဖွင့်လို့ရတယ်။")
                else:
                    st.error(f"Failed: {model}")

        st.divider()
        st.subheader("💾 Saved analyses")
        saved = sorted(ANALYSIS_DIR.glob("*_analysis.txt"), key=lambda f: f.name)
        if saved:
            saved_names = [f.name for f in saved]
            last = st.session_state.get("last_analysis")
            default_saved = saved_names.index(last) if last in saved_names else 0
            saved_selected = st.selectbox("သိမ်းထားတာ ရွေးပါ", saved_names, index=default_saved)
            saved_text = (ANALYSIS_DIR / saved_selected).read_text(encoding="utf-8")
            if looks_like_html(saved_text):
                saved_text = html_to_text(saved_text)
            st.text(saved_text)
            st.download_button(
                "Download saved analysis (.txt)",
                data=(ANALYSIS_DIR / saved_selected).read_bytes(),
                file_name=saved_selected,
                mime="text/plain",
            )
        else:
            st.info("သိမ်းထားတဲ့ analysis မရှိသေးပါ။")

elif page == "Quantitative Analysis":
    st.title("📊 Quantitative Analysis")
    st.caption("Survey data (Google Forms / SurveyMonkey / Kobo Collect export — CSV / XLSX) → counts, %, charts, cross-tab.")

    uploaded = st.file_uploader("Survey file တင်ပါ (CSV / XLSX)", type=["csv", "xlsx"])
    df = None
    if uploaded is not None:
        try:
            if uploaded.name.endswith(".csv"):
                df = pd.read_csv(uploaded)
            else:
                df = pd.read_excel(uploaded)
        except ImportError:
            st.error("openpyxl မရှိသေးဘူး — PowerShell မှာ ဒါ run လုပ်ပါ: .\\.venv\\Scripts\\python.exe -m pip install openpyxl")
        except Exception as e:
            st.error(f"File ဖတ်မရဘူး: {e}")
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
                st.caption(f"Open text — {nunique} unique answers. နမူနာ:")
                for val in series.head(5):
                    st.write(f"- {val}")
                summary_lines.append(f"Open text — {nunique} unique answers.")
            summary_lines.append("")
        if len(cat_cols) >= 2:
            st.divider()
            st.subheader("🔀 Cross-tab")
            col_a = st.selectbox("ပထမ မေးခွန်း", cat_cols, key="xtab_a")
            col_b = st.selectbox("ဒုတိယ မေးခွန်း", cat_cols, index=1 if len(cat_cols) > 1 else 0, key="xtab_b")
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
    st.title("📑 Research Reports")

    saved = sorted(ANALYSIS_DIR.glob("*_analysis.txt"), key=lambda f: f.name)
    if not saved:
        st.info("Analysis မရှိသေးပါ — အရင် Market Analysis မှာ Analyze လုပ်ပါ။")
    else:
        saved_names = [f.name for f in saved]
        selected = st.selectbox("Analysis ရွေးပါ", saved_names)
        template = st.selectbox("Template ရွေးပါ", list(TEMPLATES.keys()))
        col_pptx, col_docx = st.columns(2)
        if col_pptx.button("Export PPTX"):
            analysis_text = (ANALYSIS_DIR / selected).read_text(encoding="utf-8")
            if looks_like_html(analysis_text):
                analysis_text = html_to_text(analysis_text)
            title_text = selected.replace("_analysis.txt", "").replace("_", " ")
            out_path = REPORTS_DIR / selected.replace("_analysis.txt", ".pptx")
            with st.spinner("PPTX ထုတ်နေပါတယ်..."):
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
            report_selected = st.selectbox("သိမ်းထားတဲ့ report ရွေးပါ", report_names)
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
            st.info("သိမ်းထားတဲ့ report မရှိသေးပါ။")

elif page == "Report Editor":
    st.title("✏️ Report Editor")
    st.caption("Google Docs မှာ ပြင် — Doc ဖွင့်, ပြင်ပြီး, Import back နှိပ်. မူရင်း file ကို မထိဘူး.")

    saved = sorted(ANALYSIS_DIR.glob("*_analysis.txt"), key=lambda f: f.name)
    if not saved:
        st.info("Analysis မရှိသေးပါ — အရင် Market Analysis မှာ Analyze လုပ်ပါ။")
    else:
        saved_names = [f.name for f in saved]
        # v3.2 fix: a widget's value may only be set BEFORE the widget is
        # created in the current run. Load buttons set a plain flag and rerun;
        # it is applied here, before the selectbox exists in this run.
        _load_target = st.session_state.pop("editor_load_target", None)
        if _load_target in saved_names:
            st.session_state["editor_select"] = _load_target
        selected = st.selectbox("Analysis ရွေးပါ", saved_names, key="editor_select")
        content = (ANALYSIS_DIR / selected).read_text(encoding="utf-8")
        if looks_like_html(content):
            content = html_to_text(content)
        edited = st.text_area("Report", value=content, height=400, key=f"ed_text_{selected}")
        stem = selected[:-len("_analysis.txt")]

        if docs_available():
            col_c, col_i = st.columns(2)
            if col_c.button("📄 Create Google Doc"):
                title = stem.replace("_", " ")
                with st.spinner("Google Doc ဖန်တီးနေပါတယ်..."):
                    try:
                        doc_id, url = drive_create_doc(title, edited)
                        st.session_state[f"editor_doc_id_{selected}"] = doc_id
                        st.session_state[f"editor_doc_url_{selected}"] = url
                        # v3.2 fix: fill the link field too — the text input
                        # keeps its old value unless its own key is set here.
                        st.session_state["editor_doc_input"] = url
                        st.success("Google Doc ရပြီ — အောက်က link နှိပ်ပြီး ပြင်ပါ.")
                    except Exception as e:
                        st.error(f"Doc ဖန်တီးမရဘူး: {e}")
            # v3.2 fix: when the selected analysis changes, show that
            # analysis's stored Doc link (or empty) instead of a stale one.
            if st.session_state.get("_editor_last_selected") != selected:
                st.session_state["_editor_last_selected"] = selected
                st.session_state["editor_doc_input"] = st.session_state.get(f"editor_doc_url_{selected}", "")
            doc_url_default = st.session_state.get(f"editor_doc_url_{selected}", "")
            doc_input = st.text_input("Google Doc link (သို့မဟုတ် ID)", value=doc_url_default, key="editor_doc_input")
            if st.session_state.get(f"editor_doc_url_{selected}"):
                st.link_button("🔗 Open Google Doc", st.session_state[f"editor_doc_url_{selected}"])
            if col_i.button("⬇️ Import back from Doc"):
                doc_id = _doc_id_from_input(doc_input)
                if not doc_id:
                    st.warning("Doc link အရင် ထည့်ပါ.")
                else:
                    with st.spinner("Doc ကနေ ပြန်ယူနေပါတယ်..."):
                        try:
                            new_text = docs_read_text(doc_id)
                            out_name = f"{stem}_v{next_version_number(stem)}_analysis.txt"
                            store_write_text("analysis", out_name, new_text)
                            st.success(f"Saved: {out_name} — မူရင်း {selected} က မပြောင်းဘူး.")
                            st.caption("Reports page သွားရင် version အသစ် ရွေးပြီး PPTX/DOCX export လုပ်လို့ရပြီ။")
                        except Exception as e:
                            st.error(f"Import မရဘူး: {e}")
        else:
            st.info("Google နဲ့ မချိတ်ရသေးဘူး — Doc ဖွင့်လို့ မရသေးဘူး. ဒီမှာ ပြင်ပြီး version အသစ် အဖြစ် သိမ်းလို့ရတယ်.")
            if st.button("💾 Save as new version"):
                out_name = f"{stem}_v{next_version_number(stem)}_analysis.txt"
                store_write_text("analysis", out_name, edited or "")
                st.success(f"Saved: {out_name} — မူရင်း {selected} က မပြောင်းဘူး.")
                st.caption("Reports page သွားရင် version အသစ် ရွေးပြီး PPTX/DOCX export လုပ်လို့ရပြီ။")

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
            st.caption("Version အသစ် မရှိသေးဘူး။")

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
