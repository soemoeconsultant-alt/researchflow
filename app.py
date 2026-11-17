import os
import re
import time
import csv
import json
import html
import datetime
import hashlib
import shutil
import subprocess
import unicodedata
from zoneinfo import ZoneInfo
from html.parser import HTMLParser
import pandas as pd
import streamlit as st
from pathlib import Path
from google import genai
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from docx import Document
from docx.shared import RGBColor as DocxRGBColor

st.set_page_config(
    page_title="ResearchFlow | AI Market Research",
    page_icon="🔮",
    layout="wide"
)

# App version — shown in the sidebar and written into backup manifests, so the
# running version can always be verified from the code itself (v3.15 had none).
APP_VERSION = "3.31"

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


def _update_ui_prefs(**kw):
    """v3.29: merge keys into the saved prefs instead of replacing the file,
    so saving theme/lang never wipes the persisted active project (and
    vice versa). No secrets, no personal data — just theme/lang/project."""
    _p = _load_ui_prefs()
    _p.update(kw)
    _save_ui_prefs(_p)


def _flash(msg, kind="success"):
    """v3.29: queue a one-shot confirmation to show once after the next
    st.rerun(). A plain st.success() before st.rerun() is wiped by the rerun
    (trial FAIL: "Brief saved." never appeared), so save handlers call this
    instead and the message is rendered once on the next run, then cleared.
    kind is "success", "error", "warning" or "info"."""
    st.session_state["_flash_msg"] = (kind, msg)


# Seed theme + language + active project once per session from the saved
# prefs (a fresh browser session starts with empty session state, so this is
# the only moment the file can be read). Later reruns keep the in-session
# choice. v3.29: the active project is persisted too — a full page reload
# used to forget it and reopen on "Default" (trial FAIL, revised: the
# step-1 Dashboard checkmarks were Default's data, not a leak). The
# persisted name is validated against the project list at the selector
# below; a name that no longer exists falls back to "Default".
if "_ui_prefs_loaded" not in st.session_state:
    _prefs = _load_ui_prefs()
    if _prefs.get("theme") in ("light", "dark"):
        st.session_state["theme"] = _prefs["theme"]
    if _prefs.get("lang") in ("EN", "MM"):
        st.session_state["lang"] = _prefs["lang"]
    if isinstance(_prefs.get("active_project"), str) and _prefs["active_project"].strip():
        st.session_state["active_project"] = _prefs["active_project"].strip()
        st.session_state["_active_from_prefs"] = True
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
BRIEF_DELIVERABLES = ["PPTX Report", "DOCX Report", "Presentation"]


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


# --- Research Design (v3.25, Stage 2): one design.json per project, in
# notes/ (Drive notes folder + local, same path as brief.json). The design is
# the feed-forward contract for later stages: the Market Analysis planner
# reads brief.json + notes/discussion_guide.txt + this file's questionnaire
# structure; Stage 3's form page will render design["questionnaire"].
# NO SAMPLE CONTENT RULE (user, 2026-11-17): the real project samples he
# shared are REFERENCE DATA for the website's STRUCTURE only — their format,
# section shapes, tables, and ordering. NO sample content may appear in this
# app: no findings, quotes, prices, dates, venues, client names, contacts,
# interviewer names, product codes, brand lists, township lists, MMK bands,
# or respondent data — not as defaults, not as placeholders, not as examples,
# not in help text, not in prompts. The design page starts EMPTY for every
# new project; the only prefill allowed is from the project's OWN saved Brief.
# The RESULT side (DG DOCX, questionnaire XLSForm, reports in v3.27) follows
# the sample FORMATS but is filled with the CURRENT project's own data only.
# Keys (all optional — empty means unset, never invented):
#   approach: "Qualitative" | "Quantitative" | "Mixed" | ""
#   methods: ["FGD"] etc. (subset of DESIGN_METHODS)
#   target_audience: copied from the brief when the design is saved
#   goal: the research goal in one short text (prefilled from the brief's
#     objectives when the design is first saved; empty when no brief).
#   target: the target in one short text (prefilled from the brief's
#     target audience the same way). Goal + target + method + N + personas
#     + sessions are the report's front matter in Stage 5 (v3.27) — that is
#     why they live here as structured fields, not free text only.
#   sample_size: int, 0 = unset (total respondents, e.g. survey n)
#   sessions_count: int, 0 = unset (number of IDI/FGD sessions)
#   participants_per_session: int, 0 = unset (respondents in one session)
#   personas: [{"name": ..., "description": ..., "value_framing": ...}] —
#     from the brief's respondent criteria; the DG draft turns these into
#     persona-specific probe blocks; "value_framing" is the segment's short
#     value label. The personas are the report's organizing spine in Stage 5,
#     so keep them structured, not free text.
#   quotas: [{"segment": ..., "share": ..., "notes": ...}] — channel /
#     segment splits (each quota is a segment name plus its share).
#   screening_criteria: text — who qualifies / who is screened out
#     (user type, behaviour, non-rejector rules from the brief).
#   recruitment_criteria: text — extra recruitment notes.
#   rationale: text — why this method was chosen, in the user's own words.
#     The proposal is drafted FROM the design (v3.28, Reports page).
#   dg_language: "English" | "Myanmar" (moderator-script language)
#   dg_text: the Discussion Guide draft (markup, see DG_DRAFT_TEMPLATE)
#   dg_file: "discussion_guide.txt" in notes/ (same text, as a file)
#   questionnaire_language: "Myanmar" | "English" (respondent language)
#   questionnaire: {"intro_en": ..., "intro_mm": ..., "closing_en": ...,
#     "closing_mm": ..., "header_fields": ["FGD date", "Interviewer",
#     "Respondent ID", "Confirmation"], "questions": [{"code": "Q1",
#     "name": "q1", "section": "Screening", "question_en": ...,
#     "question_mm": ..., "type": ..., "options_en": ..., "options_mm": ...,
#     "outcome": ..., "hint_en": ..., "hint_mm": ..., "required": false,
#     "appearance": ..., "relevant": ..., "constraint": ...,
#     "constraint_message": ...}]}
#     — the screener instrument, stored ONCE with EN + MM text in parallel
#     (same Q numbers/codes render both versions; the MM version keeps
#     English food-category labels where the instrument uses English category
#     names). "type" is in QUESTION_TYPES (single_select, multi_select,
#     ranking_top3, likert_1_5, rating, open_text, numeric, date, note,
#     observed = observe and record, do not ask; contact = name + phone).
#     "options_en"/"options_mm" are " | " separated (options contain commas).
#     "outcome" is the per-question rule in plain words: Terminate /
#     Continue / group assignment ("<answer> → Group <n>") / quota tag
#     ("<segment>"). Screening is a section + outcome rule, not a type. "name" is the XLSForm field name, derived from the code
#     ("Q1" → "q1") when empty. "hint_en"/"hint_mm" are recruiter hints;
#     "required" is a bool; "relevant" is an XLSForm expression
#     (e.g. "${q11} = 'other'"); "constraint" an XPath check; "appearance"
#     is usually derived at export (likert for Likert/rating, field-list
#     for groups). "intro" is the recruiter script read to the respondent;
#     "closing" the thank-you. This structure maps cleanly to a Kobo
#     XLSForm (questionnaire_to_xlsform) and is what Stage 3's form page
#     will render.
#   fieldwork_plan: [{"session": 1, "group": ..., "method": ..., "date":
#     "YYYY-MM-DD", "time": ..., "seats": 0, "recruited": 0, "confirmed": 0,
#     "moderator": ..., "moderator_seat": "Seat 1", "venue": ...}] — one
#     session per persona segment; "seats" is the number of respondents who
#     sit in the session and "recruited" the over-recruited count (extra
#     recruits cover no-shows);
#     "confirmed" is how many confirmed to attend. The moderator sits at
#     Seat 1 by the seating-plan template convention (filled in on save when
#     a moderator is set and the seat is empty). Dates/times are plain text,
#     never pickers (an unset date must never invent today's date, same rule
#     as the Brief timeline).
#     Participant rows (names/phones) are NOT stored here — that is Stage 3
#     interview management; the future import shape is one sheet per group
#     with screener answers + contact columns. A seating-plan DOCX export
#     (one card per confirmed participant: name, age, gender, education,
#     occupation, SEC, food type, frequency, outlet; moderator at the top)
#     belongs to Stage 3, not v3.25.
DESIGN_APPROACHES = ["Qualitative", "Quantitative", "Mixed"]
DESIGN_METHODS = ["IDI", "FGD", "Survey"]
DESIGN_DG_LANGUAGES = ["English", "Myanmar"]
DESIGN_Q_LANGUAGES = ["Myanmar", "English"]
QUESTION_TYPES = ["single_select", "multi_select", "ranking_top3", "likert_1_5", "rating", "open_text", "numeric", "date", "note", "observed", "contact"]
QUESTION_TYPE_LABELS = {
    "single_select": "Single-select",
    "multi_select": "Multi-select",
    "ranking_top3": "Top-3 ranking",
    "likert_1_5": "Likert 1–5",
    "rating": "Rating",
    "open_text": "Open text",
    "numeric": "Numeric",
    "date": "Date",
    "note": "Note (display text)",
    "observed": "Observed (do not ask)",
    "contact": "Contact capture",
}


def load_design():
    """Saved design for the current project, or {} when none exists."""
    try:
        data = json.loads((NOTES_DIR / "design.json").read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_design(data):
    """Save the design (Drive + local, via the notes folder)."""
    return store_write_text("notes", "design.json", json.dumps(data, ensure_ascii=False, indent=2))


# --- v3.26 Stage 3 Fieldworks helpers ---
# NO SAMPLE CONTENT RULE (user, 2026-11-17): the samples he sent (brief,
# proposal, DG, recruitment list/forms, report, Kobo form, seating plan) are
# STRUCTURE reference only — zero sample content (no findings, quotes,
# prices, dates, venues, client/interviewer names, product codes,
# brand/township lists, MMK bands, respondent data) anywhere in this file,
# not as defaults, placeholders, examples, help text, or prompts. Every value
# below comes from the project's OWN design.json, the user's own input, or a
# file he uploads. The Fieldworks page starts EMPTY for a new project.
# Fieldwork state lives in notes/fieldwork.json (Drive + local, same pattern
# as design.json): participants per session, field notes, survey-file
# registry, per-session tracking status. Form responses live in
# notes/responses.csv (one row per submitted form). Participant rows
# (names/phones) are stored here in Stage 3 — they were explicitly NOT in
# design.json. A public link-only respondent route needs an auth exception
# and is NOT built in v3.26 (open decision for him); the form runs here
# behind the login gate.
FW_FIELDWORK_FILE = "fieldwork.json"
FW_RESPONSES_FILE = "responses.csv"
FW_SESSION_STATUSES = ["Not started", "Recruiting", "Confirmed", "Completed"]
FW_PARTICIPANT_FIELDS = ["name", "phone", "age_band", "gender", "marital",
                         "education", "occupation", "income_band", "township",
                         "sec", "food_type", "frequency", "outlet", "confirmed"]
FW_PARTICIPANT_LABELS = {"name": "Name", "phone": "Phone", "age_band": "Age band",
                         "gender": "Gender", "marital": "Marital",
                         "education": "Education", "occupation": "Occupation",
                         "income_band": "Income band", "township": "Township",
                         "sec": "SEC", "food_type": "Food type",
                         "frequency": "Frequency", "outlet": "Outlet",
                         "confirmed": "Confirmed"}
# Normalised import header -> participant field (recruitment-list XLSX:
# one sheet per group; any column not listed here is kept as a screener
# answer under its original header).
FW_IMPORT_HEADER_MAP = {
    "name": "name", "respondentname": "name", "respondent": "name",
    "phone": "phone", "phonenumber": "phone", "contact": "phone", "mobile": "phone",
    "gender": "gender", "sex": "gender",
    "age": "age_band", "ageband": "age_band",
    "marital": "marital", "maritalstatus": "marital",
    "education": "education",
    "occupation": "occupation", "job": "occupation", "jobtitle": "occupation",
    "position": "occupation",
    "income": "income_band", "householdincome": "income_band",
    "householdmonthlyincome": "income_band", "incomeband": "income_band",
    "township": "township", "sec": "sec",
    "foodtype": "food_type", "retfoodtype": "food_type",
    "foodtypetheyeat": "food_type",
    "frequency": "frequency", "outlet": "outlet", "outlettheygo": "outlet",
}


def load_fieldwork():
    """Saved fieldwork state for the current project, or a fresh empty one."""
    try:
        data = json.loads((NOTES_DIR / FW_FIELDWORK_FILE).read_text(encoding="utf-8"))
        if isinstance(data, dict):
            data.setdefault("participants", {})
            data.setdefault("notes", [])
            data.setdefault("survey_files", [])
            data.setdefault("session_status", {})
            return data
    except Exception:
        pass
    return {"participants": {}, "notes": [], "survey_files": [], "session_status": {}}


def save_fieldwork(data):
    """Save the fieldwork state (Drive + local, via the notes folder)."""
    return store_write_text("notes", FW_FIELDWORK_FILE, json.dumps(data, ensure_ascii=False, indent=2))


def _fw_plan_sessions(design):
    """Fieldwork-plan sessions from design.json, normalised (never invented:
    an empty plan returns an empty list)."""
    plan = (design or {}).get("fieldwork_plan") or []
    out = []
    for i, r in enumerate(plan):
        if not isinstance(r, dict):
            continue
        out.append({"session": int(r.get("session") or i + 1),
                    "group": str(r.get("group") or "").strip(),
                    "method": str(r.get("method") or "").strip(),
                    "date": str(r.get("date") or "").strip(),
                    "time": str(r.get("time") or "").strip(),
                    "seats": int(r.get("seats") or 0),
                    "recruited": int(r.get("recruited") or 0),
                    "confirmed": int(r.get("confirmed") or 0),
                    "moderator": str(r.get("moderator") or "").strip(),
                    "moderator_seat": str(r.get("moderator_seat") or "").strip(),
                    "venue": str(r.get("venue") or "").strip()})
    return out


def _fw_session_key(sess):
    """Stable key for a plan session: 'session|group' (group may be empty)."""
    return f"{sess.get('session', '')}|{sess.get('group', '')}"


def _fw_session_label(sess):
    """Short human label for a plan session, from its own fields only."""
    parts = [f"Session {sess.get('session', '')}"]
    if sess.get("group"):
        parts.append(f"— {sess['group']}")
    if sess.get("date"):
        parts.append(f"({sess['date']}" + (f" {sess['time']}" if sess.get("time") else "") + ")")
    return " ".join(parts)


def _fw_question_label(q, lang):
    """Question text in the questionnaire language (fallback: the other one)."""
    if lang == "Myanmar":
        return (q.get("question_mm") or q.get("question_en") or "").strip()
    return (q.get("question_en") or q.get("question_mm") or "").strip()


def _fw_question_options(q, lang):
    """Option list in the questionnaire language (' | '-separated in storage)."""
    raw = q.get("options_mm") if lang == "Myanmar" else q.get("options_en")
    if not raw:
        raw = q.get("options_en") or q.get("options_mm") or ""
    return [o.strip() for o in str(raw).split(" | ") if o.strip()]


def _fw_matches(needle, sel_en, sel_mm):
    """Does the outcome-rule trigger name the selected answer? Substring
    match either way, against the English and Myanmar labels."""
    n = (needle or "").strip().lower()
    if not n:
        return False
    for sel in (sel_en, sel_mm):
        s = (sel or "").strip().lower()
        if s and (n in s or s in n):
            return True
    return False


def _resolve_outcome(outcome, sel_en, sel_mm):
    """Resolve one question's outcome rule against the selected answer.
    Outcome rules come from the v3.25 questionnaire draft format, e.g.
    'Yes → Terminate; No → Continue' or '<answer> → Group <n>' or a bare
    quota tag. Returns {'terminates': bool, 'group': str, 'quota': str}.
    A bare 'Terminate' (no answer named) ends the interview on any answer —
    drafters should write 'X → Terminate' instead (the v3.25 template asks
    for that form); ambiguous rules never invent a group or quota."""
    res = {"terminates": False, "group": "", "quota": ""}
    if not outcome:
        return res
    for seg in str(outcome).split(";"):
        seg = seg.strip()
        if not seg:
            continue
        low = seg.lower()
        if "terminate" in low:
            trigger = None
            if "→" in seg:
                trigger = seg.split("→", 1)[0].strip()
            else:
                m = re.search(r"terminat\w*\s*(?:if|when|on|:)\s*(.+)$", seg, re.IGNORECASE)
                if m:
                    trigger = m.group(1).strip()
            if trigger is None or trigger == "":
                res["terminates"] = True
            elif _fw_matches(trigger, sel_en, sel_mm):
                res["terminates"] = True
            continue
        if "→" in seg:
            left, right = seg.split("→", 1)
            if _fw_matches(left.strip(), sel_en, sel_mm):
                right = right.strip()
                if "group" in right.lower():
                    res["group"] = right
                elif "continue" not in right.lower():
                    res["quota"] = right
            continue
        if low != "continue":
            res["quota"] = seg
    return res


def _responses_header(qdata):
    """CSV header for the responses file: fixed columns, header-field
    columns, then one column per question (contact expands to two)."""
    header = ["submitted_at", "session", "terminated", "terminate_reason",
              "group_assignment", "quota_tag"]
    for h in (qdata or {}).get("header_fields") or []:
        header.append("header_" + _xlsform_slug(h))
    for q in (qdata or {}).get("questions") or []:
        if not isinstance(q, dict):
            continue
        name = (q.get("name") or _xlsform_slug(q.get("code", "")) or "q").strip()
        if q.get("type") == "contact":
            header.append(name + "_name")
            header.append(name + "_phone")
        elif q.get("type") == "note":
            continue
        else:
            header.append(name)
    return header


def _responses_append_csv(existing_text, header, row):
    """Pure CSV append: parse existing text, merge the header (new columns
    are added, old rows backfilled with ''), append the row. Returns the
    new CSV text. Never drops an existing row."""
    import io
    rows = []
    old_header = []
    if existing_text:
        try:
            reader = csv.reader(io.StringIO(existing_text))
            parsed = list(reader)
            if parsed:
                old_header = [c.strip() for c in parsed[0]]
                for r in parsed[1:]:
                    if any(str(c).strip() for c in r):
                        rows.append(r)
        except Exception:
            rows, old_header = [], []
    new_header = list(old_header)
    for c in header:
        if c not in new_header:
            new_header.append(c)
    out_rows = []
    for r in rows:
        d = dict(zip(old_header, [str(c) for c in r]))
        out_rows.append([d.get(c, "") for c in new_header])
    out_rows.append([str(row.get(c, "")) for c in new_header])
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(new_header)
    w.writerows(out_rows)
    return buf.getvalue()


def read_responses():
    """(header, rows) from notes/responses.csv, or ([], []) when none yet."""
    import io
    path = NOTES_DIR / FW_RESPONSES_FILE
    if not path.exists():
        return [], []
    try:
        reader = csv.DictReader(io.StringIO(path.read_text(encoding="utf-8")))
        header = reader.fieldnames or []
        return header, [dict(r) for r in reader]
    except Exception:
        return [], []


def append_response_row(header, row):
    """Append one response row to notes/responses.csv (Drive + local).
    Returns the saved path."""
    existing = ""
    path = NOTES_DIR / FW_RESPONSES_FILE
    if path.exists():
        existing = path.read_text(encoding="utf-8")
    return store_write_text("notes", FW_RESPONSES_FILE,
                            _responses_append_csv(existing, header, row))


def fieldwork_counts(design=None, fw=None):
    """Real Stage 3 counts for stage_status(): responses, participants,
    field notes, survey files, sessions with confirmed > 0."""
    if design is None:
        design = load_design()
    if fw is None:
        fw = load_fieldwork()
    _h, _rows = read_responses()
    participants = fw.get("participants") or {}
    total_part = sum(len(v) for v in participants.values() if isinstance(v, list))
    confirmed_part = sum(1 for v in participants.values() if isinstance(v, list)
                         for p in v if isinstance(p, dict) and p.get("confirmed"))
    plan = _fw_plan_sessions(design)
    return {"responses": len(_rows),
            "responses_terminated": sum(1 for r in _rows if str(r.get("terminated")) == "1"),
            "participants": total_part,
            "participants_confirmed": confirmed_part,
            "notes": len(fw.get("notes") or []),
            "survey_files": len(fw.get("survey_files") or []),
            "sessions": len(plan),
            "sessions_confirmed": sum(1 for s in plan if s["confirmed"] > 0)}


def import_recruitment_xlsx(data_bytes):
    """Import a recruitment-list XLSX (one sheet per group) into participant
    rows. Returns (groups, stats): groups maps sheet name -> [participant
    dicts]; stats counts sheets/imported/skipped. Only rows present in the
    file are imported — nothing is invented; a row needs a name or a phone,
    and a (name, phone) pair already present in its sheet is skipped.
    Columns not in FW_IMPORT_HEADER_MAP are kept as screener answers under
    their original header."""
    import io
    try:
        import openpyxl
    except ImportError as e:
        return {}, {"sheets": [], "imported": 0, "skipped": 0, "error": str(e)}
    groups = {}
    imported = 0
    skipped = 0
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)
    except Exception as e:
        return {}, {"sheets": [], "imported": 0, "skipped": 0, "error": str(e)}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header_row = None
        for r in rows:
            if r and any(c is not None and str(c).strip() for c in r):
                header_row = r
                break
        if header_row is None:
            continue
        headers = [(str(c).strip() if c is not None else "") for c in header_row]
        norm = [re.sub(r"[\s_]+", "", h.lower()) for h in headers]
        start = rows.index(header_row) + 1
        sheet_parts = []
        seen = set()
        for r in rows[start:]:
            if not r or not any(c is not None and str(c).strip() for c in r):
                skipped += 1
                continue
            part = {f: "" for f in FW_PARTICIPANT_FIELDS if f != "confirmed"}
            part["confirmed"] = False
            part["screener"] = {}
            for idx, val in enumerate(r):
                if idx >= len(headers) or val is None:
                    continue
                v = str(val).strip()
                if not v:
                    continue
                field = FW_IMPORT_HEADER_MAP.get(norm[idx])
                if field:
                    part[field] = v
                else:
                    part["screener"][headers[idx]] = v
            if not part["name"] and not part["phone"]:
                skipped += 1
                continue
            key = (part["name"].lower(), part["phone"])
            if key in seen:
                skipped += 1
                continue
            seen.add(key)
            sheet_parts.append(part)
            imported += 1
        if sheet_parts:
            groups[ws.title] = sheet_parts
    return groups, {"sheets": list(groups.keys()), "imported": imported,
                    "skipped": skipped, "error": ""}


def seating_plan_to_docx(title_text, session, participants, out_path):
    """Seating-plan DOCX for one session, following the sample's layout
    (structure only): title, group line, date line, room line, then one
    2-column table — a merged moderator row on top, then 4 pairs of
    (Respondent NO. header row, detail-card row) = 8 seats. Confirmed
    participants fill the seats in order (seats 1–6, spares 7–8); empty
    seats stay blank. Card fields: Name, Age, Gender, Education,
    Occupation, SEC, Food type, Frequency, Outlet."""
    doc = Document()
    doc.add_paragraph(title_text).runs[0].bold = True
    _s_no = session.get("session", "")
    doc.add_paragraph(f"Group{_s_no}: {session.get('group', '')}")
    _date_line = f"FGD Date: {session.get('date', '')}"
    if session.get("time"):
        _date_line += f", {session['time']}"
    doc.add_paragraph(_date_line)
    doc.add_paragraph(f"FGD Room: {session.get('venue', '')}")
    table = doc.add_table(rows=9, cols=2)
    table.style = "Table Grid"
    _mod = session.get("moderator", "")
    _seat = session.get("moderator_seat", "") or "Seat 1"
    _mcell = table.cell(0, 0).merge(table.cell(0, 1))
    _mcell.text = f"Moderator | {_mod} | {_seat}" if _mod else "Moderator"
    _confirmed = [p for p in (participants or []) if isinstance(p, dict) and p.get("confirmed")]
    _seats = _confirmed[:8]
    _card_fields = [("Name", "name"), ("Age", "age_band"), ("Gender", "gender"),
                    ("Education", "education"), ("Occupation", "occupation"),
                    ("SEC", "sec"), ("Food type", "food_type"),
                    ("Frequency", "frequency"), ("Outlet", "outlet")]
    for _pair in range(4):
        _r = 1 + _pair * 2
        for _c in range(2):
            _n = _pair * 2 + _c + 1
            table.cell(_r, _c).text = f"Respondent NO. {_n}"
            _p = _seats[_n - 1] if _n - 1 < len(_seats) else None
            _cell = table.cell(_r + 1, _c)
            _cell.text = ""
            for _i, (_label, _field) in enumerate(_card_fields):
                _val = str(_p.get(_field, "")) if _p else ""
                _line = f"{_label}: {_val}" if _val else f"{_label}:"
                if _i == 0:
                    _cell.paragraphs[0].text = _line
                else:
                    _cell.add_paragraph(_line)
    doc.save(str(out_path))
    return out_path


# --- v3.27 AI Research Intelligence helpers ---
# NO SAMPLE CONTENT RULE (user, 2026-11-17): the samples he sent (brief,
# proposal, DG, recruitment list/forms, report, Kobo form, seating plan) are
# STRUCTURE reference only — zero sample content (no findings, quotes,
# prices, dates, venues, client/interviewer names, product codes,
# brand/township lists, MMK bands, respondent data) anywhere in this file,
# not as defaults, placeholders, examples, help text, or prompts. Every
# value below comes from the project's OWN saved files or the user's own
# input. Every new page/section starts EMPTY for a new project.
# Evidence verification is deterministic pure-Python matching (no Gemini
# call, no quota): a finding is VERIFIED only when its verbatim quote is
# found word-for-word in a transcript and its claimed n matches what the
# transcripts actually contain. A failed finding stays flagged — nothing is
# auto-fixed or invented. QC corrections never overwrite an original: the
# corrected copy is saved alongside and every correction is audit-logged.
QC_FILE = "qc.json"
QC_AUDIT_FILE = "qc_audit.json"
INSIGHTS_FILE = "insights.json"
QC_STATUSES = ["Unchecked", "Checked", "Needs correction", "Corrected"]


def _ev_normalise(text):
    """Lowercase, drop punctuation/symbols, collapse whitespace — for
    matching only. The words themselves are never changed. Punctuation is
    dropped by Unicode category (not [^\\w\\s]) so Myanmar combining marks
    — which are part of the word — survive intact."""
    t = str(text or "").lower()
    t = "".join(ch if unicodedata.category(ch)[0] not in ("P", "S", "C") else " "
                for ch in t)
    return re.sub(r"\s+", " ", t).strip()


def _ev_quote_match(quote, transcript_text):
    """("exact" | "near" | "none", coverage). Exact = the normalised quote
    is a substring of the normalised transcript. Near = at least 80% of the
    quote's tokens appear in the transcript (a quote trimmed at the edges
    or with a word changed). Near never counts as verified."""
    q = _ev_normalise(quote)
    if not q:
        return "none", 0.0
    t = _ev_normalise(transcript_text)
    if q in t:
        return "exact", 1.0
    qt = q.split()
    if not qt:
        return "none", 0.0
    tset = set(t.split())
    hit = sum(1 for w in qt if w in tset)
    cov = hit / len(qt)
    return ("near", cov) if cov >= 0.8 else ("none", cov)


_EV_TURN_RE = re.compile(r"(?m)^(Speaker\s*\d+|Moderator|Respondent|Interviewer)\s*:", re.IGNORECASE)


def _ev_speakers_with(transcript_text, quote):
    """Speaker labels whose turns contain the quote (exact or near match).
    Turns are 'Speaker N:' blocks; text before the first label is checked
    as its own turn ('Unlabelled')."""
    text = str(transcript_text or "")
    parts = _EV_TURN_RE.split(text)
    speakers = set()
    if parts and _ev_quote_match(quote, parts[0])[0] in ("exact", "near"):
        speakers.add("Unlabelled")
    for i in range(1, len(parts) - 1, 2):
        label = re.sub(r"\s+", " ", parts[i].strip())
        if _ev_quote_match(quote, parts[i + 1])[0] in ("exact", "near"):
            speakers.add(label)
    return speakers


_EV_QUOTE_RE = re.compile(r'>\s*"([^"]+)"\s*(?:—|-)\s*Speaker\s*(\d+)?\s*,?\s*(.*)$')
_EV_INLINE_QUOTE_RE = re.compile(r'"([^"]+)"\s*(?:—|-)\s*Speaker\s*(\d+)?\s*,?\s*(.*)$')
_EV_N_RE = re.compile(r"n\s*=\s*(\d+)\s*(?:of|/)\s*(\d+)", re.IGNORECASE)


def _ev_parse_findings(analysis_text):
    """One entry per finding line: {section, text, claimed_n, claimed_y,
    quotes: [{quote, speaker, source}], no_evidence}. A quote line
    ('> "..." — Speaker N, <file>') attaches to the finding directly above
    it in the same section. A line that is exactly 'No evidence in
    transcripts' is kept with no_evidence=True — an honest empty statement,
    not a claim."""
    out = []
    for sec in parse_sections(analysis_text or ""):
        current = None
        for line in sec["lines"]:
            if not line:
                continue
            if line.strip() == "No evidence in transcripts":
                out.append({"section": sec["title"], "text": line.strip(),
                            "claimed_n": None, "claimed_y": None,
                            "quotes": [], "no_evidence": True})
                current = None
                continue
            m = _EV_QUOTE_RE.match(line)
            if m:
                quote = {"quote": m.group(1).strip(),
                         "speaker": m.group(2) or "",
                         "source": (m.group(3) or "").strip()}
                if current is not None:
                    current["quotes"].append(quote)
                else:
                    out.append({"section": sec["title"], "text": "",
                                "claimed_n": None, "claimed_y": None,
                                "quotes": [quote], "no_evidence": False})
                continue
            nm = _EV_N_RE.search(line)
            entry = {"section": sec["title"], "text": line,
                     "claimed_n": int(nm.group(1)) if nm else None,
                     "claimed_y": int(nm.group(2)) if nm else None,
                     "quotes": [], "no_evidence": False}
            im = _EV_INLINE_QUOTE_RE.search(line)
            if im:
                entry["quotes"].append({"quote": im.group(1).strip(),
                                       "speaker": im.group(2) or "",
                                       "source": (im.group(3) or "").strip()})
            out.append(entry)
            current = entry
    return out


def verify_analysis(analysis_text, transcript_texts):
    """Verify every finding in a saved analysis against the project's
    transcripts ({filename: text}). Returns (results, (verified, total)).
    VERIFIED only on a word-for-word quote + a claimed n the transcripts
    support; anything else is UNVERIFIED with the reason."""
    findings = _ev_parse_findings(analysis_text)
    results = []
    verified = 0
    for f in findings:
        row = {"section": f["section"], "finding": f["text"],
               "quote": "", "speaker": "", "source": "",
               "claimed_n": f["claimed_n"], "claimed_y": f["claimed_y"],
               "found_n": 0, "found_files": [], "status": "UNVERIFIED",
               "reason": ""}
        if f.get("no_evidence"):
            row["status"] = "VERIFIED"
            row["reason"] = "States no evidence — nothing to check"
            results.append(row)
            verified += 1
            continue
        if not f["quotes"]:
            row["reason"] = "No verbatim quote given"
            results.append(row)
            continue
        q = f["quotes"][0]
        row["quote"] = q["quote"]
        row["speaker"] = q["speaker"]
        row["source"] = q["source"]
        best = "none"
        best_cov = 0.0
        files_hit = []
        speakers = set()
        for fname, ttext in (transcript_texts or {}).items():
            m, cov = _ev_quote_match(q["quote"], ttext)
            if m in ("exact", "near"):
                files_hit.append(fname)
                # (file, speaker) pairs: the same "Speaker 1" label in two
                # files is two different respondents.
                speakers |= {(fname, s) for s in _ev_speakers_with(ttext, q["quote"])}
                if m == "exact":
                    best = "exact"
                elif best != "exact":
                    best = "near"
                best_cov = max(best_cov, cov)
        row["found_files"] = files_hit
        row["found_n"] = len(speakers) if speakers else len(files_hit)
        if best == "none":
            row["reason"] = "Quote not found in any transcript"
            if q["source"] and q["source"] not in (transcript_texts or {}):
                row["reason"] += " (cited source file is not in this project)"
        elif best == "near":
            row["reason"] = (f"Quote is a near match only ({best_cov:.0%} of words) "
                             "— not word-for-word")
        else:
            if f["claimed_n"] is None:
                row["status"] = "VERIFIED"
                row["reason"] = "Quote found word-for-word"
            elif row["found_n"] >= f["claimed_n"]:
                row["status"] = "VERIFIED"
                row["reason"] = (f"Quote found; {row['found_n']} speaker(s) "
                                 f"vs claimed n={f['claimed_n']}")
            else:
                row["reason"] = (f"Claimed n={f['claimed_n']}, only "
                                 f"{row['found_n']} speaker(s) found saying it")
        if row["status"] == "VERIFIED":
            verified += 1
        results.append(row)
    return results, (verified, len(results))


def evidence_report_text(analysis_name, results, score):
    """Plain-text evidence-check report, saved to notes/."""
    lines = [f"Evidence check — {analysis_name}",
             f"Verified: {score[0]} of {score[1]} findings", ""]
    for r in results:
        lines.append(f"[{r['status']}] {r['section']} — {r['finding'][:160]}")
        if r["quote"]:
            lines.append(f"  Quote: \"{r['quote']}\"")
            if r["speaker"]:
                lines.append(f"  Cited: Speaker {r['speaker']}, {r['source']}")
            else:
                lines.append(f"  Cited: {r['source']}")
        if r["claimed_n"] is not None:
            lines.append(f"  Claimed: n={r['claimed_n']} of {r['claimed_y']}")
        lines.append(f"  Found: {r['found_n']} speaker(s) in {len(r['found_files'])} file(s) — {r['reason']}")
        lines.append("")
    return "\n".join(lines)


def run_evidence_check(analysis_name):
    """Load a saved analysis + all full transcripts, verify, save the
    report to notes/evidence_check_<stem>.txt (Drive + local). Returns
    (results, score, report_path)."""
    text = (ANALYSIS_DIR / analysis_name).read_text(encoding="utf-8", errors="replace")
    if looks_like_html(text):
        text = html_to_text(text)
    full_files, _pending = get_full_transcripts()
    ttexts = {f.name: f.read_text(encoding="utf-8", errors="replace") for f in full_files}
    results, score = verify_analysis(text, ttexts)
    stem = analysis_name[:-4] if analysis_name.endswith(".txt") else analysis_name
    report = evidence_report_text(analysis_name, results, score)
    path = store_write_text("notes", f"evidence_check_{stem}.txt", report)
    return results, score, path


# --- Stage 4 QC helpers (v3.27) ---
_QC_SPEAKER_RE = re.compile(r"^(Speaker\s*\d+|Moderator|Respondent|Interviewer)\s*:", re.IGNORECASE)
_QC_UNCLEAR_RE = re.compile(r"\[unclear\]", re.IGNORECASE)


def qc_check_text(name, text):
    """Checks on one transcript's text. Returns a list of issue dicts
    {type, count, detail}. Pure counts — the text is never edited here."""
    lines = str(text or "").splitlines()
    issues = []
    unclear = sum(len(_QC_UNCLEAR_RE.findall(ln)) for ln in lines)
    if unclear:
        issues.append({"type": "unclear", "count": unclear,
                       "detail": f"{unclear} [unclear] marker(s)"})
    no_spk = 0
    short = 0
    for ln in lines:
        s = ln.strip()
        if not s or s.startswith("===") or s.startswith("---"):
            continue
        if _QC_SPEAKER_RE.match(s):
            continue
        no_spk += 1
        if len(s) < 10:
            short += 1
    if no_spk:
        issues.append({"type": "no_speaker", "count": no_spk,
                       "detail": f"{no_spk} line(s) without a speaker label"})
    if short:
        issues.append({"type": "short_line", "count": short,
                       "detail": f"{short} very short line(s) — possible truncation"})
    return issues


def load_qc():
    """Saved QC state for the current project, or a fresh empty one."""
    try:
        data = json.loads((NOTES_DIR / QC_FILE).read_text(encoding="utf-8"))
        if isinstance(data, dict):
            data.setdefault("files", {})
            return data
    except Exception:
        pass
    return {"files": {}}


def save_qc(data):
    """Save the QC state (Drive + local, via the notes folder)."""
    return store_write_text("notes", QC_FILE, json.dumps(data, ensure_ascii=False, indent=2))


def qc_run_all():
    """Run the checks over every transcript file; keep existing statuses,
    add 'Unchecked' for new files; return the qc dict."""
    qc = load_qc()
    files = qc.setdefault("files", {})
    for f in get_transcript_files():
        entry = files.setdefault(f.name, {"status": "Unchecked", "checked_at": "", "issues": []})
        entry["issues"] = qc_check_text(f.name, f.read_text(encoding="utf-8", errors="replace"))
        entry["checked_at"] = datetime.datetime.now(ZoneInfo("Asia/Rangoon")).strftime("%Y-%m-%d %H:%M")
    save_qc(qc)
    return qc


def qc_load_audit():
    """Audit trail entries (newest last), or [] when none exist."""
    try:
        data = json.loads((NOTES_DIR / QC_AUDIT_FILE).read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def qc_log_audit(entry):
    """Append one entry to the audit trail (Drive + local)."""
    audit = qc_load_audit()
    audit.append(entry)
    store_write_text("notes", QC_AUDIT_FILE, json.dumps(audit, ensure_ascii=False, indent=2))
    return audit


def _qc_first_diff(before_text, after_text):
    """First differing line (before, after), truncated — for the audit trail."""
    b_lines = str(before_text or "").splitlines()
    a_lines = str(after_text or "").splitlines()
    for i in range(max(len(b_lines), len(a_lines))):
        b = b_lines[i] if i < len(b_lines) else ""
        a = a_lines[i] if i < len(a_lines) else ""
        if b != a:
            return b[:160], a[:160]
    return "", ""


def qc_save_correction(name, corrected_text, note=""):
    """Save a corrected copy — the ORIGINAL file is never overwritten.
    Corrected copy: <stem>_corrected.txt in the transcripts folder, so it
    can be picked for analysis like any full transcript. Every correction
    is logged to the audit trail and the file's status becomes Corrected.
    Returns (path_or_None, message)."""
    orig = None
    for d in (OUTPUT_DIR, MERGED_DIR):
        p = d / name
        if p.exists():
            orig = p
            break
    if orig is None:
        return None, "Original file not found."
    original_text = orig.read_text(encoding="utf-8", errors="replace")
    if original_text == corrected_text:
        return None, "No changes — nothing saved."
    stem = name[:-4] if name.endswith(".txt") else name
    corr_name = f"{stem}_corrected.txt"
    out_path = store_write_text("transcripts", corr_name, corrected_text)
    before, after = _qc_first_diff(original_text, corrected_text)
    qc_log_audit({
        "timestamp": datetime.datetime.now(ZoneInfo("Asia/Rangoon")).strftime("%Y-%m-%d %H:%M"),
        "file": name, "corrected_file": corr_name,
        "action": "correction",
        "before": before, "after": after, "note": note or "",
    })
    qc = load_qc()
    entry = qc.setdefault("files", {}).setdefault(name, {"status": "Unchecked", "checked_at": "", "issues": []})
    entry["status"] = "Corrected"
    save_qc(qc)
    return out_path, f"Saved corrected copy: {corr_name} (original untouched)."


def _qc_summary():
    """(checked, corrected, needs_correction, total_files, audit_entries)
    for stage_status() — from notes/qc.json + notes/qc_audit.json."""
    qc = load_qc()
    files = qc.get("files", {})
    checked = sum(1 for e in files.values()
                  if isinstance(e, dict) and e.get("status") in ("Checked", "Needs correction", "Corrected"))
    corrected = sum(1 for e in files.values()
                    if isinstance(e, dict) and e.get("status") == "Corrected")
    needs = sum(1 for e in files.values()
                if isinstance(e, dict) and e.get("status") == "Needs correction")
    return checked, corrected, needs, len(files), len(qc_load_audit())


# --- Insights builder helpers (v3.27) ---
def load_insights():
    """Saved insights for the current project, or a fresh empty list."""
    try:
        data = json.loads((NOTES_DIR / INSIGHTS_FILE).read_text(encoding="utf-8"))
        if isinstance(data, dict):
            data.setdefault("insights", [])
            return data
    except Exception:
        pass
    return {"insights": []}


def save_insights(data):
    """Save the insights (Drive + local, via the notes folder)."""
    return store_write_text("notes", INSIGHTS_FILE, json.dumps(data, ensure_ascii=False, indent=2))


def _ins_qual_findings():
    """Qual findings from SAVED analysis files: each entry {statement,
    quotes: [{quote, speaker, source}], section, analysis_file}. Only from
    saved files — nothing is generated here."""
    out = []
    for f in sorted(ANALYSIS_DIR.glob("*_analysis.txt")):
        text = f.read_text(encoding="utf-8", errors="replace")
        if looks_like_html(text):
            text = html_to_text(text)
        for entry in _ev_parse_findings(text):
            if entry.get("no_evidence") or not entry.get("text"):
                continue
            out.append({"statement": entry["text"], "quotes": entry.get("quotes", []),
                        "section": entry.get("section", ""), "analysis_file": f.name})
    return out


def _ins_quant_numbers():
    """Quant numbers from saved quant summaries (quant_*.txt): each entry
    {label, value, source}. Numbers come from the Quant page's computed
    tables — never invented here. Interpretation files are skipped."""
    out = []
    for f in sorted(ANALYSIS_DIR.glob("quant_*.txt")):
        if f.name.endswith("_interpretation.txt"):
            continue
        text = f.read_text(encoding="utf-8", errors="replace")
        for sec in parse_sections(text):
            for line in sec["lines"]:
                m = re.match(r"(.+?):\s*(\d+)\s*\(([\d.]+)%\)", line)
                if m:
                    out.append({"label": f"{sec['title']} — {m.group(1).strip()}",
                                "value": f"{m.group(2)} ({m.group(3)}%)",
                                "source": f.name})
    return out


def _ins_dg_sections():
    """Section titles from the saved Discussion Guide
    (notes/discussion_guide.txt) — what an insight can answer."""
    titles = []
    dg_path = NOTES_DIR / "discussion_guide.txt"
    if dg_path.exists():
        for b in _dg_parse(dg_path.read_text(encoding="utf-8", errors="replace")):
            if b.get("kind") in ("h2", "h3") and b.get("text"):
                titles.append(b["text"])
            elif b.get("kind") == "para" and re.match(r"^Section \d+\s*:", b.get("text", "")):
                titles.append(b["text"])
    return titles


def cross_interview_themes():
    """Theme x transcript presence from saved SINGLE-transcript analyses.
    Returns (themes, files): themes[title][filename] = finding count.
    Multi-interview analyses are excluded — they already combine
    interviews."""
    files = [f for f in sorted(ANALYSIS_DIR.glob("*_analysis.txt"))
             if not f.name.startswith("multi_")]
    themes = {}
    for f in files:
        text = f.read_text(encoding="utf-8", errors="replace")
        if looks_like_html(text):
            text = html_to_text(text)
        for sec in parse_sections(text):
            n = len([ln for ln in sec["lines"]
                     if ln and ln != "No evidence in transcripts"
                     and not ln.startswith("> ")])
            themes.setdefault(sec["title"], {})[f.name] = n
    return themes, [f.name for f in files]


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
            with st.spinner(T("Saving to Google Drive...", "Google Drive မှာ သိမ်းနေပါတယ်...")):
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
            with st.spinner(T("Saving to Google Drive...", "Google Drive မှာ သိမ်းနေပါတယ်...")):
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
            with st.spinner(T("Saving to Google Drive...", "Google Drive မှာ သိမ်းနေပါတယ်...")):
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

# --- Stage 2 AI-draft templates (v3.25). The Discussion Guide template
# follows the house FGD guide format: title + method line, conversational
# Welcome script, section heading rows "Section N: <Name> (<X> Mins)" with
# a bold objective line, numbered question groups with italic moderator
# instructions, bulleted main questions with nested probe sub-bullets,
# italic moderator notes, persona-specific probe blocks, tables where they
# fit the flow, timings summing to the session length (FGD 120 min / IDI
# 60 min), closing with thank-you + incentive line. A 2h FGD guide is about
# 2,000-2,500 words; an IDI guide about 1,000-1,500. The plain-text markup
# below is what the DG editor shows and what dg_to_docx() turns into a Word
# file — the AI must output exactly this markup, no other format.
DG_DRAFT_TEMPLATE = (
    "You are a senior market research moderator. Draft a Discussion Guide for "
    "this project from the client brief and the research design below. The "
    "moderator will read this guide word for word in the room, so write it as "
    "a spoken script, not a formal document.\n"
    "Rules: Base everything ONLY on the brief and the design. Do not invent "
    "objectives, brands, personas, prices, or questions the brief does not "
    "support. If a detail is missing (brand name, moderator name, incentive), "
    "write it as a [bracketed placeholder] such as [Brand X], [Moderator Name], "
    "[Incentive/Gift] — never guess a real name. Never invent a recording "
    "setup, dates, venues, participant names, or fees either — use "
    "[bracketed placeholders] such as [Recording — confirm], [Venue], "
    "[Date] or leave them out. Number sections correctly "
    "1, 2, 3, ... in order — never repeat a section number. Section timings "
    "must SUM to the session length: FGD = 120 minutes, IDI = 60 minutes. "
    "Keep the tone conversational and colloquial, like a friendly host — never "
    "formal. Mostly English with an occasional Myanmar line where the moderator "
    "needs it; never mix English and Myanmar in the same line. Build "
    "persona-specific probe blocks ONLY from the personas listed in the design; "
    "if the design lists no personas, do not invent persona groups. Adapt the "
    "format to this project — use the sample format as a starting structure "
    "only and write fresh questions from this project's own brief, objectives "
    "and target; do not copy sample questions word for word. Go deep: use "
    "laddering follow-ups ('why is that important?') and at least one "
    "projective question where the objectives allow it; add segment-specific "
    "probes where the design lists personas or bands. Before finishing, "
    "verify: every brief objective has a section; section timings sum to the "
    "session length; no invented moderator, recording, venue, date, or fee. "
    "Fix any failure before returning the guide. Length: "
    "FGD about 2,000-2,500 words, IDI about 1,000-1,500 words.\n\n"
    "Write the guide EXACTLY in this markup (the app turns it into a Word file):\n"
    "# DISCUSSION GUIDE FOR <Project name> (<Client name>)\n"
    "Qualitative Market Research (FGD) — or (IDI) for IDI projects\n"
    "## Welcome\n"
    "<welcome script: greeting, moderator intro ('I am [Moderator Name]'), "
    "'no right or wrong answer', recording note ONLY if the brief or design "
    "states recording (audio/video) — otherwise write '[Recording — confirm]' "
    "and do not assume the session is recorded, session length>\n"
    "## Section 1: <Name> (10 Mins)\n"
    "Objective: <one sentence — what this section must find out>\n"
    "### 1. <Question group name>\n"
    "> <moderator instruction, spoken to the room>\n"
    "- <main question>\n"
    "  - <nested probe>\n"
    "> (Note: <moderator note, e.g. what an answer signals>)\n"
    "[TABLE]\n"
    "<Header A> | <Header B> | <Header C>\n"
    "<cell text> | <cell text> | <cell text>\n"
    "[/TABLE]\n"
    "(use a [TABLE] block only where a table fits the flow, e.g. a daily timeline)\n"
    "## Section N: Closing & Thank You (5 Mins)\n"
    "<thank the group, ask if anything was missed>\n"
    "Please collect your [Incentive/Gift] at door.\n"
    "Thank you.\n"
)

QUESTIONNAIRE_DRAFT_TEMPLATE = (
    "You are a senior market research analyst. Draft a recruitment screener / "
    "questionnaire for this project from the client brief and the research "
    "design below. A recruiter will read it to real respondents, so write the "
    "questions in their everyday spoken language.\n"
    "Rules: Base everything ONLY on the brief and the design. Do not invent "
    "brands, products, prices, or screening rules the brief does not support — "
    "use [Brand X] style placeholders where a real name is missing. Never "
    "invent dates, venues, fees, interviewer or moderator names, or a "
    "recording setup — use [Venue], [Date], [Interviewer] style placeholders "
    "or leave them out. Never invent an exclusion criterion (for example a "
    "prior-research time window) — only use exclusions stated in the brief or "
    "in the design's screening criteria. Never assign a group unless the "
    "design lists quotas or groups — otherwise the outcome is only Continue "
    "or Terminate. Sensitive demographic questions (for example gender, "
    "observed) are recorded, never Terminate. The mode block after this "
    "template (screener-only, or full questionnaire) says which sections to "
    "build — follow it exactly. Question types: single_select, "
    "multi_select, ranking_top3 (pick the top 3), likert_1_5 (a statement, "
    "answered 1-5), rating (1-5 or 1-10), open_text, numeric, observed "
    "(observe and record — do not ask, e.g. gender with a quota note), "
    "contact (name + phone capture). No complex skip logic. Include an intro "
    "script the recruiter reads (study intro, session length, incentive, "
    "venue as [Venue]) and a short closing thank-you.\n"
    "Language: store every question ONCE with BOTH English and Myanmar text in "
    "parallel (same Q numbers and codes for both versions). Write the Myanmar "
    "in everyday colloquial Myanmar, no English mixed in — except food-category "
    "labels, which stay in English where the instrument uses English category "
    "names. If the design says English "
    "only, write the Myanmar text the same as the English text.\n"
    "XLSForm fields: set 'required' true for screening and essential questions, "
    "false otherwise. Where an outcome implies a condition, also write it as an "
    "XLSForm 'relevant' expression using the question names, e.g. "
    "\"${q11} = 'other'\" for an Other-specify field. Put recruiter guidance in "
    "'hint_en'/'hint_mm'. Leave 'appearance', 'constraint', and "
    "'constraint_message' empty unless a question needs one.\n\n"
    "Output JSON ONLY — one object, no other text:\n"
    '{"intro_en": "...", "intro_mm": "...", "closing_en": "...", "closing_mm": "...", '
    '"header_fields": ["FGD date", "Interviewer", "Respondent ID", "Confirmation"], '
    '"questions": [{"code": "Q1", "name": "q1", "section": "Screening", "question_en": "...", '
    '"question_mm": "...", "type": "single_select|multi_select|ranking_top3|likert_1_5|rating|open_text|numeric|date|note|observed|contact", '
    '"options_en": "option 1 | option 2", "options_mm": "option 1 | option 2", '
    '"outcome": "Yes → Terminate; No → Continue", "hint_en": "", "hint_mm": "", '
    '"required": true, "appearance": "", "relevant": "", "constraint": "", "constraint_message": ""}]}\n'
    "Options are separated by ' | ' (options contain commas, so never use commas as the separator). "
    "For numeric, date, note, observed, and contact questions, leave options empty.\n"
)

# v3.30 (P1/P2): the questionnaire draft is mode-aware. Qualitative IDI/FGD
# studies get a screener ONLY — the old template's "sections that follow the
# brief objectives" instruction put 14 discussion questions into the trial
# screener (Q8-Q21), and its group-assignment instruction made the model
# invent a six-group split the brief left to-confirm (Q21). Survey studies
# keep the full questionnaire. _run_q_draft() appends the matching block to
# the prompt; both blocks are also shown as Prompt Library variants.
QUESTIONNAIRE_SCREENER_MODE = (
    "MODE: SCREENER ONLY (Qualitative IDI/FGD, no Survey). Produce a "
    "screening section ONLY — as many questions as the screening criteria "
    "require (few criteria, about 10-15 questions; many criteria, 30-40 or "
    "more), eligibility only: who qualifies for the interview and who is "
    "screened out (category use, frequency, and any age/area/education/spend "
    "rule the brief states). Do not fix the count — follow the criteria. "
    "Do NOT add sections that follow the brief objectives — motivations, "
    "aspirations, culture, comparisons and opinions belong in the Discussion "
    "Guide, not in a screener. EVERY question carries an outcome rule in "
    "plain words: Terminate (which answer ends the interview) or Continue — "
    "nothing else. No group assignment, no quota tag. Use the design's "
    "screening criteria for those rules."
)
QUESTIONNAIRE_SURVEY_MODE = (
    "MODE: FULL QUESTIONNAIRE (Survey present). Start with a short screening "
    "section (simple fail to end: category use, frequency, and any "
    "age/area/education/spend rule the brief states). Then sections that "
    "follow the brief objectives, using the brief's respondent criteria for "
    "who the questions fit. EVERY question carries an outcome rule in plain "
    "words: Terminate (which answer ends the interview), Continue, a group "
    "assignment (e.g. '<answer> → Group <n>'), or a quota tag (e.g. "
    "'<segment>'). Use the design's personas, quotas, and screening criteria "
    "for those rules. About 15-30 questions in total — do not pad."
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
    "discussion_guide": {
        "name": "8. Discussion Guide Draft Engine",
        "icon": "💬",
        "purpose": "Drafts a moderator Discussion Guide from the saved brief + research design, in the house format (Welcome script, timed sections with objectives, numbered question groups, italic moderator instructions, nested probes, persona blocks, closing). One Gemini request per draft.",
        "used_by": "Research Design page (AI draft Discussion Guide)",
        "template": DG_DRAFT_TEMPLATE,
    },
    "questionnaire": {
        "name": "9. Questionnaire Draft Engine",
        "icon": "📋",
        "purpose": "Drafts a questionnaire structure from the saved brief + research design as JSON (section, question, type, options). Screener-only for Qualitative IDI/FGD (eligibility only; question count follows the screening criteria); screening first, then objective sections, when a Survey is present. Respondent language: everyday colloquial Myanmar unless the design says English. One Gemini request per draft.",
        "used_by": "Research Design page (AI draft Questionnaire)",
        "template": QUESTIONNAIRE_DRAFT_TEMPLATE,
        "variants": {
            "Screener only (Qualitative IDI/FGD)": QUESTIONNAIRE_SCREENER_MODE,
            "Full questionnaire (Survey)": QUESTIONNAIRE_SURVEY_MODE,
        },
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


def analysis_to_pptx(analysis_text, title_text, out_path, template="Business Clean",
                    subtitle=None, footer=None):
    """PPTX export — the main report deliverable (his rule: reports are
    mostly PowerPoint). subtitle goes on the title slide (client + month);
    footer (e.g. 'Base: N=82' for Quant reports) is printed small at the
    bottom of every content slide, so the N base is stated on every slide.
    Existing templates are unchanged."""
    t = TEMPLATES.get(template, TEMPLATES["Business Clean"])
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    _set_bg(slide, t["bg"])
    slide.shapes.title.text = title_text
    _color_runs(slide.shapes.title.text_frame.paragraphs, t["title"])
    slide.placeholders[1].text = subtitle or "Market Research AI — ResearchFlow"
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
            if footer:
                _fp = tf.add_paragraph()
                _fp.text = footer
                for run in _fp.runs:
                    run.font.size = Pt(9)
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


# --- Professional Reporting (v3.28, Phase 3 Report Quality) ---
# NO SAMPLE CONTENT RULE (user, 2026-11-17): the real project samples he
# shared are REFERENCE DATA for the website's STRUCTURE only — their format,
# section shapes, tables, and ordering. NO sample content may appear in this
# app: no findings, quotes, prices, dates, venues, client names, contacts,
# interviewer names, product codes, brand lists, township lists, MMK bands,
# or respondent data — not as defaults, not as placeholders, not as examples,
# not in help text, not in prompts. Every value below comes from the
# project's OWN brief.json / design.json / saved analysis, or from text the
# user writes on the page. New sections start EMPTY for a new project.


def _exec_summary_sections(stem, analysis_text):
    """Executive Summary for a report export. Returns (sections, source):
    'topline' — from the saved <stem>_topline.txt (Reports page 'Create
    executive summary'); 'key_findings' — from the analysis's own Key
    Findings section (first 5 lines); None — no summary available. The
    ORIGINAL analysis file is never changed — the summary is prepended only
    in the exported file."""
    topline_path = ANALYSIS_DIR / f"{stem}_topline.txt"
    if topline_path.exists():
        try:
            ttext = topline_path.read_text(encoding="utf-8", errors="replace")
            if looks_like_html(ttext):
                ttext = html_to_text(ttext)
            lines = []
            for sec in parse_sections(ttext):
                lines.extend(sec["lines"])
            if lines:
                return [{"title": "Executive Summary", "lines": lines}], "topline"
        except Exception:
            pass
    for sec in parse_sections(analysis_text):
        if "key finding" in sec["title"].lower() and sec["lines"]:
            return [{"title": "Executive Summary", "lines": sec["lines"][:5]}], "key_findings"
    return [], None


def _report_text_with_summary(analysis_text, stem, fallback_text=None):
    """Compose the export text: an Executive Summary section first, then the
    full analysis. Returns (text, source). When the analysis already has an
    Executive Summary section, nothing is prepended (source 'existing').
    fallback_text (the ORIGINAL analysis, when analysis_text is an assembled
    report) is where the Key Findings fallback looks — the assembled text
    may not have a Key Findings section any more."""
    if any("executive summary" in s["title"].lower() for s in parse_sections(analysis_text)):
        return str(analysis_text or ""), "existing"
    summary, source = _exec_summary_sections(stem, fallback_text if fallback_text is not None else analysis_text)
    if not summary:
        return str(analysis_text or ""), None
    parts = []
    for sec in summary:
        parts.append(f"## {sec['title']}")
        for line in sec["lines"]:
            parts.append(line if line.startswith("> ") else f"- {line}")
        parts.append("")
    parts.append(str(analysis_text or "").strip())
    return "\n".join(parts), source


_NON_EN_RE = re.compile(r"[\u1000-\u109F\u0400-\u04FF\u4E00-\u9FFF\u3040-\u30FF\u0600-\u06FF\u0E00-\u0E7F]")


def report_english_check(text):
    """Lines containing non-Latin script (Myanmar, Cyrillic, CJK, Arabic,
    Thai). Reports are English-only — typographic punctuation (—, –, “ ”)
    is fine and is NOT flagged. Returns [(line_no, line)]."""
    out = []
    for i, line in enumerate(str(text or "").splitlines(), start=1):
        if _NON_EN_RE.search(line):
            out.append((i, line.strip()[:160]))
    return out


def report_evidence_warnings(analysis_text):
    """Run the v3.27 verifier on the ORIGINAL analysis text (not the
    summary-prepended export, so findings are not double-counted). Returns
    (unverified_rows, total, note). note is set when there are no
    transcripts in the project to check against."""
    full_files, _pending = get_full_transcripts()
    if not full_files:
        return [], 0, "No transcripts in this project to check against"
    ttexts = {f.name: f.read_text(encoding="utf-8", errors="replace") for f in full_files}
    results, score = verify_analysis(analysis_text, ttexts)
    unverified = [r for r in results if r["status"] != "VERIFIED"]
    return unverified, len(results), ""


# Report structure template (v3.28, Phase 3 Report Quality): the section
# order of the house Qual report, from the sample report's structure —
# Agenda → Goal → Segment Introduction → Competitive Landscape & Strategy →
# Targeting Strategy → Summary of Segments' Insight → Deep Dive per segment
# → Winning Menu → Action Plan. _assemble_report() reorders the project's
# OWN analysis/brief/design/insights data into this shape for the PPTX/DOCX
# exports; report_template_coverage() shows which sections have source
# data and which are left out. NOTHING is invented: a template section with
# no source data is omitted from the export and listed as missing.
# What the template does NOT do (honest limit): the sample's strategic
# synthesis, archetype framing, images, and visual layout are the user's
# craft — the exporter produces the STRUCTURE with the project's real
# findings; he refines the synthesis/design in Google Docs/PowerPoint.
REPORT_TEMPLATE_SECTIONS = ["Agenda", "Goal", "Segment Introduction",
                            "Competitive Landscape & Strategy",
                            "Targeting Strategy",
                            "Summary of Segments' Insight",
                            "Deep Dive per segment", "Winning Menu",
                            "Action Plan"]

# Keyword buckets for deep-dive lines (heuristic — the exporter groups the
# project's OWN lines; it never writes new findings). Quote lines default
# to Motivations; anything else unmatched is Findings.
def _has_word(text, word):
    """Word match with a boundary at the START of the word, so 'action'
    matches 'Actions' but NOT 'satisfaction', and 'want' matches 'wanted'.
    Plain substring matching misfiled 'Satisfaction' sections under
    Action Plan (v3.28 fix)."""
    return re.search(r"\b" + re.escape(word), text or "") is not None


_DIVE_BUCKETS = [
    ("Occasions", ("breakfast", "lunch", "dinner", "morning", "noon",
                    "afternoon", "evening", "night", "midnight", "occasion",
                    "meal", "takeaway", "on-the-go")),
    ("Wishlist", ("want", "wish", "would like", "hope", "wishlist", "future")),
    ("Landscape & Channels", ("shop", "store", "competitor", "street",
                              "vendor", "channel", "outlet", "brand",
                              "market", "price", "location", "visit")),
    ("Need to Action", ("gap", "need", "action", "dream", "bridge",
                        "missing", "problem", "pain", "barrier")),
    ("Ranked Favorites", ("favorite", "favourite", "menu", "ranked",
                          "top 3", "most consumed", "most often")),
    ("Motivations", ("because", "motivation", "reason", "feel", "i want",
                     "love", "enjoy", "prefer", "makes me")),
]


def _dive_bucket(line):
    """Keyword bucket for one deep-dive line. Quote lines default to
    Motivations; anything else unmatched is Findings."""
    text = line[2:].strip().lower() if line.startswith("> ") else line.lower()
    for label, words in _DIVE_BUCKETS:
        if any(_has_word(text, w) for w in words):
            return label
    return "Motivations" if line.startswith("> ") else "Findings"


def _report_slots(analysis_text, brief, design, insights):
    """Map the project's OWN data onto the house Qual report template.
    Returns {"slots": {slot: {"lines": [...], "source": str}},
             "unmatched": [{"title": ..., "lines": [...]}]}.
    Each analysis section is placed in exactly ONE slot (no duplication, no
    loss). insights is the dict from load_insights() (or a plain list).
    Slot lines are already rendered as export markdown ('- ' bullets,
    '> ' quotes, plain label lines)."""
    brief = brief or {}
    design = design or {}
    if isinstance(insights, dict):
        insights = insights.get("insights", []) or []
    insights = insights or []
    sections = parse_sections(analysis_text or "")
    personas = [str((p or {}).get("name") or "").strip()
                for p in (design.get("personas") or [])]
    personas = [n for n in personas if n]
    slots = {}
    unmatched = []

    def _put(slot, lines, source):
        if not lines:
            return
        s = slots.setdefault(slot, {"lines": [], "source": source})
        s["lines"].extend(lines)
        if s["source"] != source:
            s["source"] = s["source"] + " + " + source

    def _bullet(line):
        return line if line.startswith("> ") else f"- {line}"

    # --- Goal: brief objectives + design method/N + the UNDERSTAND / DEFINE /
    # OPPORTUNITY framing (structure from the sample, content from his data).
    goal_lines = []
    if brief.get("background"):
        goal_lines.append(brief["background"])
    if brief.get("objectives"):
        goal_lines.append(f"UNDERSTAND — {brief['objectives']}")
    if personas:
        goal_lines.append(f"DEFINE — Personas: {', '.join(personas)}")
    _rat = (design.get("rationale") or design.get("goal") or "").strip()
    if _rat:
        goal_lines.append(f"OPPORTUNITY — {_rat}")
    _methods = design.get("methods") or []
    _approach = design.get("approach") or ""
    if _approach or _methods:
        goal_lines.append(f"Method: {(_approach + ' — ') if _approach else ''}{', '.join(_methods) if _methods else '—'}")
    _sessions = design.get("sessions_count") or 0
    _per = design.get("participants_per_session") or 0
    _sample = design.get("sample_size") or 0
    if _sessions and _per:
        goal_lines.append(f"Target: {_sessions * _per} respondents ({_sessions} groups × {_per})")
    elif _sample:
        goal_lines.append(f"Target: {_sample} respondents")
    elif (design.get("target") or "").strip():
        goal_lines.append(f"Target: {design['target'].strip()}")
    elif (brief.get("target_audience") or "").strip():
        goal_lines.append(f"Target: {brief['target_audience'].strip()}")
    _put("goal", goal_lines, "Client Brief + Research Design")

    # --- Segment Introduction + Targeting Strategy: from the design personas.
    _intro_lines = []
    if (brief.get("target_audience") or "").strip():
        _intro_lines.append(brief["target_audience"].strip())
    _target_lines = []
    for p in (design.get("personas") or []):
        _nm = str((p or {}).get("name") or "").strip()
        if not _nm:
            continue
        _ds = str((p or {}).get("description") or "").strip()
        _vf = str((p or {}).get("value_framing") or "").strip()
        _intro_lines.append(f"{_nm}: {_ds}" if _ds else _nm)
        _target_lines.append(f"{_nm} — {_ds if _ds else '—'} (Value: {_vf if _vf else '—'})")
    for q in (design.get("quotas") or []):
        _sg = str((q or {}).get("segment") or "").strip()
        _sh = str((q or {}).get("share") or "").strip()
        if _sg or _sh:
            _intro_lines.append(f"Quota: {_sg} — {_sh}")
            _target_lines.append(f"Quota: {_sg} — {_sh}")
    _put("segment_intro", _intro_lines, "Research Design personas + Brief")
    _put("targeting", _target_lines, "Research Design")

    # --- Summary: saved insights first, then summary-classified sections.
    _sum_lines = []
    for ins in insights:
        _st = str((ins or {}).get("statement") or "").strip()
        if not _st:
            continue
        _sum_lines.append(f"- {_st}")
        for _q in ((ins or {}).get("qual_evidence") or [])[:1]:
            _qt = str((_q or {}).get("quote") or "").strip()
            if _qt:
                _sum_lines.append(f"> \"{_qt}\"")
    if _sum_lines:
        _put("summary", _sum_lines, f"{len(_sum_lines)} saved insight(s)")

    # --- Classify the analysis sections (title keywords; persona-name match
    # takes priority over every keyword slot).
    def _classify(title):
        t = title.lower()
        if "executive summary" in t:
            return "executive_summary"
        for n in personas:
            if n.lower() in t:
                return ("deep_dive", n)
        if any(_has_word(t, w) for w in ("goal", "objective")):
            return "goal"
        if any(_has_word(t, w) for w in ("summary", "insight")):
            return "summary"
        if any(_has_word(t, w) for w in ("landscape", "competitive", "competitor", "positioning", "gap", "price")):
            return "landscape"
        if any(_has_word(t, w) for w in ("winning", "menu")):
            return "winning_menu"
        if _has_word(t, "strategy"):
            return "landscape"
        if any(_has_word(t, w) for w in ("action", "recommendation", "next step")):
            return "action_plan"
        if any(_has_word(t, w) for w in ("deep dive", "segment")):
            return "deep_dive_generic"
        return None

    for sec in sections:
        _title = sec["title"]
        _lines = [_bullet(x) for x in sec["lines"]]
        _slot = _classify(_title)
        if _slot == "executive_summary":
            _put("executive_summary", _lines, "Analysis")
        elif isinstance(_slot, tuple):
            _put(f"deep_dive:{_slot[1]}", _lines, f"Analysis: {_title}")
        elif _slot == "goal":
            _put("goal", _lines, f"Analysis: {_title}")
        elif _slot == "summary":
            _put("summary", _lines, f"Analysis: {_title}")
        elif _slot == "landscape":
            _put("landscape", _lines, f"Analysis: {_title}")
        elif _slot == "winning_menu":
            _put("winning_menu", _lines, f"Analysis: {_title}")
        elif _slot == "action_plan":
            _put("action_plan", _lines, f"Analysis: {_title}")
        elif _slot == "deep_dive_generic":
            _mentioned = [n for n in personas
                          if n.lower() in " ".join(sec["lines"]).lower()]
            if len(_mentioned) == 1:
                _put(f"deep_dive:{_mentioned[0]}", _lines, f"Analysis: {_title}")
            else:
                _put("deep_dive_generic", _lines, f"Analysis: {_title}")
        else:
            # Unmatched section: if its lines mention exactly one persona,
            # the whole section belongs to that persona's deep dive.
            _mentioned = [n for n in personas
                          if n.lower() in " ".join(sec["lines"]).lower()]
            if len(_mentioned) == 1:
                _put(f"deep_dive:{_mentioned[0]}", _lines, f"Analysis: {_title}")
            else:
                unmatched.append({"title": _title, "lines": sec["lines"]})

    # --- Deep dives: profile lines first, then the persona's lines bucketed
    # (Occasions → Motivations → Need to Action → Landscape & Channels →
    # Wishlist → Ranked Favorites → Findings). Insights for the persona join.
    _bucket_order = ["Occasions", "Motivations", "Need to Action",
                     "Landscape & Channels", "Wishlist", "Ranked Favorites",
                     "Findings"]
    for p in (design.get("personas") or []):
        _nm = str((p or {}).get("name") or "").strip()
        if not _nm:
            continue
        _key = f"deep_dive:{_nm}"
        _raw = list(slots.get(_key, {}).get("lines", []))
        for ins in insights:
            if str((ins or {}).get("persona") or "").strip() == _nm:
                _st = str((ins or {}).get("statement") or "").strip()
                if _st:
                    _raw.append(f"- {_st}")
        if not _raw and not str((p or {}).get("description") or "").strip():
            continue
        _buckets = {b: [] for b in _bucket_order}
        for _ln in _raw:
            _buckets[_dive_bucket(_ln)].append(_ln)
        _dlines = []
        _ds = str((p or {}).get("description") or "").strip()
        _vf = str((p or {}).get("value_framing") or "").strip()
        if _ds:
            _dlines.append(f"Profile: {_ds}" + (f" (Value: {_vf})" if _vf else ""))
        for _b in _bucket_order:
            if _buckets[_b]:
                _dlines.append(f"{_b}:")
                _dlines.extend(_buckets[_b])
        slots[_key] = {"lines": _dlines, "source": "Research Design + Analysis"}
    return {"slots": slots, "unmatched": unmatched}


def _assemble_report(analysis_text, brief, design, insights):
    """Assemble the project's data into the house Qual report structure.
    Returns (text, included, missing): text is the report in template order
    (Executive Summary first when the analysis has one, then Agenda, Goal →
    ... → Action Plan, then Additional Findings for sections that fit no
    slot); included/missing are template section labels. Sections with no
    source data are omitted (never invented). When nothing maps at all, the
    original analysis text is returned unchanged (safe fallback)."""
    data = _report_slots(analysis_text, brief, design, insights)
    slots = data["slots"]
    unmatched = data["unmatched"]
    labels = {"goal": "Goal", "segment_intro": "Segment Introduction",
              "landscape": "Competitive Landscape & Strategy",
              "targeting": "Targeting Strategy",
              "summary": "Summary of Segments' Insight",
              "winning_menu": "Winning Menu", "action_plan": "Action Plan"}
    order = ["goal", "segment_intro", "landscape", "targeting", "summary",
             "winning_menu", "action_plan"]
    personas = [str((p or {}).get("name") or "").strip()
                for p in ((design or {}).get("personas") or [])]
    personas = [n for n in personas if n]
    parts = []
    included = []
    missing = []
    if "executive_summary" in slots:
        parts.append("## Executive Summary")
        parts.extend(slots["executive_summary"]["lines"])
        parts.append("")
    # Agenda lists only the sections actually included.
    _agenda = []
    for _key in order:
        if _key in slots:
            _agenda.append(labels[_key])
    for _n in personas:
        if f"deep_dive:{_n}" in slots:
            _agenda.append(f"Deep Dive: {_n}")
    if "deep_dive_generic" in slots:
        _agenda.append("Deep Dive per segment")
    if unmatched:
        _agenda.append("Additional Findings")
    if _agenda:
        parts.append("## Agenda")
        parts.extend(f"{i}. {_item}" for i, _item in enumerate(_agenda, start=1))
        parts.append("")
        included.append("Agenda")
    for _key in order:
        if _key in slots:
            parts.append(f"## {labels[_key]}")
            parts.extend(slots[_key]["lines"])
            parts.append("")
            included.append(labels[_key])
        else:
            missing.append(labels[_key])
    for _n in personas:
        _key = f"deep_dive:{_n}"
        if _key in slots:
            parts.append(f"## Deep Dive: {_n}")
            parts.extend(slots[_key]["lines"])
            parts.append("")
            included.append(f"Deep Dive: {_n}")
        else:
            missing.append(f"Deep Dive: {_n}")
    if "deep_dive_generic" in slots:
        parts.append("## Deep Dive per segment")
        parts.extend(slots["deep_dive_generic"]["lines"])
        parts.append("")
        included.append("Deep Dive per segment")
    if unmatched:
        parts.append("## Additional Findings")
        for _sec in unmatched:
            parts.append(f"{_sec['title']}:")
            parts.extend(f"- {x}" if not x.startswith("> ") else x for x in _sec["lines"])
        parts.append("")
        included.append("Additional Findings")
    if not included:
        return str(analysis_text or ""), [], [l for l in REPORT_TEMPLATE_SECTIONS if l != "Agenda"]
    if not personas and "Deep Dive per segment" not in included:
        missing.append("Deep Dive per segment")
    return "\n".join(parts).strip() + "\n", included, missing


def report_template_coverage(analysis_text, brief, design, insights=None):
    """Coverage of the Qual report template in one analysis. Returns a list
    of {label, present, detail}. Returns [] when the design says
    Quantitative. Detail names the source, or where the data would come
    from when the section is missing."""
    if (design or {}).get("approach") == "Quantitative":
        return []
    _text, included, missing = _assemble_report(analysis_text, brief, design, insights)
    data = _report_slots(analysis_text, brief, design, insights)
    slots = data["slots"]
    _hints = {
        "Goal": "Add objectives in the Client Brief, or a Goal/Objective section in the analysis.",
        "Segment Introduction": "Add personas in the Research Design, or target audience in the Brief.",
        "Competitive Landscape & Strategy": "Add a Competitive Landscape / Strategy section in the analysis.",
        "Targeting Strategy": "Add personas in the Research Design.",
        "Summary of Segments' Insight": "Save insights on the Insights page, or add a Summary section in the analysis.",
        "Winning Menu": "Add a Winning Menu / Strategy section in the analysis.",
        "Action Plan": "Add an Action Plan / Recommendations section in the analysis.",
        "Deep Dive per segment": "Add findings for this segment in the analysis.",
    }
    rows = []
    for _label in REPORT_TEMPLATE_SECTIONS:
        if _label == "Agenda":
            rows.append({"label": "Agenda", "present": bool(included),
                         "detail": "Lists the sections included in this export"})
            continue
        if _label == "Deep Dive per segment":
            _personas = [str((p or {}).get("name") or "").strip()
                         for p in ((design or {}).get("personas") or [])]
            _personas = [n for n in _personas if n]
            if _personas:
                for _n in _personas:
                    _key = f"deep_dive:{_n}"
                    _present = f"Deep Dive: {_n}" in included
                    rows.append({"label": f"Deep Dive: {_n}", "present": _present,
                                 "detail": slots.get(_key, {}).get("source", "") or "Add findings for this persona in the analysis."})
            else:
                _present = "Deep Dive per segment" in included
                rows.append({"label": "Deep Dive per segment", "present": _present,
                             "detail": slots.get("deep_dive_generic", {}).get("source", "") or _hints["Deep Dive per segment"]})
            continue
        _present = _label in included
        _src = ""
        _slot_key = {"Goal": "goal", "Segment Introduction": "segment_intro",
                     "Competitive Landscape & Strategy": "landscape",
                     "Targeting Strategy": "targeting",
                     "Summary of Segments' Insight": "summary",
                     "Winning Menu": "winning_menu",
                     "Action Plan": "action_plan"}.get(_label, "")
        if _present and _slot_key:
            _src = slots.get(_slot_key, {}).get("source", "")
        rows.append({"label": _label, "present": _present,
                     "detail": _src or _hints.get(_label, "")})
    if "Additional Findings" in included:
        rows.append({"label": "Additional Findings", "present": True,
                     "detail": "Sections that fit no template slot — kept verbatim at the end"})
    return rows


# Quant report template (v3.28, Phase 3 Report Quality): the section order
# of the house Quant report, from the Quant sample's STRUCTURE only —
# Actionable Insights → Project Timeline → Program Overview → Executive
# Summary → Question Analysis → VOC Feedback → Overall Satisfaction → Tier
# Evaluation → Customer Sentiment & Feedback → Recommendations / Action
# Plan → Closing. Numbers come ONLY from the saved quant summaries
# (quant_*.txt, computed in Python on the Quant page) and the project's own
# analysis/brief/design — never AI-generated, never sample content. Top-Box
# = share rating 4–5 on a 1–5 scale, computed here from the saved counts.
# QUANT CONVENTIONS followed: N base stated on every PPTX slide (the page
# passes it as the exporter's footer); cross-tabs split Total + segments;
# every finding keeps its question code [Q: ...] when the question name is
# a code; VOC direct quotes stay verbatim; counts + % together.
# What the template does NOT do (honest limit): the per-tier comparison
# table, VOC theme shares, and archetype framing need the project's tier
# structure and coded open text — the exporter produces the STRUCTURE with
# the project's real numbers; he builds those tables in PowerPoint.
QUANT_TEMPLATE_SECTIONS = ["Actionable Insights", "Project Timeline",
    "Program Overview", "Executive Summary", "Question Analysis",
    "VOC Feedback", "Overall Satisfaction", "Tier Evaluation",
    "Customer Sentiment & Feedback", "Recommendations / Action Plan", "Closing"]


def _load_quant_texts():
    """Saved quant summaries (quant_*.txt, interpretation files excluded):
    [(name, text)]. Empty when the project has no survey data."""
    out = []
    for f in sorted(ANALYSIS_DIR.glob("quant_*.txt")):
        if f.name.endswith("_interpretation.txt"):
            continue
        try:
            out.append((f.name, f.read_text(encoding="utf-8", errors="replace")))
        except Exception:
            pass
    return out


def _quant_total_n(quant_texts):
    """Total survey base N from the quant files' 'Rows: N' line (the first
    one found). None when there is no quant data."""
    for _name, text in (quant_texts or []):
        for line in str(text or "").splitlines():
            m = re.match(r"\s*Rows:\s*(\d+)\s*$", line)
            if m:
                return int(m.group(1))
    return None


def _quant_question_stats(quant_texts):
    """Per-question stats from the saved quant summaries. Returns a list of
    {question, n, counts, percents, mean, top_box, source} for each
    '## <question> (n=N)' section. Counts come from the 'val: count (pct%)'
    lines (parse_sections() already stripped the '- ' bullets); mean from
    the 'Mean: X.XX' line when present, else the weighted mean of numeric
    values; top_box = share of values 4–5 on a 1–5 scale (None when the
    question is not a 1–5 numeric scale)."""
    out = []
    for _name, text in (quant_texts or []):
        for sec in parse_sections(text):
            m = re.match(r"(.+?)\s*\(n=(\d+)\)\s*$", sec["title"])
            if not m:
                continue
            counts, percents, mean = {}, {}, None
            for line in sec["lines"]:
                mm = re.match(r"(.+?):\s*(\d+)\s*\(([\d.]+)%\)\s*$", line)
                if mm:
                    try:
                        counts[mm.group(1).strip()] = int(mm.group(2))
                        percents[mm.group(1).strip()] = float(mm.group(3))
                    except ValueError:
                        pass
                mn = re.match(r"Mean:\s*([\d.]+)\s*$", line)
                if mn:
                    try:
                        mean = float(mn.group(1))
                    except ValueError:
                        pass
            if not counts and mean is None:
                continue
            if mean is None:
                try:
                    vals = [(float(v), c) for v, c in counts.items()]
                    tot = sum(c for _v, c in vals)
                    if tot:
                        mean = round(sum(v * c for v, c in vals) / tot, 2)
                except ValueError:
                    pass
            top_box = None
            try:
                vals = {float(v): c for v, c in counts.items()}
                if vals and set(vals) <= {1.0, 2.0, 3.0, 4.0, 5.0}:
                    tot = sum(vals.values())
                    if tot:
                        top_box = round((vals.get(4.0, 0) + vals.get(5.0, 0)) / tot * 100, 1)
            except ValueError:
                pass
            out.append({"question": m.group(1).strip(), "n": int(m.group(2)),
                        "counts": counts, "percents": percents,
                        "mean": mean, "top_box": top_box, "source": _name})
    return out


def _quant_interpretation_text():
    """Text of the saved quant interpretation files
    (quant_*_interpretation.txt), joined. Empty string when there is none."""
    parts = []
    for f in sorted(ANALYSIS_DIR.glob("quant_*_interpretation.txt")):
        try:
            t = f.read_text(encoding="utf-8", errors="replace")
            if looks_like_html(t):
                t = html_to_text(t)
            parts.append(t)
        except Exception:
            pass
    return "\n".join(parts)


def _pick_report_template(design):
    """Which report structure template the PPTX export uses: 'quant' when
    the design says Quantitative; 'qual' otherwise. Mixed uses the methods:
    a survey-only Mixed project gets the Quant template, anything with
    FGD/IDI gets Qual (its Quant numbers still reach the Summary through
    saved insights). Returns 'qual' or 'quant'."""
    design = design or {}
    if design.get("approach") == "Quantitative":
        return "quant"
    if design.get("approach") == "Mixed":
        methods = [str(m) for m in (design.get("methods") or [])]
        if methods and all(m == "Survey" for m in methods):
            return "quant"
    return "qual"


def _assemble_quant_report(analysis_text, brief, design, insights=None, stem=""):
    """Assemble the project's data into the house Quant report structure.
    Returns (text, included, missing): text is the report in template order
    (Actionable Insights → ... → Closing); included/missing are template
    section labels. Sections with no source data are omitted (never
    invented). When nothing maps at all, the original analysis text is
    returned unchanged (safe fallback). Each analysis section is used at
    most once across the VOC / Sentiment / Tier / Recommendations slots."""
    brief = brief or {}
    design = design or {}
    quant_texts = _load_quant_texts()
    total_n = _quant_total_n(quant_texts)
    stats = _quant_question_stats(quant_texts)
    interp = _quant_interpretation_text()
    _secs = parse_sections(analysis_text)
    _used = set()

    def _match(*words, skip_used=True):
        found = []
        for _i, _s in enumerate(_secs):
            if skip_used and _i in _used:
                continue
            if any(_has_word(_s["title"].lower(), w) for w in words):
                found.append((_i, _s))
        return found

    def _take(pairs):
        lines = []
        for _i, _s in pairs:
            _used.add(_i)
            lines.extend(_s["lines"])
        return lines

    def _dedupe(lines):
        seen, out = set(), []
        for ln in lines:
            if ln not in seen:
                seen.add(ln)
                out.append(ln)
        return out

    slots = {}
    # Key Findings section index — its lines feed Actionable Insights and
    # the Executive Summary fallback; mark it used so those lines are not
    # repeated again under Additional Findings below.
    _kf_idx = None
    for _i, _s in enumerate(_secs):
        if "key finding" in _s["title"].lower() and _s["lines"]:
            _kf_idx = _i
            break
    # 1. Actionable Insights — headline from the quant interpretation, else
    # the analysis Key Findings; the validation sample N first.
    _ai = []
    if interp:
        for _s in parse_sections(interp):
            if _s["lines"]:
                _ai = _s["lines"][:5]
                break
    if not _ai and _kf_idx is not None:
        _ai = _secs[_kf_idx]["lines"][:5]
        _used.add(_kf_idx)
    if total_n:
        _ai = [f"Validation sample: N = {total_n}"] + _ai
    if _ai:
        slots["insights"] = _dedupe(_ai)
    # 2. Project Timeline — brief dates + fieldwork plan session dates ONLY.
    _tl = []
    if brief.get("timeline_start") or brief.get("timeline_end"):
        _tl.append(f"Project timeline: {brief.get('timeline_start') or '—'} to {brief.get('timeline_end') or '—'}")
    for _s in (design.get("fieldwork_plan") or []):
        _d = str((_s or {}).get("date") or "").strip()
        if _d:
            _tl.append(f"{_d} — {str((_s or {}).get('session') or (_s or {}).get('group') or 'Session').strip()}")
    if _tl:
        slots["timeline"] = _tl
    # 3. Program Overview — brief background/scope + design goal + analysis
    # overview sections.
    _ov = []
    if brief.get("background"):
        _ov.append(brief["background"])
    if brief.get("scope"):
        _ov.append(f"Scope: {brief['scope']}")
    if design.get("goal"):
        _ov.append(f"Goal: {design['goal']}")
    _ov_pairs = []
    for _i, _s in _match("overview", "current"):
        _t = _s["title"].lower()
        # A "Current ... Feedback" section is feedback, not overview — leave
        # it for the VOC slot below. ("program" is deliberately NOT a match
        # word here: it would eat "New Program Sentiment".)
        if "current" in _t and any(_has_word(_t, w) for w in ("feedback", "voc", "pain", "friction", "sentiment", "satisfaction")):
            continue
        _ov_pairs.append((_i, _s))
    _ov.extend(_take(_ov_pairs))
    if _dedupe(_ov):
        slots["overview"] = _dedupe(_ov)
    # 4. Executive Summary — KPI lines from the quant stats (N, mean/5.0,
    # Top-Box) + the topline / Key Findings narrative.
    _es = []
    if total_n:
        _es.append(f"Validation sample: N = {total_n}")
    for _q in stats:
        if _q["mean"] is not None:
            _line = f"{_q['question']} (n={_q['n']}): mean {_q['mean']:.2f}/5.0"
            if _q["top_box"] is not None:
                _line += f" · Top-Box (4–5): {_q['top_box']}%"
            _es.append(_line)
    _sum_secs, _sum_src = _exec_summary_sections(stem, analysis_text)
    if _sum_src == "key_findings" and _kf_idx is not None:
        _used.add(_kf_idx)
    for _sec in _sum_secs:
        _es.extend(_sec["lines"])
    if _dedupe(_es):
        slots["exec_summary"] = _dedupe(_es)
    # 5. Question Analysis — one section per survey question (counts + %,
    # mean, Top-Box) plus the saved cross-tab tables verbatim.
    _qa = []
    for _q in stats:
        _lines = [f"Base: n={_q['n']}"]
        for _v, _c in _q["counts"].items():
            _pct = _q["percents"].get(_v)
            if _pct is None and _q["n"]:
                _pct = round(_c / _q["n"] * 100, 1)
            _lines.append(f"{_v}: {_c} ({_pct}%)")
        if _q["mean"] is not None:
            _lines.append(f"Mean: {_q['mean']:.2f}")
        if _q["top_box"] is not None:
            _lines.append(f"Top-Box (4–5): {_q['top_box']}%")
        _qn = _q["question"]
        # Question codes stay cited: a code-like question name gets [Q: ...].
        _title = f"{_qn} [Q: {_qn}]" if re.match(r"^[A-Za-z]{1,3}\d+[A-Za-z0-9]*$", _qn) else _qn
        _qa.append((_title, _lines))
    for _name, _text in quant_texts:
        for _s in parse_sections(_text):
            if _s["title"].strip().lower() == "cross-tab" and _s["lines"]:
                _qa.append((f"Cross-tab: {_s['lines'][0]}", _s["lines"][1:]))
    if _qa:
        slots["questions"] = _qa
    # 6. VOC Feedback — analysis feedback sections + open-text counts.
    _voc = _take(_match("voc", "voice of customer", "feedback", "friction",
                        "pain", "complaint", "inconvenient"))
    for _name, _text in quant_texts:
        for _s in parse_sections(_text):
            for _ln in _s["lines"]:
                if _ln.startswith("Open text"):
                    _voc.append(f"{_s['title']}: {_ln}")
    if _dedupe(_voc):
        slots["voc"] = _dedupe(_voc)
    # 7. Overall Satisfaction — satisfaction-titled questions (mean +
    # Top-Box + most common answer, which is the prefer-% when the question
    # asks what respondents prefer).
    _sat = []
    for _q in stats:
        if any(w in _q["question"].lower() for w in ("satisf", "prefer", "overall", "value", "rating")):
            _line = f"{_q['question']} (n={_q['n']})"
            if _q["mean"] is not None:
                _line += f": mean {_q['mean']:.2f}/5.0"
            if _q["top_box"] is not None:
                _line += f" · Top-Box (4–5): {_q['top_box']}%"
            if _q["counts"]:
                _top_v = max(_q["counts"], key=lambda v: _q["counts"][v])
                _top_pct = _q["percents"].get(_top_v)
                if _top_pct is None and _q["n"]:
                    _top_pct = round(_q["counts"][_top_v] / _q["n"] * 100, 1)
                _line += f" · Most common: {_top_v} — {_q['counts'][_top_v]} ({_top_pct}%)"
            _sat.append(_line)
    if _sat:
        slots["satisfaction"] = _sat
    # 8. Tier Evaluation — analysis tier/benefit sections + tier-titled
    # survey questions (mean + Top-Box per benefit group).
    _tier = _take(_match("tier", "benefit"))
    for _q in stats:
        if any(w in _q["question"].lower() for w in ("tier", "benefit", "feature", "reward")):
            _line = f"{_q['question']} (n={_q['n']})"
            if _q["mean"] is not None:
                _line += f": mean {_q['mean']:.2f}/5.0"
            if _q["top_box"] is not None:
                _line += f" · Top-Box (4–5): {_q['top_box']}%"
            _tier.append(_line)
    if _dedupe(_tier):
        slots["tier"] = _dedupe(_tier)
    # 9. Customer Sentiment & Feedback — analysis sentiment sections +
    # saved insights (statement + first verbatim quote).
    _sent = _take(_match("sentiment", "new program", "improve", "change", "suggest"))
    for _ins in ((insights or {}).get("insights") or []):
        _st = str((_ins or {}).get("statement") or "").strip()
        if _st:
            _sent.append(_st)
        for _ev in ((_ins or {}).get("qual_evidence") or [])[:1]:
            _qt = str((_ev or {}).get("quote") or "").strip()
            if _qt:
                _sent.append(f"> \"{_qt}\"")
    if _dedupe(_sent):
        slots["sentiment"] = _dedupe(_sent)
    # 10. Recommendations / Action Plan.
    _rec = _take(_match("recommendation", "action plan", "action", "next step"))
    if _dedupe(_rec):
        slots["action_plan"] = _dedupe(_rec)
    # 11. Closing — client contact from the brief.
    _close = []
    if brief.get("client_name"):
        _close.append(f"Prepared for: {brief['client_name']}")
    if brief.get("client_contact"):
        _close.append(f"Contact: {brief['client_contact']}")
    if brief.get("client_email"):
        _close.append(f"Email: {brief['client_email']}")
    if _close:
        slots["closing"] = _close

    labels = {"insights": "Actionable Insights", "timeline": "Project Timeline",
              "overview": "Program Overview", "exec_summary": "Executive Summary",
              "questions": "Question Analysis", "voc": "VOC Feedback",
              "satisfaction": "Overall Satisfaction", "tier": "Tier Evaluation",
              "sentiment": "Customer Sentiment & Feedback",
              "action_plan": "Recommendations / Action Plan", "closing": "Closing"}
    order = ["insights", "timeline", "overview", "exec_summary", "questions",
             "voc", "satisfaction", "tier", "sentiment", "action_plan", "closing"]
    parts, included, missing = [], [], []
    for _key in order:
        if _key == "closing":
            continue  # emitted last, after Additional Findings
        if _key not in slots:
            missing.append(labels[_key])
            continue
        included.append(labels[_key])
        if _key == "questions":
            for _title, _lines in slots["questions"]:
                parts.append(f"## {_title}")
                parts.extend(_lines)
                parts.append("")
        else:
            parts.append(f"## {labels[_key]}")
            parts.extend(f"- {x}" if not x.startswith("> ") else x for x in slots[_key])
            parts.append("")
    # Unmatched analysis sections are kept verbatim (no loss) under
    # Additional Findings — like the Qual assembler.
    _extra = [(_i, _s) for _i, _s in enumerate(_secs) if _i not in _used]
    if _extra:
        parts.append("## Additional Findings")
        for _i, _s in _extra:
            parts.append(f"{_s['title']}:")
            parts.extend(f"- {x}" if not x.startswith("> ") else x for x in _s["lines"])
        parts.append("")
        included.append("Additional Findings")
    if "closing" in slots:
        parts.append("## Closing")
        parts.extend(f"- {x}" if not x.startswith("> ") else x for x in slots["closing"])
        parts.append("")
        included.append("Closing")
    else:
        missing.append("Closing")
    if not included:
        return str(analysis_text or ""), [], list(QUANT_TEMPLATE_SECTIONS)
    return "\n".join(parts).strip() + "\n", included, missing


def report_template_coverage_quant(analysis_text, brief, design, insights=None, stem=""):
    """Coverage of the Quant report template in one analysis. Returns a list
    of {label, present, detail}. Returns [] when the design says
    Qualitative. Detail names the source, or where the data would come from
    when the section is missing."""
    if (design or {}).get("approach") == "Qualitative":
        return []
    _text, included, missing = _assemble_quant_report(analysis_text, brief, design, insights, stem)
    _qtexts = _load_quant_texts()
    _qn = _quant_total_n(_qtexts)
    _stats = _quant_question_stats(_qtexts)
    _interp = _quant_interpretation_text()
    _hints = {
        "Actionable Insights": "Save a quant interpretation, or add a Key Findings section in the analysis.",
        "Project Timeline": "Add timeline dates in the Client Brief, or session dates in the Research Design fieldwork plan.",
        "Program Overview": "Add background / scope in the Client Brief.",
        "Executive Summary": "Create a topline on the Reports page, or add a Key Findings section.",
        "Question Analysis": "Upload survey data and save a quant summary on the Quantitative Analysis page.",
        "VOC Feedback": "Add a Feedback / VOC section in the analysis.",
        "Overall Satisfaction": "Add a satisfaction question to the survey, or a Satisfaction section in the analysis.",
        "Tier Evaluation": "Add a tier / benefit evaluation section in the analysis — per-tier comparison tables are built in PowerPoint.",
        "Customer Sentiment & Feedback": "Add a Sentiment / Feedback section in the analysis, or save insights with quotes.",
        "Recommendations / Action Plan": "Add a Recommendations / Action Plan section in the analysis.",
        "Closing": "Add client contact in the Client Brief.",
    }
    _src = {
        "Actionable Insights": ("Headline from the quant interpretation" if _interp else "Headline from the analysis Key Findings") + (f" · N = {_qn}" if _qn else ""),
        "Project Timeline": "Dates from the Client Brief / fieldwork plan",
        "Program Overview": "Background / scope from the Client Brief",
        "Executive Summary": "KPI cards from the saved quant summaries" + (" + topline" if _sum_src_topline(stem) else ""),
        "Question Analysis": f"{len(_stats)} question(s) from the saved quant summaries",
        "VOC Feedback": "Feedback sections from the analysis",
        "Overall Satisfaction": "Mean + Top-Box from the saved quant summaries",
        "Tier Evaluation": "Tier / benefit sections from the analysis",
        "Customer Sentiment & Feedback": "Sentiment sections from the analysis / saved insights",
        "Recommendations / Action Plan": "Recommendations from the analysis",
        "Closing": "Contact from the Client Brief",
    }
    rows = []
    for _label in QUANT_TEMPLATE_SECTIONS:
        _present = _label in included
        rows.append({"label": _label, "present": _present,
                     "detail": (_src.get(_label, "") if _present else _hints.get(_label, ""))})
    if "Additional Findings" in included:
        rows.append({"label": "Additional Findings", "present": True,
                     "detail": "Sections that fit no template slot — kept verbatim at the end"})
    return rows


def _sum_src_topline(stem):
    """True when a saved topline exists for this analysis stem."""
    return bool(stem) and (ANALYSIS_DIR / f"{stem}_topline.txt").exists()


# Proposal (v3.28, draft only): sections follow the house proposal's shape —
# Goal → Research Design → Target → Rationale → Timeline → Deliverables →
# Costs → Terms. Content comes ONLY from the project's own brief.json +
# design.json. The Costs section is EMPTY on purpose: fee/tax/payment fields
# are blank for him to fill — NEVER prefill sample fees or terms.
PROPOSAL_COST_ROWS = 4  # blank fee lines in the Costs table


def _proposal_session_minutes(methods):
    """Session length in minutes from the design's methods (house convention:
    FGD 120 min, IDI 60 min). Returns a list of (method, minutes) pairs."""
    out = []
    for m in (methods or []):
        if m == "FGD":
            out.append(("FGD", 120))
        elif m == "IDI":
            out.append(("IDI", 60))
    return out


def proposal_build(brief, design):
    """Build the proposal content from the project's own brief + design.
    Returns (blocks, missing): blocks = [{"kind": "title"/"h1"/"para"/
    "table", "text": ..., "rows": [[...]]}], missing = list of inputs that
    are absent (brief/design). When brief AND design are both empty, blocks
    is [] and missing names them — the page then points to Stage 1/2 and
    never invents a proposal. Unset values show as '—' or an explicit
    '[fill in]' line, never a guessed number, date, or fee."""
    brief = brief or {}
    design = design or {}
    missing = []
    if not brief:
        missing.append("Client Brief")
    if not design:
        missing.append("Research Design")
    if missing:
        return [], missing
    blocks = []
    title = (brief.get("project_code") or brief.get("client_name") or "Research Proposal").strip()
    blocks.append({"kind": "title", "text": f"Proposal — {title}"})
    blocks.append({"kind": "para", "text": "Market Research Proposal"})
    if brief.get("client_name"):
        blocks.append({"kind": "para", "text": f"Prepared for: {brief['client_name']}"})
    if brief.get("client_contact"):
        blocks.append({"kind": "para", "text": f"Contact: {brief['client_contact']}"})
    blocks.append({"kind": "para", "text": "Prepared by: [Your name]"})
    # 1. Goal
    blocks.append({"kind": "h1", "text": "1. Goal"})
    if brief.get("objectives"):
        blocks.append({"kind": "para", "text": brief["objectives"]})
    if brief.get("background"):
        blocks.append({"kind": "para", "text": brief["background"]})
    if not brief.get("objectives") and not brief.get("background"):
        blocks.append({"kind": "para", "text": "[Add the goal — fill in the Client Brief objectives.]"})
    # 2. Research Design
    blocks.append({"kind": "h1", "text": "2. Research Design"})
    methods = design.get("methods") or []
    approach = design.get("approach") or ""
    sessions = design.get("sessions_count") or 0
    per = design.get("participants_per_session") or 0
    sample = design.get("sample_size") or 0
    rows = [["Item", "Detail"]]
    rows.append(["Approach", approach or "—"])
    rows.append(["Method", ", ".join(methods) if methods else "—"])
    rows.append(["Sessions", str(sessions) if sessions else "—"])
    rows.append(["Respondents per session", str(per) if per else "—"])
    if sessions and per:
        rows.append(["Total respondents", str(sessions * per)])
    elif sample:
        rows.append(["Total respondents", str(sample)])
    else:
        rows.append(["Total respondents", "—"])
    minutes = _proposal_session_minutes(methods)
    if minutes:
        rows.append(["Session length", ", ".join(f"{m}: {n} min" for m, n in minutes)])
    else:
        rows.append(["Session length", "—"])
    blocks.append({"kind": "table", "rows": rows})
    if design.get("rationale"):
        blocks.append({"kind": "para", "text": design["rationale"]})
    # 3. Target
    blocks.append({"kind": "h1", "text": "3. Target"})
    if brief.get("target_audience"):
        blocks.append({"kind": "para", "text": brief["target_audience"]})
    personas = design.get("personas") or []
    if personas:
        prows = [["Persona", "Description", "Value framing"]]
        for p in personas:
            prows.append([str((p or {}).get("name") or "—"),
                          str((p or {}).get("description") or "—"),
                          str((p or {}).get("value_framing") or "—")])
        blocks.append({"kind": "table", "rows": prows})
    quotas = design.get("quotas") or []
    if quotas:
        qrows = [["Segment", "Share", "Notes"]]
        for q in quotas:
            qrows.append([str((q or {}).get("segment") or "—"),
                          str((q or {}).get("share") or "—"),
                          str((q or {}).get("notes") or "—")])
        blocks.append({"kind": "table", "rows": qrows})
    if design.get("screening_criteria"):
        blocks.append({"kind": "para", "text": f"Screening: {design['screening_criteria']}"})
    if not brief.get("target_audience") and not personas and not quotas and not design.get("screening_criteria"):
        blocks.append({"kind": "para", "text": "[Add the target — fill in the Research Design sampling.]"})
    # 4. Rationale
    blocks.append({"kind": "h1", "text": "4. Rationale"})
    blocks.append({"kind": "para", "text": (design.get("rationale") or "").strip()
                   or "[Add rationale — write why this method was chosen.]"})
    # 5. Timeline — brief dates + fieldwork plan session dates ONLY.
    blocks.append({"kind": "h1", "text": "5. Timeline"})
    if brief.get("timeline_start") or brief.get("timeline_end"):
        blocks.append({"kind": "para", "text": f"Project timeline: {brief.get('timeline_start') or '—'} to {brief.get('timeline_end') or '—'}"})
    plan = design.get("fieldwork_plan") or []
    dated = [s for s in plan if str((s or {}).get("date") or "").strip()]
    if dated:
        trows = [["Session", "Group", "Date"]]
        for s in dated:
            trows.append([str(s.get("session") or "—"), str(s.get("group") or "—"),
                          str(s.get("date") or "—")])
        blocks.append({"kind": "table", "rows": trows})
    if not brief.get("timeline_start") and not brief.get("timeline_end") and not dated:
        blocks.append({"kind": "para", "text": "Timeline to be confirmed — add dates in the Brief and Fieldwork Plan."})
    # 6. Deliverables
    blocks.append({"kind": "h1", "text": "6. Deliverables"})
    deliverables = brief.get("deliverables") or []
    if deliverables:
        for d in deliverables:
            blocks.append({"kind": "para", "text": f"- {d}"})
    else:
        blocks.append({"kind": "para", "text": "To be confirmed."})
    # 7. Costs — EMPTY on purpose. No fees, tax, or terms are ever prefilled.
    blocks.append({"kind": "h1", "text": "7. Costs"})
    crows = [["Item", "Amount (MMK)"]]
    for _ in range(PROPOSAL_COST_ROWS):
        crows.append(["", ""])
    crows.append(["Total (excl. tax)", ""])
    blocks.append({"kind": "table", "rows": crows})
    blocks.append({"kind": "para", "text": "Tax: [fill in]"})
    blocks.append({"kind": "para", "text": "Fill in before sending — no fees are prefilled."})
    # 8. Terms — blank for him to fill.
    blocks.append({"kind": "h1", "text": "8. Terms"})
    blocks.append({"kind": "para", "text": "Payment terms: [fill in]"})
    blocks.append({"kind": "para", "text": "Valid until: [fill in]"})
    return blocks, []


def proposal_to_docx(blocks):
    """Render proposal blocks to DOCX bytes (python-docx). Tables get the
    'Table Grid' style with a bold header row; Costs rows stay blank."""
    doc = Document()
    for b in blocks:
        kind = b.get("kind")
        if kind == "title":
            doc.add_heading(b["text"], level=0)
        elif kind == "h1":
            doc.add_heading(b["text"], level=1)
        elif kind == "para":
            doc.add_paragraph(b["text"])
        elif kind == "table":
            rows = b.get("rows") or []
            if not rows:
                continue
            table = doc.add_table(rows=1, cols=len(rows[0]))
            table.style = "Table Grid"
            for ri, row in enumerate(rows):
                cells = table.rows[0].cells if ri == 0 else table.add_row().cells
                for ci, val in enumerate(row):
                    cells[ci].text = str(val or "")
                    if ri == 0:
                        for p in cells[ci].paragraphs:
                            for run in p.runs:
                                run.bold = True
    import io
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# --- Discussion Guide (v3.25, Stage 2) ---
# The DG lives as plain-text markup (what the AI draft produces and what the
# editor shows) and is rendered to Word by dg_to_docx(). The markup mirrors
# the house FGD guide format:
#   "# DISCUSSION GUIDE FOR <Project> (<Client>)" — title line
#   "Qualitative Market Research (FGD)" — method line (or (IDI))
#   "## Welcome" — welcome script heading
#   "## Section N: <Name> (<X> Mins)" — section heading row (amber table)
#   "Objective: <one sentence>" — bold objective line under the heading
#   "### N. <Group name>" — bold numbered question-group heading
#   "> <text>" — moderator instruction (bold+italic), or "(Note: ...)"
#     moderator note (italic only)
#   "- <question>" — bulleted main question
#   "  - <probe>" — nested probe sub-bullet
#   "[TABLE]" ... "[/TABLE]" — pipe-separated table (header row + rows)
#   closing lines as plain paragraphs, ending with
#   "Please collect your [Incentive/Gift] at door." + "Thank you."
DG_SECTION_FILL = "FFC000"  # amber section-heading fill (house format)


# --- Stage 2 self-checks (v3.31, P6–P8) ---
# Deterministic programmatic checks on AI drafts and saved instruments.
# No Gemini is used here. Every result is a dict with id / severity /
# status / description / evidence / recommendation. Results are warnings
# and recommendations only — they NEVER block saving or exporting. The
# researcher is the final judge. Results are computed fresh on every
# render from the current text; nothing is stored, so a verdict can never
# go stale.
_DG_SECTION_TIMING_RE = re.compile(r"^##\s*Section\s+(\d+)\s*:.*?\((\d+)\s*mins?\)\s*$", re.IGNORECASE)
_DG_SECTION_HEAD_RE = re.compile(r"^##\s*Section\s+(\d+)\s*:", re.IGNORECASE)
_DG_DATE_RE = re.compile(r"\b\d{4}-\d{2}-\d{2}\b")
_DG_CURRENCY_RE = re.compile(r"(?i)(?:\$\s*\d|\b\d[\d,]*\s*(?:mmk|kyats?|usd|dollars?)\b)")
_DG_RECORDING_RE = re.compile(r"(?i)\b(recorded|recording|audio|video)\b")
_Q_DURATION_RE = re.compile(r"\b(\d+)\s*[- ]?\s*(months?|years?|weeks?|days?)\b", re.IGNORECASE)
_Q_GROUP_RE = re.compile(r"(?i)(?:\bgroup\s*\d+\b|→\s*group|->\s*group)")

_STOPWORDS = {
    "what", "which", "when", "where", "who", "whom", "whose", "why", "how",
    "is", "are", "was", "were", "be", "been", "being", "do", "does", "did",
    "have", "has", "had", "you", "your", "yours", "we", "our", "ours",
    "they", "their", "theirs", "he", "she", "his", "her", "hers", "it", "its",
    "the", "and", "for", "with", "from", "about", "this", "that", "these",
    "those", "there", "here", "will", "would", "can", "could", "should",
    "please", "per", "each", "every", "any", "some", "all", "more", "most",
    "other", "another", "often", "usually", "currently", "still", "ever",
    "never", "before", "after", "during", "over", "under", "into", "onto",
    "upon", "between", "among", "within", "without", "through", "across",
    "along", "toward", "towards", "against", "whether", "either", "neither",
    "both", "just", "only", "also", "very", "really", "much", "many", "few",
    "less", "least", "first", "last", "next", "own", "same", "different",
    "not", "no", "yes", "one", "two", "three", "four", "five", "six",
    "seven", "eight", "nine", "ten", "a", "an", "by", "of", "to", "in",
    "on", "at", "as", "or", "if", "but", "so", "than", "then", "once",
    "twice", "via", "understand", "explore", "identify", "assess", "examine",
    "determine", "learn", "research", "study", "project", "employee",
    "employees", "respondent", "respondents", "client",
}
_GENERIC_FIRST_WORDS = {"use", "like", "want", "need", "think", "feel",
                        "work", "buy", "get", "take", "make", "give"}


def _self_check_result(check_id, severity, status, description, evidence, recommendation):
    """One self-check result. severity/status are 'fail' or 'warning' for
    issues; only issues are returned (an empty list means all checks
    passed)."""
    return {"id": check_id, "severity": severity, "status": status,
            "description": description, "evidence": evidence,
            "recommendation": recommendation}


def _collect_reference_text(brief, design):
    """All brief + design free text, lowercased — the only place an AI draft
    may take facts from. A value found here is approved; anything else the
    draft states is possibly invented."""
    parts = []
    for src in (brief or {}, design or {}):
        if not isinstance(src, dict):
            continue
        for key in ("client_name", "project_code", "background", "objectives",
                    "target_audience", "scope", "discussion_guide",
                    "timeline_start", "timeline_end", "goal", "target",
                    "approach", "rationale", "screening_criteria",
                    "recruitment_criteria", "deliverables"):
            v = src.get(key)
            if isinstance(v, str) and v.strip():
                parts.append(v)
            elif isinstance(v, (list, tuple)):
                parts.extend(str(x) for x in v if str(x).strip())
        for coll in ("personas", "quotas"):
            for item in src.get(coll) or []:
                if isinstance(item, dict):
                    parts.extend(str(x) for x in item.values() if str(x).strip())
    return " ".join(parts).lower()


def _content_words(text, min_len=3):
    """Lowercase content words (English) — function words removed. Used for
    the weak keyword-overlap checks only; a match is a signal, never proof
    of semantic correctness."""
    words = re.findall(r"[a-z0-9]+", str(text or "").lower())
    return [w for w in words if len(w) >= min_len and w not in _STOPWORDS]


def _objective_items(objectives):
    """Brief objectives as a list of strings (string with newlines or a
    list both work); leading numbering stripped."""
    if isinstance(objectives, str):
        items = [x.strip() for x in re.split(r"[\n;]+", objectives) if x.strip()]
    elif isinstance(objectives, (list, tuple)):
        items = [str(x).strip() for x in objectives if str(x).strip()]
    else:
        items = []
    out = []
    for it in items:
        it = re.sub(r"^\s*(?:\d+[.)]|[-•])\s*", "", it)
        if it:
            out.append(it)
    return out


def _parse_ymd(value):
    """YYYY-MM-DD text → datetime.date, or None when empty/unparseable."""
    try:
        return datetime.date.fromisoformat(str(value or "").strip()[:10])
    except Exception:
        return None


def _show_self_check_results(results, label_en, label_mm):
    """Render self-check results under an editor (warnings only — never
    blocks saving). Mirrors the pre-export checks pattern: issues as
    warnings/errors, a success line when everything passed."""
    if not results:
        st.success(T(f"{label_en}: all checks passed.", f"{label_mm}: စစ်ဆေးမှု အားလုံး အောင်တယ်။"))
        return
    _fails = [r for r in results if r.get("severity") == "fail"]
    _warns = [r for r in results if r.get("severity") == "warning"]
    st.caption(T(f"{label_en}: {len(_fails)} failed · {len(_warns)} warning(s) — checks are recommendations only; saving is never blocked.",
                 f"{label_mm}: {len(_fails)} ခု မအောင်ဘူး · {len(_warns)} ခု သတိပေးချက် — စစ်ဆေးမှုက အကြံပြုချက်ပဲ၊ သိမ်းတာကို ဘယ်တော့မှ မတားဘူး။"))
    for r in results:
        _msg = f"{r.get('id', '')}: {r.get('description', '')} Evidence: {r.get('evidence', '')} Fix: {r.get('recommendation', '')}"
        if r.get("severity") == "fail":
            st.error(_msg)
        else:
            st.warning(_msg)


def _dg_self_check(dg_text, brief, design, session_minutes):
    """Deterministic checks on a Discussion Guide draft / saved guide.
    Returns a list of issue dicts (fail/warning); an empty list means all
    checks passed. Never uses Gemini."""
    results = []
    text = str(dg_text or "")
    low = text.lower()
    ref = _collect_reference_text(brief, design)
    # (a) section timings sum to the session length; (b) sequential numbers
    timed = []
    untimed = []
    for line in text.splitlines():
        line = line.strip()
        m = _DG_SECTION_TIMING_RE.match(line)
        if m:
            timed.append((int(m.group(1)), int(m.group(2))))
            continue
        m2 = _DG_SECTION_HEAD_RE.match(line)
        if m2:
            untimed.append(int(m2.group(1)))
    if not timed and not untimed:
        results.append(_self_check_result("DG-SECTIONS", "warning", "warning",
            "No numbered sections found — the guide has no '## Section N: ... (X Mins)' headings, so timings cannot be verified.",
            "No '## Section' headings in the guide.",
            "Write the guide in numbered sections with a timing in each heading, e.g. '## Section 1: Welcome (10 Mins)'."))
    else:
        all_nums = sorted([n for n, _ in timed] + untimed)
        if all_nums != list(range(1, len(all_nums) + 1)):
            results.append(_self_check_result("DG-SECTIONS", "fail", "fail",
                "Section numbers are not sequential.",
                f"Section numbers found: {all_nums}.",
                "Renumber the sections 1, 2, 3 … with no gaps or repeats."))
        if untimed:
            results.append(_self_check_result("DG-TIMING", "warning", "warning",
                "Some sections have no timing in the heading.",
                f"Sections without a timing: {sorted(untimed)}.",
                "Add a timing to every section heading, e.g. '(15 Mins)'."))
        if timed:
            total = sum(m for _, m in timed)
            if session_minutes and total != session_minutes:
                results.append(_self_check_result("DG-TIMING", "fail", "fail",
                    f"Section timings sum to {total} minutes, not the session length of {session_minutes} minutes.",
                    f"Timed sections: {timed}; sum = {total}.",
                    "Adjust the section timings so they sum exactly to the session length."))
    # (c) objective coverage — weak keyword overlap, warning only. A match
    # is a signal, not proof the objective is really covered.
    for obj in _objective_items((brief or {}).get("objectives")):
        keys = _content_words(obj, min_len=5)
        if not keys:
            continue
        if not any(k in low for k in keys):
            results.append(_self_check_result("DG-OBJECTIVES", "warning", "warning",
                "A brief objective is not reflected in the guide (keyword check — verify by reading).",
                f"Objective: {obj[:120]}",
                "Add a section or probes that cover this objective, or check that the wording differs."))
    # (d) invented dates / currency — a value the brief/design never states
    for d in sorted(set(_DG_DATE_RE.findall(text))):
        if d.lower() not in ref:
            results.append(_self_check_result("DG-INVENTED-DATE", "fail", "fail",
                "A date appears in the guide that the brief/design does not state.",
                f"Date found: {d}.",
                "Remove the date or replace it with a [Date] placeholder."))
    for cur in sorted(set(m.group(0) for m in _DG_CURRENCY_RE.finditer(text))):
        if cur.lower() not in ref:
            results.append(_self_check_result("DG-INVENTED-CURRENCY", "fail", "fail",
                "An amount appears in the guide that the brief/design does not state.",
                f"Amount found: {cur}.",
                "Remove the amount or replace it with a placeholder."))
    if _DG_RECORDING_RE.search(text):
        if not _DG_RECORDING_RE.search(ref) and "[Recording — confirm]" not in text:
            results.append(_self_check_result("DG-RECORDING", "warning", "warning",
                "The guide mentions recording, but the brief/design does not state recording and there is no [Recording — confirm] placeholder.",
                f"First match: {_DG_RECORDING_RE.search(text).group(0)}.",
                "Confirm whether the session is recorded; if unknown, use the [Recording — confirm] placeholder."))
    # (e) same-line EN+MM mixing. House placeholders [..] are English by
    # design, so they are stripped before counting; headings are structural.
    for i, line in enumerate(text.splitlines(), start=1):
        stripped = re.sub(r"\[[^\]]*\]", "", line)
        if stripped.strip().startswith("#"):
            continue
        mm = len(re.findall(r"[\u1000-\u109F]", stripped))
        lat = len(re.findall(r"[A-Za-z]", stripped))
        if mm >= 5 and lat >= 5:
            results.append(_self_check_result("DG-LANGUAGE", "warning", "warning",
                "One line mixes English and Myanmar — keep one language per line.",
                f"Line {i}: {line.strip()[:120]}",
                "Split the line or rewrite it in one language."))
    # (f) venue / date placeholders (v3.30 live re-test: the guide had none)
    if "[Venue]" not in text and "venue" not in ref:
        results.append(_self_check_result("DG-PLACEHOLDER-VENUE", "warning", "warning",
            "Venue not stated — the guide has no [Venue] placeholder.",
            "No [Venue] in the guide and no venue in the brief/design.",
            "Add a [Venue] placeholder to the guide."))
    if "[Date]" not in text and not ((brief or {}).get("timeline_start") or (brief or {}).get("timeline_end")):
        results.append(_self_check_result("DG-PLACEHOLDER-DATE", "warning", "warning",
            "Date not stated — the guide has no [Date] placeholder.",
            "No [Date] in the guide and no timeline dates in the brief.",
            "Add a [Date] placeholder to the guide."))
    return results


def _dg_parse(dg_text):
    """Parse DG markup into a list of blocks. Unknown lines become plain
    paragraphs — the parser never drops text."""
    blocks = []
    lines = str(dg_text or "").splitlines()
    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.strip()
        i += 1
        if not line:
            continue
        if line == "[TABLE]":
            rows = []
            while i < len(lines):
                tline = lines[i].strip()
                i += 1
                if tline == "[/TABLE]":
                    break
                if tline:
                    rows.append([c.strip() for c in tline.split("|")])
            if rows:
                blocks.append({"kind": "table", "rows": rows})
            continue
        if line.startswith("# "):
            blocks.append({"kind": "title", "text": line[2:].strip()})
        elif line.startswith("## "):
            blocks.append({"kind": "h2", "text": line[3:].strip()})
        elif line.startswith("### "):
            blocks.append({"kind": "h3", "text": line[4:].strip()})
        elif line.lower().startswith("objective:"):
            blocks.append({"kind": "objective", "text": line[len("Objective:"):].strip()})
        elif line.startswith("> "):
            blocks.append({"kind": "instruction", "text": line[2:].strip()})
        elif raw.startswith("  - ") or raw.startswith("\t- "):
            blocks.append({"kind": "bullet", "text": line[2:].strip(), "level": 1})
        elif line.startswith("- "):
            blocks.append({"kind": "bullet", "text": line[2:].strip(), "level": 0})
        else:
            blocks.append({"kind": "para", "text": line})
    return blocks


def _dg_shade_cell(cell, fill):
    """Fill a table cell (amber section heading, as in the sample guide)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def dg_to_docx(dg_text, title_text, out_path):
    """Render DG markup to a Word file in the house format: title, method
    line, Welcome script, amber section-heading table rows with bold
    objectives, bold group headings, bold+italic instructions, italic
    notes, bulleted questions with nested probes, tables, closing."""
    doc = Document()
    blocks = _dg_parse(dg_text)
    seen_title = False
    for b in blocks:
        kind = b["kind"]
        if kind == "title":
            doc.add_heading(b["text"], level=0)
            seen_title = True
        elif kind == "h2":
            text = b["text"]
            if text.lower().startswith("section "):
                table = doc.add_table(rows=1, cols=1)
                table.style = "Table Grid"
                cell = table.rows[0].cells[0]
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(text)
                run.bold = True
                _dg_shade_cell(cell, DG_SECTION_FILL)
            else:
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.bold = True
        elif kind == "objective":
            p = doc.add_paragraph()
            run = p.add_run(b["text"])
            run.bold = True
        elif kind == "h3":
            p = doc.add_paragraph()
            run = p.add_run(b["text"])
            run.bold = True
        elif kind == "instruction":
            p = doc.add_paragraph()
            run = p.add_run(b["text"])
            run.italic = True
            if not b["text"].lstrip().startswith("(Note:"):
                run.bold = True
        elif kind == "bullet":
            style = "List Bullet 2" if b.get("level") == 1 else "List Bullet"
            try:
                doc.add_paragraph(b["text"], style=style)
            except Exception:
                doc.add_paragraph(("  " if b.get("level") == 1 else "") + "• " + b["text"])
        elif kind == "table":
            rows = b["rows"]
            table = doc.add_table(rows=len(rows), cols=max(len(r) for r in rows))
            table.style = "Table Grid"
            for ri, row in enumerate(rows):
                for ci, val in enumerate(row):
                    cell = table.rows[ri].cells[ci]
                    cell.text = ""
                    p = cell.paragraphs[0]
                    run = p.add_run(val)
                    if ri == 0:
                        run.bold = True
        else:  # para
            doc.add_paragraph(b["text"])
    if not seen_title:
        doc.add_heading(title_text, level=0)
        doc.add_paragraph("Market Research AI — ResearchFlow")
    doc.save(out_path)
    return out_path


def _dg_to_markdown(dg_text):
    """Light conversion of DG markup to Markdown for the in-app preview.
    Section headings become ### headings, objectives bold, instructions
    quotes, tables become Markdown tables."""
    out = []
    in_table = False
    for b in _dg_parse(dg_text):
        kind = b["kind"]
        if kind == "title":
            out.append(f"# {b['text']}")
        elif kind == "h2":
            out.append(f"### {b['text']}")
        elif kind == "objective":
            out.append(f"**{b['text']}**")
        elif kind == "h3":
            out.append(f"**{b['text']}**")
        elif kind == "instruction":
            out.append(f"> *{b['text']}*")
        elif kind == "bullet":
            out.append(("  - " if b.get("level") == 1 else "- ") + b["text"])
        elif kind == "table":
            rows = b["rows"]
            if rows:
                out.append("| " + " | ".join(rows[0]) + " |")
                out.append("| " + " | ".join(["---"] * len(rows[0])) + " |")
                for row in rows[1:]:
                    out.append("| " + " | ".join(row) + " |")
            in_table = False
        else:
            out.append(b["text"])
    return "\n\n".join(out)


def _xlsform_slug(text):
    """Slug for XLSForm names: lowercase, non-alphanumeric → underscore."""
    return re.sub(r"[^a-z0-9]+", "_", str(text or "").strip().lower()).strip("_")


def _q_normalize_text(s):
    """Lowercased question text with punctuation/whitespace stripped — for
    duplicate detection."""
    s = str(s or "").lower()
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    return " ".join(s.split())


def _q_content_stem(text):
    """First 3 content words of a question — the 'stem' used for the
    duplicate-topic check. A stem match is a heuristic signal, not proof
    two questions are the same."""
    return _content_words(text, min_len=3)[:3]


def _q_self_check(qdata, brief, design):
    """Deterministic checks on a questionnaire / screener draft or saved
    questionnaire. Returns a list of issue dicts (fail/warning); an empty
    list means all checks passed. Never uses Gemini."""
    results = []
    if isinstance(qdata, list):
        questions = [q for q in qdata if isinstance(q, dict)]
    elif isinstance(qdata, dict):
        questions = [q for q in (qdata.get("questions") or []) if isinstance(q, dict)]
    else:
        questions = []
    methods = (design or {}).get("methods") or []
    screener_mode = "Survey" not in methods
    has_groups = bool((design or {}).get("quotas"))
    ref = _collect_reference_text(brief, design)
    q_lang = str((design or {}).get("questionnaire_language", "Myanmar") or "Myanmar")
    for idx, q in enumerate(questions):
        code = str(q.get("code") or f"Q{idx + 1}").strip() or f"Q{idx + 1}"
        section = str(q.get("section") or "").strip()
        q_en = str(q.get("question_en") or "").strip()
        q_mm = str(q.get("question_mm") or "").strip()
        outcome = str(q.get("outcome") or "").strip()
        qtype = str(q.get("type") or "").strip()
        # (a) discussion question in a screener (Qual/FGD without Survey)
        if screener_mode:
            if not section:
                results.append(_self_check_result("Q-SECTION-MISSING", "warning", "warning",
                    "A question has no section — in a screener every question belongs to the Screening section.",
                    f"{code}: {q_en[:100]}",
                    "Set the section to Screening, or move the question to the Discussion Guide."))
            elif section.lower() != "screening":
                results.append(_self_check_result("Q-SCREENER-SECTION", "fail", "fail",
                    "A discussion question is in the screener — it belongs in the Discussion Guide, not the screener.",
                    f"{code} (section: {section}): {q_en[:100]}",
                    "Move this question to the Discussion Guide."))
        # (b) missing outcome rule
        if not outcome:
            results.append(_self_check_result("Q-OUTCOME-MISSING", "fail", "fail",
                "A question has no outcome rule — every question must say Continue or Terminate.",
                f"{code}: {q_en[:100]}",
                "Add an outcome rule (Continue / Terminate) to this question."))
        # (c) group assignment without approved quotas/groups
        if outcome and _Q_GROUP_RE.search(outcome) and not has_groups:
            results.append(_self_check_result("Q-GROUP-ASSIGNMENT", "fail", "fail",
                "A group assignment appears in an outcome, but the design lists no quotas or groups.",
                f"{code}: outcome '{outcome[:120]}'",
                "Remove the group assignment (outcome Continue / Terminate only) or add the approved groups to the design."))
        # (d) possibly invented duration criterion — a number + time unit
        # (e.g. '6 months') that the brief/design never states
        _dur_text = " ".join([q_en, str(q.get("options_en") or ""), outcome,
                              str(q.get("hint_en") or "")])
        for m in _Q_DURATION_RE.finditer(_dur_text):
            _dur = f"{m.group(1)} {m.group(2).lower()}"
            _dur_sing = f"{m.group(1)} {m.group(2).lower().rstrip('s')}"
            _dur_hyp = f"{m.group(1)}-{m.group(2).lower()}"
            if _dur not in ref and _dur_sing not in ref and _dur_hyp not in ref:
                results.append(_self_check_result("Q-INVENTED-CRITERION", "warning", "warning",
                    "A time window in this question is not stated in the brief/design — it may be invented.",
                    f"{code}: '{m.group(0)}' in '{_dur_text[:120]}'",
                    "Remove the time window, or confirm it with the brief and add it to the design's screening criteria."))
                break
        # (e) sensitive demographic question terminating respondents
        _sensitive = (qtype == "observed") or ("gender" in q_en.lower()) or ("gender" in q_mm.lower())
        if _sensitive and "terminate" in outcome.lower():
            results.append(_self_check_result("Q-SENSITIVE-TERMINATE", "fail", "fail",
                "A sensitive demographic question disqualifies respondents.",
                f"{code} (type: {qtype or '—'}): outcome '{outcome[:100]}'",
                "Record sensitive demographics (gender, observed) — never Terminate on them unless the brief says so."))
        # (f) Myanmar text (Myanmar language only — in English mode the
        # Myanmar text is expected to be the same as the English text)
        if q_lang == "Myanmar" and (q_en or q_mm):
            if not q_mm:
                results.append(_self_check_result("Q-LANGUAGE-MM", "warning", "warning",
                    "A question has no Myanmar text.",
                    f"{code}: {q_en[:100]}",
                    "Add the Myanmar translation of this question."))
            elif q_mm == q_en:
                results.append(_self_check_result("Q-LANGUAGE-MM", "warning", "warning",
                    "The Myanmar text is identical to the English text.",
                    f"{code}: {q_en[:100]}",
                    "Translate this question into Myanmar."))
            elif not _NON_EN_RE.search(q_mm):
                results.append(_self_check_result("Q-LANGUAGE-MM", "warning", "warning",
                    "The Myanmar text contains no Myanmar script.",
                    f"{code}: {q_mm[:100]}",
                    "Translate this question into Myanmar."))
    # (g) duplicate questions — identical normalized text, or the same
    # first content word with highly overlapping stems (e.g. numeric age +
    # age bracket). Warning only: stems are a heuristic.
    _norms = []
    for idx, q in enumerate(questions):
        code = str(q.get("code") or f"Q{idx + 1}").strip() or f"Q{idx + 1}"
        _t = _q_normalize_text(q.get("question_en") or q.get("question_mm"))
        _norms.append((code, _t, _q_content_stem(q.get("question_en") or q.get("question_mm"))))
    for i in range(len(_norms)):
        for j in range(i + 1, len(_norms)):
            _ci, _ti, _si = _norms[i]
            _cj, _tj, _sj = _norms[j]
            if _ti and _ti == _tj:
                results.append(_self_check_result("Q-DUPLICATE", "warning", "warning",
                    "Two questions have identical text.",
                    f"{_ci} and {_cj}: '{_ti[:100]}'",
                    "Keep one question and remove the duplicate."))
            elif (_si and _sj and _si[0] == _sj[0]
                    and _si[0] not in _GENERIC_FIRST_WORDS
                    and len(set(_si) | set(_sj)) <= 4):
                results.append(_self_check_result("Q-DUPLICATE", "warning", "warning",
                    "Two questions may ask the same thing (same topic stem).",
                    f"{_ci} and {_cj} both ask about '{_si[0]}'.",
                    "Check whether both are needed; merge them if they ask the same thing."))
    return results


def _questionnaire_empty_save_guard(new_questions, pending_draft, saved_questionnaire):
    """Mirror the DG empty-save guard (v3.29): never destroy a draft or a
    saved questionnaire with an empty save. Returns (allow_save, en_text,
    mm_text): allow_save False means warn and keep everything."""
    if new_questions:
        return True, None, None
    _pending_qs = []
    if isinstance(pending_draft, dict):
        _pending_qs = pending_draft.get("questions") or []
    if _pending_qs:
        return False, ("The grid is empty but a new AI draft exists — it was NOT saved. "
                       "Click “✨ AI Draft Questionnaire” again, or put the questions back in the grid, then Save."), (
                       "Grid ဗလာဖြစ်နေတယ်၊ ဒါပေမယ့် AI draft အသစ် ရှိတယ် — မသိမ်းလိုက်ဘူး။ "
                       "“✨ AI Draft Questionnaire” ကို ပြန်နှိပ်ပါ၊ ဒါမှမဟုတ် မေးခွန်းတွေ grid ထဲ ပြန်ထည့်ပြီး Save နှိပ်ပါ။")
    _saved_qs = []
    if isinstance(saved_questionnaire, dict):
        _saved_qs = saved_questionnaire.get("questions") or []
    elif isinstance(saved_questionnaire, list):
        _saved_qs = saved_questionnaire
    if _saved_qs:
        return False, "The grid is empty — the saved questionnaire was kept, not erased.", \
            "Grid ဗလာဖြစ်နေတယ် — သိမ်းပြီးသား questionnaire ကို ဖျက်မပစ်ဘူး၊ ထားထားတယ်။"
    return False, "Nothing to save — the questionnaire is empty.", \
        "သိမ်းစရာ မရှိဘူး — questionnaire ဗလာဖြစ်နေတယ်။"


def _fieldwork_self_check(design, brief):
    """Deterministic checks on a saved Fieldwork Plan against Sampling and
    the Brief timeline. Returns a list of issue dicts (fail/warning); an
    empty list means all checks passed. Never uses Gemini."""
    results = []
    plan = [r for r in ((design or {}).get("fieldwork_plan") or []) if isinstance(r, dict)]
    sessions = 0
    try:
        sessions = int((design or {}).get("sessions_count") or 0)
    except Exception:
        sessions = 0
    methods = (design or {}).get("methods") or []
    # (a) planned rows vs approved session count
    if sessions > 0 and len(plan) != sessions:
        results.append(_self_check_result("FW-SESSION-COUNT", "warning", "warning",
            "The plan has a different number of sessions than Sampling approved.",
            f"Plan rows: {len(plan)}; sessions in Sampling: {sessions}.",
            "Add or remove plan rows so the count matches Sampling."))
    _start = _parse_ymd((brief or {}).get("timeline_start"))
    _end = _parse_ymd((brief or {}).get("timeline_end"))
    for row in plan:
        _sess = row.get("session")
        _label = f"Session {_sess}" if _sess not in (None, "") else "A plan row"
        # (b) method matches the approved methods
        _m = str(row.get("method") or "").strip()
        if _m and methods and _m not in methods:
            results.append(_self_check_result("FW-METHOD", "warning", "warning",
                "A plan row uses a method the design does not list.",
                f"{_label}: method '{_m}'; approved methods: {methods}.",
                "Change the row method to an approved one, or update the design methods."))
        # (c) date inside the brief timeline
        _d = _parse_ymd(row.get("date"))
        if _d:
            if _start and _d < _start:
                results.append(_self_check_result("FW-DATE-RANGE", "warning", "warning",
                    "A plan date is before the project timeline starts.",
                    f"{_label}: date {row.get('date')}; timeline starts {(brief or {}).get('timeline_start')}.",
                    "Move the session inside the timeline, or fix the timeline in the Brief."))
            elif _end and _d > _end:
                results.append(_self_check_result("FW-DATE-RANGE", "warning", "warning",
                    "A plan date is after the project timeline ends.",
                    f"{_label}: date {row.get('date')}; timeline ends {(brief or {}).get('timeline_end')}.",
                    "Move the session inside the timeline, or fix the timeline in the Brief."))
        # (d) moderator set without a seat (the save handler defaults new
        # rows to Seat 1; this catches stored rows from before that rule)
        if str(row.get("moderator") or "").strip() and not str(row.get("moderator_seat") or "").strip():
            results.append(_self_check_result("FW-MODERATOR-SEAT", "warning", "warning",
                "A moderator is set but no seat is given.",
                f"{_label}: moderator '{str(row.get('moderator')).strip()}'.",
                "Set the moderator seat to Seat 1."))
    return results


def _parse_questionnaire_json(text):
    """Parse the AI questionnaire draft (JSON object) into the stored
    structure. Returns (data, error): data is a dict with intro_en/intro_mm/
    closing_en/closing_mm/header_fields/questions, error is "" on success.
    Tolerates a bare list (old format) and JSON wrapped in other text."""
    raw = str(text or "").strip()
    data = None
    try:
        data = json.loads(raw)
    except Exception:
        start = raw.find("{")
        end = raw.rfind("}")
        if start >= 0 and end > start:
            try:
                data = json.loads(raw[start:end + 1])
            except Exception:
                data = None
    if data is None:
        start = raw.find("[")
        end = raw.rfind("]")
        if start >= 0 and end > start:
            try:
                data = {"questions": json.loads(raw[start:end + 1])}
            except Exception:
                data = None
    if not isinstance(data, dict):
        return None, "Could not read the questionnaire JSON — the draft is shown below; fix it and save."
    questions = data.get("questions")
    if not isinstance(questions, list):
        return None, "The questionnaire has no questions list — the draft is shown below; fix it and save."
    clean = []
    for q in questions:
        if not isinstance(q, dict):
            continue
        qtext_en = str(q.get("question_en", q.get("question", "")) or "").strip()
        qtext_mm = str(q.get("question_mm", "") or "").strip()
        if not qtext_en and not qtext_mm:
            continue
        qtype = str(q.get("type", "open_text") or "open_text").strip()
        if qtype not in QUESTION_TYPES:
            qtype = "open_text"
        _req = q.get("required", False)
        if isinstance(_req, str):
            _req = _req.strip().lower() in {"true", "yes", "1"}
        clean.append({
            "code": str(q.get("code", "") or "").strip(),
            "name": _xlsform_slug(q.get("name", "") or q.get("code", "")),
            "section": str(q.get("section", "") or "").strip(),
            "question_en": qtext_en,
            "question_mm": qtext_mm,
            "type": qtype,
            "options_en": str(q.get("options_en", q.get("options", "")) or "").strip(),
            "options_mm": str(q.get("options_mm", "") or "").strip(),
            "outcome": str(q.get("outcome", "") or "").strip(),
            "hint_en": str(q.get("hint_en", "") or "").strip(),
            "hint_mm": str(q.get("hint_mm", "") or "").strip(),
            "required": bool(_req),
            "appearance": str(q.get("appearance", "") or "").strip(),
            "relevant": str(q.get("relevant", "") or "").strip(),
            "constraint": str(q.get("constraint", "") or "").strip(),
            "constraint_message": str(q.get("constraint_message", "") or "").strip(),
        })
    header = data.get("header_fields")
    if not isinstance(header, list):
        header = []
    header = [str(h).strip() for h in header if str(h).strip()]
    return {
        "intro_en": str(data.get("intro_en", "") or ""),
        "intro_mm": str(data.get("intro_mm", "") or ""),
        "closing_en": str(data.get("closing_en", "") or ""),
        "closing_mm": str(data.get("closing_mm", "") or ""),
        "header_fields": header,
        "questions": clean,
    }, ""


def _ed_bool(v):
    """data_editor checkbox cell → bool (NaN/None become False — note that
    bool(float('nan')) is True, so NaN must be caught first)."""
    try:
        if pd.isna(v):
            return False
    except Exception:
        pass
    return bool(v)


def _ed_str(v):
    """data_editor cell → clean string (NaN/None become '')."""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip() if v is not None else ""


def _ed_int(v, default=0):
    """data_editor cell → int (NaN/None/'' become default)."""
    try:
        if pd.isna(v):
            return default
    except Exception:
        pass
    try:
        return int(float(v))
    except Exception:
        return default


# --- Kobo XLSForm export (v3.25, Stage 2) ---
# The questionnaire maps to a KoboToolbox XLSForm workbook (survey +
# choices + settings) in the format of the CLT sample form: survey columns
# type / name / label::English (en) / label::Myanmar (my) / hint::English
# (en) / hint::Myanmar (my) / required / appearance / relevant /
# repeat_count / constraint / constraint_message; choices columns
# list_name / name / label::English (en) / label::Myanmar (my) / order;
# settings id_string / version / default language. Sections become
# begin_group/end_group (field-list); a section named "Repeat: ..." becomes
# begin_repeat/end_repeat with a blank repeat_count (set the count in Kobo).
XLSFORM_SURVEY_HEADER = ["type", "name", "label::English (en)", "label::Myanmar (my)",
                         "hint::English (en)", "hint::Myanmar (my)", "required", "appearance",
                         "relevant", "repeat_count", "constraint", "constraint_message"]
XLSFORM_CHOICES_HEADER = ["list_name", "name", "label::English (en)", "label::Myanmar (my)", "order"]


def _design_prompt_context(brief, design):
    """Shared brief + design context block for the Stage 2 AI drafts.
    Only non-empty fields are included — nothing is invented."""
    b = []
    for key, label in [("client_name", "Client"), ("project_code", "Project code"),
                       ("background", "Background"), ("objectives", "Objectives"),
                       ("target_audience", "Target audience"), ("scope", "Scope"),
                       ("discussion_guide", "Discussion Guide outline from the brief")]:
        v = str(brief.get(key, "") or "").strip()
        if v:
            b.append(f"{label}: {v}")
    _tl = [str(brief.get("timeline_start", "") or "").strip(),
           str(brief.get("timeline_end", "") or "").strip()]
    _tl = [x for x in _tl if x]
    if _tl:
        b.append("Timeline: " + " – ".join(_tl))
    if brief.get("deliverables"):
        b.append("Deliverables: " + ", ".join(brief["deliverables"]))
    d = []
    if str(design.get("goal", "") or "").strip():
        d.append("Research goal: " + str(design["goal"]).strip())
    if str(design.get("target", "") or "").strip():
        d.append("Target: " + str(design["target"]).strip())
    if design.get("approach"):
        d.append("Approach: " + str(design["approach"]))
    if design.get("methods"):
        d.append("Methods: " + ", ".join(design["methods"]))
    if str(design.get("rationale", "") or "").strip():
        d.append("Why this approach: " + str(design["rationale"]).strip())
    if (design.get("sample_size") or 0) > 0:
        d.append(f"Sample size: {int(design['sample_size'])}")
    if (design.get("sessions_count") or 0) > 0:
        d.append(f"Sessions: {int(design['sessions_count'])}")
    if (design.get("participants_per_session") or 0) > 0:
        d.append(f"Participants per session: {int(design['participants_per_session'])}")
    for p in design.get("personas") or []:
        if isinstance(p, dict) and str(p.get("name", "") or "").strip():
            _pd = f"Persona: {p['name'].strip()}"
            if str(p.get("description", "") or "").strip():
                _pd += f" — {p['description'].strip()}"
            if str(p.get("value_framing", "") or "").strip():
                _pd += f" (value framing: {p['value_framing'].strip()})"
            d.append(_pd)
    for qq in design.get("quotas") or []:
        if isinstance(qq, dict) and str(qq.get("segment", "") or "").strip():
            _qd = f"Quota: {qq['segment'].strip()}"
            if str(qq.get("share", "") or "").strip():
                _qd += f" — {qq['share'].strip()}"
            if str(qq.get("notes", "") or "").strip():
                _qd += f" ({qq['notes'].strip()})"
            d.append(_qd)
    if str(design.get("screening_criteria", "") or "").strip():
        d.append("Screening criteria: " + str(design["screening_criteria"]).strip())
    if str(design.get("recruitment_criteria", "") or "").strip():
        d.append("Recruitment notes: " + str(design["recruitment_criteria"]).strip())
    out = ["Client brief:\n" + ("\n".join(b) if b else "(no brief saved yet)"),
           "Research design:\n" + ("\n".join(d) if d else "(no design saved yet)")]
    return "\n\n".join(out)


def questionnaire_to_xlsform(qdata, title_text, out_path):
    """Export the saved questionnaire as a Kobo XLSForm workbook.
    Returns (out_path, stats) or (None, error) when openpyxl is missing.
    Type mapping: single_select → select_one; multi_select → select_multiple;
    ranking_top3 → select_multiple + constraint count-selected(.) = 3;
    likert_1_5 / rating → select_one + appearance likert (choices 1-5 unless
    options given); open_text → text; numeric → integer; date → date;
    note → note; observed → text (or select_one when options exist — the
    recruiter records what they see); contact → two text rows (<name>_name,
    <name>_phone). The screener's Terminate outcome has no XLSForm
    equivalent — it stays in the question's outcome text and in the relevant
    expression when one is given; no skip logic is invented. Header fields
    export as date (FGD date), select_one Yes/No (Confirmation), else text.
    Intro/closing export as note rows."""
    try:
        import openpyxl
    except ImportError:
        return None, "openpyxl is not installed — run this in PowerShell: .\\.venv\\Scripts\\python.exe -m pip install openpyxl"
    if isinstance(qdata, list):
        qdata = {"questions": qdata}
    if not isinstance(qdata, dict):
        qdata = {}
    questions = qdata.get("questions") or []
    if not isinstance(questions, list):
        questions = []
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "survey"
    ws.append(XLSFORM_SURVEY_HEADER)
    ch = wb.create_sheet("choices")
    ch.append(XLSFORM_CHOICES_HEADER)
    st_ = wb.create_sheet("settings")
    st_.append(["form_title", "id_string", "version", "default_language"])
    st_.append([title_text or "ResearchFlow Questionnaire",
                _xlsform_slug(title_text) or "researchflow", "1", "English (en)"])
    used_names = set()
    choice_rows = 0
    notes = []

    def _unique_name(base, fallback):
        name = _xlsform_slug(base) or fallback
        if name not in used_names:
            used_names.add(name)
            return name
        i = 2
        while f"{name}_{i}" in used_names:
            i += 1
        used_names.add(f"{name}_{i}")
        return f"{name}_{i}"

    def _add_choices(list_name, opts_en, opts_mm):
        nonlocal choice_rows
        for i, en in enumerate(opts_en):
            mm = opts_mm[i] if i < len(opts_mm) else en
            ch.append([list_name, str(i + 1), en, mm, ""])
            choice_rows += 1

    def _row(qtype, name, en="", mm="", hint_en="", hint_mm="", required="false",
             appearance="", relevant="", repeat_count="", constraint="", constraint_message=""):
        ws.append([qtype, name, en, mm, hint_en, hint_mm, required, appearance,
                   relevant, repeat_count, constraint, constraint_message])

    _row("start", "start")
    _row("end", "end")
    _row("start-geopoint", "start-geopoint")
    # Header / operational fields.
    for field in qdata.get("header_fields") or []:
        slug = _xlsform_slug(field)
        if not slug:
            continue
        if "date" in slug:
            _row("date", _unique_name(slug, "field"), field, field, required="false")
        elif "confirm" in slug:
            lname = _unique_name(slug, "field")
            _row(f"select_one {lname}", lname, field, field, required="false")
            _add_choices(lname, ["Yes", "No"], ["ဟုတ်ကဲ့", "မဟုတ်ဘူး"])
        else:
            _row("text", _unique_name(slug, "field"), field, field, required="false")
    # Intro script as a note row.
    if (qdata.get("intro_en") or "").strip() or (qdata.get("intro_mm") or "").strip():
        _row("note", _unique_name("intro", "intro"),
             qdata.get("intro_en", ""), qdata.get("intro_mm", ""), required="false")
    # Questions, grouped by section in first-appearance order.
    sections = []
    for q in questions:
        if not isinstance(q, dict):
            continue
        sec = str(q.get("section", "") or "").strip()
        if sec not in sections:
            sections.append(sec)
    has_terminate = False
    has_repeat = False
    has_contact = False
    has_observed = False
    for sec in sections:
        sec_questions = [q for q in questions if isinstance(q, dict)
                         and str(q.get("section", "") or "").strip() == sec]
        if not sec_questions:
            continue
        gname = _unique_name(_xlsform_slug(sec) or "section", "section")
        is_repeat = sec.lower().startswith("repeat:")
        if is_repeat:
            has_repeat = True
            _row("begin_repeat", gname, sec, sec, required="false", appearance="field-list")
        else:
            _row("begin_group", gname, sec, sec, required="false", appearance="field-list")
        for qi, q in enumerate(sec_questions):
            qname = _unique_name(q.get("name") or q.get("code") or f"q{qi + 1}", f"q{qi + 1}")
            q_en = str(q.get("question_en", "") or "")
            q_mm = str(q.get("question_mm", "") or "")
            qtype = str(q.get("type", "open_text") or "open_text")
            if qtype not in QUESTION_TYPES:
                qtype = "open_text"
            opts_en = [o.strip() for o in str(q.get("options_en", "") or "").split("|") if o.strip()]
            opts_mm = [o.strip() for o in str(q.get("options_mm", "") or "").split("|") if o.strip()]
            if not opts_mm:
                opts_mm = list(opts_en)
            hint_en = str(q.get("hint_en", "") or "")
            hint_mm = str(q.get("hint_mm", "") or "")
            required = "true" if q.get("required") else "false"
            appearance = str(q.get("appearance", "") or "")
            relevant = str(q.get("relevant", "") or "")
            constraint = str(q.get("constraint", "") or "")
            constraint_message = str(q.get("constraint_message", "") or "")
            outcome = str(q.get("outcome", "") or "")
            if "terminate" in outcome.lower():
                has_terminate = True
            if qtype == "single_select":
                _row(f"select_one {qname}", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
                _add_choices(qname, opts_en, opts_mm)
            elif qtype == "multi_select":
                _row(f"select_multiple {qname}", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
                _add_choices(qname, opts_en, opts_mm)
            elif qtype == "ranking_top3":
                if not constraint:
                    constraint = "count-selected(.) = 3"
                if not constraint_message:
                    constraint_message = "Please select exactly 3."
                _row(f"select_multiple {qname}", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
                _add_choices(qname, opts_en, opts_mm)
            elif qtype in ("likert_1_5", "rating"):
                if not appearance:
                    appearance = "likert"
                if not opts_en:
                    opts_en = ["1", "2", "3", "4", "5"]
                    opts_mm = list(opts_en)
                _row(f"select_one {qname}", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
                _add_choices(qname, opts_en, opts_mm)
            elif qtype == "open_text":
                _row("text", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
            elif qtype == "numeric":
                _row("integer", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
            elif qtype == "date":
                _row("date", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
            elif qtype == "note":
                _row("note", qname, q_en, q_mm, hint_en, hint_mm,
                     "false", appearance, relevant, "", "", "")
            elif qtype == "observed":
                has_observed = True
                if opts_en:
                    _row(f"select_one {qname}", qname, q_en, q_mm, hint_en, hint_mm,
                         required, appearance, relevant, "", constraint, constraint_message)
                    _add_choices(qname, opts_en, opts_mm)
                else:
                    _row("text", qname, q_en, q_mm, hint_en, hint_mm,
                         required, appearance, relevant, "", constraint, constraint_message)
            elif qtype == "contact":
                has_contact = True
                _row("text", f"{qname}_name", q_en + " — Name", (q_mm + " — နာမည်") if q_mm else q_en + " — Name",
                     hint_en, hint_mm, required, appearance, relevant, "", constraint, constraint_message)
                _row("text", f"{qname}_phone", q_en + " — Phone", (q_mm + " — ဖုန်း") if q_mm else q_en + " — Phone",
                     hint_en, hint_mm, required, appearance, relevant, "", constraint, constraint_message)
            else:
                _row("text", qname, q_en, q_mm, hint_en, hint_mm,
                     required, appearance, relevant, "", constraint, constraint_message)
        if is_repeat:
            _row("end_repeat", "", "", "", "", "", "", "", "", "", "", "")
        else:
            _row("end_group", "", "", "", "", "", "", "", "", "", "", "")
    # Closing thank-you as a note row.
    if (qdata.get("closing_en") or "").strip() or (qdata.get("closing_mm") or "").strip():
        _row("note", _unique_name("closing", "closing"),
             qdata.get("closing_en", ""), qdata.get("closing_mm", ""), required="false")
    if has_terminate:
        notes.append("Terminate outcomes are kept as text in the outcome column — "
                     "no skip logic was invented. Add relevant expressions in Kobo where a "
                     "Terminate must hide the rest of the form.")
    if has_repeat:
        notes.append("Repeat sections export with a blank repeat_count — set the count in Kobo.")
    if has_contact:
        notes.append("Contact questions export as two text rows (<name>_name, <name>_phone).")
    if has_observed:
        notes.append("Observed questions export as text (or select_one when options exist) — "
                     "the recruiter records what they see; nothing is asked.")
    wb.save(out_path)
    stats = {"questions": len(questions), "choice_rows": choice_rows,
             "groups": len(sections), "notes": notes}
    return out_path, stats


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
.stApp { font-family: 'IBM Plex Sans', 'Noto Sans Myanmar', 'Myanmar Text', 'Pyidaungsu', -apple-system, 'Segoe UI', sans-serif; font-synthesis: none; }
/* v3.24.5 (#2): font-synthesis:none — if a bold face is ever missing for the
   font in use, the browser must NOT fake bold by stroking the glyphs: faux
   bold is what crushes Myanmar stacked/medial marks together. Real bold
   faces exist for Noto Sans Myanmar (variable wght 100-900) and IBM Plex
   Sans, so this changes nothing when the fonts load; it only protects the
   fallback case. */
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
   it visible on load so it is always reachable.
   v3.24.5 (#3): the Light gap above the logo is ~2x the Dark gap (~95-105px
   vs ~48-50px) and NO theme-dependent geometry rule exists in code — every
   gap rule is theme-independent and DARK_CSS touches no sidebar-header
   geometry, so the cause could not be identified statically. Fix by
   construction: pin the header row to exactly 2.5rem (fixed height,
   overflow hidden, zero padding/margin) in BOTH theme blocks, so both
   themes MUST render the same 40px header + 0.5rem padding = ~48px gap.
   If the Light gap is still ~100px live, the next step is a live DOM
   inspection of the header's children in Light mode. */
[data-testid="stSidebarHeader"] { height: 2.5rem !important; min-height: 2.5rem !important; max-height: 2.5rem !important; padding: 0 !important; margin-bottom: 0 !important; overflow: hidden !important; justify-content: flex-end !important; }
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
   document (not iframed). No native-navigation refactor was needed.
   v3.24.5 (#1b): free manual scrolling can still stop with content
   half-under the fixed toolbar — scroll-padding/scroll-margin only affect
   anchor jumps, not free scroll stops. The CSS mechanism for scroll stops
   is scroll-snap: `y proximity` on the real scroll container + `start`
   alignment on headings, the stage card, radio rows and stat cards. With
   proximity (not mandatory) the snap only engages when the stop is already
   near a snap point, so normal scrolling feels unchanged; a heading/radio/
   card that would rest half-under the toolbar instead settles with its top
   6rem below the viewport top (scroll-padding), i.e. just below the 60px
   toolbar. If the user dislikes the feel, delete the two scroll-snap lines
   below — everything else still works.
   v3.24.5 (#1a): main vertical-block gap 1rem -> 0.65rem. The stage pages
   stack search bar + stage card + two radio rows + step title; at 743px
   viewport the title sat half below the fold. Together with the compacted
   stage card above this reclaims ~35px so the title fits in the first
   viewport. 0.65rem is still generous spacing everywhere else. */
h1, h2, h3, .rf-sh, [data-testid="stRadio"], .rf-metric { scroll-margin-top: 6rem; }
html { scroll-padding-top: 6rem; }
[data-testid="stMain"] { scroll-padding-top: 6rem; scroll-snap-type: y proximity; }
h1, h2, h3, .rf-sh, [data-testid="stRadio"], .rf-metric { scroll-snap-align: start; }
[data-testid="stMainBlockContainer"] [data-testid="stVerticalBlock"] { gap: 0.65rem; }
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
  border-radius: 10px; padding: 8px 18px; margin: 0 0 6px; }
/* v3.24.5 (#1a): stage card compacted (padding 12px->8px, margin 10px->6px)
   so the step title below the Section/Step radio rows fits in the first
   viewport at 743px height instead of sitting half below the fold. */
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
/* v3.24.5 (#3): pin the sidebar header row to exactly 2.5rem in Dark too,
   so both theme blocks render the same ~48px gap above the logo. */
[data-testid="stSidebarHeader"] { height: 2.5rem !important; min-height: 2.5rem !important; max-height: 2.5rem !important; padding: 0 !important; margin-bottom: 0 !important; overflow: hidden !important; }
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

# v3.24.5 (#2): Myanmar diacritics — full diagnosis, verified 2026-11-16.
# What was RULED OUT (with evidence):
# (a) Font weights: Google Fonts serves Noto Sans Myanmar as a VARIABLE font
#     (fvar wght 100-900, GPOS/GSUB present) — one woff2 per subset covers
#     ALL weights 400-800, each with the Myanmar subset (U+1000-109F).
#     Verified with curl: HTTP 200, 39 @font-face blocks, all five weights
#     of the Myanmar subset point at the same file. Bold faces are REAL,
#     not synthesised — weights are NOT the problem.
# (b) Shaping: HarfBuzz shaping of the three problem strings ("မြန်မာ",
#     "တင်ပြမှု", "ခွဲခြမ်းစိတ်ဖြာမှု") with the real variable font gives
#     IDENTICAL glyph sequences, offsets and advances at wght 400 and 700
#     — the shaper positions marks identically at both weights, so Noto
#     itself shapes correctly.
# (c) Loader gating: the v3.24.4 loader sets _font_link_loaded before
#     emitting st.html, guards duplicates via link[data-rf-font], uses the
#     exact @import URL, and unsafe_allow_javascript=True is set. Scripts
#     DO execute (the scroll-to-top script on the same mechanism works
#     live). No gating bug found.
# (d) Font stack: the .stApp stack covers buttons/headings/labels; the only
#     font-family rules without Noto are the Material Symbols icon rules
#     (correct). No stack bug found.
# CONCLUSION: the code-side font pipeline is correct, yet the v3.24.4 loader
# produced NO visible change — so the font almost certainly never reaches
# the user's browser (fonts.googleapis.com / fonts.gstatic.com unreachable
# or blocked from his browser/network) and the fallback font renders bold
# Myanmar marks tight. The fix is to stop depending on Google Fonts:
# BUNDLED FONT. The Noto Sans Myanmar variable woff2 files are served from
# the project's own GitHub Pages (same host as the privacy page, already
# proven reachable) via @font-face rules injected by this script as a
# script-created <style> (never sanitised). The Google Fonts <link> stays
# as well: the bundled face is appended first, so CSS font matching tries
# it first and falls through to Google's face if the bundled file 404s —
# the app works before AND after the files are uploaded.
# *** FILES THE USER MUST UPLOAD TO THE REPO (GitHub web UI, Add file): ***
#   static/fonts/noto-sans-myanmar-myanmar.woff2  (153,980 bytes, Myanmar subset, wght 100-900)
#   static/fonts/noto-sans-myanmar-latin.woff2    (31,596 bytes, latin subset, wght 100-900)
#   Staging copies are in the workspace at files/static/fonts/ (same names).
#   Until they are uploaded, the @font-face rules 404 harmlessly and the
#   Google Fonts link remains the only source (current behaviour).
# If the marks are STILL tight after the files are uploaded and the app is
# reloaded, the cause is NOT the font at all — then check live in the
# browser: getComputedStyle(el).fontFamily on the "မြန်မာ" button and
# document.fonts.check('700 16px "Noto Sans Myanmar"', 'မြ') — that tells
# whether Noto is actually the used font.
# v3.24.5 also adds font-synthesis:none on .stApp (see CUSTOM_CSS): if a
# bold face is ever missing, the browser must not fake bold by stroking
# glyphs — faux bold is what crushes stacked marks.
if "_font_link_loaded" not in st.session_state:
    st.session_state["_font_link_loaded"] = True
    st.html(
        """<script>(function(){if(!document.querySelector('style[data-rf-font-face]')){var s=document.createElement('style');s.setAttribute('data-rf-font-face','1');s.textContent="@font-face{font-family:'Noto Sans Myanmar';font-style:normal;font-weight:100 900;font-display:swap;src:url('https://soemoeconsultant-alt.github.io/researchflow/static/fonts/noto-sans-myanmar-myanmar.woff2') format('woff2');unicode-range:U+1000-109F,U+200C-200D,U+25CC,U+A92E,U+A9E0-A9FE,U+AA60-AA7F,U+116D0-116E3;}"+"@font-face{font-family:'Noto Sans Myanmar';font-style:normal;font-weight:100 900;font-display:swap;src:url('https://soemoeconsultant-alt.github.io/researchflow/static/fonts/noto-sans-myanmar-latin.woff2') format('woff2');unicode-range:U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD;}";document.head.appendChild(s);}if(!document.querySelector('link[data-rf-font]')){var l=document.createElement('link');l.rel='stylesheet';l.setAttribute('data-rf-font','1');l.href='https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=Noto+Sans+Myanmar:wght@400;500;600;700;800&display=swap';document.head.appendChild(l);}})();</script>""",
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

# v3.29: per-project editor/widget state lives in session state keyed by
# widget name, NOT per project — switching projects without clearing it
# would show the old project's text in the new project's editors (brief
# form, design editors, DG/questionnaire editors). Clear it on every
# project change; the editors re-seed from the new project's saved files.
def _clear_project_editor_state():
    for _k in ("dg_editor", "dg_pending_draft", "dg_confirm_overwrite",
               "q_pending_draft", "q_confirm_overwrite", "q_editor_version",
               "q_intro_en", "q_intro_mm", "q_closing_en", "q_closing_mm",
               "q_header", "q_json_editor", "q_lang", "dg_lang",
               "brief_client_name", "brief_client_contact", "brief_client_email",
               "brief_project_code", "brief_background", "brief_objectives",
               "brief_target_audience", "brief_scope", "brief_start", "brief_end",
               "brief_deliverables", "brief_dg", "brief_approval",
               "design_approach", "design_methods", "design_rationale",
               "design_goal", "design_target", "design_sample_size",
               "design_sessions", "design_pps", "design_personas",
               "design_quotas", "design_screening", "design_recruitment"):
        st.session_state.pop(_k, None)
    for _k in [k for k in list(st.session_state.keys()) if k.startswith("q_editor_v")]:
        st.session_state.pop(_k, None)


# Project selector: list once per session (Drive call), refresh after creating one.
if "project_list" not in st.session_state:
    st.session_state["project_list"] = list_projects()
projects = st.session_state["project_list"]
current = st.session_state.get("active_project", "Default")
if current not in projects and st.session_state.pop("_active_from_prefs", None):
    # v3.29: the persisted project no longer exists — fall back to Default
    # instead of opening a project with no folder.
    current = "Default"
    st.session_state["active_project"] = "Default"
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
    # v3.29: the dropdown truncates long names — the full name is shown
    # here as a caption (it wraps; no fragile CSS on the native selectbox).
    st.sidebar.caption("📁 " + selected)
if selected != current:
    st.session_state["active_project"] = selected
    st.session_state["drive_synced"] = False
    _clear_project_editor_state()
    _update_ui_prefs(active_project=selected)
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
        _update_ui_prefs(theme=st.session_state.get("theme", "light"), lang="EN")
        st.rerun()
with _lc2:
    if st.button("မြန်မာ", key="lang_mm", type="primary" if st.session_state.get("lang", "EN") == "MM" else "secondary", use_container_width=True):
        st.session_state["lang"] = "MM"
        _update_ui_prefs(theme=st.session_state.get("theme", "light"), lang="MM")
        st.rerun()
st.sidebar.caption(T("Theme", "အပြင်အဆင်"))
_tc1, _tc2 = st.sidebar.columns(2)
with _tc1:
    if st.button("☀️ Light", key="theme_light", type="primary" if st.session_state.get("theme", "light") == "light" else "secondary", use_container_width=True):
        st.session_state["theme"] = "light"
        _update_ui_prefs(theme="light", lang=st.session_state.get("lang", "EN"))
        st.rerun()
with _tc2:
    if st.button("🌙 Dark", key="theme_dark", type="primary" if st.session_state.get("theme", "light") == "dark" else "secondary", use_container_width=True):
        st.session_state["theme"] = "dark"
        _update_ui_prefs(theme="dark", lang=st.session_state.get("lang", "EN"))
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
    proposal_n = 1 if (NOTES_DIR / "proposal.docx").exists() else 0
    quant_n = len(list(ANALYSIS_DIR.glob("quant_*.txt")))
    multi_n = len(list(ANALYSIS_DIR.glob("multi_*_analysis.txt")))
    note_n = len([f for f in NOTES_DIR.glob("*")
                  if f.is_file() and f.name not in {"brief.json", "project.json"}])
    any_files = audio_n + tr_n + an_n + rep_n > 0
    brief = load_brief()
    design = load_design()
    _d_methods = design.get("methods") or []
    _d_q = design.get("questionnaire") or {}
    _d_qs = _d_q.get("questions") if isinstance(_d_q, dict) else []
    _d2_items = [
        _stage_item("Approach (Qual / Quant / Mixed)", "Approach (Qual / Quant / Mixed)",
                    "done" if design.get("approach") in DESIGN_APPROACHES else "todo"),
        _stage_item("IDI / FGD / survey", "IDI / FGD / survey",
                    "done" if _d_methods else "todo"),
        _stage_item("Sampling", "Sampling",
                    "done" if (design.get("sample_size") or 0) > 0 or (design.get("sessions_count") or 0) > 0
                    or design.get("personas") or design.get("quotas") else "todo"),
        _stage_item("DG / questionnaire / screener", "DG / questionnaire / screener",
                    "done" if design.get("dg_text") or _d_qs else "todo"),
        _stage_item("Fieldwork plan", "ကွင်းဆင်းအစီအစဉ်",
                    "done" if design.get("fieldwork_plan") else "todo"),
    ]
    _fw_counts = fieldwork_counts(design, load_fieldwork())
    _fw_q = design.get("questionnaire") or {}
    _fw_qs = _fw_q.get("questions") if isinstance(_fw_q, dict) else []
    _qc_checked, _qc_corrected, _qc_needs, _qc_total, _qc_audited = _qc_summary()
    _ins_saved_n = len(load_insights().get("insights", []))
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
         all(_s == "done" for _e, _m, _s in _d2_items),
         _d2_items),
        (_sl("3 · Fieldworks"),
         T("Collect data in the field", "ကွင်းဆင်းပြီး ဒေတာ ကောက်ယူခြင်း"),
         _fw_counts["responses"] > 0 or _fw_counts["participants"] > 0 or audio_n > 0,
         [
             _stage_item("Form", "Form", "done" if _fw_qs else "todo"),
             _stage_item("F2F discussion", "F2F discussion", "done" if audio_n > 0 else "todo"),
             _stage_item("Submit data", "ဒေတာ တင်ခြင်း", "done" if _fw_counts["responses"] > 0 else "todo"),
             _stage_item("Record", "အသံသွင်းခြင်း", "done" if audio_n > 0 else "todo"),
             _stage_item("Interview management", "Interview စီမံခန့်ခွဲခြင်း", "done" if _fw_counts["participants"] > 0 else "todo"),
             _stage_item("Survey data", "Survey ဒေတာ", "done" if quant_n > 0 or _fw_counts["survey_files"] > 0 else "todo"),
             _stage_item("Field notes", "ကွင်းဆင်းမှတ်စု", "done" if _fw_counts["notes"] > 0 else "todo"),
             _stage_item("Tracking", "ခြေရာခံခြင်း", "done" if _fw_counts["sessions_confirmed"] > 0 or _fw_counts["responses"] > 0 else "todo"),
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
             _stage_item("QC: check", "QC: စစ်ဆေးခြင်း", "done" if _qc_checked > 0 else "todo"),
             _stage_item("QC: correct", "QC: ပြင်ဆင်ခြင်း", "done" if _qc_corrected > 0 else "todo"),
             _stage_item("QC: status", "QC: အခြေအနေ", "done" if _qc_checked > 0 else "todo"),
             _stage_item("QC: audit trail", "QC: audit trail", "done" if _qc_audited > 0 else "todo"),
         ]),
        (_sl("5 · Analysis & Reports"),
         T("Generate insights & reports", "Insight နဲ့ report ထုတ်လုပ်ခြင်း"),
         an_n > 0 or rep_n > 0,
         [
             _stage_item("Qual analysis", "Qual ခွဲခြမ်းစိတ်ဖြာမှု", "done" if an_n > 0 else "todo"),
             _stage_item("Quant analysis", "Quant ခွဲခြမ်းစိတ်ဖြာမှု", "done" if quant_n > 0 else "todo"),
             _stage_item("Insights (Qual + Quant)", "Insights (Qual + Quant)", "done" if _ins_saved_n > 0 else "todo"),
             _stage_item("Report (PPTX / DOCX)", "Report (PPTX / DOCX)", "done" if rep_n > 0 else "todo"),
             _stage_item("Proposal (draft)", "Proposal (မူကြမ်း)", "done" if proposal_n > 0 else "todo"),
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

# v3.29 one-shot confirmation (see _flash): a save handler queues its
# message here and it is shown once on the run after the rerun, then
# cleared. Rendered here, above every page's content.
if st.session_state.get("_flash_msg"):
    _fm_kind, _fm_text = st.session_state.pop("_flash_msg")
    if _fm_kind == "error":
        st.error(_fm_text)
    elif _fm_kind == "warning":
        st.warning(_fm_text)
    elif _fm_kind == "info":
        st.info(_fm_text)
    else:
        st.success(_fm_text)

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
    # v3.24.5 (#1): second attempt at 300ms — cheap insurance in case
    # Streamlit restores the old scroll position after the 60ms attempt on a
    # slow rerun. Gated on a real page change only, and it runs within 300ms
    # of the page-change render, so the user cannot have scrolled yet.
    st.html(
        """<script>(function(){function t(){var m=document.querySelector('[data-testid="stMain"]');if(m){m.scrollTop=0;}window.scrollTo(0,0);}setTimeout(t,60);setTimeout(t,300);})();</script>""",
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
            # v3.29: clear the old project's editor state, remember this
            # project across reloads, and confirm (a plain st.success here
            # would be wiped by the rerun below).
            _clear_project_editor_state()
            _update_ui_prefs(active_project=_created)
            _flash(T(f"Project created: {_created}.", f"Project ဖွင့်ပြီးပြီ: {_created}။"))
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
        # v3.29: these are TEXT fields, not a calendar picker and not a
        # dropdown — say so, so nobody waits for a picker that never comes.
        st.caption(T("Type the date as text (YYYY-MM-DD). No calendar picker — an empty date stays empty, never today's date.", "ရက်စွဲကို စာရိုက်ထည့်ပါ (YYYY-MM-DD)။ Calendar မရှိဘူး — မထည့်ရင် ဗလာအတိုင်း ကျန်မယ်, ဒီနေ့ရက် အလိုအလျောက် မထည့်ဘူး။"))
        _bt1, _bt2 = st.columns(2)
        _start = _bt1.text_input(T("Start date (YYYY-MM-DD, type as text)", "စမည့်ရက် (YYYY-MM-DD, စာရိုက်ထည့်ပါ)"), value=_brief.get("timeline_start", ""), key="brief_start")
        _end = _bt2.text_input(T("End date (YYYY-MM-DD, type as text)", "ပြီးမည့်ရက် (YYYY-MM-DD, စာရိုက်ထည့်ပါ)"), value=_brief.get("timeline_end", ""), key="brief_end")
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
            _flash(T("Brief saved.", "Brief သိမ်းပြီးပြီ။"))
            st.rerun()

elif page == "Research Design":
    # Stage 2 builder (v3.25): method, sampling, Discussion Guide +
    # questionnaire (AI draft), fieldwork plan — all saved per project as
    # design.json in notes/ (Drive + local). The page starts EMPTY for a new
    # project; the only prefill is from the project's OWN saved Brief.
    st.title(T("📐 Research Design", "📐 သုတေသန ဒီဇိုင်း"))
    _brief = load_brief()
    _design = load_design()
    _has_brief = bool(_brief)
    if not _has_brief:
        st.warning(
            T("No Client Brief saved yet — fill in the Client Brief first so the AI drafts have something to work from. Method and sampling below can still be set.",
              "Client Brief မသိမ်းရသေးဘူး — အရင် Client Brief ဖြည့်ပါ၊ AI draft တွေက Brief ကနေ ရေးမှာ။ အောက်က Method နဲ့ Sampling ကိုတော့ အခု သတ်မှတ်လို့ရတယ်။")
        )
    # Saved-design summary (what is set, what is still empty).
    _sum_items = [
        ("Method", "Method", bool(_design.get("approach")) or bool(_design.get("methods"))),
        ("Sampling", "Sampling", bool(_design.get("sample_size") or _design.get("sessions_count") or _design.get("personas") or _design.get("quotas") or _design.get("screening_criteria"))),
        ("Discussion Guide", "Discussion Guide", bool((_design.get("dg_text") or "").strip())),
        ("Questionnaire", "Questionnaire", bool((_design.get("questionnaire") or {}).get("questions") if isinstance(_design.get("questionnaire"), dict) else _design.get("questionnaire"))),
        ("Fieldwork Plan", "Fieldwork Plan", bool(_design.get("fieldwork_plan"))),
    ]
    st.caption(" · ".join(("✅ " if _done else "○ ") + T(_en, _mm) for _en, _mm, _done in _sum_items))

    # ---- 1. Method ----
    st.subheader(T("1. Method", "၁။ Method"))
    _approach_idx = DESIGN_APPROACHES.index(_design.get("approach")) if _design.get("approach") in DESIGN_APPROACHES else None
    _approach = st.radio(T("Research approach", "သုတေသန approach"), DESIGN_APPROACHES, index=_approach_idx, key="design_approach")
    _methods = st.multiselect(T("Method", "Method"), DESIGN_METHODS, default=[m for m in _design.get("methods", []) if m in DESIGN_METHODS], key="design_methods")
    if _approach == "Qualitative" and "Survey" in _methods:
        st.warning(T("Qualitative with a survey is unusual — a survey is usually Quantitative or Mixed.", "Qualitative နဲ့ Survey တွဲတာ မများဘူး — Survey က Quantitative ဒါမှမဟုတ် Mixed မှာ ပိုများတယ်။"))
    if _approach == "Quantitative" and ("IDI" in _methods or "FGD" in _methods):
        st.warning(T("Quantitative with IDI / FGD is unusual — interviews and groups are usually Qualitative or Mixed.", "Quantitative နဲ့ IDI / FGD တွဲတာ မများဘူး — IDI / FGD က Qualitative ဒါမှမဟုတ် Mixed မှာ ပိုများတယ်။"))
    if _approach and _approach != "Mixed" and len(_methods) > 1:
        st.info(T("More than one method — consider Mixed.", "Method တစ်ခုထက်ပိုရင် Mixed စဉ်းစားကြည့်ပါ။"))
    _rationale = st.text_area(T("Why this method? (rationale)", "ဘာလို့ ဒီ method ရွေးတာလဲ။"), value=_design.get("rationale", ""), height=80, key="design_rationale")
    _goal_default = _design.get("goal", "") or _brief.get("objectives", "")
    _goal = st.text_area(T("Research goal", "သုတေသန ရည်ရွယ်ချက်"), value=_goal_default, height=80, key="design_goal")
    _target_default = _design.get("target", "") or _brief.get("target_audience", "")
    _target = st.text_area(T("Target", "Target"), value=_target_default, height=80, key="design_target")
    if st.button(T("Save method", "Method သိမ်းမယ်"), key="design_save_method", type="primary"):
        _design.update({"approach": _approach or "", "methods": list(_methods),
                        "rationale": _rationale.strip(), "goal": _goal.strip(),
                        "target": _target.strip(), "target_audience": _brief.get("target_audience", "")})
        save_design(_design)
        _flash(T("Method saved.", "Method သိမ်းပြီးပြီ။"))
        st.rerun()

    # ---- 2. Sampling ----
    st.subheader(T("2. Sampling", "၂။ Sampling"))
    if _has_brief:
        st.text_input(T("Target audience (from the brief)", "Target audience (Brief ကနေ)"), value=_brief.get("target_audience", ""), disabled=True, key="design_audience_ro")
    _sc1, _sc2, _sc3 = st.columns(3)
    _sample_size = _sc1.number_input(T("Sample size (total respondents)", "Sample size (ဖြေဆိုသူ စုစုပေါင်း)"), value=_ed_int(_design.get("sample_size"), 0), min_value=0, step=1, key="design_sample_size")
    _sessions = _sc2.number_input(T("Sessions (groups / interviews)", "Sessions (အဖွဲ့ / အင်တာဗျူး အရေအတွက်)"), value=_ed_int(_design.get("sessions_count"), 0), min_value=0, step=1, key="design_sessions")
    _pps = _sc3.number_input(T("Participants per session", "တစ်ဖွဲ့မှာ လူအရေအတွက်"), value=_ed_int(_design.get("participants_per_session"), 0), min_value=0, step=1, key="design_pps")
    st.caption(T("0 means unset — empty is empty, never guessed.", "0 ဆိုတာ မသတ်မှတ်ရသေးဘူး — မရှိရင် မရှိဘူး၊ ခန့်မှန်းမထည့်ဘူး။"))
    st.caption(T("Personas (from the brief's respondent criteria)", "Personas (Brief ရဲ့ respondent criteria ကနေ)"))
    _personas = _design.get("personas") or []
    _p_df = pd.DataFrame([{"Name": _ed_str(p.get("name")), "Description": _ed_str(p.get("description")), "Value framing": _ed_str(p.get("value_framing"))} for p in _personas if isinstance(p, dict)], columns=["Name", "Description", "Value framing"])
    _p_edited = st.data_editor(_p_df, num_rows="dynamic", use_container_width=True, key="design_personas")
    st.caption(T("Quotas (channel / segment splits)", "Quotas (channel / segment ခွဲဝေမှု)"))
    _quotas = _design.get("quotas") or []
    _qt_df = pd.DataFrame([{"Segment": _ed_str(q.get("segment")), "Share": _ed_str(q.get("share")), "Notes": _ed_str(q.get("notes"))} for q in _quotas if isinstance(q, dict)], columns=["Segment", "Share", "Notes"])
    _qt_edited = st.data_editor(_qt_df, num_rows="dynamic", use_container_width=True, key="design_quotas")
    _screening = st.text_area(T("Screening criteria — who qualifies, who is screened out", "Screening criteria — ဘယ်သူ ပါမလဲ၊ ဘယ်သူ မပါဘူး"), value=_design.get("screening_criteria", ""), height=80, key="design_screening")
    _recruitment = st.text_area(T("Recruitment criteria / notes", "Recruitment criteria / မှတ်ချက်"), value=_design.get("recruitment_criteria", ""), height=80, key="design_recruitment")
    # Current editor values (used by the Save button below and by the AI drafts).
    _new_personas = []
    for _r in _p_edited.to_dict("records"):
        _nm = _ed_str(_r.get("Name"))
        if _nm or _ed_str(_r.get("Description")) or _ed_str(_r.get("Value framing")):
            _new_personas.append({"name": _nm, "description": _ed_str(_r.get("Description")), "value_framing": _ed_str(_r.get("Value framing"))})
    _new_quotas = []
    for _r in _qt_edited.to_dict("records"):
        _sg = _ed_str(_r.get("Segment"))
        if _sg or _ed_str(_r.get("Share")) or _ed_str(_r.get("Notes")):
            _new_quotas.append({"segment": _sg, "share": _ed_str(_r.get("Share")), "notes": _ed_str(_r.get("Notes"))})
    if st.button(T("Save sampling", "Sampling သိမ်းမယ်"), key="design_save_sampling", type="primary"):
        _design.update({"sample_size": int(_sample_size), "sessions_count": int(_sessions),
                        "participants_per_session": int(_pps), "personas": _new_personas,
                        "quotas": _new_quotas, "screening_criteria": _screening.strip(),
                        "recruitment_criteria": _recruitment.strip(),
                        "target_audience": _brief.get("target_audience", "")})
        save_design(_design)
        _flash(T("Sampling saved.", "Sampling သိမ်းပြီးပြီ။"))
        st.rerun()
    # Design as the AI drafts see it: saved values + what is on screen now.
    _draft_design = dict(_design)
    _draft_design.update({"approach": _approach or "", "methods": list(_methods),
                          "rationale": _rationale.strip(), "goal": _goal.strip(),
                          "target": _target.strip(), "personas": _new_personas,
                          "quotas": _new_quotas, "screening_criteria": _screening.strip(),
                          "recruitment_criteria": _recruitment.strip()})

    # ---- 3. Discussion Guide ----
    st.subheader(T("3. Discussion Guide", "၃။ Discussion Guide"))
    _dg_lang = st.selectbox(T("Discussion Guide language", "Discussion Guide ဘာသာစကား"), DESIGN_DG_LANGUAGES, index=DESIGN_DG_LANGUAGES.index(_design.get("dg_language", "English")) if _design.get("dg_language") in DESIGN_DG_LANGUAGES else 0, key="dg_lang")
    _survey_only = list(_methods) == ["Survey"]
    if _survey_only:
        st.info(T("A survey-only project has no Discussion Guide — the questionnaire below is the instrument.", "Survey သီးသန့် project မှာ Discussion Guide မရှိဘူး — အောက်က questionnaire ကပဲ instrument ဖြစ်တယ်။"))

    def _run_dg_draft():
        _sess = "120 minutes (FGD)" if "FGD" in _draft_design.get("methods", []) else "60 minutes (IDI)"
        _lang_note = (" — write the whole guide in everyday colloquial Myanmar; never mix English and Myanmar in the same line." if _dg_lang == "Myanmar" else " — mostly English with an occasional Myanmar line where the moderator needs it.")
        _prompt = (PROMPT_LIBRARY["discussion_guide"]["template"] + "\n\n"
                   + _design_prompt_context(_brief, _draft_design)
                   + "\n\nDiscussion Guide language: " + _dg_lang + _lang_note
                   + "\nSession length: " + _sess)
        with st.spinner(T("Drafting the Discussion Guide…", "Discussion Guide ရေးနေတယ်…")):
            _draft, _model_used = ask_gemini(_prompt)
        if _draft:
            st.session_state["dg_pending_draft"] = _draft
            # v3.29: set the editor value directly through the Session State
            # API — the reliable way to set a widget value. The old code
            # popped the widget key and passed the draft as value= to the
            # text_area, but once the key exists value= is ignored on later
            # reruns, so the editor stayed empty while the Preview showed
            # the draft (trial FAIL) and Save stored nothing.
            st.session_state["dg_editor"] = _draft
            st.session_state.pop("dg_confirm_overwrite", None)
            st.rerun()
        else:
            st.error(T(f"AI draft failed: {_model_used}", f"AI draft မရဘူး: {_model_used}"))

    if st.button(T("✨ AI Draft Discussion Guide", "✨ AI နဲ့ Discussion Guide ရေးခိုင်းမယ်"), key="dg_draft", type="primary", disabled=(not _has_brief) or _survey_only):
        _has_saved_dg = bool((_design.get("dg_text") or "").strip())
        if _has_saved_dg and not st.session_state.get("dg_confirm_overwrite"):
            st.session_state["dg_confirm_overwrite"] = True
            st.rerun()
        else:
            _run_dg_draft()
    if st.session_state.get("dg_confirm_overwrite"):
        st.warning(T("A Discussion Guide is already saved. Drafting again will REPLACE the text in the editor — the saved guide stays in design.json until you press Save.", "Discussion Guide သိမ်းပြီးသား ရှိတယ်။ ထပ်ရေးခိုင်းရင် editor ထဲက စာသား အစားထိုးသွားမယ် — Save မနှိပ်မချင်း သိမ်းပြီးသား guide က design.json မှာ ရှိနေသေးတယ်။"))
        _cy, _cn = st.columns(2)
        if _cy.button(T("Yes, draft again", "ဟုတ်ကဲ့၊ ထပ်ရေးခိုင်းမယ်"), key="dg_confirm_yes", type="primary"):
            _run_dg_draft()
        if _cn.button(T("Keep my edits", "ကိုယ့်စာသား ထားမယ်"), key="dg_confirm_no"):
            st.session_state.pop("dg_confirm_overwrite", None)
            st.rerun()
    if st.session_state.get("dg_pending_draft"):
        st.info(T("New AI draft in the editor — check it and edit it, then Save.", "AI draft အသစ် editor ထဲမှာ ရှိတယ် — စစ်ပြီး ပြင်ပါ၊ ပြီးရင် Save နှိပ်ပါ။"))
    # v3.29: the editor value lives in session state — seeded once from the
    # saved guide (or empty for a new project) and set directly when a
    # draft arrives. Never pass value= here: it is ignored once the widget
    # key exists, which is what desynced the editor. Project switches clear
    # this key (see _clear_project_editor_state), so a new project re-seeds.
    if "dg_editor" not in st.session_state:
        st.session_state["dg_editor"] = _design.get("dg_text", "")
    st.text_area(T("Discussion Guide (editable)", "Discussion Guide (ပြင်လို့ရတယ်)"), height=400, key="dg_editor")
    # v3.31 (P6): self-check — warnings only, never blocks saving. Runs on
    # the current editor text on every render, so a fresh AI draft is
    # checked as soon as it lands (the draft flow reruns into this render)
    # and a saved guide is checked whenever it is shown. Results are never
    # stored — they are recomputed from the current text.
    _dg_text_now = st.session_state.get("dg_editor", "")
    if str(_dg_text_now or "").strip():
        _dg_minutes = 120 if "FGD" in _draft_design.get("methods", []) else 60
        _show_self_check_results(_dg_self_check(_dg_text_now, _brief, _draft_design, _dg_minutes),
                                 "Discussion Guide checks", "Discussion Guide စစ်ဆေးမှု")
    _dc1, _dc2 = st.columns(2)
    if _dc1.button(T("Save Discussion Guide", "Discussion Guide သိမ်းမယ်"), key="dg_save", type="primary"):
        _dg_final = st.session_state.get("dg_editor", "")
        if not _dg_final.strip():
            # v3.29: never destroy a draft (or a saved guide) with an empty
            # save — warn and keep everything. No rerun on this path, so the
            # warning stays visible.
            if (st.session_state.get("dg_pending_draft") or "").strip():
                st.warning(T("The editor is empty but a new AI draft exists — it was NOT saved. Click “✨ AI Draft Discussion Guide” again, or paste the draft into the editor, then Save.", "Editor ဗလာဖြစ်နေတယ်၊ ဒါပေမယ့် AI draft အသစ် ရှိတယ် — မသိမ်းလိုက်ဘူး။ “✨ AI Draft Discussion Guide” ကို ပြန်နှိပ်ပါ၊ ဒါမှမဟုတ် draft ကို editor ထဲ paste လုပ်ပြီး Save နှိပ်ပါ။"))
            elif (_design.get("dg_text") or "").strip():
                st.warning(T("The editor is empty — the saved Discussion Guide was kept, not erased.", "Editor ဗလာဖြစ်နေတယ် — သိမ်းပြီးသား Discussion Guide ကို ဖျက်မပစ်ဘူး၊ ထားထားတယ်။"))
            else:
                st.warning(T("Nothing to save — the Discussion Guide is empty.", "သိမ်းစရာ မရှိဘူး — Discussion Guide ဗလာဖြစ်နေတယ်။"))
        else:
            _design["dg_text"] = _dg_final
            _design["dg_language"] = _dg_lang
            _design["dg_file"] = "discussion_guide.txt"
            save_design(_design)
            store_write_text("notes", "discussion_guide.txt", _dg_final)
            st.session_state.pop("dg_pending_draft", None)
            _flash(T("Discussion Guide saved.", "Discussion Guide သိမ်းပြီးပြီ။"))
            st.rerun()
    if _dc2.button(T("Export DOCX", "DOCX ထုတ်မယ်"), key="dg_export"):
        _dg_final = st.session_state.get("dg_editor", "")
        if not _dg_final.strip():
            st.warning(T("Nothing to export — the Discussion Guide is empty.", "ထုတ်စရာ မရှိဘူး — Discussion Guide ဗလာဖြစ်နေတယ်။"))
        else:
            _out = NOTES_DIR / "discussion_guide.docx"
            dg_to_docx(_dg_final, f"Discussion Guide — {current_project()}", _out)
            store_upload_existing("notes", "discussion_guide.docx")
            st.success(T("DOCX exported.", "DOCX ထုတ်ပြီးပြီ။"))
    if (NOTES_DIR / "discussion_guide.docx").exists():
        st.download_button(T("Download discussion_guide.docx", "discussion_guide.docx ဒေါင်းလုပ်ဆွဲမယ်"), data=(NOTES_DIR / "discussion_guide.docx").read_bytes(), file_name="discussion_guide.docx", key="dg_download")
    with st.expander(T("Preview", "အစမ်းကြည့်မယ်")):
        _dg_prev = st.session_state.get("dg_editor", "")
        if _dg_prev.strip():
            st.markdown(_dg_to_markdown(_dg_prev))
        else:
            st.caption(T("Empty — no Discussion Guide yet.", "ဗလာဖြစ်နေတယ် — Discussion Guide မရှိသေးဘူး။"))

    # ---- 4. Questionnaire / Screener ----
    st.subheader(T("4. Questionnaire / Screener", "၄။ Questionnaire / Screener"))
    _q_lang = st.selectbox(T("Questionnaire language", "Questionnaire ဘာသာစကား"), DESIGN_Q_LANGUAGES, index=DESIGN_Q_LANGUAGES.index(_design.get("questionnaire_language", "Myanmar")) if _design.get("questionnaire_language") in DESIGN_Q_LANGUAGES else 0, key="q_lang")
    _q_saved = _design.get("questionnaire") or {}
    if isinstance(_q_saved, list):
        _q_saved = {"questions": _q_saved}
    if st.session_state.get("q_pending_draft") is not None:
        _q_saved = st.session_state["q_pending_draft"]
    _qs = _q_saved.get("questions") or []
    if not isinstance(_qs, list):
        _qs = []

    def _run_q_draft():
        _lang_note = (" — write the questions in everyday colloquial Myanmar (respondent language); store both EN and MM text for every question." if _q_lang == "Myanmar" else " — write the questions in English; store the Myanmar text the same as the English text.")
        # v3.30 (P1): the questionnaire draft is mode-aware. Qualitative
        # IDI/FGD studies get a screener ONLY (eligibility; question count
        # follows the screening criteria); a Survey method keeps the full
        # questionnaire (screening + objective sections). No method selected
        # defaults to the narrower screener mode.
        _q_methods = _draft_design.get("methods", []) or []
        _q_mode = QUESTIONNAIRE_SURVEY_MODE if "Survey" in _q_methods else QUESTIONNAIRE_SCREENER_MODE
        _prompt = (PROMPT_LIBRARY["questionnaire"]["template"] + "\n\n" + _q_mode + "\n\n"
                   + _design_prompt_context(_brief, _draft_design)
                   + "\n\nQuestionnaire language: " + _q_lang + _lang_note)
        with st.spinner(T("Drafting the questionnaire…", "Questionnaire ရေးနေတယ်…")):
            _draft, _model_used = ask_gemini(_prompt)
        if not _draft:
            st.error(T(f"AI draft failed: {_model_used}", f"AI draft မရဘူး: {_model_used}"))
            return
        _parsed, _perr = _parse_questionnaire_json(_draft)
        if _parsed is None:
            st.error(T(f"AI draft could not be read as JSON: {_perr}", f"AI draft ကို JSON အဖြစ် ဖတ်မရဘူး: {_perr}"))
            with st.expander(T("Raw AI draft", "AI draft အစိမ်း")):
                st.code(_draft)
            return
        st.session_state["q_pending_draft"] = _parsed
        # v3.29: same fix as the DG editor — the grid + intro/closing fields
        # are value+key widgets, so a new draft must be pushed into session
        # state directly (and the grid given a fresh key) instead of relying
        # on pop-key + value=, which desyncs on later reruns.
        st.session_state["q_editor_version"] = st.session_state.get("q_editor_version", 0) + 1
        st.session_state["q_intro_en"] = _parsed.get("intro_en", "")
        st.session_state["q_intro_mm"] = _parsed.get("intro_mm", "")
        st.session_state["q_closing_en"] = _parsed.get("closing_en", "")
        st.session_state["q_closing_mm"] = _parsed.get("closing_mm", "")
        st.session_state["q_header"] = ", ".join(_parsed.get("header_fields", []) or [])
        st.session_state["q_json_editor"] = json.dumps(_parsed, ensure_ascii=False, indent=2)
        st.session_state.pop("q_confirm_overwrite", None)
        st.rerun()

    if st.button(T("✨ AI Draft Questionnaire", "✨ AI နဲ့ Questionnaire ရေးခိုင်းမယ်"), key="q_draft", type="primary", disabled=not _has_brief):
        if _qs and not st.session_state.get("q_confirm_overwrite"):
            st.session_state["q_confirm_overwrite"] = True
            st.rerun()
        else:
            _run_q_draft()
    if st.session_state.get("q_confirm_overwrite"):
        st.warning(T("A questionnaire is already saved. Drafting again will REPLACE the questions in the editor — the saved questionnaire stays in design.json until you press Save.", "Questionnaire သိမ်းပြီးသား ရှိတယ်။ ထပ်ရေးခိုင်းရင် editor ထဲက မေးခွန်းတွေ အစားထိုးသွားမယ် — Save မနှိပ်မချင်း သိမ်းပြီးသား questionnaire က design.json မှာ ရှိနေသေးတယ်။"))
        _qy, _qn = st.columns(2)
        if _qy.button(T("Yes, draft again", "ဟုတ်ကဲ့၊ ထပ်ရေးခိုင်းမယ်"), key="q_confirm_yes", type="primary"):
            _run_q_draft()
        if _qn.button(T("Keep my edits", "ကိုယ့်စာသား ထားမယ်"), key="q_confirm_no"):
            st.session_state.pop("q_confirm_overwrite", None)
            st.rerun()
    if st.session_state.get("q_pending_draft") is not None:
        st.info(T("New AI draft in the editor — check it and edit it, then Save.", "AI draft အသစ် editor ထဲမှာ ရှိတယ် — စစ်ပြီး ပြင်ပါ၊ ပြီးရင် Save နှိပ်ပါ။"))
    # v3.29: seed these fields from the saved/draft questionnaire on first
    # view, then let the widgets own their values (same fix as the DG
    # editor — value= is ignored once the key exists).
    if "q_intro_en" not in st.session_state:
        st.session_state["q_intro_en"] = _q_saved.get("intro_en", "")
    if "q_intro_mm" not in st.session_state:
        st.session_state["q_intro_mm"] = _q_saved.get("intro_mm", "")
    if "q_closing_en" not in st.session_state:
        st.session_state["q_closing_en"] = _q_saved.get("closing_en", "")
    if "q_closing_mm" not in st.session_state:
        st.session_state["q_closing_mm"] = _q_saved.get("closing_mm", "")
    if "q_header" not in st.session_state:
        st.session_state["q_header"] = ", ".join(_q_saved.get("header_fields", []) or [])
    _q_intro_en = st.text_area(T("Intro script (English)", "Intro script (English)"), height=70, key="q_intro_en")
    _q_intro_mm = st.text_area(T("Intro script (Myanmar)", "Intro script (မြန်မာ)"), height=70, key="q_intro_mm")
    _q_closing_en = st.text_area(T("Closing script (English)", "Closing script (English)"), height=60, key="q_closing_en")
    _q_closing_mm = st.text_area(T("Closing script (Myanmar)", "Closing script (မြန်မာ)"), height=60, key="q_closing_mm")
    _q_header = st.text_input(T("Header fields (comma-separated)", "Header fields (ကော်မာနဲ့ ခွဲပါ)"), key="q_header")
    _q_cols = ["Code", "Section", "Question (EN)", "Question (MM)", "Type", "Options (EN)", "Options (MM)", "Outcome", "Required", "Relevant"]
    _q_rows = []
    for _qq in _qs:
        if not isinstance(_qq, dict):
            continue
        _qt = _ed_str(_qq.get("type"))
        _q_rows.append({"Code": _ed_str(_qq.get("code")), "Section": _ed_str(_qq.get("section")),
                        "Question (EN)": _ed_str(_qq.get("question_en")), "Question (MM)": _ed_str(_qq.get("question_mm")),
                        "Type": _qt if _qt in QUESTION_TYPES else "open_text",
                        "Options (EN)": _ed_str(_qq.get("options_en")), "Options (MM)": _ed_str(_qq.get("options_mm")),
                        "Outcome": _ed_str(_qq.get("outcome")), "Required": _ed_bool(_qq.get("required")),
                        "Relevant": _ed_str(_qq.get("relevant"))})
    _q_df = pd.DataFrame(_q_rows, columns=_q_cols)
    # v3.29: the grid key carries a version that bumps on every new draft,
    # so the grid is rebuilt from the draft instead of showing stale rows.
    _q_editor_key = f"q_editor_v{st.session_state.get('q_editor_version', 0)}"
    _q_edited = st.data_editor(_q_df, num_rows="dynamic", use_container_width=True, key=_q_editor_key, column_config={
        "Type": st.column_config.SelectboxColumn(T("Type", "Type"), options=QUESTION_TYPES),
        "Required": st.column_config.CheckboxColumn(T("Required", "Required")),
    })
    # v3.31 (P7): self-check on the questions currently shown (a pending
    # draft replaces the saved list — see _qs above). Warnings only, never
    # blocks saving; recomputed on every render, never stored.
    if _qs:
        _show_self_check_results(_q_self_check({"questions": _qs}, _brief, _draft_design),
                                 "Questionnaire checks", "Questionnaire စစ်ဆေးမှု")
    if st.button(T("Save questionnaire", "Questionnaire သိမ်းမယ်"), key="q_save", type="primary"):
        _new_qs = []
        for _i, _r in enumerate(_q_edited.to_dict("records")):
            _en = _ed_str(_r.get("Question (EN)"))
            _mm = _ed_str(_r.get("Question (MM)"))
            if not _en and not _mm:
                continue
            _hidden = _qs[_i] if _i < len(_qs) and isinstance(_qs[_i], dict) else {}
            _code = _ed_str(_r.get("Code")) or _ed_str(_hidden.get("code"))
            _tt = _ed_str(_r.get("Type"))
            _new_qs.append({
                "code": _code,
                "name": _ed_str(_hidden.get("name")) or _xlsform_slug(_code) or f"q{_i + 1}",
                "section": _ed_str(_r.get("Section")),
                "question_en": _en, "question_mm": _mm,
                "type": _tt if _tt in QUESTION_TYPES else "open_text",
                "options_en": _ed_str(_r.get("Options (EN)")),
                "options_mm": _ed_str(_r.get("Options (MM)")),
                "outcome": _ed_str(_r.get("Outcome")),
                "hint_en": _ed_str(_hidden.get("hint_en")),
                "hint_mm": _ed_str(_hidden.get("hint_mm")),
                "required": _ed_bool(_r.get("Required")),
                "appearance": _ed_str(_hidden.get("appearance")),
                "relevant": _ed_str(_r.get("Relevant")),
                "constraint": _ed_str(_hidden.get("constraint")),
                "constraint_message": _ed_str(_hidden.get("constraint_message")),
            })
        # v3.31 (P15): empty-save guard — parity with the DG guard (v3.29).
        _allow_q, _warn_q_en, _warn_q_mm = _questionnaire_empty_save_guard(
            _new_qs, st.session_state.get("q_pending_draft"), _design.get("questionnaire"))
        if not _allow_q:
            # Never destroy a draft or a saved questionnaire with an empty
            # save — warn and keep everything. No rerun here, so the
            # warning stays visible.
            st.warning(T(_warn_q_en, _warn_q_mm))
        else:
            _header_fields = [h.strip() for h in _q_header.split(",") if h.strip()]
            _design["questionnaire"] = {"intro_en": _q_intro_en.strip(), "intro_mm": _q_intro_mm.strip(),
                                        "closing_en": _q_closing_en.strip(), "closing_mm": _q_closing_mm.strip(),
                                        "header_fields": _header_fields, "questions": _new_qs}
            _design["questionnaire_language"] = _q_lang
            save_design(_design)
            st.session_state.pop("q_pending_draft", None)
            # v3.29: the JSON view is derived from the saved questionnaire —
            # drop it so it re-seeds from the just-saved design.
            st.session_state.pop("q_json_editor", None)
            _flash(T("Questionnaire saved.", "Questionnaire သိမ်းပြီးပြီ။"))
            st.rerun()
    if st.button(T("Export Kobo XLSForm (.xlsx)", "Kobo XLSForm (.xlsx) ထုတ်မယ်"), key="q_export_kobo"):
        _q_now = _design.get("questionnaire") or {}
        if not isinstance(_q_now, dict) or not (_q_now.get("questions") or []):
            st.warning(T("Nothing to export — save a questionnaire first.", "ထုတ်စရာ မရှိဘူး — အရင် questionnaire သိမ်းပါ။"))
        else:
            _out = NOTES_DIR / "questionnaire_kobo.xlsx"
            _res, _stats = questionnaire_to_xlsform(_q_now, f"Questionnaire — {current_project()}", _out)
            if _res is None:
                st.error(_stats)
            else:
                store_upload_existing("notes", "questionnaire_kobo.xlsx")
                st.success(T(f"Kobo XLSForm exported: {_stats['questions']} questions, {_stats['choice_rows']} choices, {_stats['groups']} groups.", f"Kobo XLSForm ထုတ်ပြီးပြီ: မေးခွန်း {_stats['questions']} ခု၊ choices {_stats['choice_rows']} ခု၊ groups {_stats['groups']} ခု။"))
                for _n in _stats.get("notes", []):
                    st.caption("• " + _n)
    if (NOTES_DIR / "questionnaire_kobo.xlsx").exists():
        st.download_button(T("Download questionnaire_kobo.xlsx", "questionnaire_kobo.xlsx ဒေါင်းလုပ်ဆွဲမယ်"), data=(NOTES_DIR / "questionnaire_kobo.xlsx").read_bytes(), file_name="questionnaire_kobo.xlsx", key="q_kobo_download")
    with st.expander(T("Questionnaire JSON (advanced — XLSForm fields)", "Questionnaire JSON (အဆင့်မြင့် — XLSForm fields)")):
        st.caption(T("Hint, appearance, relevant, and constraint live here — edit the JSON and save it.", "Hint, appearance, relevant, constraint တို့ ဒီမှာ ရှိတယ် — JSON ပြင်ပြီး သိမ်းပါ။"))
        if "q_json_editor" not in st.session_state:
            st.session_state["q_json_editor"] = json.dumps(_q_saved, ensure_ascii=False, indent=2)
        _q_json = st.text_area(T("Questionnaire JSON", "Questionnaire JSON"), height=200, key="q_json_editor")
        if st.button(T("Save questionnaire JSON", "Questionnaire JSON သိမ်းမယ်"), key="q_json_save"):
            _parsed, _perr = _parse_questionnaire_json(_q_json)
            if _parsed is None:
                st.error(T(f"Could not read the JSON: {_perr}", f"JSON ဖတ်မရဘူး: {_perr}"))
            else:
                _design["questionnaire"] = _parsed
                save_design(_design)
                st.session_state.pop("q_pending_draft", None)
                _flash(T("Questionnaire JSON saved.", "Questionnaire JSON သိမ်းပြီးပြီ။"))
                st.rerun()

    # ---- 5. Fieldwork Plan ----
    st.subheader(T("5. Fieldwork Plan", "၅။ Fieldwork Plan"))
    _fw = _design.get("fieldwork_plan") or []
    _fw_cols = ["Session", "Group", "Method", "Date", "Time", "Seats", "Recruited", "Confirmed", "Moderator", "Moderator seat", "Venue"]
    _fw_rows = []
    for _i, _r in enumerate(_fw):
        if not isinstance(_r, dict):
            continue
        _fw_rows.append({"Session": _ed_int(_r.get("session"), _i + 1), "Group": _ed_str(_r.get("group")),
                         "Method": _ed_str(_r.get("method")), "Date": _ed_str(_r.get("date")),
                         "Time": _ed_str(_r.get("time")), "Seats": _ed_int(_r.get("seats"), 0),
                         "Recruited": _ed_int(_r.get("recruited"), 0), "Confirmed": _ed_int(_r.get("confirmed"), 0),
                         "Moderator": _ed_str(_r.get("moderator")), "Moderator seat": _ed_str(_r.get("moderator_seat")),
                         "Venue": _ed_str(_r.get("venue"))})
    _fw_df = pd.DataFrame(_fw_rows, columns=_fw_cols)
    _fw_edited = st.data_editor(_fw_df, num_rows="dynamic", use_container_width=True, key="fw_editor", column_config={
        "Method": st.column_config.SelectboxColumn(T("Method", "Method"), options=["", "IDI", "FGD", "Survey"]),
    })
    st.caption(T("Dates are plain text (YYYY-MM-DD) — an unset date never invents today's date. Seats = respondents in the session; Recruited = over-recruited count for no-shows; the moderator sits at Seat 1.", "ရက်စွဲက စာသားပဲ (YYYY-MM-DD) — မသတ်မှတ်ရင် ဒီနေ့ရက် အလိုအလျောက် မထည့်ဘူး။ Seats = အဖွဲ့ထဲမှာ ထိုင်မယ့်သူ၊ Recruited = မလာသူ အတွက် ပိုခေါ်ထားတာ၊ moderator က Seat 1 မှာ ထိုင်တယ်။"))
    # v3.31 (P8): self-check on the saved plan — warnings only, never
    # blocks saving; recomputed on every render, never stored.
    if _fw:
        _show_self_check_results(_fieldwork_self_check(_draft_design, _brief),
                                 "Fieldwork Plan checks", "Fieldwork Plan စစ်ဆေးမှု")
    if st.button(T("Save fieldwork plan", "Fieldwork Plan သိမ်းမယ်"), key="fw_save", type="primary"):
        _new_fw = []
        for _i, _r in enumerate(_fw_edited.to_dict("records")):
            _grp = _ed_str(_r.get("Group"))
            _mth = _ed_str(_r.get("Method"))
            _mod = _ed_str(_r.get("Moderator"))
            if not _grp and not _mth and not _mod and not _ed_str(_r.get("Date")) and not _ed_str(_r.get("Venue")):
                continue
            _seat = _ed_str(_r.get("Moderator seat"))
            if _mod and not _seat:
                _seat = "Seat 1"
            _new_fw.append({"session": _ed_int(_r.get("Session"), _i + 1), "group": _grp, "method": _mth,
                            "date": _ed_str(_r.get("Date")), "time": _ed_str(_r.get("Time")),
                            "seats": _ed_int(_r.get("Seats"), 0), "recruited": _ed_int(_r.get("Recruited"), 0),
                            "confirmed": _ed_int(_r.get("Confirmed"), 0), "moderator": _mod,
                            "moderator_seat": _seat, "venue": _ed_str(_r.get("Venue"))})
        _design["fieldwork_plan"] = _new_fw
        save_design(_design)
        _flash(T("Fieldwork plan saved.", "Fieldwork plan သိမ်းပြီးပြီ။"))
        st.rerun()

elif page == "Data Cleansing":
    # Stage 4 (v3.27): Data Preparation status from real counts + real
    # Quality Control below (checks, per-file status, correction with
    # audit trail). Originals are never overwritten.
    st.title(T("🧹 Data Cleansing", "🧹 ဒေတာ သန့်စင်ခြင်း"))

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

    st.divider()
    st.subheader(T("Quality Control", "အရည်အသွေး စစ်ဆေးခြင်း"))
    _qc_fulls, _qc_pending = get_full_transcripts()
    if _qc_pending:
        st.warning(T("Not merged yet: " + ", ".join(f"{k} ({n} chunks)" for k, n in _qc_pending) + " — merge on the Transcripts page.",
                     "Merge မလုပ်ရသေးဘူး: " + ", ".join(f"{k} ({n} chunks)" for k, n in _qc_pending) + " — Transcripts page မှာ Merge လုပ်ပါ။"))
    if st.button(T("Run QC checks", "QC စစ်မယ်"), key="qc_run"):
        _qc = qc_run_all()
        st.success(T(f"QC checks done — {len(_qc.get('files', {}))} transcript(s) checked.",
                     f"QC စစ်ပြီးပြီ — transcript {len(_qc.get('files', {}))} ခု စစ်ပြီးပြီ။"))
    _qc = load_qc()
    _qc_files = _qc.get("files", {})
    if not _qc_files:
        st.info(T("No transcripts checked yet — transcribe first, then press Run QC checks.",
                  "Transcript မစစ်ရသေးဘူး — အရင် transcribe လုပ်ပြီး Run QC checks နှိပ်ပါ။"))
    else:
        _qc_s = _qc_summary()
        st.caption(T(f"Checked: {_qc_s[0]} of {_qc_s[3]} | Corrected: {_qc_s[1]} | Needs correction: {_qc_s[2]} | Audit entries: {_qc_s[4]}",
                     f"စစ်ပြီး: {_qc_s[0]} / {_qc_s[3]} | ပြင်ပြီး: {_qc_s[1]} | ပြင်ရန်လို: {_qc_s[2]} | Audit: {_qc_s[4]}"))
        _qc_names = sorted(_qc_files)
        for _qi, _qname in enumerate(_qc_names):
            _entry = _qc_files[_qname] if isinstance(_qc_files[_qname], dict) else {}
            _issues = _entry.get("issues", []) if isinstance(_entry.get("issues"), list) else []
            _issue_txt = "; ".join(str(i.get("detail", "")) for i in _issues) if _issues else T("No issues found", "ပြဿနာ မတွေ့ဘူး")
            st.write(f"**{_qname}** — {_issue_txt}")
            _cur_status = _entry.get("status", "Unchecked")
            _new_status = st.selectbox(T("Status", "အခြေအနေ"), QC_STATUSES,
                                       index=QC_STATUSES.index(_cur_status) if _cur_status in QC_STATUSES else 0,
                                       key=f"qc_status_{_qi}")
            if st.button(T("Save status", "အခြေအနေ သိမ်းမယ်"), key=f"qc_save_status_{_qi}"):
                _qc2 = load_qc()
                _e2 = _qc2.setdefault("files", {}).setdefault(_qname, {"status": "Unchecked", "checked_at": "", "issues": []})
                _e2["status"] = _new_status
                save_qc(_qc2)
                st.success(T(f"Status saved: {_qname} → {_new_status}", f"သိမ်းပြီး: {_qname} → {_new_status}"))
            with st.expander(T(f"Correct {_qname}", f"{_qname} ကို ပြင်မယ်")):
                _orig_path = None
                for _d in (OUTPUT_DIR, MERGED_DIR):
                    if (_d / _qname).exists():
                        _orig_path = _d / _qname
                        break
                if _orig_path is None:
                    st.warning(T("Original file not found.", "မူရင်း file မတွေ့ဘူး။"))
                else:
                    _orig_text = _orig_path.read_text(encoding="utf-8", errors="replace")
                    _corr_text = st.text_area(T("Corrected text (the original is never overwritten — a _corrected copy is saved)",
                                                "ပြင်ထားတဲ့ စာသား (မူရင်းကို ဘယ်တော့မှ မရေးဘူး — _corrected copy သိမ်းမယ်)"),
                                              value=_orig_text, height=300, key=f"qc_correct_{_qi}")
                    _corr_note = st.text_input(T("What did you change? (note for the audit trail)",
                                                 "ဘာပြင်လိုက်လဲ? (audit trail အတွက် မှတ်ချက်)"), key=f"qc_note_{_qi}")
                    if st.button(T("Save correction", "ပြင်ဆင်ချက် သိမ်းမယ်"), key=f"qc_save_corr_{_qi}"):
                        _out, _msg = qc_save_correction(_qname, _corr_text, _corr_note)
                        if _out is None:
                            st.warning(_msg)
                        else:
                            st.success(_msg)
        with st.expander(T("Audit trail", "Audit trail")):
            _audit = qc_load_audit()
            if not _audit:
                st.info(T("No corrections logged yet.", "ပြင်ဆင်ချက် မှတ်တမ်း မရှိသေးပါ။"))
            else:
                for _a in reversed(_audit[-10:]):
                    st.write(f"- {_a.get('timestamp', '')} — {_a.get('file', '')} → {_a.get('corrected_file', '')}")
                    if _a.get("before") or _a.get("after"):
                        st.caption(f"Before: {_a.get('before', '')}")
                        st.caption(f"After: {_a.get('after', '')}")
                    if _a.get("note"):
                        st.caption(f"Note: {_a.get('note', '')}")

elif page == "Insights":
    # Stage 5 (v3.27): build insight cards from SAVED outputs — Qual
    # analysis files + Quant summaries. Nothing is invented: a missing
    # input is shown as missing, with a link to the page that makes it.
    st.title(T("💡 Insights", "💡 Insights"))
    _ins_qual = _ins_qual_findings()
    _ins_quant = _ins_quant_numbers()
    _ins_design = load_design()
    _ins_personas = [p.get("name") for p in (_ins_design.get("personas") or []) if isinstance(p, dict) and p.get("name")]
    _ins_sections = _ins_dg_sections()
    st.write(("✅ " if _ins_qual else "○ ") + T(f"Qualitative findings saved: {len(_ins_qual)}", f"သိမ်းထားတဲ့ Qualitative finding: {len(_ins_qual)}"))
    st.write(("✅ " if _ins_quant else "○ ") + T(f"Quantitative numbers saved: {len(_ins_quant)}", f"သိမ်းထားတဲ့ Quantitative နံပါတ်: {len(_ins_quant)}"))
    if not _ins_qual and not _ins_quant:
        st.info(T("No analysis saved yet — analyze on the Market Analysis page (Qual) and save a quant summary on the Quantitative Analysis page first.",
                  "Analysis မသိမ်းရသေးဘူး — အရင် Market Analysis မှာ Qual analyse လုပ်ပြီး Quantitative Analysis မှာ quant summary သိမ်းပါ။"))
        if st.button(T("Open Market Analysis", "Market Analysis ဖွင့်မယ်"), key="ins_goto_ma"):
            st.session_state["page"] = "Market Analysis"
        if st.button(T("Open Quantitative Analysis", "Quantitative Analysis ဖွင့်မယ်"), key="ins_goto_quant"):
            st.session_state["page"] = "Quantitative Analysis"
    else:
        st.subheader(T("Build an insight", "Insight တစ်ခု ဆောက်မယ်"))
        _q_labels = [f"{q['section']} — {q['statement'][:90]}" for q in _ins_qual]
        _qi = 0
        if _q_labels:
            _qi = st.selectbox(T("Pick a Qual finding", "Qual finding ရွေးပါ"), range(len(_q_labels)),
                               format_func=lambda i: _q_labels[i], key="ins_pick_qual")
        _picked = _ins_qual[_qi] if _ins_qual else None
        if _picked:
            st.caption(T("Qual evidence:", "Qual evidence:"))
            for _qq in _picked.get("quotes", []):
                st.write(f"> \"{_qq.get('quote', '')}\" — Speaker {_qq.get('speaker', '')}, {_qq.get('source', '')}")
            if not _picked.get("quotes"):
                st.caption(T("This finding has no verbatim quote attached.", "ဒီ finding မှာ verbatim quote မပါဘူး။"))
        _qnum_labels = [f"{n['label']}: {n['value']} ({n['source']})" for n in _ins_quant]
        _picked_nums = []
        if _qnum_labels:
            _sel_nums = st.multiselect(T("Attach Quant numbers (optional)", "Quant နံပါတ်တွေ တွဲမယ် (မတွဲလည်းရတယ်)"), range(len(_qnum_labels)),
                                       format_func=lambda i: _qnum_labels[i], key="ins_pick_quant")
            _picked_nums = [_ins_quant[i] for i in _sel_nums]
        _sec_opts = _ins_sections if _ins_sections else [T("General", "အထွေထွေ")]
        _sec = st.selectbox(T("DG section it answers", "ဘယ် DG section ကို ဖြေတာလဲ"), _sec_opts, key="ins_section")
        _per_opts = _ins_personas if _ins_personas else [T("All", "အားလုံး")]
        _per = st.selectbox(T("Persona", "Persona"), _per_opts, key="ins_persona")
        _stmt_default = _picked["statement"] if _picked else ""
        _stmt = st.text_area(T("Insight statement (edit freely)", "Insight စာသား (စိတ်ကြိုက် ပြင်လို့ရတယ်)"),
                             value=_stmt_default, height=120, key=f"ins_statement_{_qi}")
        if st.button(T("Save insight", "Insight သိမ်းမယ်"), key="ins_save"):
            if not _stmt.strip():
                st.warning(T("Write the insight statement first.", "Insight စာသား အရင် ရေးပါ။"))
            else:
                _data = load_insights()
                _now = datetime.datetime.now(ZoneInfo("Asia/Rangoon"))
                _data.setdefault("insights", []).append({
                    "id": _now.strftime("%Y%m%d%H%M%S%f"),
                    "statement": _stmt.strip(),
                    "qual_evidence": _picked.get("quotes", []) if _picked else [],
                    "qual_source": _picked.get("analysis_file", "") if _picked else "",
                    "quant_evidence": _picked_nums,
                    "section": _sec, "persona": _per,
                    "created": _now.strftime("%Y-%m-%d %H:%M"),
                })
                save_insights(_data)
                st.success(T("Insight saved.", "Insight သိမ်းပြီးပြီ။"))
    _saved_ins = load_insights().get("insights", [])
    if _saved_ins:
        st.divider()
        st.subheader(T("Saved insights", "သိမ်းထားတဲ့ insight များ"))
        for _ins in _saved_ins:
            st.markdown(f"**{_ins.get('statement', '')}**")
            st.caption(T(f"Section: {_ins.get('section', '')} | Persona: {_ins.get('persona', '')} | {_ins.get('created', '')}",
                         f"Section: {_ins.get('section', '')} | Persona: {_ins.get('persona', '')} | {_ins.get('created', '')}"))
            for _qq in _ins.get("qual_evidence", []):
                st.write(f"> \"{_qq.get('quote', '')}\" — Speaker {_qq.get('speaker', '')}, {_qq.get('source', '')}")
            for _qn in _ins.get("quant_evidence", []):
                st.write(f"- {_qn.get('label', '')}: {_qn.get('value', '')} ({_qn.get('source', '')})")
            if st.button(T("Delete insight", "Insight ဖျက်မယ်"), key=f"ins_del_{_ins.get('id', '')}"):
                _data = load_insights()
                _data["insights"] = [x for x in _data.get("insights", []) if x.get("id") != _ins.get("id")]
                save_insights(_data)
                st.rerun()

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
    # Stage 3 builder (v3.26): respondent form from design["questionnaire"],
    # submit -> notes/responses.csv, interview management + recruitment-list
    # import, seating-plan DOCX, survey data, field notes, F2F discussion,
    # tracking, plus the existing chunk recorder (unchanged, last section).
    # The form runs here behind the login gate; a public link-only route for
    # respondents needs an auth exception and is NOT built in v3.26.
    st.title(T("🎙️ Fieldworks", "🎙️ ကွင်းဆင်းလုပ်ငန်း"))
    _fw_design = load_design()
    _fw = load_fieldwork()
    _fw_counts = fieldwork_counts(_fw_design, _fw)
    _fw_q = _fw_design.get("questionnaire") or {}
    if not isinstance(_fw_q, dict):
        _fw_q = {}
    _fw_qs = _fw_q.get("questions") or []
    if not isinstance(_fw_qs, list):
        _fw_qs = []
    _fw_lang = _fw_design.get("questionnaire_language", "Myanmar")
    _fw_plan = _fw_plan_sessions(_fw_design)
    _fw_sum = [("✅ " if _fw_qs else "○ ") + T("Form", "Form"),
               T(f"Responses: {_fw_counts['responses']}", f"Responses: {_fw_counts['responses']}"),
               T(f"Participants: {_fw_counts['participants']}", f"Participants: {_fw_counts['participants']}"),
               T(f"Notes: {_fw_counts['notes']}", f"Notes: {_fw_counts['notes']}")]
    st.caption(" · ".join(_fw_sum))
    st.info(T("Offline: no internet in the field? Record or take notes first, then fill in and upload when you are back online. Responses submitted here are saved live to Drive (CSV).",
              "Offline: ကွင်းမှာ internet မရှိရင် အရင် record လုပ်ပါ / မှတ်စုရေးပါ၊ online ပြန်ရမှ form ဖြည့် / တင်ပါ။ ဒီမှာ တင်လိုက်တဲ့ response တွေ Drive (CSV) ထဲ တိုက်ရိုက် သိမ်းမယ်။"))

    # ---- 1. Respondent form ----
    st.subheader(T("1. Respondent form", "၁။ ဖြေဆိုသူ form"))
    st.caption(T("This form runs here while you are logged in. A public link-only form for respondents needs a separate login exception — not built in this version.",
                "ဒီ form က login ဝင်ပြီးသား ဒီမှာပဲ သုံးရတယ်။ ဖြေဆိုသူ link တစ်ခုတည်းနဲ့ ဝင်လို့ရတဲ့ form ကတော့ login exception သပ်သပ် လိုတယ် — ဒီ version မှာ မပါသေးဘူး။"))
    if not _fw_qs:
        st.warning(T("No questionnaire saved yet — build one on the Research Design page first. Nothing is invented here.",
                    "Questionnaire မသိမ်းရသေးဘူး — အရင် Research Design page မှာ ဆောက်ပါ။ ဒီမှာ ဘာမှ အလိုအလျောက် မထည့်ဘူး။"))
        if st.button(T("Go to Research Design", "Research Design သွားမယ်"), key="fw_goto_design"):
            st.session_state["page"] = "Research Design"
            st.rerun()
    else:
        _fw_run = st.session_state.get("fw_run", 0)
        if st.session_state.pop("fw_just_submitted", False):
            _ls = st.session_state.get("fw_last_submit") or {}
            _closing = _fw_q.get("closing_mm") if _fw_lang == "Myanmar" else _fw_q.get("closing_en")
            if not _closing:
                _closing = _fw_q.get("closing_en") or _fw_q.get("closing_mm") or ""
            if _ls.get("terminated"):
                st.info((_closing + " " if _closing else "") + T(f"(Screened out at {_ls.get('reason', '')} — the answer was saved.)", f"({_ls.get('reason', '')} မှာ screening ပြီးသွားတယ် — အဖြေ သိမ်းပြီးပြီ။)"))
            else:
                st.success((_closing + " " if _closing else "") + T("Response saved.", "Response သိမ်းပြီးပြီ။"))
        _intro = _fw_q.get("intro_mm") if _fw_lang == "Myanmar" else _fw_q.get("intro_en")
        if not _intro:
            _intro = _fw_q.get("intro_en") or _fw_q.get("intro_mm") or ""
        if _intro.strip():
            st.info(_intro)
        _fw_sess_opts = [""] + [_fw_session_key(s) for s in _fw_plan]
        _fw_sess = st.selectbox(T("Session (optional — which session is this response for?)", "Session (မဖြစ်မနေ မဟုတ် — ဒီ response ဘယ် session အတွက်လဲ)"), _fw_sess_opts,
                                format_func=lambda k: T("(not assigned)", "(မသတ်မှတ်ရသေး)") if not k else next((_fw_session_label(s) for s in _fw_plan if _fw_session_key(s) == k), k),
                                key=f"fw_run{_fw_run}_session")
        _fw_header_vals = {}
        for _h in (_fw_q.get("header_fields") or []):
            _hl = str(_h).strip().lower()
            if "confirmation" in _hl:
                _fw_header_vals[_h] = st.selectbox(str(_h), ["", "Yes", "No"], key=f"fw_run{_fw_run}_hdr_{_xlsform_slug(str(_h))}")
            else:
                _fw_header_vals[_h] = st.text_input(str(_h), key=f"fw_run{_fw_run}_hdr_{_xlsform_slug(str(_h))}")
        _fw_answers = {}
        _fw_cur_section = None
        for _q in _fw_qs:
            if not isinstance(_q, dict):
                continue
            _sec = str(_q.get("section") or "").strip()
            if _sec and _sec != _fw_cur_section:
                _fw_cur_section = _sec
                st.markdown(f"**{_sec}**")
            _qlabel = _fw_question_label(_q, _fw_lang)
            _qcode = str(_q.get("code") or "")
            _qtype = _q.get("type") or "open_text"
            _qname = str(_q.get("name") or _xlsform_slug(_qcode) or "q").strip()
            _opts = _fw_question_options(_q, _fw_lang)
            _opts_en = _fw_question_options(_q, "English")
            _opts_mm = _fw_question_options(_q, "Myanmar")
            _key = f"fw_run{_fw_run}_q_{_qname}"
            _ql = f"{_qcode}. {_qlabel}" if _qcode else _qlabel
            if _q.get("required"):
                _ql += " *"
            _val_en = ""
            _val_mm = ""
            _disp = ""
            _name_part = ""
            _phone_part = ""
            if _qtype == "note":
                st.info(_qlabel)
                continue
            elif _qtype in ("single_select", "observed"):
                if _qtype == "observed":
                    st.caption(T("Observe — do not ask.", "ကြည့်မှတ်ပါ — မမေးပါနဲ့။"))
                if _opts:
                    _sel = st.selectbox(_ql, [""] + _opts, key=_key)
                    if _sel and _sel in _opts:
                        _i = _opts.index(_sel)
                        _val_en = _opts_en[_i] if _i < len(_opts_en) else _sel
                        _val_mm = _opts_mm[_i] if _i < len(_opts_mm) else _sel
                        _disp = _sel
                else:
                    _sel = st.text_input(_ql, key=_key)
                    _val_en = _val_mm = _disp = _sel.strip()
            elif _qtype == "multi_select":
                _sel = st.multiselect(_ql, _opts, key=_key)
                _val_en = "; ".join(_opts_en[_opts.index(o)] if o in _opts and _opts.index(o) < len(_opts_en) else o for o in _sel)
                _val_mm = "; ".join(_opts_mm[_opts.index(o)] if o in _opts and _opts.index(o) < len(_opts_mm) else o for o in _sel)
                _disp = "; ".join(_sel)
            elif _qtype == "ranking_top3":
                _sel = st.multiselect(_ql + T(" (pick the top 3)", " (အကောင်းဆုံး ၃ ခု ရွေးပါ)"), _opts, max_selections=3, key=_key)
                _val_en = "; ".join(_opts_en[_opts.index(o)] if o in _opts and _opts.index(o) < len(_opts_en) else o for o in _sel)
                _val_mm = "; ".join(_opts_mm[_opts.index(o)] if o in _opts and _opts.index(o) < len(_opts_mm) else o for o in _sel)
                _disp = "; ".join(_sel)
            elif _qtype in ("likert_1_5", "rating"):
                _sel = st.selectbox(_ql, ["", "1", "2", "3", "4", "5"], key=_key)
                _val_en = _val_mm = _disp = _sel
            elif _qtype == "numeric":
                _sel = st.text_input(_ql, key=_key)
                _val_en = _val_mm = _disp = _sel.strip()
            elif _qtype == "date":
                _sel = st.text_input(_ql + " (YYYY-MM-DD)", key=_key)
                _val_en = _val_mm = _disp = _sel.strip()
            elif _qtype == "contact":
                _c1, _c2 = st.columns(2)
                _name_part = _c1.text_input(T("Name", "နာမည်"), key=_key + "_name").strip()
                _phone_part = _c2.text_input(T("Phone", "ဖုန်း"), key=_key + "_phone").strip()
                _val_en = _val_mm = f"{_name_part} | {_phone_part}"
                _disp = _val_en if (_name_part or _phone_part) else ""
            else:
                _sel = st.text_area(_ql, key=_key, height=80)
                _val_en = _val_mm = _disp = _sel.strip()
            _fw_answers[_qname] = {"code": _qcode, "q": _q, "en": _val_en, "mm": _val_mm,
                                   "disp": _disp, "name_part": _name_part, "phone_part": _phone_part}
            if _disp and _q.get("outcome"):
                _oc = _resolve_outcome(_q.get("outcome"), _val_en, _val_mm)
                if _oc["terminates"]:
                    st.warning(T(f"This answer ends the screening here ({_qcode}) — submit to finish.", f"ဒီအဖြေနဲ့ screening ဒီမှာပြီးတယ် ({_qcode}) — Submit နှိပ်ပြီး အဆုံးသတ်ပါ။"))
            _hint = _q.get("hint_mm") if _fw_lang == "Myanmar" else _q.get("hint_en")
            if not _hint:
                _hint = _q.get("hint_en") or _q.get("hint_mm") or ""
            if str(_hint).strip():
                st.caption(str(_hint))
        if st.button(T("Submit response", "Response တင်မယ်"), key=f"fw_run{_fw_run}_submit", type="primary"):
            _missing = []
            for _qn, _a in _fw_answers.items():
                if _a["q"].get("required") and not _a["disp"]:
                    _missing.append(_a["code"] or _qn)
            _bad_num = []
            for _qn, _a in _fw_answers.items():
                if _a["q"].get("type") == "numeric" and _a["disp"]:
                    try:
                        float(_a["disp"])
                    except ValueError:
                        _bad_num.append(_a["code"] or _qn)
            if _missing:
                st.error(T(f"Please answer the required questions: {', '.join(_missing)}", f"မဖြစ်မနေ ဖြေရမယ့် မေးခွန်းတွေ ကျန်နေတယ်: {', '.join(_missing)}"))
            elif _bad_num:
                st.error(T(f"Numbers only: {', '.join(_bad_num)}", f"ဂဏန်းပဲ ထည့်ပါ: {', '.join(_bad_num)}"))
            else:
                _terminated = 0
                _term_reason = ""
                _grp = ""
                _quota = ""
                for _qn, _a in _fw_answers.items():
                    if not _a["disp"]:
                        continue
                    _oc = _resolve_outcome(_a["q"].get("outcome"), _a["en"], _a["mm"])
                    if _oc["terminates"] and not _terminated:
                        _terminated = 1
                        _term_reason = f"{_a['code']}: {_a['disp']}"
                    if not _grp and _oc["group"]:
                        _grp = _oc["group"]
                    if not _quota and _oc["quota"]:
                        _quota = _oc["quota"]
                _header = _responses_header(_fw_q)
                _row = {"submitted_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "session": _fw_sess, "terminated": str(_terminated),
                        "terminate_reason": _term_reason, "group_assignment": _grp, "quota_tag": _quota}
                for _h in (_fw_q.get("header_fields") or []):
                    _row["header_" + _xlsform_slug(str(_h))] = _fw_header_vals.get(_h, "")
                for _qn, _a in _fw_answers.items():
                    if _a["q"].get("type") == "contact":
                        _row[_qn + "_name"] = _a.get("name_part", "")
                        _row[_qn + "_phone"] = _a.get("phone_part", "")
                    elif _a["q"].get("type") == "note":
                        continue
                    else:
                        _row[_qn] = _a["en"] if _fw_lang == "English" else _a["mm"]
                append_response_row(_header, _row)
                st.session_state["fw_run"] = _fw_run + 1
                st.session_state["fw_last_submit"] = {"terminated": _terminated, "reason": _term_reason}
                st.session_state["fw_just_submitted"] = True
                st.rerun()

    # ---- 2. Responses ----
    st.subheader(T("2. Responses", "၂။ Responses"))
    _resp_header, _resp_rows = read_responses()
    if not _resp_rows:
        st.caption(T("No responses yet — submit the form above and they appear here live.", "Response မရှိသေးဘူး — အပေါ်က form တင်လိုက်ရင် ဒီမှာ တိုက်ရိုက် ပေါ်မယ်။"))
    else:
        _resp_term = sum(1 for _r in _resp_rows if str(_r.get("terminated")) == "1")
        st.write(T(f"Responses: {len(_resp_rows)} (screened out: {_resp_term})", f"Responses: {len(_resp_rows)} ခု (screening ပြီးသွားတာ: {_resp_term} ခု)"))
        st.dataframe(pd.DataFrame(_resp_rows), use_container_width=True)
        st.download_button(T("Download responses.csv", "responses.csv ဒေါင်းလုပ်ဆွဲမယ်"), data=(NOTES_DIR / FW_RESPONSES_FILE).read_bytes(), file_name=FW_RESPONSES_FILE, key="fw_resp_download")

    # ---- 3. Interview management ----
    st.subheader(T("3. Interview management", "၃။ Interview စီမံခန့်ခွဲခြင်း"))
    if not _fw_plan:
        st.warning(T("No fieldwork plan yet — add sessions on the Research Design page first. A recruitment list can still be imported below; its sheets become separate groups until they are linked to a plan.",
                    "Fieldwork Plan မရှိသေးဘူး — အရင် Research Design page မှာ session တွေ ထည့်ပါ။ Recruitment list ကိုတော့ အောက်မှာ import လုပ်လို့ရတယ်; plan နဲ့ မချိတ်ရသေးတဲ့ sheet တွေက သပ်သပ် group အနေနဲ့ ပေါ်မယ်။"))
    _fw_all_keys = [_fw_session_key(s) for s in _fw_plan]
    _fw_bucket_keys = [k for k in (_fw.get("participants") or {}).keys() if k not in _fw_all_keys]
    _fw_sel_opts = _fw_all_keys + _fw_bucket_keys
    if _fw_sel_opts:
        _fw_sel = st.selectbox(T("Session", "Session"), _fw_sel_opts,
                               format_func=lambda k: next((_fw_session_label(s) for s in _fw_plan if _fw_session_key(s) == k), T(f"{k} (not in fieldwork plan)", f"{k} (Fieldwork Plan ထဲမှာ မရှိဘူး)")),
                               key="fw_im_session")
        _fw_sel_sess = next((s for s in _fw_plan if _fw_session_key(s) == _fw_sel), None)
        _fw_sel_parts = (_fw.get("participants") or {}).get(_fw_sel, [])
        if _fw_sel_sess is not None:
            st.caption(T(f"Seats: {_fw_sel_sess['seats']} · Recruited: {_fw_sel_sess['recruited']} · Confirmed (plan): {_fw_sel_sess['confirmed']} · Moderator: {_fw_sel_sess['moderator'] or '—'} · Venue: {_fw_sel_sess['venue'] or '—'}",
                       f"Seats: {_fw_sel_sess['seats']} · Recruited: {_fw_sel_sess['recruited']} · Confirmed (plan): {_fw_sel_sess['confirmed']} · Moderator: {_fw_sel_sess['moderator'] or '—'} · Venue: {_fw_sel_sess['venue'] or '—'}"))
            _fw_conf_key = "fw_im_confirmed_" + re.sub(r"[^\w]+", "_", _fw_sel)
            _fw_new_conf = st.number_input(T("Confirmed (plan) — how many confirmed to attend", "Confirmed (plan) — လာမယ်လို့ အတည်ပြုသူ အရေအတွက်"), value=int(_fw_sel_sess["confirmed"]), min_value=0, step=1, key=_fw_conf_key)
            if st.button(T("Save confirmed count", "Confirmed အရေအတွက် သိမ်းမယ်"), key="fw_im_conf_save"):
                _d = load_design()
                for _s in (_d.get("fieldwork_plan") or []):
                    if isinstance(_s, dict) and _fw_session_key(_s) == _fw_sel:
                        _s["confirmed"] = int(_fw_new_conf)
                save_design(_d)
                _flash(T("Confirmed count saved.", "Confirmed အရေအတွက် သိမ်းပြီးပြီ။"))
                st.rerun()
            _fw_conf_parts = sum(1 for _p in _fw_sel_parts if isinstance(_p, dict) and _p.get("confirmed"))
            st.caption(T(f"Participants in this session: {len(_fw_sel_parts)} (confirmed: {_fw_conf_parts}) — over-recruitment shows when Recruited is above Seats.",
                       f"ဒီ session မှာ participants: {len(_fw_sel_parts)} ယောက် (confirmed: {_fw_conf_parts} ယောက်) — Recruited က Seats ထက် များနေရင် ပိုခေါ်ထားတာ။"))
        _im_rows = []
        for _p in _fw_sel_parts:
            if not isinstance(_p, dict):
                continue
            _im_rows.append({FW_PARTICIPANT_LABELS[f]: str(_p.get(f, "") or "") for f in FW_PARTICIPANT_FIELDS if f != "confirmed"})
            _im_rows[-1]["Confirmed"] = bool(_p.get("confirmed"))
        _im_df = pd.DataFrame(_im_rows, columns=[FW_PARTICIPANT_LABELS[f] for f in FW_PARTICIPANT_FIELDS])
        _im_key = "fw_im_editor_" + re.sub(r"[^\w]+", "_", _fw_sel)
        _im_edited = st.data_editor(_im_df, num_rows="dynamic", use_container_width=True, key=_im_key,
                                     column_config={"Confirmed": st.column_config.CheckboxColumn(T("Confirmed", "Confirmed"))})
        if st.button(T("Save participants", "Participants သိမ်းမယ်"), key="fw_im_save", type="primary"):
            _new_parts = []
            for _i, _r in enumerate(_im_edited.to_dict("records")):
                _old = _fw_sel_parts[_i] if _i < len(_fw_sel_parts) and isinstance(_fw_sel_parts[_i], dict) else {}
                _nm = str(_r.get("Name") or "").strip()
                _ph = str(_r.get("Phone") or "").strip()
                if not _nm and not _ph and not str(_r.get("Age band") or "").strip():
                    continue
                _np = {"screener": _old.get("screener") or {}}
                for _f in FW_PARTICIPANT_FIELDS:
                    if _f == "confirmed":
                        _np["confirmed"] = bool(_r.get("Confirmed"))
                    else:
                        _np[_f] = str(_r.get(FW_PARTICIPANT_LABELS[_f]) or "").strip()
                _new_parts.append(_np)
            _fw.setdefault("participants", {})[_fw_sel] = _new_parts
            save_fieldwork(_fw)
            _flash(T("Participants saved.", "Participants သိမ်းပြီးပြီ။"))
            st.rerun()
        with st.expander(T("Screener answers (from import)", "Screener အဖြေများ (import ကနေ)")):
            _any_scr = False
            for _p in _fw_sel_parts:
                if isinstance(_p, dict) and _p.get("screener"):
                    _any_scr = True
                    st.write(f"- {_p.get('name') or _p.get('phone') or '?'}: " + "; ".join(f"{k}={v}" for k, v in _p["screener"].items()))
            if not _any_scr:
                st.caption(T("No screener answers stored for this session.", "ဒီ session အတွက် screener အဖြေ မရှိသေးဘူး။"))
    st.markdown("**" + T("Import recruitment list (.xlsx)", "Recruitment list (.xlsx) import လုပ်မယ်") + "**")
    st.caption(T("One sheet per group. Columns: name, gender, age band, marital, education, occupation / job title, income band, phone, township — plus screener answers. Only rows in the file are imported; a row needs a name or a phone.",
                "တစ်ဖွဲ့ကို sheet တစ်ခု။ ကော်လံ: name, gender, age band, marital, education, occupation / job title, income band, phone, township — ပြီးတော့ screener အဖြေတွေ။ File ထဲမှာ ရှိတဲ့ row ပဲ import လုပ်တယ်; row တစ်ခုမှာ name ဒါမှမဟုတ် phone ပါရမယ်။"))
    _fw_up = st.file_uploader(T("Upload a recruitment-list XLSX", "Recruitment-list XLSX တင်ပါ"), type=["xlsx"], key="fw_import")
    if _fw_up is not None and st.button(T("Import", "Import လုပ်မယ်"), key="fw_import_go"):
        _groups, _stats = import_recruitment_xlsx(_fw_up.getvalue())
        if _stats.get("error"):
            st.error(T(f"Could not read the file: {_stats['error']}", f"File ဖတ်မရဘူး: {_stats['error']}"))
        elif not _groups:
            st.warning(T("Nothing imported — no participant rows found in the file.", "ဘာမှ import မရဘူး — file ထဲမှာ participant row မတွေ့ဘူး။"))
        else:
            _plan_by_group = {}
            for _s in _fw_plan:
                _plan_by_group[_s["group"].strip().lower()] = _fw_session_key(_s)
            _added = 0
            _unlinked = []
            for _sheet, _parts in _groups.items():
                _key = _plan_by_group.get(_sheet.strip().lower())
                if _key is None:
                    for _g, _k in _plan_by_group.items():
                        if _g and (_g in _sheet.strip().lower() or _sheet.strip().lower() in _g):
                            _key = _k
                            break
                if _key is None:
                    _key = f"sheet|{_sheet}"
                    _unlinked.append(_sheet)
                _existing = (_fw.get("participants") or {}).get(_key, [])
                _seen = {(str(_p.get("name", "")).lower(), str(_p.get("phone", ""))) for _p in _existing if isinstance(_p, dict)}
                for _p in _parts:
                    _pk = (str(_p.get("name", "")).lower(), str(_p.get("phone", "")))
                    if _pk in _seen:
                        continue
                    _seen.add(_pk)
                    _existing.append(_p)
                    _added += 1
                _fw.setdefault("participants", {})[_key] = _existing
            save_fieldwork(_fw)
            # v3.29: the success AND the unlinked-sheets note below are both
            # wiped by the rerun — queue them together as one flash message.
            _imp_msg = T(f"Imported {_added} participants from {len(_groups)} sheet(s) ({_stats['skipped']} rows skipped).", f"Sheet {len(_groups)} ခုကနေ participants {_added} ယောက် import ပြီးပြီ (row {_stats['skipped']} ခု ကျော်သွားတယ်)။")
            if _unlinked:
                _imp_msg += " " + T(f"Sheets not matching any planned group (kept as separate groups): {', '.join(_unlinked)}", f"Plan နဲ့ မကိုက်တဲ့ sheet တွေ (သပ်သပ် group အနေနဲ့ သိမ်းထားတယ်): {', '.join(_unlinked)}")
            _flash(_imp_msg)
            st.rerun()

    # ---- 4. Seating plan ----
    st.subheader(T("4. Seating plan", "၄။ Seating plan"))
    if not _fw_plan:
        st.caption(T("No fieldwork plan yet — add sessions on the Research Design page first.", "Fieldwork Plan မရှိသေးဘူး — အရင် Research Design page မှာ session တွေ ထည့်ပါ။"))
    else:
        _sp_sel = st.selectbox(T("Session", "Session"), [_fw_session_key(s) for s in _fw_plan],
                               format_func=lambda k: next((_fw_session_label(s) for s in _fw_plan if _fw_session_key(s) == k), k),
                               key="fw_sp_session")
        _sp_sess = next((s for s in _fw_plan if _fw_session_key(s) == _sp_sel), None)
        _sp_parts = (_fw.get("participants") or {}).get(_sp_sel, [])
        _sp_conf = [p for p in _sp_parts if isinstance(p, dict) and p.get("confirmed")]
        st.caption(T(f"Confirmed participants: {len(_sp_conf)} of {len(_sp_parts)} — seats 1–6, spares 7–8. Empty seats stay blank.",
                    f"Confirmed participants: {len(_sp_parts)} ယောက်ထဲက {len(_sp_conf)} ယောက် — ထိုင်ခုံ ၁–၆, spare ၇–၈။ လွတ်နေတဲ့ ခုံက ဗလာပဲ ထားမယ်။"))
        _sp_name = f"seating_plan_session{_sp_sess['session']}.docx"
        if st.button(T("Export seating plan (.docx)", "Seating plan (.docx) ထုတ်မယ်"), key="fw_sp_export"):
            _sp_out = NOTES_DIR / _sp_name
            seating_plan_to_docx(f"{current_project()} — FGD Seating Plan", _sp_sess, _sp_parts, _sp_out)
            store_upload_existing("notes", _sp_name)
            st.success(T("Seating plan exported.", "Seating plan ထုတ်ပြီးပြီ။"))
        if (NOTES_DIR / _sp_name).exists():
            st.download_button(T("Download seating plan (.docx)", "Seating plan (.docx) ဒေါင်းလုပ်ဆွဲမယ်"), data=(NOTES_DIR / _sp_name).read_bytes(), file_name=_sp_name, key="fw_sp_download")

    # ---- 5. Survey data ----
    st.subheader(T("5. Survey data", "၅။ Survey ဒေတာ"))
    _sv_up = st.file_uploader(T("Upload a survey file (CSV / XLSX)", "Survey file တင်ပါ (CSV / XLSX)"), type=["csv", "xlsx"], key="fw_survey_upload")
    if _sv_up is not None and st.button(T("Save survey file", "Survey file သိမ်းမယ်"), key="fw_survey_save"):
        try:
            _sv_df = pd.read_csv(_sv_up) if _sv_up.name.lower().endswith(".csv") else pd.read_excel(_sv_up)
            _sv_safe = re.sub(r"[^\w\-.]+", "_", _sv_up.name).strip("_") or "survey.xlsx"
            store_write_bytes("notes", _sv_safe, _sv_up.getvalue())
            _fw["survey_files"] = [e for e in (_fw.get("survey_files") or []) if e.get("name") != _sv_safe]
            _fw["survey_files"].append({"name": _sv_safe, "rows": int(len(_sv_df)), "columns": int(len(_sv_df.columns)),
                                        "uploaded": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")})
            save_fieldwork(_fw)
            _flash(T(f"Survey file saved: {_sv_safe} ({len(_sv_df)} rows, {len(_sv_df.columns)} columns).", f"Survey file သိမ်းပြီးပြီ: {_sv_safe} ({len(_sv_df)} rows, {len(_sv_df.columns)} columns)။"))
            st.rerun()
        except Exception as _e:
            st.error(T(f"Could not read the file: {_e}", f"File ဖတ်မရဘူး: {_e}"))
    if _fw.get("survey_files"):
        st.dataframe(pd.DataFrame(_fw["survey_files"]), use_container_width=True)
        for _sf in _fw["survey_files"]:
            _sf_path = NOTES_DIR / str(_sf.get("name", ""))
            if _sf_path.exists():
                st.download_button(T(f"Download {_sf.get('name')}", f"{_sf.get('name')} ဒေါင်းလုပ်ဆွဲမယ်"), data=_sf_path.read_bytes(), file_name=str(_sf.get("name")), key=f"fw_sv_dl_{_sf.get('name')}")
    else:
        st.caption(T("No survey files saved yet.", "Survey file မရှိသေးဘူး။"))
    if st.button(T("Open Quantitative Analysis", "Quantitative Analysis ဖွင့်မယ်"), key="fw_goto_quant"):
        st.session_state["page"] = "Quantitative Analysis"
        st.rerun()
    st.caption(T("Analysis itself runs on the Quantitative Analysis page — this section only stores the files and their counts.",
                "ခွဲခြမ်းစိတ်ဖြာမှုက Quantitative Analysis page မှာ လုပ်တယ် — ဒီနေရာက file တွေ သိမ်းပြီး အရေအတွက် ပြပေးရုံပဲ။"))

    # ---- 6. Field notes ----
    st.subheader(T("6. Field notes", "၆။ ကွင်းဆင်းမှတ်စု"))
    _fn_sess_opts = [""] + [_fw_session_key(s) for s in _fw_plan]
    _fn_sess = st.selectbox(T("Session", "Session"), _fn_sess_opts,
                            format_func=lambda k: T("(whole project)", "(project တစ်ခုလုံး)") if not k else next((_fw_session_label(s) for s in _fw_plan if _fw_session_key(s) == k), k),
                            key="fw_fn_session")
    _fn_text = st.text_area(T("Note", "မှတ်စု"), key="fw_fn_text", height=120)
    if st.button(T("Save note", "မှတ်စု သိမ်းမယ်"), key="fw_fn_save", type="primary"):
        if not _fn_text.strip():
            st.warning(T("Nothing to save — the note is empty.", "သိမ်းစရာ မရှိဘူး — မှတ်စု ဗလာဖြစ်နေတယ်။"))
        else:
            _fw.setdefault("notes", []).append({"session": _fn_sess, "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "text": _fn_text.strip()})
            save_fieldwork(_fw)
            st.session_state.pop("fw_fn_text", None)
            _flash(T("Note saved.", "မှတ်စု သိမ်းပြီးပြီ။"))
            st.rerun()
    _fn_list = [n for n in reversed(_fw.get("notes", [])) if not _fn_sess or n.get("session") == _fn_sess]
    if _fn_list:
        for _n in _fn_list:
            _fn_when = str(_n.get("timestamp", ""))
            _fn_sess_label = next((_fw_session_label(s) for s in _fw_plan if _fw_session_key(s) == _n.get("session", "")), "") if _n.get("session") else ""
            st.write(f"- {_fn_when}" + (f" · {_fn_sess_label}" if _fn_sess_label else "") + f" — {str(_n.get('text', ''))}")
    else:
        st.caption(T("No notes yet.", "မှတ်စု မရှိသေးဘူး။"))

    # ---- 7. F2F discussion ----
    st.subheader(T("7. F2F discussion", "၇။ F2F discussion"))
    if not _fw_plan:
        st.caption(T("No fieldwork plan yet — add sessions on the Research Design page first.", "Fieldwork Plan မရှိသေးဘူး — အရင် Research Design page မှာ session တွေ ထည့်ပါ။"))
    else:
        _f2f_sel = st.selectbox(T("Session", "Session"), [_fw_session_key(s) for s in _fw_plan],
                                format_func=lambda k: next((_fw_session_label(s) for s in _fw_plan if _fw_session_key(s) == k), k),
                                key="fw_f2f_session")
        _f2f_sess = next((s for s in _fw_plan if _fw_session_key(s) == _f2f_sel), None)
        _f2f_parts = (_fw.get("participants") or {}).get(_f2f_sel, [])
        _f2f_conf = sum(1 for _p in _f2f_parts if isinstance(_p, dict) and _p.get("confirmed"))
        st.write(T(f"Participants: {len(_f2f_parts)} (confirmed: {_f2f_conf}) · Moderator: {_f2f_sess['moderator'] or '—'} · Venue: {_f2f_sess['venue'] or '—'}",
                   f"Participants: {len(_f2f_parts)} ယောက် (confirmed: {_f2f_conf} ယောက်) · Moderator: {_f2f_sess['moderator'] or '—'} · Venue: {_f2f_sess['venue'] or '—'}"))
        if _f2f_parts:
            st.dataframe(pd.DataFrame([{FW_PARTICIPANT_LABELS[f]: str(_p.get(f, "") or "") for f in FW_PARTICIPANT_FIELDS if f != "confirmed"} for _p in _f2f_parts if isinstance(_p, dict)]), use_container_width=True)
        _dg_path = NOTES_DIR / "discussion_guide.txt"
        with st.expander(T("Discussion Guide", "Discussion Guide")):
            if _dg_path.exists():
                st.text(_dg_path.read_text(encoding="utf-8"))
            else:
                st.caption(T("No Discussion Guide saved yet — draft one on the Research Design page.", "Discussion Guide မသိမ်းရသေးဘူး — Research Design page မှာ ရေးပါ။"))
        with st.expander(T("Session notes", "Session မှတ်စုများ")):
            _f2f_notes = [n for n in (_fw.get("notes") or []) if n.get("session") == _f2f_sel]
            if _f2f_notes:
                for _n in _f2f_notes:
                    st.write(f"- {str(_n.get('timestamp', ''))} — {str(_n.get('text', ''))}")
            else:
                st.caption(T("No notes for this session yet.", "ဒီ session အတွက် မှတ်စု မရှိသေးဘူး။"))
        if st.button(T("Open Transcripts", "Transcripts ဖွင့်မယ်"), key="fw_goto_transcripts"):
            st.session_state["page"] = "Transcripts"
            st.rerun()

    # ---- 8. Tracking ----
    st.subheader(T("8. Tracking", "၈။ Tracking"))
    if not _fw_plan:
        st.caption(T("No fieldwork plan yet — tracking appears once sessions exist.", "Fieldwork Plan မရှိသေးဘူး — session ရှိလာရင် tracking ဒီမှာ ပေါ်မယ်။"))
    else:
        _resp_by_sess = {}
        for _r in _resp_rows:
            _resp_by_sess[str(_r.get("session", ""))] = _resp_by_sess.get(str(_r.get("session", "")), 0) + 1
        _trk_rows = []
        for _s in _fw_plan:
            _k = _fw_session_key(_s)
            _parts = (_fw.get("participants") or {}).get(_k, [])
            _trk_rows.append({"Session": _s["session"], "Group": _s["group"], "Date": _s["date"],
                             "Seats": _s["seats"], "Recruited": _s["recruited"],
                             "Confirmed (plan)": _s["confirmed"],
                             "Confirmed participants": sum(1 for _p in _parts if isinstance(_p, dict) and _p.get("confirmed")),
                             "Responses": _resp_by_sess.get(_k, 0),
                             "Status": (_fw.get("session_status") or {}).get(_k, "Not started")})
        st.dataframe(pd.DataFrame(_trk_rows), use_container_width=True)
        _trk_sel = st.selectbox(T("Session", "Session"), [_fw_session_key(s) for s in _fw_plan],
                                format_func=lambda k: next((_fw_session_label(s) for s in _fw_plan if _fw_session_key(s) == k), k),
                                key="fw_trk_session")
        _trk_cur = (_fw.get("session_status") or {}).get(_trk_sel, "Not started")
        _trk_new = st.selectbox(T("Status", "Status"), FW_SESSION_STATUSES,
                                index=FW_SESSION_STATUSES.index(_trk_cur) if _trk_cur in FW_SESSION_STATUSES else 0,
                                key="fw_trk_status")
        if st.button(T("Save status", "Status သိမ်းမယ်"), key="fw_trk_save"):
            _fw.setdefault("session_status", {})[_trk_sel] = _trk_new
            save_fieldwork(_fw)
            _flash(T("Status saved.", "Status သိမ်းပြီးပြီ။"))
            st.rerun()
        st.caption(T(f"Responses: {_fw_counts['responses']} · Participants: {_fw_counts['participants']} · Notes: {_fw_counts['notes']} · Survey files: {_fw_counts['survey_files']}",
                    f"Responses: {_fw_counts['responses']} · Participants: {_fw_counts['participants']} · Notes: {_fw_counts['notes']} · Survey files: {_fw_counts['survey_files']}"))

    # ---- 9. Record (existing chunk recorder — unchanged) ----
    st.subheader(T("9. Record", "၉။ အသံသွင်းခြင်း"))
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
            # v3.29: both messages below are wiped by the rerun — collect
            # them and flash once (the error case flashes as an error).
            _rec_msgs = [f"Saved: {chunk_path.name}"]
            txt_path = OUTPUT_DIR / (chunk_path.stem + ".txt")
            with st.spinner(T(f"Transcribing chunk {next_n}... (long audio can take a few minutes)", f"Chunk {next_n} ကို transcribing... (audio ကြာရင် မိနစ်အနည်းငယ် ကြာနိုင်တယ်)")):
                text, model = transcribe_audio(chunk_path)
            if text:
                store_write_text("transcripts", txt_path.name, text)
                _rec_msgs.append(f"Done — model: {model}")
                _flash(" ".join(_rec_msgs))
            else:
                _flash(T(f"Failed: {model} — retry with Transcribe pending below.", f"Failed: {model} — အောက်က Transcribe pending နဲ့ ပြန်လုပ်လို့ရတယ်။"), kind="error")
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
            _ec_stem = saved_selected[:-4] if saved_selected.endswith(".txt") else saved_selected
            _ec_path = NOTES_DIR / f"evidence_check_{_ec_stem}.txt"
            if _ec_path.exists():
                _ec_lines = _ec_path.read_text(encoding="utf-8", errors="replace").splitlines()
                _ec_score = _ec_lines[1] if len(_ec_lines) > 1 else ""
                st.caption("Evidence: " + _ec_score)
            else:
                st.caption(T("Evidence: not yet verified.", "Evidence: မစစ်ရသေးဘူး။"))
            if st.button(T("🔍 Verify evidence", "🔍 Evidence စစ်မယ်"), key="ma_verify_evidence"):
                with st.spinner(T("Verifying quotes against transcripts... (no AI call)", "Transcript တွေနဲ့ quote တွေ တိုက်စစ်နေတယ်... (AI မသုံးဘူး)")):
                    _ev_results, _ev_score, _ev_path = run_evidence_check(saved_selected)
                st.success(T(f"Verified: {_ev_score[0]} of {_ev_score[1]} findings",
                             f"စစ်ပြီးပြီ: finding {_ev_score[1]} ခု ထဲက {_ev_score[0]} ခု အတည်ပြုနိုင်တယ်"))
                _ev_rows = [{
                    "Finding": r["finding"][:120],
                    "Quote": r["quote"][:120],
                    "Source": r["source"],
                    "Claimed n": r["claimed_n"] if r["claimed_n"] is not None else "",
                    "Found n": r["found_n"],
                    "Status": r["status"],
                    "Reason": r["reason"],
                } for r in _ev_results]
                st.dataframe(pd.DataFrame(_ev_rows))
                st.download_button(
                    "Download evidence check (.txt)",
                    data=_ev_path.read_bytes(),
                    file_name=_ev_path.name,
                    mime="text/plain",
                    key="ma_dl_evidence",
                )
        else:
            st.info(T("No saved analysis yet.", "သိမ်းထားတဲ့ analysis မရှိသေးပါ။"))

        st.divider()
        st.subheader(T("🔀 Cross-interview themes", "🔀 Theme တွေ — transcript တိုင်းမှာ ပါမပါ"))
        _th_themes, _th_files = cross_interview_themes()
        if not _th_themes:
            st.info(T("No single-transcript analyses yet — analyze transcripts one by one first.",
                      "Transcript တစ်ခုချင်းစီ analyse လုပ်ထားတာ မရှိသေးဘူး — အရင် တစ်ခုချင်းစီ analyse လုပ်ပါ။"))
        else:
            _th_rows = {}
            for _th, _counts in _th_themes.items():
                _th_rows[_th] = {f: (f"✓ {_counts[f]}" if _counts.get(f) else "—") for f in _th_files}
            st.dataframe(pd.DataFrame.from_dict(_th_rows, orient="index"))
            st.caption(T("✓ n = findings in that theme for that transcript. Multi-interview analyses are not listed here — they already combine interviews.",
                         "✓ n = အဲဒီ transcript မှာ အဲဒီ theme အောက် finding အရေအတွက်။ Multi-interview analysis တွေ ဒီမှာ မပါဘူး — သူတို့က interview တွေ ပေါင်းပြီးသား။"))

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
    # v3.28: page-level brief + design — the export wiring above and the
    # Proposal (draft) below both use them.
    _pbrief = load_brief()
    _pdesign = load_design()

    saved = sorted(ANALYSIS_DIR.glob("*_analysis.txt"), key=lambda f: f.name)
    if not saved:
        st.info(T("No analysis yet — Analyze on the Market Analysis page first.", "Analysis မရှိသေးပါ — အရင် Market Analysis မှာ Analyze လုပ်ပါ။"))
    else:
        saved_names = [f.name for f in saved]
        selected = st.selectbox(T("Pick an analysis", "Analysis ရွေးပါ"), saved_names)
        template = st.selectbox(T("Pick a template", "Template ရွေးပါ"), list(TEMPLATES.keys()))
        _stem = selected[:-len("_analysis.txt")] if selected.endswith("_analysis.txt") else selected
        _analysis_text = (ANALYSIS_DIR / selected).read_text(encoding="utf-8")
        if looks_like_html(_analysis_text):
            _analysis_text = html_to_text(_analysis_text)
        title_text = selected.replace("_analysis.txt", "").replace("_", " ")
        # v3.28: PPTX is the main deliverable — it uses the assembled report
        # structure template (Qual or Quant, picked from the design method);
        # DOCX stays topline/summary — the original analysis + Executive
        # Summary. The summary comes from the saved topline when it exists,
        # else from the analysis's Key Findings.
        _insights = load_insights()
        _struct = _pick_report_template(_pdesign)
        if _struct == "quant":
            _assembled_text, _rep_included, _rep_missing = _assemble_quant_report(_analysis_text, _pbrief, _pdesign, _insights, _stem)
            _qn = _quant_total_n(_load_quant_texts())
            _pptx_footer = f"Base: N={_qn}" if _qn else None
            st.caption(T("Structure template: Quantitative report — picked from the design method.", "ပုံစံ: Quantitative report — design method ကနေ ရွေးထားတာ။"))
        else:
            _assembled_text, _rep_included, _rep_missing = _assemble_report(_analysis_text, _pbrief, _pdesign, _insights)
            _pptx_footer = None
            st.caption(T("Structure template: Qualitative report — picked from the design method.", "ပုံစံ: Qualitative report — design method ကနေ ရွေးထားတာ။"))
        _pptx_text, _summary_src = _report_text_with_summary(_assembled_text, _stem, _analysis_text)
        _docx_text, _ = _report_text_with_summary(_analysis_text, _stem)
        if _struct == "quant" and _summary_src == "existing":
            st.caption(T("Executive Summary is in the template order, with KPI cards from the survey data.", "Executive Summary က ပုံစံအစီအစဉ်အတိုင်း ပါတယ်, survey data က KPI cards နဲ့အတူ။"))
        elif _summary_src == "topline":
            st.caption(T("Executive Summary at the top is from the saved topline.", "အပေါ်က Executive Summary က သိမ်းထားတဲ့ topline ကနေ ယူထားတာ။"))
        elif _summary_src == "key_findings":
            st.caption(T("Executive Summary at the top is from the analysis Key Findings.", "အပေါ်က Executive Summary က analysis ရဲ့ Key Findings ကနေ ယူထားတာ။"))
        elif _summary_src == "existing":
            st.caption(T("This analysis already has an Executive Summary section.", "ဒီ analysis မှာ Executive Summary အပိုင်း ရှိနေပြီးသားမို့ ထပ်မထည့်ဘူး။"))
        else:
            st.caption(T("No Executive Summary available — create a topline below, or add a Key Findings section.", "Executive Summary မရှိသေးဘူး — အောက်မှာ topline လုပ်ပါ၊ ဒါမှမဟုတ် Key Findings အပိုင်း ထည့်ပါ။"))

        # v3.28: pre-export checks — warnings only, exports are never blocked.
        st.subheader(T("Pre-export checks", "ထုတ်ခင် စစ်ဆေးခြင်း"))
        if st.button(T("🔍 Run pre-export checks", "🔍 ထုတ်ခင် စစ်မယ်"), key="report_precheck_btn"):
            _en_issues = report_english_check(_pptx_text)
            _unverified, _ev_total, _ev_note = report_evidence_warnings(_analysis_text)
            st.session_state["report_precheck"] = {"en": _en_issues, "unverified": _unverified, "total": _ev_total, "note": _ev_note, "stem": _stem}
        _pc = st.session_state.get("report_precheck") or {}
        if _pc.get("stem") == _stem:
            if _pc.get("en"):
                _en_list = "; ".join(f"line {n}: {t[:80]}" for n, t in _pc["en"][:5])
                st.warning(T(f"English-only check: {len(_pc['en'])} non-English line(s) found.", f"English-only စစ်ဆေးမှု: English မဟုတ်တဲ့ စာကြောင်း {len(_pc['en'])} ကြောင်း တွေ့တယ်။") + " " + _en_list)
            else:
                st.success(T("English-only check: no non-English lines.", "English-only စစ်ဆေးမှု: English မဟုတ်တဲ့ စာကြောင်း မရှိဘူး။"))
            if _pc.get("note"):
                st.info(T(_pc["note"], "စစ်ဆေးဖို့ transcript မရှိသေးဘူး။"))
            elif _pc.get("total"):
                _uv = _pc["unverified"]
                if _uv:
                    _uv_list = "; ".join(r["finding"][:80] for r in _uv[:5])
                    st.warning(T(f"Evidence check: {len(_uv)} of {_pc['total']} finding(s) UNVERIFIED.", f"Evidence စစ်ဆေးမှု: finding {_pc['total']} ခုမှာ {len(_uv)} ခု အတည်မပြုနိုင်သေးဘူး။") + " " + _uv_list)
                else:
                    st.success(T(f"Evidence check: all {_pc['total']} finding(s) verified.", f"Evidence စစ်ဆေးမှု: finding {_pc['total']} ခု လုံး အတည်ပြုပြီးပြီ။"))

        # PPTX title slide: client + export month, like the sample title slides.
        _pptx_subtitle = "Market Research AI — ResearchFlow"
        if _pbrief.get("client_name"):
            _pptx_subtitle += f" · {_pbrief['client_name']}"
        _pptx_subtitle += f" · {datetime.date.today().strftime('%b %Y')}"
        col_pptx, col_docx = st.columns(2)
        if col_pptx.button("Export PPTX"):
            out_path = REPORTS_DIR / selected.replace("_analysis.txt", ".pptx")
            with st.spinner(T("Exporting PPTX...", "PPTX ထုတ်နေပါတယ်...")):
                analysis_to_pptx(_pptx_text, title_text, out_path, template,
                                 subtitle=_pptx_subtitle, footer=_pptx_footer)
            store_upload_existing("reports", out_path.name)
            st.success(f"Done — {out_path.name} ({template})")
            st.download_button(
                "Download PPTX",
                data=out_path.read_bytes(),
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        if col_docx.button("Export DOCX"):
            out_path = REPORTS_DIR / selected.replace("_analysis.txt", ".docx")
            with st.spinner("DOCX ထုတ်နေပါတယ်..."):
                analysis_to_docx(_docx_text, title_text, out_path)
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
            if report_selected.endswith(".pptx"):
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            else:
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            st.download_button(
                "Download saved report",
                data=(REPORTS_DIR / report_selected).read_bytes(),
                file_name=report_selected,
                mime=mime,
            )
        else:
            st.info(T("No saved report yet.", "သိမ်းထားတဲ့ report မရှိသေးပါ။"))

        st.divider()
        with st.expander(T(f"Report structure template ({'Quant' if _struct == 'quant' else 'Qual'})", f"Report ပုံစံ ({'Quant' if _struct == 'quant' else 'Qual'})")):
            if _struct == "quant":
                st.caption(T("Template order: Actionable Insights → Timeline → Overview → Executive Summary → Question Analysis → VOC → Satisfaction → Tier Evaluation → Sentiment → Action Plan.", "ပုံစံ အစီအစဉ်: Actionable Insights → Timeline → Overview → Executive Summary → Question Analysis → VOC → Satisfaction → Tier Evaluation → Sentiment → Action Plan။"))
                _rows = report_template_coverage_quant(_analysis_text, _pbrief, _pdesign, _insights, _stem)
            else:
                st.caption(T("Template order: Goal → Personas → Deep Dive per segment → Strategy → Action Plan.", "ပုံစံ အစီအစဉ်: Goal → Personas → Deep Dive → Strategy → Action Plan။"))
                _rows = report_template_coverage(_analysis_text, _pbrief, _pdesign, _insights)
            if not _rows:
                st.info(T("This report template does not apply to this project's method.", "ဒီ report ပုံစံက ဒီ project ရဲ့ method နဲ့ မကိုက်ဘူး။"))
            else:
                for _row in _rows:
                    st.write(("✅ " if _row["present"] else "○ ") + _row["label"]
                             + (f" — {_row['detail']}" if _row.get("detail") else ""))
            if _rep_missing:
                st.caption(T("Sections with no source data are left out of the export — add them in Google Docs/PowerPoint.", "Source data မရှိတဲ့ အပိုင်းတွေ export မှာ မပါဘူး — Google Docs/PowerPoint မှာ ဖြည့်ပါ။"))
            st.caption(T("The template covers structure only. Strategic synthesis, archetype framing, images and layout still need your manual refinement.", "ပုံစံက structure ပဲ ပါတယ်။ Strategic synthesis, archetype framing, ပုံ နဲ့ layout က ကိုယ်တိုင် ပြင်ရဦးမယ်။"))

    # v3.28: Proposal (draft) — needs brief + design only, no analysis.
    st.divider()
    st.subheader(T("Proposal (draft)", "Proposal (မူကြမ်း)"))
    st.caption(T("Draft only — check and edit before sending. Costs are always blank.", "မူကြမ်းပဲ — ပို့ခင် စစ်ပြီး ပြင်ပါ။ Costs က အမြဲ ဗလာပဲ။"))
    _pblocks, _pmissing = proposal_build(_pbrief, _pdesign)
    if _pmissing:
        st.info(T("To draft a proposal, save first: " + " + ".join(_pmissing) + ".", "Proposal ရေးဖို့ အရင် သိမ်းပါ: " + " + ".join(_pmissing) + "။"))
        _pc1, _pc2 = st.columns(2)
        if _pc1.button(T("Go to Client Brief", "Client Brief သွားမယ်"), key="prop_go_brief"):
            st.session_state["page"] = "Client Brief"
            st.rerun()
        if _pc2.button(T("Go to Research Design", "Research Design သွားမယ်"), key="prop_go_design"):
            st.session_state["page"] = "Research Design"
            st.rerun()
    else:
        _rat = st.text_area(T("Rationale — why this method", "Rationale — ဘာကြောင့် ဒီ method လဲ"), value=_pdesign.get("rationale", ""), height=100, key="proposal_rationale")
        if st.button(T("Generate proposal (DOCX)", "Proposal (DOCX) ထုတ်မယ်"), key="proposal_generate"):
            _pdesign["rationale"] = _rat.strip()
            save_design(_pdesign)
            _pblocks, _ = proposal_build(_pbrief, _pdesign)
            _prop_bytes = proposal_to_docx(_pblocks)
            _prop_path = store_write_bytes("notes", "proposal.docx", _prop_bytes)
            st.success(T(f"Done — {_prop_path.name} (draft, costs blank)", f"ပြီးပြီ — {_prop_path.name} (မူကြမ်း, costs ဗလာ)"))
            st.download_button(
                "Download proposal (DOCX)",
                data=_prop_bytes,
                file_name=_prop_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

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
