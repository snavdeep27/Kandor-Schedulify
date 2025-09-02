# Kandor Schedulify — mobile-polished (light theme, top-right signout, responsive calendar)
# ----------------------------------------------------------------------------------------
# ENV: MONGO_URI, MONGO_DB_NAME (opt), MS_CLIENT_ID, MS_CLIENT_SECRET,
#      MS_AUTHORITY, MS_REDIRECT_URI, BASE_URL, LOGO_PATH (opt)
# Run: streamlit run app.py

import os
import re
import uuid
import time
import calendar
import datetime as dt
from typing import List, Tuple, Optional
from zoneinfo import ZoneInfo

import streamlit as st
from streamlit.components.v1 import html as st_html
from pymongo import MongoClient
from pymongo.errors import ConfigurationError
from dotenv import load_dotenv
import requests
import msal

# ---------- Page ----------
st.set_page_config(page_title="Kandor Schedulify", page_icon="🗓️", layout="wide")

# ---------- Hard-force LIGHT theme & mobile layout fixes ----------
st.markdown("""
<style>
:root { color-scheme: light !important; }
html, body, .stApp { background:#ffffff !important; color:#111827 !important; }

/* Reduce the Streamlit top padding so UI isn't hidden under Streamlit Cloud header */
.block-container { padding-top: 20px !important; padding-bottom: 20px !important; }

/* Top bar: logo + title left, Sign out pinned to the right */
.topbar { display:flex; align-items:center; justify-content:space-between;
          gap:12px; padding:6px 4px 0 4px; }
.topbar-left { display:flex; align-items:center; gap:10px; min-width:0; }
.appname { font-weight:800; letter-spacing:-0.02em; font-size:20px; white-space:nowrap; }
.topbar-actions { margin-left:auto; display:flex; align-items:center; }

/* Make every Streamlit button a tidy pill */
.stButton > button {
  border-radius: 999px !important;
  border: 1px solid #E5E7EB !important;
  background:#f8fafc !important;
  color:#111827 !important;
  padding: 8px 16px !important;
}
.stButton > button:hover { background:#eef2ff !important; border-color:#c7ccff !important; }

/* ---- Booking calendar ---- */
.cal-wrap { margin-top: 6px; }
.cal-navrow { display:grid; grid-template-columns: 42px 1fr 42px; gap:10px; align-items:center; }
.cal-navbtn { border:1px solid #E5E7EB; background:#F8FAFC; border-radius:999px; padding:10px 0; text-align:center; font-weight:700; }
.cal-month { text-align:center; font-weight:800; font-size:1.05rem; }

/* Day-of-week header and grid share the same 7-col grid */
.cal-dow, .cal-grid { display:grid; grid-template-columns: repeat(7, 1fr); gap:8px; }
.cal-dow div { text-align:center; font-size:12px; color:#6B7280; padding:2px 0; }

/* Day cell */
.cal-day { text-align:center; padding:10px 0; border-radius: 999px;
           border:1px solid #E5E7EB; background:#ffffff; color:#111827; user-select:none; }
.cal-day.today { border-color:#c7ccff; background:#f5f7ff; }
.cal-day.sel { background:#4f46e5; color:#ffffff; border-color:#4f46e5; }
.cal-day.dim { color:#9CA3AF; background:#F9FAFB; border-color:#F1F5F9; }
.cal-day.btn { cursor:pointer; }
.cal-day.btn:hover { border-color:#9aa2ff; background:#eef2ff; }

/* Slots column pills inherit Streamlit button style; make them denser on phones */
@media (max-width: 600px){
  .stButton > button { padding: 8px 12px !important; }
}

/* Prevent Streamlit's column min-width from breaking 7-col grids on mobile */
.cal-scope [data-testid="column"] { min-width: 0 !important; flex: 1 1 0 !important; padding: 0 4px !important; }

/* Cards (how-to etc.) sit on white like desktop */
.howto-card{ background:#fff; border:1px solid rgba(0,0,0,.06); border-radius:14px; padding:14px 16px; box-shadow:0 4px 14px rgba(0,0,0,.04); }

/* Booking link card */
.link-card {
  display:flex; align-items:center; gap:10px; padding:12px 14px;
  border-radius:999px; border:1px solid #e5e7eb; background:#f8fafc;
}
.link-card input { border:none; background:transparent; outline:none; width:100%;
  font-family: ui-monospace, Menlo, monospace; font-size:14px; color:#111827; }
.link-card button{
  padding:10px 16px; border-radius:999px; border:none; color:#fff; font-weight:800; letter-spacing:.2px; cursor:pointer;
  background:linear-gradient(90deg,#6366f1 0%, #a855f7 100%);
  box-shadow:0 8px 18px rgba(99,102,241,.25);
}
.link-card button:hover{ filter:brightness(1.05); }

/* Ensure Streamlit dark header overlay never tints our pages */
[data-testid="stHeader"] { background:transparent !important; }
</style>
""", unsafe_allow_html=True)

# ---------- ENV ----------
load_dotenv()
MONGO_URI = os.getenv("MONGO_URI")
MONGO_DB_NAME = os.getenv("MONGO_DB_NAME") or None
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
MS_AUTHORITY = os.getenv("MS_AUTHORITY", "https://login.microsoftonline.com/common")
MS_REDIRECT_URI = os.getenv("MS_REDIRECT_URI", "http://localhost:8501")
BASE_URL = os.getenv("BASE_URL", "http://localhost:8501").rstrip("/")
LOGO_PATH = os.getenv("LOGO_PATH", "assets/kandor-logo.png")
MS_SCOPE = ["User.Read", "Calendars.ReadWrite", "Mail.Send"]
GRAPH = "https://graph.microsoft.com/v1.0"

# ---------- Policy ----------
BUSY_STATUSES = {"busy", "oof", "workingelsewhere", "tentative"}
MIN_LEAD_MINUTES = 5
RECHECK_EACH_SLOT = True

# ---------- Mongo ----------
@st.cache_resource
def _mongo():
    client = MongoClient(MONGO_URI); client.admin.command("ping"); return client

def _db():
    c = _mongo()
    if MONGO_DB_NAME: return c[MONGO_DB_NAME]
    try:
        d = c.get_default_database()
        if d is not None: return d
    except ConfigurationError:
        pass
    return c["schedulify"]

def users_col(): return _db().users
def flows_col(): return _db().auth_flows

# ---------- Helpers ----------
def slugify_email(email: str, fallback: str) -> str:
    base = (email.split("@")[0] if email else fallback).lower()
    base = re.sub(r"[^a-z0-9\-]+", "-", base).strip("-")
    base = re.sub(r"-{2,}", "-", base)
    return base or fallback.lower()

def ensure_unique_slug(base: str, oid: str) -> str:
    existing = users_col().find_one({"slug": base})
    if existing and existing.get("oid") != oid:
        return f"{base}-{oid[-6:].lower()}"
    return base

# ---------- MSAL ----------
def build_msal_app(msal_cache: Optional[msal.SerializableTokenCache] = None):
    return msal.ConfidentialClientApplication(
        MS_CLIENT_ID, authority=MS_AUTHORITY,
        client_credential=MS_CLIENT_SECRET, token_cache=msal_cache
    )

def persist_cache_for_user(user_key: dict, cache):
    try:
        changed = cache.has_state_changed  # type: ignore[attr-defined]
    except Exception:
        changed = True
    if changed:
        try: users_col().update_one(user_key, {"$set": {"msal_cache": cache.serialize()}})
        except Exception: pass

def create_auth_url():
    app = build_msal_app()
    flow = app.initiate_auth_code_flow(scopes=MS_SCOPE, redirect_uri=MS_REDIRECT_URI)
    flows_col().insert_one({"_id": flow["state"], "flow": flow, "created_utc": dt.datetime.utcnow()})
    return flow["auth_uri"]

def finish_auth_redirect():
    state = st.query_params.get("state"); code = st.query_params.get("code")
    if not state or not code: st.error("Missing authorization response parameters."); return
    doc = flows_col().find_one({"_id": state})
    if not doc: st.error("Sign-in session expired. Please try again."); return
    flow = doc["flow"]

    cache = msal.SerializableTokenCache()
    app = build_msal_app(cache)
    try:
        result = app.acquire_token_by_auth_code_flow(flow, st.query_params.to_dict())
    except Exception as e:
        st.error(f"Failed to complete sign-in: {e}"); return
    finally:
        flows_col().delete_one({"_id": state})

    if "access_token" not in result:
        st.error(result.get("error_description", "Could not obtain access token.")); return

    claims = result.get("id_token_claims", {}) or {}
    oid = claims.get("oid") or claims.get("sub")
    email = claims.get("preferred_username") or ""
    name = claims.get("name") or (email.split("@")[0] if email else "User")
    if not oid: st.error("Missing Microsoft account ID (oid)."); return

    u = users_col().find_one({"oid": oid})
    if not u:
        base_slug = slugify_email(email, name); slug = ensure_unique_slug(base_slug, oid)
        users_col().insert_one({
            "oid": oid, "email": email, "name": name, "slug": slug,
            "zoom_link": "", "meeting_duration": 30,
            "available_days": ["Monday","Tuesday","Wednesday","Thursday","Friday"],
            "start_time": "09:00", "end_time": "17:00", "timezone": "Asia/Kolkata",
        })
    else:
        slug = u.get("slug") or ensure_unique_slug(slugify_email(u.get("email",""), u.get("name","user")), oid)
        users_col().update_one({"oid": oid}, {"$set": {"slug": slug, "email": email or u.get("email",""), "name": name or u.get("name","")}})

    users_col().update_one({"oid": oid}, {"$set": {"msal_cache": cache.serialize()}})
    st.session_state["oid"] = oid
    st.success("Signed in with Outlook!")
    time.sleep(0.4)
    st.query_params.clear(); st.query_params["page"] = "dashboard"; st.rerun()

def get_user_by_slug_or_email(value: str):
    u = users_col().find_one({"slug": value})
    if not u and "@" in value: u = users_col().find_one({"email": value})
    return u

def get_access_token_for_user_doc(user_doc) -> Optional[str]:
    cache = msal.SerializableTokenCache()
    serialized = user_doc.get("msal_cache")
    if serialized:
        try: cache.deserialize(serialized)
        except Exception: pass
    app = build_msal_app(cache)
    accounts = app.get_accounts()
    token_result = app.acquire_token_silent(MS_SCOPE, account=accounts[0]) if accounts else None
    if token_result and "access_token" in token_result:
        persist_cache_for_user({"oid": user_doc["oid"]}, cache)
        return token_result["access_token"]
    return None

# ---------- Graph ----------
def graph_headers(token: str, tzname: Optional[str] = None):
    h = {"Authorization": f"Bearer {token}"}; h["Prefer"] = f'outlook.timezone="{tzname or "UTC"}"'; return h

def graph_day_view(token: str, start_utc: dt.datetime, end_utc: dt.datetime, tzname: str) -> List[dict]:
    params = {"startDateTime": start_utc.isoformat(), "endDateTime": end_utc.isoformat(),
              "$select": "subject,start,end,showAs,isCancelled", "$orderby": "start/dateTime ASC"}
    r = requests.get(f"{GRAPH}/me/calendarView", headers=graph_headers(token, tzname), params=params, timeout=20)
    if r.status_code == 200: return r.json().get("value", [])
    st.warning(f"Graph calendarView failed ({r.status_code}). Treating day as free."); return []

def is_interval_free(token: str, start_local: dt.datetime, end_local: dt.datetime, tzname: str) -> bool:
    start_utc = start_local.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
    end_utc   = end_local.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
    for ev in graph_day_view(token, start_utc, end_utc, tzname):
        if ev.get("isCancelled"): continue
        if (ev.get("showAs") or "busy").lower() in BUSY_STATUSES: return False
    return True

def graph_create_event(token: str, subject: str, html_body: str,
                       start_local: dt.datetime, end_local: dt.datetime,
                       attendees: List[str], timezone_name: str):
    payload = {
        "subject": subject, "showAs": "busy", "transactionId": str(uuid.uuid4()),
        "body": {"contentType": "HTML", "content": html_body},
        "start": {"dateTime": start_local.isoformat(), "timeZone": timezone_name},
        "end": {"dateTime": end_local.isoformat(), "timeZone": timezone_name},
        "attendees": [{"emailAddress": {"address": e}, "type": "required"} for e in attendees]
    }
    r = requests.post(f"{GRAPH}/me/events", headers=graph_headers(token, timezone_name), json=payload, timeout=20)
    return (r.status_code in (200, 201), (r.json() if r.headers.get("content-type","").startswith("application/json") else {}))

def graph_send_mail(token: str, to_addr: str, subject: str, html_body: str, tzname: str) -> bool:
    url = f"{GRAPH}/me/sendMail"
    payload = {"message": {"subject": subject, "body": {"contentType": "HTML", "content": html_body},
                           "toRecipients": [{"emailAddress": {"address": to_addr}}]}, "saveToSentItems": True}
    r = requests.post(url, headers=graph_headers(token, tzname), json=payload, timeout=20)
    return r.status_code in (202, 200)

# ---------- Slot math ----------
def overlap(a_start: dt.time, a_end: dt.time, b_start: dt.time, b_end: dt.time) -> bool:
    return max(a_start, b_start) < min(a_end, b_end)

def build_slots(day: dt.date, work_start: dt.time, work_end: dt.time,
                meeting_minutes: int, busy: List[Tuple[dt.time, dt.time]]) -> List[dt.time]:
    step = dt.timedelta(minutes=meeting_minutes)
    t = dt.datetime.combine(day, work_start); end = dt.datetime.combine(day, work_end)
    out = []
    while t + step <= end:
        s = t.time(); e = (t + step).time()
        if not any(overlap(s, e, bs, be) for bs, be in busy): out.append(s)
        t += step
    return out

# ---------- Header ----------
def topbar():
    c1, c2 = st.columns([7, 5])
    with c1:
        st.markdown('<div class="topbar"><div class="topbar-left">', unsafe_allow_html=True)
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=32)
        st.markdown('<div class="appname">Kandor Schedulify</div></div>', unsafe_allow_html=True)
        st.markdown('<div class="topbar-actions">', unsafe_allow_html=True)
        st.markdown('</div></div>', unsafe_allow_html=True)
    with c2:
        right = st.container()
        with right:
            r1, r2 = st.columns([4,1])
            with r2:
                if "oid" in st.session_state:
                    if st.button("Sign out", key="signout_btn"):
                        st.session_state.pop("oid", None)
                        st.success("Signed out.")
                        time.sleep(0.3)
                        st.query_params.clear(); st.rerun()

# ---------- Landing ----------
def landing():
    topbar()
    st.markdown("### Your Personal Calendly Clone by Kandor")
    st.caption("Sign in with your Outlook account, set availability, and share a simple booking link.")
    if st.button("Sign in with Outlook", type="primary", use_container_width=True):
        st.query_params["page"] = "signin"; st.rerun()

    # simple how-to on white cards (unchanged logic)
    st.markdown("#### How to use Kandor Schedulify")
    cols = st.columns(2)
    with cols[0]:
        st.markdown('<div class="howto-card">🔐 <b>Sign in with Outlook</b><br/>Grant Calendars.ReadWrite & Mail.Send.</div>', unsafe_allow_html=True)
        st.markdown('<div class="howto-card">🔗 <b>Share your booking link</b><br/>Copy from the Dashboard.</div>', unsafe_allow_html=True)
        st.markdown('<div class="howto-card">✉️ <b>Automatic invites</b><br/>Guests+you get an invite & you get an email.</div>', unsafe_allow_html=True)
    with cols[1]:
        st.markdown('<div class="howto-card">⚙️ <b>Configure settings</b><br/>Duration, working hours/days, time zone, video link.</div>', unsafe_allow_html=True)
        st.markdown('<div class="howto-card">📅 <b>Clients pick a slot</b><br/>We show only open times—no double booking.</div>', unsafe_allow_html=True)
        st.markdown('<div class="howto-card">🛠️ <b>Manage in Outlook</b><br/>Reschedule/cancel in Outlook anytime.</div>', unsafe_allow_html=True)

# ---------- Dashboard ----------
def dashboard():
    topbar()

    if "oid" not in st.session_state:
        st.info("Please sign in with Outlook to manage your settings.")
        if st.button("Sign in with Outlook", type="primary"):
            st.query_params["page"] = "signin"; st.rerun()
        return

    user = users_col().find_one({"oid": st.session_state["oid"]})
    if not user:
        st.error("User session lost. Please sign in again."); st.session_state.pop("oid", None); return

    st.success(f"Welcome, **{user.get('name','User')}** 👋")
    col1, col2, col3 = st.columns(3, gap="large")

    with col1:
        st.markdown("##### Calendar Connection")
        st.markdown('✅ Outlook Calendar Connected' if get_access_token_for_user_doc(user)
                    else '❌ Not connected. Click “Sign in with Outlook”.')

    with col2:
        st.markdown("##### Your Booking Link")
        slug = user.get("slug", "unknown"); booking_url = f"{BASE_URL}/?page=book&user={slug}"
        st_html(f"""
            <div class="link-card">
              <input id="k-copy-input" value="{booking_url}" readonly />
              <button id="k-copy-btn">Copy</button>
            </div>
            <script>
              const btn = document.getElementById('k-copy-btn');
              const inp = document.getElementById('k-copy-input');
              if (btn && inp) {{
                btn.addEventListener('click', () => {{
                  inp.select(); inp.setSelectionRange(0, 99999);
                  try {{ navigator.clipboard.writeText(inp.value); }} catch (e) {{ document.execCommand('copy'); }}
                }});
              }}
            </script>
        """, height=68)

    with col3:
        st.markdown("##### Quick Stats")
        st.write(f"**Meeting Duration:** {user.get('meeting_duration', 30)} min")
        st.write(f"**Working Days:** {len(user.get('available_days', []))} days")

    st.divider()
    with st.form("settings"):
        st.markdown("#### Meeting Settings")
        c1, c2 = st.columns(2)
        with c1:
            video_link = st.text_input("Zoom / Google Meet / Cisco Link", value=user.get("zoom_link", ""))
        with c2:
            dur = st.number_input("Meeting Duration (minutes)", min_value=15, max_value=180,
                                  value=int(user.get("meeting_duration", 30)), step=15)

        st.markdown("#### Working Hours & Time Zone")
        days = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        avail_days = st.multiselect("Working Days", options=days, default=user.get("available_days", days[:5]))

        user_tz = user.get("timezone", "Asia/Kolkata")
        common_tz = ["UTC","Asia/Kolkata","Asia/Dubai","Asia/Singapore","Asia/Tokyo",
                     "Europe/London","Europe/Berlin","Europe/Paris",
                     "America/Los_Angeles","America/New_York","America/Chicago","America/Toronto",
                     "Australia/Sydney","Africa/Johannesburg"]
        if user_tz not in common_tz: common_tz.insert(0, user_tz)
        tz = st.selectbox("Time Zone", options=common_tz, index=common_tz.index(user_tz))

        t1, t2 = st.columns(2)
        with t1:
            start_val = dt.datetime.strptime(user.get("start_time", "09:00"), "%H:%M").time()
            start_time = st.time_input("Work Day Starts At", value=start_val)
        with t2:
            end_val = dt.datetime.strptime(user.get("end_time", "17:00"), "%H:%M").time()
            end_time = st.time_input("Work Day Ends At", value=end_val)

        if st.form_submit_button("Save Settings", type="primary"):
            users_col().update_one({"oid": user["oid"]}, {"$set": {
                "zoom_link": video_link, "meeting_duration": int(dur),
                "available_days": avail_days, "start_time": start_time.strftime("%H:%M"),
                "end_time": end_time.strftime("%H:%M"), "timezone": tz,
            }})
            st.success("Settings saved!"); st.rerun()

# ---------- Month Calendar (responsive) ----------
def _month_start(d: dt.date) -> dt.date: return d.replace(day=1)
def _next_month(d: dt.date) -> dt.date: return dt.date(d.year+1,1,1) if d.month==12 else dt.date(d.year, d.month+1, 1)
def _prev_month(d: dt.date) -> dt.date: return dt.date(d.year-1,12,1) if d.month==1 else dt.date(d.year, d.month-1, 1)

def calendar_widget(selected: dt.date, working_days: List[str]) -> dt.date:
    """Pure-Streamlit (but mobile-safe) grid using HTML + query params for clicks."""
    if "k_view_month" not in st.session_state:
        st.session_state["k_view_month"] = _month_start(selected or dt.date.today())
    view = st.session_state["k_view_month"]

    # Handle clicks from params
    picked_param = st.query_params.get("pick")
    nav = st.query_params.get("nav")
    if nav == "prev":
        st.session_state["k_view_month"] = _prev_month(view); st.query_params.pop("nav", None)
        st.rerun()
    elif nav == "next":
        st.session_state["k_view_month"] = _next_month(view); st.query_params.pop("nav", None)
        st.rerun()
    if picked_param:
        try:
            selected = dt.date.fromisoformat(picked_param)
        except Exception:
            pass
        finally:
            st.query_params.pop("pick", None)

    # Build month cells
    first_weekday, days_in_month = calendar.monthrange(view.year, view.month)
    lead_blank = (first_weekday + 1) % 7
    today = dt.date.today()

    # Render
    st.markdown('<div class="cal-wrap cal-scope">', unsafe_allow_html=True)
    # nav row
    st_html(f"""
      <div class="cal-navrow">
        <a class="cal-navbtn" href="?page=book&user={st.query_params.get('user','')}&nav=prev">‹</a>
        <div class="cal-month">{view.strftime("%B %Y")}</div>
        <a class="cal-navbtn" href="?page=book&user={st.query_params.get('user','')}&nav=next">›</a>
      </div>
    """, height=40)

    # DOW header
    st.markdown('<div class="cal-dow">' + ''.join(f'<div>{d}</div>' for d in ["SUN","MON","TUE","WED","THU","FRI","SAT"]) + '</div>',
                unsafe_allow_html=True)

    # Grid (42 cells)
    cells = []
    day_num = 1
    for i in range(42):
        if i < lead_blank or day_num > days_in_month:
            cells.append('<div class="cal-day dim">&nbsp;</div>')
        else:
            d = dt.date(view.year, view.month, day_num)
            classes = ["cal-day"]
            if d == today: classes.append("today")
            if d == selected: classes.append("sel")
            enabled = d.strftime("%A") in working_days and d >= today
            if enabled:
                url = f'?page=book&user={st.query_params.get("user","")}&pick={d.isoformat()}'
                cells.append(f'<a class="cal-day btn {" ".join(classes)}" href="{url}">{day_num}</a>')
            else:
                classes.append("dim")
                cells.append(f'<div class="{" ".join(classes)}">{day_num}</div>')
            day_num += 1

    st.markdown('<div class="cal-grid">' + ''.join(cells) + '</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    return selected

# ---------- Booking ----------
def booking_page():
    topbar()

    slug_or_email = st.query_params.get("user", "")
    if not slug_or_email:
        st.info(f"Add '?user=<slug>' to the URL, e.g., `{BASE_URL}/?page=book&user=navdeep`"); return

    user = get_user_by_slug_or_email(slug_or_email)
    if not user: st.error("This booking link is invalid."); return

    left, middle, right = st.columns([1, 2, 1.3], gap="large")

    with left:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=36)
        st.markdown(f"### {user.get('name','Host')}")
        st.caption(f"{user.get('meeting_duration',30)} min")

    tzname = user.get("timezone", "UTC"); tz = ZoneInfo(tzname)

    with middle:
        st.markdown("#### Select a Date & Time")
        picked = st.session_state.get("picked_date", dt.date.today())
        picked = calendar_widget(picked, user.get("available_days", []))
        st.session_state["picked_date"] = picked
        st.caption(f"Time zone: **{tzname}**")

    selected_date = st.session_state["picked_date"]
    if selected_date.strftime("%A") not in user.get("available_days", []):
        with right: st.info("Choose a working day (enabled dates) to see available times.")
        return

    token = get_access_token_for_user_doc(user)
    if not token:
        with right: st.error("The host hasn't connected their Outlook calendar (or the connection expired).")
        return

    start_local_day = dt.datetime.combine(selected_date, dt.time(0,0,0, tzinfo=tz))
    end_local_day   = dt.datetime.combine(selected_date, dt.time(23,59,59, tzinfo=tz))
    start_utc = start_local_day.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
    end_utc   = end_local_day.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
    events = graph_day_view(token, start_utc, end_utc, tzname)

    busy: List[Tuple[dt.time, dt.time]] = []
    for ev in events:
        if ev.get("isCancelled"): continue
        if (ev.get("showAs") or "busy").lower() not in BUSY_STATUSES: continue
        try:
            sdt = dt.datetime.fromisoformat(ev["start"]["dateTime"])
            edt = dt.datetime.fromisoformat(ev["end"]["dateTime"])
            busy.append((sdt.time(), edt.time()))
        except Exception:
            continue

    meet_min = int(user.get("meeting_duration", 30))
    work_start = dt.datetime.strptime(user.get("start_time", "09:00"), "%H:%M").time()
    work_end   = dt.datetime.strptime(user.get("end_time", "17:00"), "%H:%M").time()

    slots = build_slots(selected_date, work_start, work_end, meet_min, busy)

    now_local = dt.datetime.now(tz)
    if selected_date == now_local.date():
        cutoff = (now_local + dt.timedelta(minutes=MIN_LEAD_MINUTES)).time()
        slots = [t for t in slots if t > cutoff]

    if RECHECK_EACH_SLOT and slots:
        filtered = []
        for t in slots:
            s_local = dt.datetime.combine(selected_date, t, tzinfo=tz)
            e_local = s_local + dt.timedelta(minutes=meet_min)
            if is_interval_free(token, s_local, e_local, tzname): filtered.append(t)
        slots = filtered

    with right:
        st.markdown(f"#### {selected_date.strftime('%A, %B %d')}")
        if not slots:
            st.info("No available time slots for this day."); return

        cols = st.columns(3)
        chosen_key = "selected_slot_time"
        chosen_time: Optional[dt.time] = st.session_state.get(chosen_key)

        def pretty(t: dt.time) -> str:
            fmt = "%-I:%M %p" if os.name != "nt" else "%#I:%M %p"
            return dt.datetime.combine(dt.date.today(), t).strftime(fmt)

        for i, t in enumerate(slots):
            with cols[i % 3]:
                if st.button(pretty(t), key=f"slot_{t}", use_container_width=True):
                    st.session_state[chosen_key] = t; chosen_time = t

        st.markdown("---")
        st.write(f"**Selected time:** {pretty(chosen_time) if chosen_time else '—'}")

        with st.form("book"):
            name = st.text_input("Your Name")
            email = st.text_input("Your Email")
            agenda = st.text_area("Agenda / context (optional)", height=100)
            if st.form_submit_button("Confirm Booking", type="primary"):
                if not (name.strip() and email.strip() and chosen_time):
                    st.error("Please pick a time and enter your name & email."); return
                start_dt_local = dt.datetime.combine(selected_date, chosen_time, tzinfo=tz)
                end_dt_local = start_dt_local + dt.timedelta(minutes=meet_min)

                if not is_interval_free(token, start_dt_local, end_dt_local, tzname):
                    st.error("Someone just booked this slot. Please pick another time."); st.rerun()

                agenda_snip = (agenda.strip()[:80] + "…") if (agenda and len(agenda.strip()) > 80) else (agenda.strip() if agenda else "")
                subject = f"Meeting with {name}" + (f" — {agenda_snip}" if agenda_snip else "")
                body = f"""
                    <p>Meeting scheduled via Kandor Schedulify.</p>
                    <ul>
                      <li><b>Guest:</b> {name} &lt;{email.strip()}&gt;</li>
                      <li><b>When:</b> {start_dt_local.strftime("%A, %B %d %Y %I:%M %p")} ({tzname})</li>
                      <li><b>Duration:</b> {meet_min} minutes</li>
                      <li><b>Video link:</b> {user.get('zoom_link','N/A')}</li>
                      {"<li><b>Agenda:</b> " + agenda.strip() + "</li>" if (agenda and agenda.strip()) else ""}
                    </ul>
                """
                ok, _ = graph_create_event(token, subject, body, start_dt_local, end_dt_local, [email.strip()], tzname)
                if ok:
                    host_email = user.get("email")
                    human_time = start_dt_local.strftime("%A, %B %d at %I:%M %p")
                    mail_subj = f"[New booking] {name} <{email.strip()}> on {human_time} ({tzname})" + (f" — {agenda_snip}" if agenda_snip else "")
                    mail_body = f"""
                        <p>New booking created.</p>
                        <ul>
                          <li><b>Guest:</b> {name} &lt;{email.strip()}&gt;</li>
                          <li><b>When:</b> {human_time} ({tzname})</li>
                          <li><b>Duration:</b> {meet_min} minutes</li>
                          <li><b>Video link:</b> {user.get('zoom_link','N/A')}</li>
                          {"<li><b>Agenda:</b> " + agenda.strip() + "</li>" if (agenda and agenda.strip()) else ""}
                        </ul>
                    """
                    if host_email: graph_send_mail(token, host_email, mail_subj, mail_body, tzname)
                    st.success(f"Meeting booked! Invite sent to {email.strip()}."); st.balloons()
                else:
                    st.error("That time is no longer available. Please choose another slot."); st.rerun()

# ---------- Sign-in ----------
def signin_page():
    topbar()
    st.markdown("### Sign in with Outlook")
    st.write("You’ll be redirected to Microsoft to sign in.")
    auth_url = create_auth_url()
    st.markdown(f"""
      <script>
        (function(){{
          const url="{auth_url}";
          try {{
            if (window.top && window.top !== window.self) {{ window.top.location.href = url; }}
            else {{ window.location.assign(url); }}
          }} catch(e){{}}
        }})();
      </script>
    """, unsafe_allow_html=True)
    st.markdown(
        f'<a href="{auth_url}" target="_blank" rel="noopener" '
        'style="display:inline-block;padding:10px 16px;border-radius:8px;'
        'background:#4f46e5;color:#fff;text-decoration:none;font-weight:700;">'
        'Continue with Microsoft</a>', unsafe_allow_html=True)

# ---------- Router ----------
def main():
    if "code" in st.query_params and "state" in st.query_params:
        finish_auth_redirect(); return
    page = st.query_params.get("page", "home")
    if page == "dashboard": dashboard()
    elif page == "book": booking_page()
    elif page == "signin": signin_page()
    else: landing()

if __name__ == "__main__":
    main()
