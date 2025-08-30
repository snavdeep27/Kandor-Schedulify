# Kandor Schedulify — Outlook sign-in, inline month calendar, agenda, no-double-booking
# -------------------------------------------------------------------------------------
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
EXTRA_CSS = """
<style>
/* Add padding so your hero/topbar clears the Streamlit Cloud header */
.block-container { padding-top: 4rem !important; }
</style>
"""
st.markdown(EXTRA_CSS, unsafe_allow_html=True)

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

# ---------- Slot policy ----------
BUSY_STATUSES = {"busy", "oof", "workingelsewhere", "tentative"}
MIN_LEAD_MINUTES = 5
RECHECK_EACH_SLOT = True

# ---------- Mongo ----------
@st.cache_resource
def _mongo():
    client = MongoClient(MONGO_URI)
    client.admin.command("ping")
    return client

def _db():
    c = _mongo()
    if MONGO_DB_NAME:
        return c[MONGO_DB_NAME]
    try:
        d = c.get_default_database()
        if d is not None:
            return d
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
    changed = True
    try:
        attr = getattr(cache, "has_state_changed", None)
        if attr is not None: changed = bool(attr)
    except Exception:
        changed = True
    if changed:
        try:
            users_col().update_one(user_key, {"$set": {"msal_cache": cache.serialize()}})
        except Exception:
            pass

def create_auth_url():
    app = build_msal_app()
    flow = app.initiate_auth_code_flow(scopes=MS_SCOPE, redirect_uri=MS_REDIRECT_URI)
    flows_col().insert_one({"_id": flow["state"], "flow": flow, "created_utc": dt.datetime.utcnow()})
    return flow["auth_uri"]

def finish_auth_redirect():
    state = st.query_params.get("state"); code = st.query_params.get("code")
    if not state or not code:
        st.error("Missing authorization response parameters."); return
    doc = flows_col().find_one({"_id": state})
    if not doc:
        st.error("Sign-in session expired. Please try again."); return
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
    if not oid:
        st.error("Missing Microsoft account ID (oid)."); return

    u = users_col().find_one({"oid": oid})
    if not u:
        base_slug = slugify_email(email, name)
        slug = ensure_unique_slug(base_slug, oid)
        users_col().insert_one({
            "oid": oid, "email": email, "name": name, "slug": slug,
            "zoom_link": "",
            "meeting_duration": 30,
            "available_days": ["Monday","Tuesday","Wednesday","Thursday","Friday"],
            "start_time": "09:00", "end_time": "17:00",
            "timezone": "Asia/Kolkata",
        })
    else:
        slug = u.get("slug") or ensure_unique_slug(slugify_email(u.get("email",""), u.get("name","user")), oid)
        users_col().update_one({"oid": oid}, {"$set": {
            "slug": slug,
            "email": email or u.get("email",""),
            "name": name or u.get("name",""),
        }})

    users_col().update_one({"oid": oid}, {"$set": {"msal_cache": cache.serialize()}})
    st.session_state["oid"] = oid
    st.success("Signed in with Outlook!")
    time.sleep(0.5)
    st.query_params.clear(); st.query_params["page"] = "dashboard"; st.rerun()

def get_user_by_slug_or_email(value: str):
    u = users_col().find_one({"slug": value})
    if not u and "@" in value:
        u = users_col().find_one({"email": value})
    return u

def get_access_token_for_user_doc(user_doc) -> Optional[str]:
    cache = msal.SerializableTokenCache()
    serialized = user_doc.get("msal_cache")
    if serialized:
        try: cache.deserialize(serialized)
        except Exception: pass
    app = build_msal_app(cache)
    accounts = app.get_accounts()
    token_result = None
    if accounts:
        token_result = app.acquire_token_silent(MS_SCOPE, account=accounts[0])
    if token_result and "access_token" in token_result:
        persist_cache_for_user({"oid": user_doc["oid"]}, cache)
        return token_result["access_token"]
    return None

# ---------- Graph ----------
def graph_headers(token: str, tzname: Optional[str] = None):
    headers = {"Authorization": f"Bearer {token}"}
    headers["Prefer"] = f'outlook.timezone="{tzname or "UTC"}"'
    return headers

def graph_day_view(token: str, start_utc: dt.datetime, end_utc: dt.datetime, tzname: str) -> List[dict]:
    params = {
        "startDateTime": start_utc.isoformat(),
        "endDateTime": end_utc.isoformat(),
        "$select": "subject,start,end,showAs,isCancelled",
        "$orderby": "start/dateTime ASC"
    }
    r = requests.get(f"{GRAPH}/me/calendarView",
                     headers=graph_headers(token, tzname),
                     params=params, timeout=20)
    if r.status_code == 200:
        return r.json().get("value", [])
    st.warning(f"Graph calendarView failed ({r.status_code}). Treating day as free."); return []

def is_interval_free(token: str, start_local: dt.datetime, end_local: dt.datetime, tzname: str) -> bool:
    start_utc = start_local.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
    end_utc   = end_local.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
    events = graph_day_view(token, start_utc, end_utc, tzname)
    for ev in events:
        if ev.get("isCancelled"):
            continue
        show_as = (ev.get("showAs") or "busy").lower()
        if show_as in {"busy", "oof", "workingelsewhere", "tentative"}:
            return False
    return True

def graph_create_event(token: str, subject: str, html_body: str,
                       start_local: dt.datetime, end_local: dt.datetime,
                       attendees: List[str], timezone_name: str) -> Tuple[bool, Optional[dict]]:
    payload = {
        "subject": subject,
        "showAs": "busy",
        "transactionId": str(uuid.uuid4()),
        "body": {"contentType": "HTML", "content": html_body},
        "start": {"dateTime": start_local.isoformat(), "timeZone": timezone_name},
        "end": {"dateTime": end_local.isoformat(), "timeZone": timezone_name},
        "attendees": [{"emailAddress": {"address": e}, "type": "required"} for e in attendees]
    }
    r = requests.post(f"{GRAPH}/me/events",
                      headers=graph_headers(token, timezone_name),
                      json=payload, timeout=20)
    if r.status_code in (201, 200):
        return True, r.json()
    return False, {"status": r.status_code, "text": r.text}

def graph_send_mail(token: str, to_addr: str, subject: str, html_body: str, tzname: str) -> bool:
    url = f"{GRAPH}/me/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_addr}}],
        },
        "saveToSentItems": True
    }
    r = requests.post(url, headers=graph_headers(token, tzname), json=payload, timeout=20)
    return r.status_code in (202, 200)

# ---------- Slot math (no buffer) ----------
def overlap(a_start: dt.time, a_end: dt.time, b_start: dt.time, b_end: dt.time) -> bool:
    return max(a_start, b_start) < min(a_end, b_end)

def build_slots(day: dt.date, work_start: dt.time, work_end: dt.time,
                meeting_minutes: int, busy: List[Tuple[dt.time, dt.time]]) -> List[dt.time]:
    step = dt.timedelta(minutes=meeting_minutes)
    t = dt.datetime.combine(day, work_start)
    end = dt.datetime.combine(day, work_end)
    result = []
    while t + step <= end:
        s = t.time(); e = (t + step).time()
        if not any(overlap(s, e, bs, be) for bs, be in busy):
            result.append(s)
        t += step
    return result

# ---------- CSS (colorful calendar & buttons) ----------
st.markdown("""
<style>
.block-container { padding-top: 1.1rem; padding-bottom: 2rem; }
.topbar { display:flex; align-items:center; justify-content:space-between; padding:8px 6px 0 6px; }
.topbar-left { display:flex; align-items:center; gap:10px; }
.appname { font-weight:800; letter-spacing:-0.02em; font-size:20px; }

/* Global button style (soft gradient pills) */
.stButton > button {
  border-radius: 999px !important;
  border: 1px solid #cfd4ff !important;
  background: linear-gradient(180deg,#ffffff 0%,#f4f6ff 100%) !important;
  color: #1f2544 !important;
  box-shadow: 0 2px 6px rgba(80, 90, 230, .06) !important;
}
.stButton > button:hover {
  border-color: #8e96ff !important;
  box-shadow: 0 6px 16px rgba(80, 90, 230, .18) !important;
}

/* Calendar */
.cal-header { display:flex; align-items:center; justify-content:space-between; margin: 6px 0 10px 0; }
.cal-title { font-weight: 700; }
.cal-grid { display:grid; grid-template-columns: repeat(7, minmax(36px,1fr)); gap:10px; }
.cal-dow { text-align:center; font-size: 12px; color:#6b7280; }
.cal-day {
  text-align:center; padding:12px 0; border-radius: 999px;
  border:1px solid #e6e9ff; background: #fbfcff; color:#3a3f6b;
}
.cal-day.today {
  background: linear-gradient(180deg,#eef3ff 0%, #f5f7ff 100%);
  border-color:#b9c2ff;
}
.cal-day.sel {
  background: linear-gradient(180deg,#6a5cff 0%, #8c7bff 100%);
  color:white; border-color:#6a5cff;
  box-shadow: 0 6px 16px rgba(106,92,255,.28);
}
.cal-day.dim { color:#a3a3a3; border-color:#f1f5f9; background:#fafbff; opacity:.65; }
.cal-nav { min-width:44px; }

/* Time-slot buttons (right column) inherit global pill styling; bump font-weight a bit */
div[data-testid="stVerticalBlock"] .stButton > button { font-weight: 700 !important; }

/* Booking link pill */
.link-card {
  display:flex; align-items:center; gap:10px; padding:12px 14px;
  border-radius:999px; border:1px solid #cfd4ff;
  background:linear-gradient(135deg,#edf1ff 0%, #f8faff 100%);
  box-shadow:0 8px 22px rgba(16,24,40,.08);
}
.link-card input {
  border:none; background:transparent; outline:none; width:100%;
  font-family: ui-monospace, Menlo, monospace;
  font-size:14px; color:#1f2544;
}
..link-card button{
  padding:10px 16px;
  border-radius:999px;
  border:none;
  color:#fff;
  font-weight:800;
  letter-spacing:.2px;
  cursor:pointer;
  background:linear-gradient(90deg,#6366f1 0%, #a855f7 100%); /* indigo → fuchsia */
  box-shadow:0 8px 18px rgba(99,102,241,.25);
  transition:transform .08s ease, box-shadow .15s ease, filter .15s ease;
}
.link-card button:hover{
  filter:brightness(1.05);
  box-shadow:0 10px 22px rgba(99,102,241,.32);
}
.link-card button:active{
  transform:translateY(1px) scale(.99);
}
</style>
""", unsafe_allow_html=True)

HOWTO_CSS = """
<style>
.howto-grid{
  display:grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap:16px;
  margin-top:18px;
}
.howto-card{
  background:#ffffff;
  border:1px solid rgba(0,0,0,.06);
  border-radius:14px;
  padding:14px 16px;
  box-shadow:0 4px 14px rgba(0,0,0,.04);
}
.howto-card h4{ margin:.2rem 0 .25rem; font-size:1rem; }
.howto-emoji{ font-size:22px; margin-right:8px; }
</style>
"""
st.markdown(HOWTO_CSS, unsafe_allow_html=True)

# ---------- Header ----------
def topbar():
    c1, c2 = st.columns([7, 5])
    with c1:
        st.markdown('<div class="topbar"><div class="topbar-left">', unsafe_allow_html=True)
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=34)
        st.markdown('<div class="appname">Kandor Schedulify</div></div></div>', unsafe_allow_html=True)
    with c2:
        r1, r2 = st.columns([4,1])
        with r2:
            if "oid" in st.session_state:
                if st.button("Sign out", type="primary"):
                    st.session_state.pop("oid", None)
                    st.success("Signed out."); time.sleep(0.3)
                    st.query_params.clear(); st.rerun()

# ---------- Landing ----------
def landing():
    topbar()
    st.markdown(
        """
        <div class="hero">
          <h3>Your Personal Calendly Clone by Kandor</h3>
          <p>Sign in with your Outlook account, set your availability, and share a simple booking link.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Sign in with Outlook", type="primary", use_container_width=True):
        st.query_params["page"] = "signin"
        st.rerun()

    st.markdown("#### How to use Kandor Schedulify")
    st.markdown(
        """
        <div class="howto-grid">
          <div class="howto-card">
            <div><span class="howto-emoji">🔐</span><b>Sign in with Outlook</b></div>
            <p>Connect securely via Microsoft and grant <i>Calendars.ReadWrite</i> & <i>Mail.Send</i>.</p>
          </div>
          <div class="howto-card">
            <div><span class="howto-emoji">⚙️</span><b>Configure settings</b></div>
            <p>Set meeting duration, working days/hours, time zone, and your video link on the Dashboard.</p>
          </div>
          <div class="howto-card">
            <div><span class="howto-emoji">🔗</span><b>Share your booking link</b></div>
            <p>Copy the personal link from the Dashboard and send it to clients.</p>
          </div>
          <div class="howto-card">
            <div><span class="howto-emoji">📅</span><b>Clients pick a slot</b></div>
            <p>We show <b>only open</b> times from your Outlook calendar in their time zone—no double booking.</p>
          </div>
          <div class="howto-card">
            <div><span class="howto-emoji">✉️</span><b>Automatic invites</b></div>
            <p>Both attendees receive a calendar invite. You also get a confirmation email with the guest’s details.</p>
          </div>
          <div class="howto-card">
            <div><span class="howto-emoji">🛠️</span><b>Manage in Outlook</b></div>
            <p>Reschedule/cancel directly in Outlook. Update your settings anytime in the Dashboard.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

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
        st.error("User session lost. Please sign in again.")
        st.session_state.pop("oid", None); return

    st.success(f"Welcome, **{user.get('name','User')}** 👋")
    col1, col2, col3 = st.columns(3, gap="large")

    with col1:
        st.markdown("##### Calendar Connection")
        token_test = get_access_token_for_user_doc(user)
        st.markdown('✅ Outlook Calendar Connected' if token_test else '❌ Not connected. Click “Sign in with Outlook”.')

    with col2:
        st.markdown("##### Your Booking Link")
        slug = user.get("slug", "unknown")
        booking_url = f"{BASE_URL}/?page=book&user={slug}"
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
                  try {{ navigator.clipboard.writeText(inp.value); }} catch (e) {{
                    document.execCommand('copy');
                  }}
                }});
              }}
            </script>
        """, height=68)

    with col3:
        st.markdown("##### Quick Stats")
        mdur = user.get("meeting_duration", 30)
        wdays = len(user.get("available_days", []))
        st.write(f"**Meeting Duration:** {mdur} min")
        st.write(f"**Working Days:** {wdays} days")

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
        common_tz = [
            "UTC","Asia/Kolkata","Asia/Dubai","Asia/Singapore","Asia/Tokyo",
            "Europe/London","Europe/Berlin","Europe/Paris",
            "America/Los_Angeles","America/New_York","America/Chicago","America/Toronto",
            "Australia/Sydney","Africa/Johannesburg",
        ]
        if user_tz not in common_tz: common_tz.insert(0, user_tz)
        tz = st.selectbox("Time Zone", options=common_tz, index=common_tz.index(user_tz))

        t1, t2 = st.columns(2)
        with t1:
            start_str = user.get("start_time", "09:00")
            start_val = dt.datetime.strptime(start_str, "%H:%M").time()
            start_time = st.time_input("Work Day Starts At", value=start_val)
        with t2:
            end_str = user.get("end_time", "17:00")
            end_val = dt.datetime.strptime(end_str, "%H:%M").time()
            end_time = st.time_input("Work Day Ends At", value=end_val)

        if st.form_submit_button("Save Settings", type="primary"):
            users_col().update_one({"oid": user["oid"]}, {"$set": {
                "zoom_link": video_link,
                "meeting_duration": int(dur),
                "available_days": avail_days,
                "start_time": start_time.strftime("%H:%M"),
                "end_time": end_time.strftime("%H:%M"),
                "timezone": tz,
            }})
            st.success("Settings saved!"); st.rerun()

# ---------- Month Calendar ----------
def _month_start(d: dt.date) -> dt.date:
    return d.replace(day=1)

def _next_month(d: dt.date) -> dt.date:
    y, m = d.year, d.month
    return dt.date(y+1, 1, 1) if m == 12 else dt.date(y, m+1, 1)

def _prev_month(d: dt.date) -> dt.date:
    y, m = d.year, d.month
    return dt.date(y-1, 12, 1) if m == 1 else dt.date(y, m-1, 1)

def calendar_widget(selected: dt.date, working_days: List[str]) -> dt.date:
    key_prefix = "kcal_"
    if "k_view_month" not in st.session_state:
        st.session_state["k_view_month"] = _month_start(selected if selected else dt.date.today())
    view = st.session_state["k_view_month"]

    c1, c2, c3 = st.columns([1, 5, 1])
    with c1:
        if st.button("‹", key=key_prefix+"prev", use_container_width=True):
            st.session_state["k_view_month"] = _prev_month(view)
            st.rerun()
    with c2:
        st.markdown(f'<div class="cal-header"><div class="cal-title">{view.strftime("%B %Y")}</div></div>', unsafe_allow_html=True)
    with c3:
        if st.button("›", key=key_prefix+"next", use_container_width=True):
            st.session_state["k_view_month"] = _next_month(view)
            st.rerun()

    cols = st.columns(7)
    for i, dow in enumerate(["SUN","MON","TUE","WED","THU","FRI","SAT"]):
        with cols[i]:
            st.markdown(f'<div class="cal-dow">{dow}</div>', unsafe_allow_html=True)

    first_weekday, days_in_month = calendar.monthrange(view.year, view.month)
    lead_blanks = (first_weekday + 1) % 7
    day_counter = 1
    today = dt.date.today()

    for _ in range(6):
        row = st.columns(7)
        for col_idx in range(7):
            with row[col_idx]:
                if lead_blanks > 0:
                    lead_blanks -= 1
                    st.markdown('<div class="cal-day dim">&nbsp;</div>', unsafe_allow_html=True)
                elif day_counter <= days_in_month:
                    d = dt.date(view.year, view.month, day_counter)
                    classes = ["cal-day"]
                    dayname = d.strftime("%A")
                    enabled = dayname in working_days and d >= today
                    if d == today: classes.append("today")
                    if d == selected: classes.append("sel")
                    if enabled:
                        if st.button(str(day_counter), key=f"{key_prefix}{d.isoformat()}", use_container_width=True):
                            selected = d
                    else:
                        classes.append("dim")
                        st.markdown(f'<div class="{" ".join(classes)}">{day_counter}</div>', unsafe_allow_html=True)
                    day_counter += 1
                else:
                    st.markdown('<div class="cal-day dim">&nbsp;</div>', unsafe_allow_html=True)
    return selected

# ---------- Booking ----------
def booking_page():
    topbar()

    slug_or_email = st.query_params.get("user", "")
    if not slug_or_email:
        st.info("Add '?user=<slug>' to the URL, e.g., "
                f"`{BASE_URL}/?page=book&user=navdeep`")
        return

    user = get_user_by_slug_or_email(slug_or_email)
    if not user:
        st.error("This booking link is invalid.")
        return

    left, middle, right = st.columns([1, 2, 1.3], gap="large")

    with left:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=36)
        st.markdown(f"### {user.get('name','Host')}")
        st.caption(f"{user.get('meeting_duration',30)} min")

    tzname = user.get("timezone", "UTC")
    tz = ZoneInfo(tzname)

    with middle:
        st.markdown("#### Select a Date & Time")
        picked = st.session_state.get("picked_date", dt.date.today())
        picked = calendar_widget(picked, user.get("available_days", []))
        st.session_state["picked_date"] = picked
        st.caption(f"Time zone: **{tzname}**")

    selected_date = st.session_state["picked_date"]

    if selected_date.strftime("%A") not in user.get("available_days", []):
        with right:
            st.info("Choose a working day (enabled dates) to see available times.")
        return

    token = get_access_token_for_user_doc(user)
    if not token:
        with right:
            st.error("The host hasn't connected their Outlook calendar (or the connection expired).")
        return

    start_local_day = dt.datetime.combine(selected_date, dt.time(0,0,0, tzinfo=tz))
    end_local_day   = dt.datetime.combine(selected_date, dt.time(23,59,59, tzinfo=tz))
    start_utc = start_local_day.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
    end_utc   = end_local_day.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)

    events = graph_day_view(token, start_utc, end_utc, tzname)

    busy: List[Tuple[dt.time, dt.time]] = []
    for ev in events:
        if ev.get("isCancelled"): continue
        show_as = (ev.get("showAs") or "busy").lower()
        if show_as not in BUSY_STATUSES: continue
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
            if is_interval_free(token, s_local, e_local, tzname):
                filtered.append(t)
        slots = filtered

    with right:
        st.markdown(f"#### {selected_date.strftime('%A, %B %d')}")
        if not slots:
            st.info("No available time slots for this day.")
            return

        cols = st.columns(3)
        chosen_key = "selected_slot_time"
        chosen_time: Optional[dt.time] = st.session_state.get(chosen_key)

        def pretty(t: dt.time) -> str:
            fmt = "%-I:%M %p" if os.name != "nt" else "%#I:%M %p"
            return dt.datetime.combine(dt.date.today(), t).strftime(fmt)

        for i, t in enumerate(slots):
            with cols[i % 3]:
                if st.button(pretty(t), key=f"slot_{t}", use_container_width=True):
                    st.session_state[chosen_key] = t
                    chosen_time = t

        st.markdown("---")
        chosen_label = pretty(chosen_time) if chosen_time else "—"
        st.write(f"**Selected time:** {chosen_label}")

        with st.form("book"):
            name = st.text_input("Your Name")
            email = st.text_input("Your Email")
            agenda = st.text_area("Agenda / context (optional)", height=100)
            confirm = st.form_submit_button("Confirm Booking", type="primary")
            if confirm:
                if not (name.strip() and email.strip() and chosen_time):
                    st.error("Please pick a time and enter your name & email.")
                    return
                start_dt_local = dt.datetime.combine(selected_date, chosen_time, tzinfo=tz)
                end_dt_local = start_dt_local + dt.timedelta(minutes=meet_min)

                if not is_interval_free(token, start_dt_local, end_dt_local, tzname):
                    st.error("Someone just booked this slot. Please pick another time.")
                    st.rerun()

                agenda_snip = (agenda.strip()[:80] + "…") if agenda and len(agenda.strip()) > 80 else (agenda.strip() if agenda else "")
                subject = f"Meeting with {name}" + (f" — {agenda_snip}" if agenda_snip else "")
                body = f"""
                    <p>Meeting scheduled via Kandor Schedulify.</p>
                    <ul>
                      <li><b>Guest:</b> {name} &lt;{email.strip()}&gt;</li>
                      <li><b>When:</b> {start_dt_local.strftime("%A, %B %d %Y %I:%M %p")} ({tzname})</li>
                      <li><b>Duration:</b> {meet_min} minutes</li>
                      <li><b>Video link:</b> {user.get('zoom_link','N/A')}</li>
                      {"<li><b>Agenda:</b> " + agenda.strip() + "</li>" if agenda.strip() else ""}
                    </ul>
                """
                ok, _ = graph_create_event(
                    token, subject, body,
                    start_dt_local, end_dt_local,
                    [email.strip()], tzname
                )
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
                          {"<li><b>Agenda:</b> " + agenda.strip() + "</li>" if agenda.strip() else ""}
                        </ul>
                    """
                    if host_email:
                        graph_send_mail(token, host_email, mail_subj, mail_body, tzname)
                    st.success(f"Meeting booked! Invite sent to {email.strip()}.")
                    st.balloons()
                else:
                    st.error("That time is no longer available. Please choose another slot.")
                    st.caption("Tip: your page may be stale. I’ve refreshed the available times.")
                    st.rerun()

# ---------- Sign-in ----------
def signin_page():
    topbar()
    st.markdown("### Sign in with Outlook")
    st.write("You’ll be redirected to Microsoft to sign in.")

    auth_url = create_auth_url()

    # Try to navigate the TOP window (not the Streamlit iframe).
    st.markdown(
        f"""
        <script>
          (function() {{
            const url = "{auth_url}";
            try {{
              if (window.top && window.top !== window.self) {{
                window.top.location.href = url;
              }} else {{
                window.location.assign(url);
              }}
            }} catch (e) {{
              console.warn("Auto-redirect suppressed:", e);
            }}
          }})();
        </script>
        """,
        unsafe_allow_html=True,
    )

    # Visible fallback for browsers/pop-up blockers
    st.markdown(
        f'<a href="{auth_url}" target="_blank" rel="noopener" '
        'style="display:inline-block;padding:10px 16px;border-radius:8px;'
        'background:#4f46e5;color:#fff;text-decoration:none;font-weight:700;">'
        'Continue with Microsoft</a>',
        unsafe_allow_html=True,
    )

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
