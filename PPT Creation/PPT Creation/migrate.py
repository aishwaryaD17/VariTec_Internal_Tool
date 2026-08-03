# REMOVE_NEW_PAGE_SIDEBAR
# SIDEBAR_CLEAN_APPLIED
# SIDEBAR_MONTH_ISSUES_FIXED
# PAGES_PANEL_RESTORED
# SUBSECTION_NAME_FIX_APPLIED
# SIDEBAR_TREE_LINES_APPLIED
# SIDEBAR_FINAL_APPLIED
# SIDEBAR_STYLE_V2_APPLIED
# SIDEBAR_3LEVEL_APPLIED
# SI_FIXES_RESTORED_V2
# SI_FIXES_RESTORED
# OLD_COLOR_SCRIPTS_REMOVED
# SI_COLORS_FINAL_20260706
# SI_COLORS_FRESH_20260706
# FIX_ALL_SI_COLORS
# SI_COLOR_AND_SAVE_FIXED
# SI_PURGE_GREEN_FIXED
# SI_HEADER_NO_GREEN_FIXED
# SI_CARD_COLORS_SOFT_FIXED
# SI_STATUS_BOLD_FIXED
# SI_DATE_BOLD_V2_FIXED
# SI_DATE_BOLD_FIXED
# PI_SUBROW_ALIGN_FIXED
# SI_CARD_COLOR_V2_FIXED
# SI_CARD_COLOR_JS_FIXED
# SI_CARD_HEADER_COLOR_FIXED
# SI_CARD_STATUS_COLOR_FIXED
# SEC_CLONE_ICON_SVG_FIXED
# CLONE_ICON_SVG_FIXED
# PASTE_LIST_MARKERS_FIXED
# PASTE_CLEAN_V3_FIXED
# PASTE_CLEAN_V2_FIXED
# PASTE_CLEAN_FIXED
# PI_PDF_OVERLAP_FIXED
# BANNER_UNIFORM_FIXED
# BANNER_UPPERCASE_FIXED
# RR_PDF_SIZE_V2_FIXED
# FA_PDF_ADDBUTTON_FIXED
# FA_PDF_TITLE_FIXED
# SI_DUPLICATE_BODY_FIXED
# RR_PDF_SIZE_FIXED
# RR_PDF_NOWRAP_FIXED
# RR_PDF_COMPACT_FIXED
# RR_PDF_FINAL_FIXED
# RR_PDF_V6_FIXED
# RR_PDF_V5_FIXED
# RR_PDF_V4_FIXED
# RR_PDF_V3_FIXED
# RR_PDF_V2_FIXED
# RR_PDF_BLANK_PAGE_FIXED
# CLONE_SECTION_ADDED
# CLONE_ICON_FIXED
# CLONE_VISIBLE_FIXED
# CLONE_PAGE_ADDED
# FOCUS_AREAS_DOUBLE_DOT_FIXED
# FOCUS_AREAS_IMPACT_ICON_FIXED
# FOCUS_AREAS_IMPACT_TOGGLE_FIXED
# FOCUS_AREAS_TITLE_HIDE_FIXED
# FOCUS_AREAS_V2_FIXED
# FOCUS_AREAS_MENU_ADDED
# FOCUS_AREAS_TEMPLATE_ADDED
# COVER_LOGO_PASTE_REARM_FIXED
# COVER_OVERLAY_FINAL_FIXED
# INSERT_IMG_CURSOR_FIXED
# COVER_OVERLAY_ATTR_FIXED
# COVER_TEXT_DRAG_FIXED
# COVER_TEXT_INLINE_FIXED
# COVER_OVERLAY_REARM_V2_FIXED
# COVER_OVERLAY_REARM_FIXED
# COVER_AFTER_PDF_V2_FIXED
# COVER_AFTER_PDF_FIXED
# SI_DRAG_REORDER_APPLIED
# SI_LIST_INDENT_FIXED
# SI_BACKSPACE_V4_FIXED
# SI_BACKSPACE_V3_FIXED
# SI_BACKSPACE_V2_FIXED
# SI_BACKSPACE_FIXED
# SI_ACTION_OVERFLOW_FIXED
# SI_ENTER_KEY_FIXED
# PI_HIDE_MAIN_STATUS_APPLIED
# PI_STATUS_INLINE_APPLIED
# PI_FLUENT_REDESIGN_APPLIED
# PI_TEMPLATE_REDESIGN_APPLIED
# PI_JS_QUOTES_FIXED
# ACTIVE_SECTION_ESCAPE_APPLIED
# PI_INLINE_FALLBACK_APPLIED
# PI_NO_ICONS_APPLIED
# FONT_SIZE_INCREASE_APPLIED
# PDF_HIDE_TITLE_APPLIED
# PI_V3_APPLIED
# PI_V2_APPLIED
# PI_TEMPLATE_APPLIED
# RR_FINAL_LIST_APPLIED
# RR_UNIFORM_APPLIED
# RR_V2_APPLIED
# RR_TEMPLATE_APPLIED
"""

VariTec Notes — Single File Setup & Launch

Run with:  python migrate.py

"""

import subprocess, sys, os, sqlite3


# ── Setup: install deps + migrate DB (only when run directly) ─────────────

def _setup():

    print("=" * 50)

    print("  VariTec Notes — Setup & Launch")

    print("=" * 50)


    print("\n[1/3] Installing dependencies...")

    for pkg in ["flask","flask-sqlalchemy","python-pptx","beautifulsoup4","lxml"]:

        r = subprocess.run([sys.executable,"-m","pip","install",pkg,"--quiet"],

                           capture_output=True, text=True)

        print(f"  {'✅' if r.returncode==0 else '⚠️ '} {pkg}")


    print("\n[2/3] Setting up database...")

    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),"notes.db")

    if os.path.exists(db_path):

        conn = sqlite3.connect(db_path); cur = conn.cursor()

        cur.execute("PRAGMA table_info(note)")

        cols = [r[1] for r in cur.fetchall()]

        if "section" not in cols:

            cur.execute("ALTER TABLE note ADD COLUMN section VARCHAR(100) DEFAULT 'General'")

            conn.commit(); print("  ✅ Database upgraded")

        else:

            print("  ✅ Database OK")

        conn.close()

    else:

        print("  ✅ Fresh database will be created on start")


    print("\n[3/3] Starting VariTec Notes...")

    print("  URL  : http://localhost:5000")

    print("  Login: sai@varitecconsulting.com")

    print("         aish@varitecconsulting.com")

    print("\n  Press Ctrl+C to stop")

    print("=" * 50 + "\n")


_setup()


# ══════════════════════════════════════════════════

#   APP CODE BELOW

# ══════════════════════════════════════════════════

import os, io, re, json, base64, sqlite3, copy

from datetime import datetime


from flask import (Flask, render_template_string, request,

                   redirect, session, jsonify, send_file)

from flask_sqlalchemy import SQLAlchemy

from pptx import Presentation

from pptx.util import Pt

from pptx.dml.color import RGBColor

from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE

from bs4 import BeautifulSoup


try:

    from openai import OpenAI as _OpenAI

    OPENAI_PKG = True

except ImportError:

    OPENAI_PKG = False


app = Flask(__name__)

# ---- PERMANENT FIX: strip stray lone-surrogate characters from any ----
# ---- response before Werkzeug tries to UTF-8 encode it. Prevents   ----
# ---- "UnicodeEncodeError: surrogates not allowed" from any route.  ----
from werkzeug.security import generate_password_hash, check_password_hash
import werkzeug.wrappers.response as _wkr_resp_patch
_orig_set_data = _wkr_resp_patch.Response.set_data
def _sanitized_set_data(self, value):
    if isinstance(value, str):
        value = ''.join(ch for ch in value if not (0xD800 <= ord(ch) <= 0xDFFF))
    elif isinstance(value, bytes):
        try:
            value.decode('utf-8')
        except UnicodeDecodeError:
            value = value.decode('utf-8', errors='ignore').encode('utf-8')
    return _orig_set_data(self, value)
_wkr_resp_patch.Response.set_data = _sanitized_set_data

app.secret_key = "varitec-notes-secret-2025"


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_FILE  = os.path.join(BASE_DIR, "notes.db")

app.config["SQLALCHEMY_DATABASE_URI"]  = f"sqlite:///{DB_FILE}"

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


db = SQLAlchemy(app)


ALLOWED_USERS = {

    "aish@varitecconsulting.com",

    "sai@varitecconsulting.com",

}


# ── Role constants ────────────────────────────────────────────────────────

ROLE_ADMIN = "admin"

ROLE_WRITE = "write"

ROLE_READ  = "read"

ROLES      = [ROLE_ADMIN, ROLE_WRITE, ROLE_READ]

DEFAULT_ADMINS = {"sai@varitecconsulting.com", "aish@varitecconsulting.com"}


TPL_PATH      = os.path.join(BASE_DIR, "default_template.pptx")

TPL_NAME_FILE = os.path.join(BASE_DIR, "default_template_name.txt")


def get_tpl_name():

    return open(TPL_NAME_FILE).read().strip() if os.path.exists(TPL_NAME_FILE) else None


def save_tpl_name(n):

    open(TPL_NAME_FILE, "w").write(n)


class Section(db.Model):

    id        = db.Column(db.Integer, primary_key=True)

    name      = db.Column(db.String(100), unique=True, nullable=False)

    order     = db.Column(db.Integer, default=0)

    parent_id = db.Column(db.Integer, db.ForeignKey('section.id'), nullable=True, default=None)


class Note(db.Model):

    id         = db.Column(db.Integer, primary_key=True)

    title      = db.Column(db.String(200), default="Untitled")

    content    = db.Column(db.Text,        default="")

    section    = db.Column(db.String(100), default="General")

    summary    = db.Column(db.Text,        default="")

    note_order = db.Column(db.Integer,     default=0)


class UserAccess(db.Model):

    """User roles managed by admins."""

    id    = db.Column(db.Integer, primary_key=True)

    email = db.Column(db.String(200), unique=True, nullable=False)

    role  = db.Column(db.String(20), default=ROLE_READ)

    password = db.Column(db.String(200), default="")

    reset_token = db.Column(db.String(100), default="")

    reset_expires = db.Column(db.Float, default=0)

    failed_attempts = db.Column(db.Integer, default=0)

    locked = db.Column(db.Boolean, default=False)

    must_change_password = db.Column(db.Boolean, default=False)


    def set_password(self, raw_password):

        self.password = generate_password_hash(raw_password)

        self.failed_attempts = 0

        self.locked = False


    def check_password(self, raw_password):

        stored = self.password or ""

        if not stored:

            return False

        # Legacy accounts may still have a plaintext password on file;

        # verify it directly and transparently upgrade it to a hash.

        if ":" not in stored:

            if stored == raw_password:

                self.set_password(raw_password)

                return True

            return False

        try:

            return check_password_hash(stored, raw_password)

        except ValueError:

            return stored == raw_password


    @staticmethod

    def get_role(email):

        if email in DEFAULT_ADMINS: return ROLE_ADMIN

        ua = UserAccess.query.filter_by(email=email).first()

        return ua.role if ua else None


    @staticmethod

    def has_access(email):

        if email in DEFAULT_ADMINS: return True

        ua = UserAccess.query.filter_by(email=email).first()

        return ua is not None


    @staticmethod

    def can_write(email):

        return UserAccess.get_role(email) in (ROLE_ADMIN, ROLE_WRITE)


    @staticmethod

    def is_admin(email):

        return UserAccess.get_role(email) == ROLE_ADMIN


class PendingSignup(db.Model):

    """Sign-up requests awaiting admin review/approval."""

    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(200), nullable=False)

    email = db.Column(db.String(200), unique=True, nullable=False)

    password_hash = db.Column(db.String(255), nullable=False)

    company = db.Column(db.String(200), default="")

    reason = db.Column(db.String(500), default="")

    status = db.Column(db.String(20), default="pending")

    created_at = db.Column(db.DateTime, default=datetime.utcnow)


def _migrate():

    if not os.path.exists(DB_FILE): return

    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()

    cur.execute("PRAGMA table_info(note)")

    cols = [r[1] for r in cur.fetchall()]

    if "section" not in cols:

        cur.execute("ALTER TABLE note ADD COLUMN section VARCHAR(100) DEFAULT 'General'")

    if "summary" not in cols:

        cur.execute("ALTER TABLE note ADD COLUMN summary TEXT DEFAULT ''")

    if "note_order" not in cols:

        cur.execute("ALTER TABLE note ADD COLUMN note_order INTEGER DEFAULT 0")

        # Seed existing notes with their current rowid order

        cur.execute("UPDATE note SET note_order = id")

    # Migrate section.parent_id
    cur.execute("PRAGMA table_info(section)")
    sec_cols = [r[1] for r in cur.fetchall()]
    if "parent_id" not in sec_cols:
        cur.execute("ALTER TABLE section ADD COLUMN parent_id INTEGER DEFAULT NULL")
    conn.commit(); conn.close()


def _seed():

    if Section.query.count() == 0:

        for i, n in enumerate(["General","Natco","Omnivium","Billing","Interview","Passwords"]):

            db.session.add(Section(name=n, order=i))

        db.session.commit()


with app.app_context():

    _migrate(); db.create_all(); _seed()


# ── Scorecard customer list ────────────────────────────────────────────────

SCORECARD_CUSTOMERS = [

    "MWI Animal Health Cencora",

    "Cencora",

    "Cardinal Health",

    "McKesson",

]


def _build_scorecard_html(customers):

    """Scorecard: editable month header + Customer/Week 1-4 table."""

    from datetime import datetime

    current_month = datetime.now().strftime("%B %Y")

    row_styles = ["#8b5cf6","#059669","#d97706","#2563eb","#db2777","#0891b2","#7c3aed"]

    th_week = "".join(f'<th>Week {w}</th>' for w in range(1,5))

    thead = f'<thead><tr><th class="sc-th-name">Customer Name</th>{th_week}</tr></thead>'

    tbody_rows = ""

    for i, customer in enumerate(customers):

        accent = row_styles[i % len(row_styles)]

        week_cells = "".join(

            f'<td class="sc-week-td"><div class="sc-cell-inner">'

            f'<span class="sc-val" contenteditable="true"></span>'

            f'<span class="sc-pct">%</span></div></td>'

            for _ in range(4)

        )

        name_cell = (f'<td class="sc-name-td" style="border-left:4px solid {accent};color:{accent};">'

                     f'{customer}</td>')

        tbody_rows += f'<tr class="sc-row">{name_cell}{week_cells}</tr>'

    SCORECARD_SVG = '<svg width="180" height="120" viewBox="0 0 180 120" xmlns="http://www.w3.org/2000/svg"><rect x="10" y="20" width="160" height="85" rx="8" fill="white" opacity=".12"/><rect x="20" y="30" width="60" height="8" rx="4" fill="white" opacity=".6"/><rect x="100" y="44" width="3" height="35" rx="1" fill="#4db2de" opacity=".7"/><rect x="120" y="52" width="3" height="27" rx="1" fill="#4db2de" opacity=".7"/><rect x="140" y="38" width="3" height="41" rx="1" fill="#4db2de" opacity=".7"/><rect x="160" y="46" width="3" height="33" rx="1" fill="#4db2de" opacity=".7"/><circle cx="150" cy="15" r="12" fill="#038dbd" opacity=".5"/><text x="145" y="20" font-size="13" fill="white">%</text></svg>'

    banner = (

        '<div style="background:linear-gradient(135deg,#0d152b 0%,#1e3868 60%,#0a2a5e 100%);'

        'border-radius:10px 10px 0 0;padding:14px 24px;display:flex;align-items:center;'

        'gap:16px;color:white;position:relative;overflow:hidden;">'

        '<div style="position:absolute;right:-30px;top:-40px;width:160px;height:160px;'

        'border-radius:50%;background:rgba(255,255,255,0.04);pointer-events:none;"></div>'

        '<div style="flex:1;position:relative;z-index:1;">'

        '<div style="font-size:22px;font-weight:900;letter-spacing:1px;text-transform:uppercase;line-height:1.1;">'

        'Customer Scorecards</div>'

        '</div>'

        '</div>'

    )

    return f'{banner}<table class="sc-table" id="scorecard-table">{thead}<tbody>{tbody_rows}</tbody></table>\n'


# All-customers template (default)

SCORECARD_HTML = _build_scorecard_html(SCORECARD_CUSTOMERS)


# Per-customer templates (one customer repeated across weeks)

SCORECARD_TEMPLATES = {

    "all":      ("Customer Scorecards",              SCORECARD_CUSTOMERS),

    "mwi":      ("Scorecard – MWI Animal Health Cencora",  ["MWI Animal Health Cencora"]),

    "cencora":  ("Scorecard – Cencora",                    ["Cencora"]),

    "cardinal": ("Scorecard – Cardinal Health",             ["Cardinal Health"]),

    "mckesson": ("Scorecard – McKesson",                   ["McKesson"]),

}


# ── API key config (stored in config.json next to app.py) ──────────────────

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


def get_config():

    try:

        return json.load(open(CONFIG_FILE))

    except Exception:

        return {}


def save_config(cfg):

    json.dump(cfg, open(CONFIG_FILE, "w"), indent=2)


def get_api_key(provider="openai"):

    """Get key: env var first, then config file."""

    env_key = {"openai": "OPENAI_API_KEY", "anthropic": "ANTHROPIC_API_KEY"}

    key = os.environ.get(env_key.get(provider, ""), "")

    if not key:

        key = get_config().get(f"{provider}_key", "")

    return key


def _generate_summary(title: str, plain_text: str) -> dict:

    """

    Generate AI summary. Tries OpenAI first, then Anthropic.

    Keys come from environment variables OR the in-app config file.

    """

    prompt = f"""Analyse the following note and return a structured JSON summary.


Note Title: {title}

Note Content:

{plain_text[:4000]}


Return ONLY a valid JSON object — no markdown, no code fences:

{{

  "small_title": "Concise title (max 10 words)",

  "issue": "Brief description of the issue or task",

  "root_cause": "Underlying cause or reason (N/A if not applicable)",

  "status": "Exactly one of: In Progress | Completed | Pending",

  "action": "Actions taken or recommended next steps"

}}"""


    # ── Try OpenAI ──────────────────────────────────────────────────────────

    openai_key = get_api_key("openai")

    if openai_key and OPENAI_PKG:

        try:

            client = _OpenAI(api_key=openai_key)

            resp = client.chat.completions.create(

                model="gpt-4o",

                messages=[

                    {"role": "system", "content": "You are a precise business analyst. Always respond with valid JSON only."},

                    {"role": "user",   "content": prompt},

                ],

                max_tokens=600, temperature=0.2,

            )

            raw = resp.choices[0].message.content.strip()

            raw = re.sub(r'^```json?\s*', '', raw); raw = re.sub(r'\s*```$', '', raw)

            return json.loads(raw)

        except Exception as e:

            print(f"OpenAI failed ({e}), trying Anthropic…")


    # ── Try Anthropic ───────────────────────────────────────────────────────

    try:

        import anthropic as _sdk

        anth_key = get_api_key("anthropic")

        client = _sdk.Anthropic(api_key=anth_key) if anth_key else _sdk.Anthropic()

        msg = client.messages.create(

            model="claude-sonnet-4-20250514",

            max_tokens=600,

            messages=[{"role": "user", "content": prompt}],

        )

        raw = msg.content[0].text.strip()

        raw = re.sub(r'^```json?\s*', '', raw); raw = re.sub(r'\s*```$', '', raw)

        return json.loads(raw)

    except Exception as e:

        print(f"Anthropic failed: {e}")


    raise ValueError(

        "No AI provider is configured. Add an API key via ⚙ Settings in the app "

        "(OpenAI key  OR  Anthropic key), or set the OPENAI_API_KEY / "

        "ANTHROPIC_API_KEY environment variable."

    )


C_NAVY=RGBColor(0x0D,0x15,0x2B); C_TEAL=RGBColor(0x03,0x8D,0xBD)

C_LTBL=RGBColor(0x4D,0xB2,0xDE); C_WHITE=RGBColor(0xFF,0xFF,0xFF)

C_GREEN=RGBColor(0x00,0xB0,0x50); C_RED=RGBColor(0xC0,0x00,0x00)

C_AMBER=RGBColor(0xFF,0xA5,0x00); C_LGRAY=RGBColor(0xF2,0xF5,0xF8)

C_DGRAY=RGBColor(0x55,0x65,0x77); C_MGRAY=RGBColor(0xCC,0xCC,0xCC)

_NS='http://schemas.openxmlformats.org/drawingml/2006/main'


def _R(slide,x,y,w,h,fill,border=False):

    s=slide.shapes.add_shape(1,int(x),int(y),int(w),int(h))

    s.fill.solid(); s.fill.fore_color.rgb=fill

    if border: s.line.color.rgb=C_MGRAY; s.line.width=Pt(0.5)

    else: s.line.fill.background()

    return s


def _T(slide,txt,x,y,w,h,size,bold=False,color=C_NAVY,align=PP_ALIGN.LEFT,wrap=True):

    tb=slide.shapes.add_textbox(int(x),int(y),int(w),int(h))

    tf=tb.text_frame; tf.word_wrap=wrap

    p=tf.paragraphs[0]; p.alignment=align

    r=p.add_run(); r.text=str(txt)

    r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color

    return tb


def _Tm(slide,lines,x,y,w,h,size=17,color=C_NAVY,autofit=False):

    tb=slide.shapes.add_textbox(int(x),int(y),int(w),int(h))

    tf=tb.text_frame; tf.word_wrap=True; first=True

    for line in (lines if isinstance(lines,list) else [lines]):

        p=tf.paragraphs[0] if first else tf.add_paragraph(); first=False

        p.space_after=Pt(4); r=p.add_run(); r.text=str(line)

        r.font.size=Pt(size); r.font.color.rgb=color

    if autofit:

        tf.auto_size=MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    return tb


def _footer(sl,SW,SH):

    _R(sl,0,SH-int(SH*.025),SW,int(SH*.005),C_TEAL)

    _T(sl,"VariTec Consulting  |  Confidential",int(SW*.02),SH-int(SH*.025),int(SW*.5),int(SH*.025),13,color=C_DGRAY)


def _hdr(sl,SW,SH,title,sub=""):

    _R(sl,0,0,SW,int(SH*.105),C_NAVY); _R(sl,0,0,int(SW*.007),SH,C_TEAL)

    _T(sl,title,int(SW*.02),int(SH*.015),int(SW*.72),int(SH*.075),34,bold=True,color=C_WHITE)

    if sub: _T(sl,sub,int(SW*.02),int(SH*.072),int(SW*.72),int(SH*.035),18,color=C_LTBL)


def _extract(html):

    soup=BeautifulSoup(html or "","html.parser"); lines=[]; images=[]; seen_imgs=set()

    for elem in soup.descendants:

        if elem.name=="img":

            src=elem.get("src","")

            if src.startswith("data:image") and src not in seen_imgs: images.append(src); seen_imgs.add(src)

        elif elem.name in ("p","div","li","h1","h2","h3"):

            t=elem.get_text(separator=" ").strip()

            if t: lines.append(t)

        elif elem.name is None:

            p=elem.parent

            if p and p.name not in ("script","style","p","div","li","span","b","i","u","strong","em","a"):

                t=str(elem).strip()

                if t: lines.append(t)

    return lines,images


def _parse_scorecard(lines):

    weeks,cur_wk,cur_ent=[],None,[]

    wr=re.compile(r'^week\s*(\d+)',re.I); sr=re.compile(r'^(.+?)\s+(100\s*%?|\d{2,3}\s*%?)\s*$',re.I)

    def flush():

        if cur_wk and cur_ent: weeks.append({"week":cur_wk,"entries":list(cur_ent)})

    for raw in lines:

        line=raw.strip()

        if not line: continue

        wm=wr.match(line)

        if wm: flush(); cur_wk=f"Week {wm.group(1)}"; cur_ent=[]

        elif re.match(r'^[-=]{3,}',line): pass

        else:

            sm=sr.match(line)

            if sm and cur_wk is not None:

                c=sm.group(1).strip().title(); s=sm.group(2).strip()

                if "%" not in s: s+="%"

                cur_ent.append({"customer":c,"score":s})

    flush(); return weeks


def _parse_issues(lines):

    full="\n".join(lines); blocks=re.split(r'\n\s*\d+\.\s+',"\n"+full); parsed=[]

    for body in blocks[1:]:

        iss={"title":"","issue":"","root_cause":"","status":"","actions":[]}

        bls=body.split("\n")

        for bl in bls:

            bl=bl.strip()

            if bl: iss["title"]=bl[6:].strip() if bl.lower().startswith("title:") else bl; break

        sec=None

        for bl in bls[1:]:

            bs=bl.strip(); lo=bs.lower()

            if lo.startswith("issue:"): sec="issue"; v=re.sub(r'^issue\s*[:\s]','',bs,flags=re.I).strip(); iss["issue"]=v if v else iss["issue"]

            elif lo.startswith("root cause"): sec="root_cause"; v=re.sub(r'^root cause\s*[:\s]','',bs,flags=re.I).strip(); iss["root_cause"]=v if v else iss["root_cause"]

            elif lo.startswith("status"): sec="status"; v=re.sub(r'^status\s*[:\s]','',bs,flags=re.I).strip(); iss["status"]=v if v else iss["status"]

            elif lo.startswith("action"): sec="actions"

            elif bs.startswith(("•","–","-","▸","*")) and sec=="actions": iss["actions"].append(bs.lstrip("•–-▸* ").strip())

            elif sec=="issue" and bs: iss["issue"]+=(" "+bs) if iss["issue"] else bs

            elif sec=="root_cause" and bs: iss["root_cause"]+=(" "+bs) if iss["root_cause"] else bs

            elif sec=="actions" and len(bs)>3: iss["actions"].append(bs.lstrip("•–-▸* ").strip())

        if iss["title"]: parsed.append(iss)

    return parsed


def _detect(title,lines):

    text=(title+" "+" ".join(lines)).lower()

    if re.search(r'week\s*\d',text) and re.search(r'\d{2,3}\s*%',text): return "scorecard"

    if re.search(r'(root cause|issue:|status:|action)',text,re.I): return "issues"

    return "general"


def _smart_replace(element,old,new):

    for t in element.findall(f'.//{{{_NS}}}t'):

        if t.text and old in t.text: t.text=t.text.replace(old,new); return True

    for para in element.findall(f'.//{{{_NS}}}p'):

        ts=para.findall(f'.//{{{_NS}}}t'); combined=''.join(t.text or '' for t in ts)

        if old in combined:

            nt=combined.replace(old,new)

            for i,t in enumerate(ts): t.text=nt if i==0 else ''

            return True

    return False


def _clear_xml(element):

    for t in element.findall(f'.//{{{_NS}}}t'):

        if t.text: t.text=''


def _fill_issue_group(group,issue):

    shapes=[]

    def collect(g):

        for s in g.shapes:

            if s.has_text_frame and s.text_frame.text.strip(): shapes.append((s.top,s))

            if s.shape_type==6: collect(s)

    collect(group); shapes.sort(key=lambda x:x[0])

    mapping=[issue["title"][:70],issue["root_cause"][:150],f"Issue: {issue['issue'][:280]}"]

    for i,(_,shape) in enumerate(shapes[:3]):

        _clear_xml(shape._element)

        ts=shape._element.findall(f'.//{{{_NS}}}t')

        if ts: ts[0].text=mapping[i] if i<len(mapping) else ''


def _generate_from_template(primary_note,notes_list):

    NS_R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    today=datetime.now().strftime("%b-%Y")

    prs=Presentation(TPL_PATH); orig=len(prs.slides); blank=prs.slide_layouts[6]

    def clone(idx):

        src=prs.slides[idx]; new=prs.slides.add_slide(blank)

        sp=new.shapes._spTree

        for c in list(sp)[2:]: sp.remove(c)

        for c in list(src.shapes._spTree)[2:]: sp.append(copy.deepcopy(c))

        return new

    def fr(slide,old,new_text): _smart_replace(slide._element,old,new_text)

    sc_data,iss_data,gen_data=[],[],[]

    for note in notes_list:

        lines,imgs=_extract(note.content); nt=_detect(note.title,lines)

        if nt=="scorecard":

            wks=_parse_scorecard(lines)

            if wks: sc_data.append((note,wks))

            else: gen_data.append((note,lines,imgs))

        elif nt=="issues":

            iss=_parse_issues(lines)

            if iss: iss_data.append((note,iss))

            else: gen_data.append((note,lines,imgs))

        else: gen_data.append((note,lines,imgs))

    all_issues=[i for _,iss in iss_data for i in iss]

    all_weeks=[wk for _,wks in sc_data for wk in wks]

    cov=clone(0)

    fr(cov,"Monthly Operations & Support Review",primary_note.section)

    fr(cov,"Support Review","")

    fr(cov,"May-2026",today)

    fr(cov,"Presented by: VariTec Consulting",f"VariTec Consulting  |  {len(notes_list)} module(s)  |  {today}")

    summ=clone(2)

    c_stubs=["LSL aggregation & interoperability gaps","Drake Pharmacy GLN/SGLN configuration mismatch under investigation with TL NSM team"]

    a_stubs=["100% scorecard performance maintained across all customers (Cencora, MWI, Cardinal, McKesson)","All in-month serialization issues resolved with no downstream compliance impact","Successful RMA and corrective actions completed for CVS, McKesson, Capital Wholesale & Cencora exceptions"]

    f_stubs=["Support LSL in fully aligning physical packaging with EPCIS aggregation hierarchy","Resolve Drake Pharmacy GLN configuration through TraceLink NSM team","Drive standardized decommissioned case handling procedures to prevent recurrence","Completion of Clipper and Unit dose Onboarding and testing."]

    our_c=[i["title"][:80] for i in all_issues[:2]]

    our_a=[f"{w['week']}: {', '.join(e['customer'] for e in w['entries'][:3])} – {w['entries'][0]['score'] if w['entries'] else ''}" for w in all_weeks[:3]]

    our_f=[f"{n.title}: {ls[0][:50] if ls else ''}" for n,ls,*_ in gen_data[:4]]

    def fill_stubs(stubs,values):

        pad=["—"]*len(stubs); merged=(values+pad)[:len(stubs)]

        for s,v in zip(stubs,merged): fr(summ,s,v or "—")

    fill_stubs(c_stubs,our_c); fill_stubs(a_stubs,our_a); fill_stubs(f_stubs,our_f)

    if all_issues:

        for bs in range(0,len(all_issues),4):

            batch=all_issues[bs:bs+4]; iss_slide=clone(5)

            fr(iss_slide,"Serialization Issues",f"Issues  {bs+1}–{bs+len(batch)}  of  {len(all_issues)}")

            fr(iss_slide,"All issues were resolved successfully with no impact to downstream shipments or compliance",

               f"Total: {len(all_issues)} issue(s)  |  {sum(1 for i in all_issues if any(w in i.get('status','').lower() for w in('closed','complet','resolved')))} resolved")

            groups=sorted([s for s in iss_slide.shapes if s.shape_type==6],key=lambda s:s.left)

            for gi,grp in enumerate(groups):

                if gi<len(batch): _fill_issue_group(grp,batch[gi])

                else: _clear_xml(grp._element)

    if all_weeks:

        sc_slide=clone(7); _smart_replace(sc_slide._element,"MWI Animal Health Cencora","All Customers")

        wl=[w["week"] for w in all_weeks]

        fr(sc_slide,"Week 2",wl[0] if wl else "Week 1")

        fr(sc_slide,"Week 4",wl[1] if len(wl)>1 else "Week 2")

    for note,lines,imgs in gen_data:

        gen_slide=clone(2); fr(gen_slide,"Executive Summary",note.title)

        for s in c_stubs+a_stubs+f_stubs: fr(gen_slide,s,"")

    clone(orig-1)

    xml_slides=prs.slides._sldIdLst

    for ref in list(xml_slides)[:orig]:

        rId=ref.get(f'{{{NS_R}}}id')

        try: prs.part.drop_rel(rId)

        except: pass

        xml_slides.remove(ref)

    out=io.BytesIO(); prs.save(out); out.seek(0); return out


def _bi_cover(prs,blank,section,total):

    SW,SH=int(prs.slide_width),int(prs.slide_height); sl=prs.slides.add_slide(blank)

    _R(sl,0,0,SW,SH,C_NAVY); _R(sl,0,0,int(SW*.012),SH,C_TEAL); _R(sl,0,SH-int(SH*.08),SW,int(SH*.08),C_TEAL)

    _T(sl,section.upper(),int(SW*.06),int(SH*.28),int(SW*.88),int(SH*.07),22,bold=True,color=C_LTBL,align=PP_ALIGN.CENTER)

    _R(sl,int(SW*.3),int(SH*.36),int(SW*.4),int(SH*.004),C_TEAL)

    _T(sl,section,int(SW*.06),int(SH*.38),int(SW*.88),int(SH*.22),58,bold=True,color=C_WHITE,align=PP_ALIGN.CENTER)

    _T(sl,f"{total} sub-module(s)  ·  VariTec Notes  ·  {datetime.now().strftime('%b %Y')}",int(SW*.06),int(SH*.63),int(SW*.88),int(SH*.06),22,color=C_LTBL,align=PP_ALIGN.CENTER)

    _T(sl,"VariTec Consulting  |  Confidential",int(SW*.06),SH-int(SH*.06),int(SW*.88),int(SH*.05),18,color=C_WHITE,align=PP_ALIGN.CENTER)


def _bi_divider(prs,blank,idx,total,title,ntype,sub=""):

    SW,SH=int(prs.slide_width),int(prs.slide_height); sl=prs.slides.add_slide(blank)

    _R(sl,0,0,SW,SH,C_NAVY); _R(sl,0,0,int(SW*.012),SH,C_TEAL)

    icons={"scorecard":("📊","SCORECARD",C_LTBL),"issues":("⚠️","ISSUES",C_AMBER),"onboarding":("📋","ONBOARDING",C_LTBL),"general":("📄","NOTES",C_LTBL)}

    ic,lbl,col=icons.get(ntype,("📄","NOTES",C_LTBL))

    bw=int(SW*.10); bh=int(SH*.06)

    _R(sl,SW-bw-int(SW*.03),int(SH*.03),bw,bh,C_TEAL)

    _T(sl,f"{idx}/{total}",SW-bw-int(SW*.03),int(SH*.03),bw,bh,20,bold=True,color=C_WHITE,align=PP_ALIGN.CENTER)

    _T(sl,f"{ic}  {lbl}",int(SW*.06),int(SH*.30),int(SW*.88),int(SH*.07),22,bold=True,color=col)

    _R(sl,int(SW*.06),int(SH*.40),int(SW*.30),int(SH*.005),C_TEAL)

    _T(sl,title,int(SW*.06),int(SH*.42),int(SW*.88),int(SH*.22),52,bold=True,color=C_WHITE)

    if sub: _T(sl,sub,int(SW*.06),int(SH*.66),int(SW*.88),int(SH*.06),22,color=C_LTBL)

    _R(sl,0,SH-int(SH*.07),SW,int(SH*.07),C_TEAL)


def _bi_scorecard(prs,blank,title,weeks):

    SW,SH=int(prs.slide_width),int(prs.slide_height); sl=prs.slides.add_slide(blank)

    _R(sl,0,0,SW,SH,C_WHITE); _hdr(sl,SW,SH,title,f"{len(weeks)} week(s)")

    n=len(weeks); mg_x=int(SW*.02); mg_y=int(SH*.14)

    col_w=(SW-2*mg_x)//n; gap=int(SW*.006); hdr_h=int(SH*.082); row_h=int(SH*.082); row_g=int(SH*.006)

    customers=[]

    for wk in weeks:

        for e in wk["entries"]:

            if e["customer"] not in customers: customers.append(e["customer"])

    for ci,wk in enumerate(weeks):

        cx=mg_x+ci*col_w; cy=mg_y

        _R(sl,cx,cy,col_w-gap,hdr_h,C_NAVY)

        _T(sl,wk["week"],cx+int(SW*.003),cy+int(SH*.015),col_w-gap-int(SW*.006),hdr_h-int(SH*.015),24,bold=True,color=C_WHITE,align=PP_ALIGN.CENTER)

        wk_sc={e["customer"]:e["score"] for e in wk["entries"]}

        for ri,cust in enumerate(customers):

            ry=cy+hdr_h+int(SH*.008)+ri*(row_h+row_g)

            bg=C_LGRAY if ri%2==0 else C_WHITE; _R(sl,cx,ry,col_w-gap,row_h,bg,border=True)

            sc=wk_sc.get(cust)

            if sc:

                try: pc=C_GREEN if float(sc.replace("%",""))>=100 else(C_AMBER if float(sc.replace("%",""))>=95 else C_RED)

                except: pc=C_TEAL

                pw=int(SW*.045); px=cx+col_w-gap-pw-int(SW*.004); py=ry+int(SH*.014); ph=row_h-int(SH*.028)

                _R(sl,px,py,pw,ph,pc)

                _T(sl,sc,px,py,pw,ph,15,bold=True,color=C_WHITE,align=PP_ALIGN.CENTER)

                _T(sl,cust,cx+int(SW*.008),ry+int(SH*.02),px-cx-int(SW*.01),row_h-int(SH*.01),18,color=C_NAVY)

            else: _T(sl,f"{cust}  —",cx+int(SW*.008),ry+int(SH*.02),col_w-gap-int(SW*.016),row_h-int(SH*.01),16,color=C_MGRAY)

    _footer(sl,SW,SH)


def _bi_issues(prs,blank,title,chunk,offset,total):

    SW,SH=int(prs.slide_width),int(prs.slide_height); sl=prs.slides.add_slide(blank)

    _R(sl,0,0,SW,SH,C_WHITE); _hdr(sl,SW,SH,title,f"Issues {offset+1}-{offset+len(chunk)} of {total}")

    n=len(chunk); mg_x=int(SW*.018); mg_y=int(SH*.125)

    gap=int(SW*.012); col_w=(SW-2*mg_x-gap*(n-1))//n; col_h=SH-mg_y-int(SH*.06)

    hc=[C_NAVY,C_TEAL,C_NAVY,C_LTBL]

    def sc(s): sl2=(s or "").lower(); return C_GREEN if any(w in sl2 for w in("closed","complet","resolved")) else(C_AMBER if any(w in sl2 for w in("progress","open","pending")) else C_TEAL)

    for ci,iss in enumerate(chunk):

        cx=mg_x+ci*(col_w+gap); cy=mg_y; h=hc[ci%4]

        _R(sl,cx,cy,col_w,col_h,C_LGRAY,border=True)

        hh=int(SH*.085); _R(sl,cx,cy,col_w,hh,h)

        bw=int(SW*.032); bh=int(SH*.048); bx=cx+int(SW*.006); by=cy+int(SH*.018)

        _R(sl,bx,by,bw,bh,C_WHITE)

        _T(sl,f"#{offset+ci+1}",bx,by,bw,bh,14,bold=True,color=h,align=PP_ALIGN.CENTER)

        _T(sl,iss["title"],bx+bw+int(SW*.006),cy+int(SH*.01),cx+col_w-bx-bw-int(SW*.012),hh-int(SH*.01),17,bold=True,color=C_WHITE,wrap=True)

        if iss.get("status"):

            pw=int(SW*.09); ph=int(SH*.038); px=cx+col_w-pw-int(SW*.006); py=cy+hh+int(SH*.01)

            _R(sl,px,py,pw,ph,sc(iss["status"]))

            _T(sl,iss["status"][:18],px,py,pw,ph,13,bold=True,color=C_WHITE,align=PP_ALIGN.CENTER)

        secs=[("📋 Issue",iss.get("issue",""),C_TEAL),("🔍 Root Cause",iss.get("root_cause",""),C_NAVY),("✅ Actions","\n".join(f"▸ {a}" for a in iss.get("actions",[]) if a),C_LTBL)]

        sy=cy+hh+int(SH*.06); sg=int(SH*.008); sp2=int(SW*.007); remain=col_h-(sy-cy)-int(SH*.01)

        lh=int(SH*.042)

        # Estimate content height per section based on character count (rough chars-per-line for this width)

        chars_per_line=max(20,int((col_w-2*sp2-int(SW*.01))/(size_to_emu_ratio:=Pt(13)*9000)))

        weights=[]

        for _,sbody,_ in secs:

            lines_count=max(1,sum(max(1,(len(l)//40)+1) for l in sbody.split("\n") if l.strip()) or 1)

            weights.append(lines_count)

        total_w=sum(weights) or 1

        for si,(slbl,sbody,scol) in enumerate(secs):

            sh2=max(int(SH*.10), int(remain*weights[si]/total_w))

            if si==len(secs)-1:

                sh2=remain-sum(max(int(SH*.10),int(remain*w/total_w)) for w in weights[:-1])

            sy2=sy+sum(max(int(SH*.10),int(remain*weights[k]/total_w)) if k<len(secs)-1 else 0 for k in range(si))+si*sg

            _R(sl,cx+sp2,sy2,col_w-2*sp2,lh,scol)

            _T(sl,slbl,cx+sp2+int(SW*.005),sy2+int(SH*.007),col_w-2*sp2-int(SW*.01),lh-int(SH*.007),15,bold=True,color=C_WHITE)

            by2=sy2+lh+int(SH*.005); bh2=sh2-lh-int(SH*.005)

            _R(sl,cx+sp2,by2,col_w-2*sp2,bh2,C_WHITE,border=True)

            _Tm(sl,[l for l in sbody.split("\n") if l.strip()],cx+sp2+int(SW*.005),by2+int(SH*.008),col_w-2*sp2-int(SW*.01),bh2-int(SH*.01))

    _footer(sl,SW,SH)


def _bi_general(prs,blank,title,lines,images=None,page_num=1):

    SW,SH=int(prs.slide_width),int(prs.slide_height); sl=prs.slides.add_slide(blank)

    _R(sl,0,0,SW,SH,C_WHITE); _hdr(sl,SW,SH,title if page_num==1 else f"{title} (cont. {page_num})")

    cy=int(SH*.13); ch=SH-cy-int(SH*.06)

    if images and page_num==1:

        tw=int(SW*.50); _R(sl,int(SW*.02),cy,tw,ch,C_LGRAY,border=True)

        _Tm(sl,[l for l in lines if l.strip()],int(SW*.03),cy+int(SH*.015),tw-int(SW*.02),ch-int(SH*.02))

        try:

            _,b64=images[0].split(",",1); sl.shapes.add_picture(io.BytesIO(base64.b64decode(b64)),int(SW*.53),cy,int(SW*.44),ch)

        except: pass

    else:

        _R(sl,int(SW*.02),cy,SW-int(SW*.04),ch,C_LGRAY,border=True)

        cl=[l for l in lines if l.strip()]

        if len(cl)>15:

            mid=len(cl)//2; cw=(SW-int(SW*.06))//2

            _Tm(sl,cl[:mid],int(SW*.03),cy+int(SH*.015),cw-int(SW*.01),ch-int(SH*.02))

            _Tm(sl,cl[mid:],int(SW*.03)+cw,cy+int(SH*.015),cw-int(SW*.01),ch-int(SH*.02))

        else: _Tm(sl,cl,int(SW*.03),cy+int(SH*.015),SW-int(SW*.06),ch-int(SH*.02))

    _footer(sl,SW,SH)


def _parse_onboarding_html(html):

    """Parse onboarding template HTML into structured data for PPT rendering."""

    soup = BeautifulSoup(html or "", "html.parser")

    header_type = "customer"

    for cls in ['ob-header-mah','ob-header-vendor','ob-header-customer']:

        if soup.find(class_=cls):

            header_type = cls.replace('ob-header-',''); break

    title_el = soup.find(class_='ob-header-title')

    title = ' '.join(title_el.get_text(' ').split()) if title_el else "Onboarding"

    month_el = soup.find(class_='ob-month-val')

    month = month_el.get_text().strip() if month_el else ""

    customers = {}

    tbody = soup.find('tbody')

    if tbody:

        for row in tbody.find_all('tr', attrs={'data-cg': True}):

            cg = row.get('data-cg','')

            if cg not in customers:

                customers[cg] = {'name': f'Customer {int(cg)+1}', 'steps': []}

            cells = row.find_all('td', recursive=False)

            name_td = row.find('td', class_='sc-name-td')

            if name_td:

                txt = name_td.get_text().strip()

                if txt and txt not in ('\xa0','&nbsp;'):

                    customers[cg]['name'] = txt

                step_cells = [c for c in cells if c is not name_td]

            else:

                step_cells = cells

            if len(step_cells) >= 2:

                s0 = step_cells[0]

                step_span = s0.find('span', style=lambda s: s and 'font-weight:500' in s)

                step_text = (step_span.get_text().strip() if step_span else s0.get_text().strip())[:80]

                status_btn = step_cells[1].find(class_='ob-status-btn') if len(step_cells) > 1 else None

                status = status_btn.get_text().strip() if status_btn else ""

                def _td(idx):

                    if len(step_cells) > idx:

                        t = step_cells[idx].get_text().strip()

                        return "" if t in ('\xa0','&nbsp;','') else t

                    return ""

                target = _td(2); completion = _td(3); comment = _td(4)

                customers[cg]['steps'].append({'step':step_text,'status':status,

                    'target':target,'completion':completion,'comment':comment})

    return {'title':title,'month':month,'header_type':header_type,

            'customers':list(customers.values())}


def _bi_onboarding(prs, blank, note_title, ob_data):

    """Render onboarding checklist as compact, page-fitting PPT slide(s)."""

    TYPE_COLORS = {

        'customer': RGBColor(0x1E,0x3A,0x8A),

        'vendor':   RGBColor(0x06,0x4E,0x3B),

        'mah':      RGBColor(0x7C,0x2D,0x12),

    }

    ST_COLORS = {

        'completed':  C_GREEN,

        'in progress': C_AMBER,

        'pending':    RGBColor(0xF5,0x9E,0x0B),

        'hold':       C_RED,

    }

    h_col = TYPE_COLORS.get(ob_data.get('header_type','customer'), C_NAVY)


    # Flatten all rows: (customer_name, step_idx, total_steps, step_dict)

    all_rows = []

    for cust in ob_data.get('customers', []):

        steps = cust.get('steps', [])

        for si, step in enumerate(steps):

            all_rows.append((cust.get('name','Customer'), si, len(steps), step))


    SW, SH = int(prs.slide_width), int(prs.slide_height)

    MG_X = int(SW * 0.018)

    MG_Y_TOP = int(SH * 0.14)

    MG_Y_BOT = int(SH * 0.06)

    TW = SW - 2 * MG_X

    TH_H = int(SH * 0.058)   # column header row

    ROW_H = int(SH * 0.072)  # data row


    AVAIL_H = SH - MG_Y_TOP - TH_H - MG_Y_BOT

    ROWS_PER_SLIDE = max(1, int(AVAIL_H / ROW_H))


    # Column proportions: Customer 14%, CheckPoint 34%, Status 12%, Target 11%, Completion 13%, Comment 16%

    col_pcts = [0.14, 0.34, 0.12, 0.11, 0.13, 0.16]

    col_ws = [int(TW * p) for p in col_pcts]

    col_ws[-1] = TW - sum(col_ws[:-1])  # absorb rounding remainder

    COL_HDRS = ['Customer','Check Point','Status','Target Date','Completion Date','Comment']


    def _draw_slide(batch, page_num, total_pages):

        sl = prs.slides.add_slide(blank)

        _R(sl, 0, 0, SW, SH, C_WHITE)

        sub = ob_data.get('month','')

        if total_pages > 1:

            sub = f"{sub}  ·  Part {page_num}/{total_pages}".strip(' ·')

        _hdr(sl, SW, SH, note_title, sub)


        # Column headers

        cx = MG_X

        for i, (hdr, cw) in enumerate(zip(COL_HDRS, col_ws)):

            _R(sl, cx, MG_Y_TOP, cw, TH_H, h_col)

            _T(sl, hdr, cx+int(SW*.004), MG_Y_TOP+int(SH*.01),

               cw-int(SW*.008), TH_H-int(SH*.01),

               10, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

            cx += cw


        # Data rows

        cy = MG_Y_TOP + TH_H

        prev_cust = None

        for cust_name, step_idx, total_steps, step in batch:

            bg = C_LGRAY if (cy - MG_Y_TOP - TH_H) // ROW_H % 2 == 0 else C_WHITE

            cx = MG_X


            # Customer name cell — draw only on first step of this customer in this batch

            if cust_name != prev_cust:

                # count how many rows this customer still has in batch

                rows_left = sum(1 for r in batch if r[0] == cust_name)

                c_h = rows_left * ROW_H

                _R(sl, cx, cy, col_ws[0], c_h, C_WHITE, border=True)

                _T(sl, cust_name, cx+int(SW*.003), cy+int(SH*.005),

                   col_ws[0]-int(SW*.006), c_h-int(SH*.005),

                   10, bold=True, color=h_col, align=PP_ALIGN.CENTER)

                prev_cust = cust_name

            cx += col_ws[0]


            # Checkpoint

            _R(sl, cx, cy, col_ws[1], ROW_H, bg, border=True)

            _T(sl, step.get('step',''), cx+int(SW*.004), cy+int(SH*.008),

               col_ws[1]-int(SW*.008), ROW_H-int(SH*.01), 9, color=C_NAVY)

            cx += col_ws[1]


            # Status badge

            _R(sl, cx, cy, col_ws[2], ROW_H, bg, border=True)

            st_text = step.get('status','')

            if st_text:

                st_col = ST_COLORS.get(st_text.lower(), C_TEAL)

                bw2 = int(col_ws[2]*0.82); bh2 = int(ROW_H*0.55)

                bx2 = cx + (col_ws[2]-bw2)//2; by2 = cy + (ROW_H-bh2)//2

                _R(sl, bx2, by2, bw2, bh2, st_col)

                _T(sl, st_text[:10], bx2, by2, bw2, bh2, 8, bold=True,

                   color=C_WHITE, align=PP_ALIGN.CENTER)

            cx += col_ws[2]


            # Target date

            _R(sl, cx, cy, col_ws[3], ROW_H, bg, border=True)

            _T(sl, step.get('target',''), cx+int(SW*.003), cy+int(SH*.008),

               col_ws[3]-int(SW*.006), ROW_H-int(SH*.01), 9,

               color=C_DGRAY, align=PP_ALIGN.CENTER)

            cx += col_ws[3]


            # Completion date

            _R(sl, cx, cy, col_ws[4], ROW_H, bg, border=True)

            _T(sl, step.get('completion',''), cx+int(SW*.003), cy+int(SH*.008),

               col_ws[4]-int(SW*.006), ROW_H-int(SH*.01), 9,

               color=C_DGRAY, align=PP_ALIGN.CENTER)

            cx += col_ws[4]


            # Comment

            _R(sl, cx, cy, col_ws[5], ROW_H, bg, border=True)

            _T(sl, step.get('comment','')[:45], cx+int(SW*.003), cy+int(SH*.008),

               col_ws[5]-int(SW*.006), ROW_H-int(SH*.01), 9, color=C_NAVY)


            cy += ROW_H

        _footer(sl, SW, SH)


    # Paginate and draw

    pages = [all_rows[i:i+ROWS_PER_SLIDE] for i in range(0, max(len(all_rows),1), ROWS_PER_SLIDE)]

    for pg_num, batch in enumerate(pages, 1):

        _draw_slide(batch, pg_num, len(pages))


def _generate_builtin(primary_note,notes_list):

    prs=Presentation(); prs.slide_width=24022050; prs.slide_height=13716000

    SW,SH=int(prs.slide_width),int(prs.slide_height); blank=prs.slide_layouts[6]; total=len(notes_list)

    _bi_cover(prs,blank,primary_note.section,total)

    def _is_onboarding(note):

        c = note.content or ''

        return any(x in c for x in ['ob-header-customer','ob-header-vendor','ob-header-mah'])


    if total>1:

        sl=prs.slides.add_slide(blank); _R(sl,0,0,SW,SH,C_WHITE); _hdr(sl,SW,SH,"Table of Contents")

        icons={"scorecard":"📊","issues":"⚠️","onboarding":"📋","general":"📄"}

        mg_x=int(SW*.04); mg_y=int(SH*.14); ih=int(SH*.095); ig=int(SH*.012)

        cols=2 if total>5 else 1; per=(total+cols-1)//cols; cw=(SW-2*mg_x-int(SW*.02))//cols

        for ci in range(cols):

            cx=mg_x+ci*(cw+int(SW*.02)); start=ci*per; end=min(start+per,total)

            for ri,note in enumerate(notes_list[start:end]):

                iy=mg_y+ri*(ih+ig); ls,_=_extract(note.content)

                nt="onboarding" if _is_onboarding(note) else _detect(note.title,ls)

                ic=icons.get(nt,"📄")

                nc=C_TEAL if nt=="scorecard" else(C_AMBER if nt in("issues","onboarding") else C_NAVY)

                nw=int(SW*.038); _R(sl,cx,iy,nw,ih,nc)

                _T(sl,str(start+ri+1),cx,iy,nw,ih,18,bold=True,color=C_WHITE,align=PP_ALIGN.CENTER)

                _R(sl,cx+nw+int(SW*.004),iy,cw-nw-int(SW*.004),ih,C_LGRAY,border=True)

                _T(sl,f"{ic}  {note.title}",cx+nw+int(SW*.012),iy+int(SH*.022),cw-nw-int(SW*.02),ih-int(SH*.01),20,color=C_NAVY)

        _footer(sl,SW,SH)

    for idx,note in enumerate(notes_list,1):

        lines,images=_extract(note.content)

        # Detect onboarding BEFORE generic detect (HTML templates won't match text patterns)

        if _is_onboarding(note):

            ntype="onboarding"

            ob_data=_parse_onboarding_html(note.content)

            total_cust=len(ob_data.get('customers',[]))

            total_steps=sum(len(c.get('steps',[])) for c in ob_data.get('customers',[]))

            sub=f"{total_cust} customer(s) · {total_steps} step(s)"

        else:

            ntype=_detect(note.title,lines)

            if ntype=="scorecard": wks=_parse_scorecard(lines); sub=f"{len(wks)} week(s)"

            elif ntype=="issues": iss=_parse_issues(lines); sub=f"{len(iss)} issue(s)"

            else: sub=f"{len([l for l in lines if l.strip()])} line(s)"

        _bi_divider(prs,blank,idx,total,note.title,ntype,sub)

        if ntype=="onboarding":

            _bi_onboarding(prs,blank,note.title,ob_data)

        elif ntype=="scorecard":

            wks=_parse_scorecard(lines)

            if wks:

                for wi in range(0,len(wks),4): _bi_scorecard(prs,blank,note.title,wks[wi:wi+4])

            else: _bi_general(prs,blank,note.title,lines,images)

        elif ntype=="issues":

            iss=_parse_issues(lines)

            if iss:

                for ii in range(0,len(iss),4): _bi_issues(prs,blank,note.title,iss[ii:ii+4],ii,len(iss))

            else: _bi_general(prs,blank,note.title,lines,images)

        else:

            cl=[l for l in lines if l.strip()]; ps=20

            for pi,pg in enumerate([cl[i:i+ps] for i in range(0,max(len(cl),1),ps)]):

                _bi_general(prs,blank,note.title,pg,images if pi==0 else None,pi+1)

    out=io.BytesIO(); prs.save(out); out.seek(0); return out



def _build_roles_responsibilities_html():
    """Roles & Responsibilities compact horizontal card grid (v2 design)."""
    RR_COLORS = [
        ('#3b82f6', '#93c5fd'),   # 1 blue
        ('#10b981', '#6ee7b7'),   # 2 green
        ('#f59e0b', '#fcd34d'),   # 3 amber
        ('#06b6d4', '#67e8f9'),   # 4 cyan
        ('#8b5cf6', '#c4b5fd'),   # 5 purple
        ('#ef4444', '#fca5a5'),   # 6 red
        ('#0ea5e9', '#7dd3fc'),   # 7 sky
        ('#f97316', '#fdba74'),   # 8 orange
    ]
    # (title, description, bold) — bold cards show a small description line below
    DEFAULTS = [
        ("Conduct weekly reconciliation of outbound shipments from 3PL to customers.",
         "", False),
        ("Monitor and resolve information exchange failures.",
         "", False),
        ("Respond promptly to customer queries related to EPCIS data.",
         "", False),
        ("Review and respond to VRS inbox messages daily.",
         "", False),
        ("Re-trigger failed or incorrect delivery events as required.",
         "", False),
        ("Coordinate with the 3PL team to resolve outbound shipment-related issues.",
         "", False),
        ("Communicate serialization-related issues to customers as needed.",
         "", False),
    ]

    cards_html = ""
    for i, (title, desc, bold) in enumerate(DEFAULTS):
        accent, accent_light = RR_COLORS[i % len(RR_COLORS)]
        num = str(i + 1).zfill(2)
        title_cls = "rr-card-title"  # uniform weight for all cards
        empty_desc_cls = " rr-empty-desc" if bold else ""
        cards_html += (
            f'<div class="rr-card{empty_desc_cls}" data-rr-card="{i}" '
            f'style="--rr-accent:{accent};--rr-accent-light:{accent_light};">'
            f'<button class="rr-del-card" type="button" onclick="rrDeleteCard(this)" '
            f'title="Delete card">&#10005;</button>'
            f'<div class="rr-num">{num}</div>'
            f'<div class="rr-card-body">'
            f'<div class="{title_cls}" contenteditable="true" '
            f'data-ph="Role or responsibility...">{title}</div>'
            f'<div class="rr-card-desc" contenteditable="true" '
            f'data-ph="Add a short note...">{desc}</div>'
            f'</div></div>'
        )

    banner = (
        '<div class="rr-banner">'
        '<div class="rr-banner-title">Roles &amp; Responsibilities</div>'
        '</div>'
    )
    add_btn = (
        '<button class="rr-add-btn" type="button" onclick="rrAddCard(this)">'
        '&#43; Add Card</button>'
    )
    return (
        '<div class="rr-wrap">'
        + banner
        + '<div class="rr-cards-area">'
        + '<div class="rr-cards-grid">'
        + cards_html
        + add_btn
        + '</div></div></div>'
    )



def _build_process_improvements_html():
    """Ongoing / Completed Process Improvements — no icons, text only."""
    PI_LABELS = {
        "pending": "Pending", "inprogress": "In Progress",
        "hold": "Hold", "completed": "Completed",
    }

    def _badge(status_key, custom_label=None):
        st = status_key or "pending"
        label = custom_label or PI_LABELS.get(st, "Pending")
        return (f'<span class="pi-pill pi-pill-{st}" data-status="{st}" '
                f'data-label="{label}">{label}</span>')

    def _status_cols(test_status, prod_status, test_label=None, prod_label=None):
        return (
            '<div class="pi-status-cols">'
            '<div class="pi-status-col"><span class="pi-status-label">Test</span>'
            f'{_badge(test_status, test_label)}</div>'
            '<div class="pi-status-col"><span class="pi-status-label">Production</span>'
            f'{_badge(prod_status, prod_label)}</div>'
            '</div>'
        )

    # (main text, [ (sub text, test_status, prod_status, prod_label), ... ],
    #  test_status_if_no_subs, prod_status_if_no_subs, prod_label_if_no_subs)
    DEFAULTS = [
        (
            "Successfully implemented automatic serial number status updates "
            "in TraceLink for the following scenarios:",
            [
                ("Void shipments", "completed", "pending", "Pending (First File)"),
                ("Salable returns", "completed", "completed", None),
                ("Samples", "completed", "completed", None),
                ("Destroy/Damage", "completed", "completed", None),
            ],
            None, None, None,
        ),
        (
            "Implemented data transmission to both Sold-To and Ship-To entities "
            "for drop-shipment scenarios",
            [],
            "completed", "pending", "Ready for Production",
        ),
    ]

    items_html = ""
    for i, (text, subs, own_test, own_prod, own_prod_label) in enumerate(DEFAULTS):
        has_content = bool(text.strip())
        main_icon_cls = "pi-main-icon" if has_content else "pi-main-icon pi-ghost"

        badge_html = (_status_cols(own_test, own_prod, None, own_prod_label)
                      if not subs else "")

        sub_rows = ""
        for si, (sub_text, t_status, p_status, p_label) in enumerate(subs):
            sub_rows += (
                '<div class="pi-sub-row">'
                '<div class="pi-sub-icon"></div>'
                '<div class="pi-sub-text" contenteditable="true" '
                f'data-ph="Sub-point...">{sub_text}</div>'
                f'{_status_cols(t_status, p_status, None, p_label)}'
                '<button class="pi-del-sub" type="button" '
                'title="Remove">&#10005;</button>'
                '</div>'
            )
        sub_html = (
            f'<div class="pi-sub-list">{sub_rows}</div>'
            '<div class="pi-add-sub-wrap">'
            '<button class="pi-add-sub" type="button" '
            'onclick="event.stopPropagation();var w=this.closest(\'.pi-add-sub-wrap\');var l=w?w.previousElementSibling:null;if(l)piAddSubRow(l,this);">'
            '&#43; Add Sub-point</button></div>'
        )
        items_html += (
            f'<div class="pi-item" data-pi-item="{i}">'
            f'<button class="pi-del-item" type="button" '
            f'title="Delete point">&#10005;</button>'
            f'<div class="pi-item-row">'
            f'<div class="{main_icon_cls}"></div>'
            f'<div class="pi-item-text" contenteditable="true" '
            f'data-ph="Process improvement point...">{text}</div>'
            f'{badge_html}'
            f'</div>'
            f'{sub_html}'
            f'</div>'
        )

    banner = (
        '<div class="pi-banner">'
        '<div class="pi-banner-title">Ongoing / Completed Process Improvements</div>'
        '</div>'
    )
    add_btn = ('<button class="pi-add-item" type="button" '
               'onclick="event.stopPropagation();piAddItemRow(this.closest(\'.pi-body\'),this);">'
               '&#43; Add Point</button>')
    return (
        '<div class="pi-wrap">'
        + banner
        + '<div class="pi-body">'
        + items_html
        + add_btn
        + '</div></div>'
    )


def generate_ppt(primary_note,notes_list=None):

    if not notes_list: notes_list=[primary_note]

    if os.path.exists(TPL_PATH):

        try: return _generate_from_template(primary_note,notes_list)

        except Exception as e:

            import traceback; traceback.print_exc()

            print(f"Template PPT failed ({e}), using builtin")

    return _generate_builtin(primary_note,notes_list)


LOGO_SRC = "/logo.png"


LOGIN_HTML="""<!DOCTYPE html><html><head><title>VariTec Notes</title>

<style>

*{margin:0;padding:0;box-sizing:border-box;font-family:'Segoe UI',system-ui,sans-serif;}

body{

  height:100vh;display:flex;align-items:center;justify-content:center;

  background:linear-gradient(135deg,#0d152b 0%,#1e3868 50%,#0a2a5e 100%);

  position:relative;overflow:hidden;

}

body::before{

  content:'';position:absolute;top:-120px;right:-120px;width:380px;height:380px;

  border-radius:50%;background:rgba(3,141,189,.18);

}

body::after{

  content:'';position:absolute;bottom:-150px;left:-100px;width:340px;height:340px;

  border-radius:50%;background:rgba(77,178,222,.12);

}

.login-shell{

  position:relative;z-index:1;display:flex;flex-direction:column;align-items:center;gap:22px;

}

.brand-logo{

  display:flex;align-items:center;gap:10px;color:white;

}

.brand-logo .mark{

  font-family:'Arial Black',Arial,sans-serif;font-size:24px;font-weight:900;

  letter-spacing:3px;text-shadow:0 0 20px rgba(56,189,248,.5);

}

.brand-logo .sub{

  font-size:12px;font-weight:300;color:#4db2de;letter-spacing:2px;

}

.card{

  background:white;padding:40px 38px 32px;border-radius:18px;

  box-shadow:0 20px 60px rgba(0,0,0,.35);width:380px;

  border:1px solid rgba(255,255,255,.08);

}

.card-header{

  display:flex;align-items:center;gap:14px;margin-bottom:22px;

}

.card-icon{

  width:52px;height:52px;border-radius:14px;flex-shrink:0;

  background:linear-gradient(135deg,#038dbd,#0284c7);

  display:flex;align-items:center;justify-content:center;

  font-size:24px;

  box-shadow:0 8px 20px rgba(3,141,189,.35);

}

.card-header-text{display:flex;flex-direction:column;justify-content:center;}

.card h2{

  font-size:22px;font-weight:800;color:#0d152b;line-height:1.2;

}

.card .tagline{

  font-size:13px;color:#94a3b8;margin-top:3px;

}

.field-label{

  font-size:12px;font-weight:700;color:#475569;text-transform:uppercase;

  letter-spacing:.6px;margin-bottom:6px;display:block;

}

input{

  width:100%;padding:13px 14px;margin-bottom:18px;

  border:1.5px solid #e2e8f0;border-radius:10px;font-size:14px;

  outline:none;transition:.15s;color:#1e293b;

}

input::placeholder{color:#cbd5e1;}

input:focus{

  border-color:#038dbd;box-shadow:0 0 0 4px rgba(3,141,189,.10);

}

button{

  width:100%;padding:13px;

  background:linear-gradient(135deg,#0d152b,#1e3868);

  color:white;border:none;border-radius:10px;font-size:14px;font-weight:700;

  letter-spacing:.4px;cursor:pointer;transition:.18s;

  box-shadow:0 6px 18px rgba(13,21,43,.25);

}

button:hover{

  background:linear-gradient(135deg,#038dbd,#0284c7);

  box-shadow:0 8px 22px rgba(3,141,189,.35);

  transform:translateY(-1px);

}

.err{

  color:#dc2626;margin-top:14px;font-size:13px;font-weight:600;

  background:#fee2e2;border:1px solid #fecaca;border-radius:8px;

  padding:10px 12px;text-align:center;

}

.footer-note{

  font-size:12px;color:rgba(255,255,255,.45);text-align:center;

  letter-spacing:.4px;

}

.signup-link{

  text-align:center;margin-top:16px;font-size:13px;color:#94a3b8;

}

.signup-link a{

  color:#038dbd;font-weight:700;text-decoration:none;

}

.signup-link a:hover{

  text-decoration:underline;

}

</style></head>

<body>

<div class="login-shell">

  

  <div class="card">

    <div class="card-header">

      <div class="card-icon">📝</div>

      <div class="card-header-text">

        <h2>VariTec Notes</h2>

        <div class="tagline">Sign in with your work email</div>

      </div>

    </div>

    <form method="post">

      <label class="field-label">Email</label>

      <input name="user_email_field" type="text" inputmode="email" placeholder="you@varitecconsulting.com" required autofocus autocomplete="username">

      <label class="field-label">Password</label>

      <input name="password" type="password" placeholder="Enter your password" required autocomplete="current-password">


      <button>Sign In</button>

      {% if error %}<div class="err">{{ error }}</div>{% endif %}

    </form>

    <div class="signup-link">New here? <a href="/signup">Request access</a></div>

  </div>

  <div class="footer-note">VariTec Consulting &middot; Confidential</div>

</div>

</body></html>"""


SIGNUP_HTML="""<!DOCTYPE html><html><head><title>Request Access &middot; VariTec Notes</title>

<style>

*{margin:0;padding:0;box-sizing:border-box;font-family:'Segoe UI',system-ui,sans-serif;}

body{

  min-height:100vh;display:flex;align-items:center;justify-content:center;

  background:linear-gradient(135deg,#0d152b 0%,#1e3868 50%,#0a2a5e 100%);

  position:relative;overflow:hidden;padding:32px 16px;

}

body::before{

  content:'';position:absolute;top:-120px;right:-120px;width:380px;height:380px;

  border-radius:50%;background:rgba(3,141,189,.18);

}

body::after{

  content:'';position:absolute;bottom:-150px;left:-100px;width:340px;height:340px;

  border-radius:50%;background:rgba(77,178,222,.12);

}

.login-shell{

  position:relative;z-index:1;display:flex;flex-direction:column;align-items:center;gap:22px;

}

.card{

  background:white;padding:40px 38px 32px;border-radius:18px;

  box-shadow:0 20px 60px rgba(0,0,0,.35);width:420px;

  border:1px solid rgba(255,255,255,.08);

}

.card-header{

  display:flex;align-items:center;gap:14px;margin-bottom:22px;

}

.card-icon{

  width:52px;height:52px;border-radius:14px;flex-shrink:0;

  background:linear-gradient(135deg,#038dbd,#0284c7);

  display:flex;align-items:center;justify-content:center;

  font-size:24px;

  box-shadow:0 8px 20px rgba(3,141,189,.35);

}

.card-header-text{display:flex;flex-direction:column;justify-content:center;}

.card h2{

  font-size:22px;font-weight:800;color:#0d152b;line-height:1.2;

}

.card .tagline{

  font-size:13px;color:#94a3b8;margin-top:3px;

}

.field-label{

  font-size:12px;font-weight:700;color:#475569;text-transform:uppercase;

  letter-spacing:.6px;margin-bottom:6px;display:block;

}

input,textarea{

  width:100%;padding:13px 14px;margin-bottom:18px;

  border:1.5px solid #e2e8f0;border-radius:10px;font-size:14px;

  outline:none;transition:.15s;color:#1e293b;font-family:inherit;

}

textarea{resize:vertical;min-height:64px;}

input::placeholder,textarea::placeholder{color:#cbd5e1;}

input:focus,textarea:focus{

  border-color:#038dbd;box-shadow:0 0 0 4px rgba(3,141,189,.10);

}

button{

  width:100%;padding:13px;

  background:linear-gradient(135deg,#0d152b,#1e3868);

  color:white;border:none;border-radius:10px;font-size:14px;font-weight:700;

  letter-spacing:.4px;cursor:pointer;transition:.18s;

  box-shadow:0 6px 18px rgba(13,21,43,.25);

}

button:hover{

  background:linear-gradient(135deg,#038dbd,#0284c7);

  box-shadow:0 8px 22px rgba(3,141,189,.35);

  transform:translateY(-1px);

}

.err{

  color:#dc2626;margin-top:-8px;margin-bottom:14px;font-size:13px;font-weight:600;

  background:#fee2e2;border:1px solid #fecaca;border-radius:8px;

  padding:10px 12px;text-align:center;

}

.ok-box{

  text-align:center;padding:8px 0 4px;

}

.ok-icon{

  font-size:40px;margin-bottom:10px;

}

.ok-box h3{

  font-size:18px;color:#0d152b;margin-bottom:8px;

}

.ok-box p{

  font-size:13px;color:#64748b;line-height:1.5;

}

.footer-note{

  font-size:12px;color:rgba(255,255,255,.45);text-align:center;

  letter-spacing:.4px;

}

.signup-link{

  text-align:center;margin-top:16px;font-size:13px;color:#94a3b8;

}

.signup-link a{

  color:#038dbd;font-weight:700;text-decoration:none;

}

.signup-link a:hover{

  text-decoration:underline;

}

</style></head>

<body>

<div class="login-shell">

  <div class="card">

    {% if submitted %}

    <div class="ok-box">

      <div class="ok-icon">✅</div>

      <h3>Request submitted</h3>

      <p>Thanks, {{ name }}. Your request has been sent to an administrator for review. You'll be able to sign in once it's approved.</p>

    </div>

    <div class="signup-link"><a href="/">← Back to Sign In</a></div>

    {% else %}

    <div class="card-header">

      <div class="card-icon">📝</div>

      <div class="card-header-text">

        <h2>Request Access</h2>

        <div class="tagline">Fill in your details to request an account</div>

      </div>

    </div>

    <form method="post">

      <label class="field-label">Full Name</label>

      <input name="name" type="text" placeholder="Jane Doe" required autofocus value="{{ name or '' }}">

      <label class="field-label">Email</label>

      <input name="email" type="text" inputmode="email" placeholder="you@varitecconsulting.com" required value="{{ email or '' }}">

      <label class="field-label">Password</label>

      <input name="password" type="password" placeholder="Create a password" required autocomplete="new-password">

      <label class="field-label">Confirm Password</label>

      <input name="confirm_password" type="password" placeholder="Re-enter your password" required autocomplete="new-password">

      <label class="field-label">Company / Team</label>

      <input name="company" type="text" placeholder="e.g. VariTec Consulting" value="{{ company or '' }}">

      <label class="field-label">Reason for Access</label>

      <textarea name="reason" placeholder="Briefly tell us why you need access">{{ reason or '' }}</textarea>

      {% if error %}<div class="err">{{ error }}</div>{% endif %}

      <button>Submit Request</button>

    </form>

    <div class="signup-link">Already have access? <a href="/">Sign in</a></div>

    {% endif %}

  </div>

  <div class="footer-note">VariTec Consulting &middot; Confidential</div>

</div>

</body></html>"""


NOTES_HTML="""<!DOCTYPE html>

<html><head><title>VariTec Notes</title>

<style>

*{margin:0;padding:0;box-sizing:border-box;font-family:'Segoe UI',system-ui,sans-serif;}

body{height:100vh;overflow:hidden;display:flex;flex-direction:column;}

.topbar{height:60px;background:#0d152b;display:flex;justify-content:space-between;align-items:center;padding:0 20px;border-bottom:2px solid #038dbd;flex-shrink:0;position:sticky;top:0;z-index:1000;box-shadow:0 2px 8px rgba(0,0,0,.35);}

.topbar-right{display:flex;align-items:center;gap:14px;}

.topbar a{color:#4db2de;text-decoration:none;font-size:13px;}.topbar a:hover{color:#fff;}

.topbar-email{font-size:13px;color:#94a3b8;}

/* ── Profile Dropdown ── */

.profile-wrap{position:relative;}

.profile-btn{display:flex;align-items:center;gap:8px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);border-radius:22px;padding:5px 12px 5px 6px;cursor:pointer;color:white;transition:.15s;}

.profile-btn:hover{background:rgba(255,255,255,.16);border-color:rgba(255,255,255,.3);}

.profile-avatar{width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,#038dbd,#0284c7);display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;color:white;flex-shrink:0;text-transform:uppercase;}

.profile-name-text{font-size:12px;font-weight:600;color:rgba(255,255,255,.9);max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}

.profile-caret{font-size:9px;color:rgba(255,255,255,.6);margin-left:2px;}

.profile-menu{display:none;position:absolute;right:0;top:calc(100% + 8px);width:240px;background:white;border-radius:14px;box-shadow:0 8px 32px rgba(0,0,0,.22);border:1px solid #e2e8f0;overflow:hidden;z-index:9999;}

.profile-menu.show{display:block;}

.profile-menu-header{background:linear-gradient(135deg,#0d152b,#1e3868);padding:16px 16px 14px;display:flex;align-items:center;gap:10px;}

.profile-avatar-lg{width:38px;height:38px;border-radius:50%;background:linear-gradient(135deg,#038dbd,#0284c7);display:flex;align-items:center;justify-content:center;font-size:17px;font-weight:800;color:white;flex-shrink:0;text-transform:uppercase;}

.profile-menu-email{font-size:12px;color:rgba(255,255,255,.75);margin-top:2px;word-break:break-all;}

.profile-menu-username{font-size:13px;font-weight:700;color:white;}

.profile-role-pill{display:inline-block;margin-top:4px;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;background:rgba(3,141,189,.3);color:#7dd3fc;letter-spacing:.5px;}

.profile-menu-item{display:flex;align-items:center;gap:10px;padding:11px 16px;font-size:13px;color:#374151;text-decoration:none;border-bottom:1px solid #f1f5f9;transition:.12s;cursor:pointer;}

.profile-menu-item:last-child{border-bottom:none;}

.profile-menu-item:hover{background:#f5f3ff;color:#5b21b6;}

.profile-menu-item.danger{color:#dc2626;}

.profile-menu-item.danger:hover{background:#fee2e2;color:#dc2626;}

.profile-menu-item-icon{font-size:15px;flex-shrink:0;}

.main{display:flex;flex:1;overflow:hidden;}

.sidebar{width:210px;background:#efebe5;display:flex;flex-direction:column;flex-shrink:0;overflow:hidden;}

.sidebar-header{padding:15px 15px 8px;font-size:11px;font-weight:700;color:#999;text-transform:uppercase;letter-spacing:.8px;}

.sections-list{flex:1;overflow-y:auto;padding:0 8px;}

.section-row{display:flex;align-items:center;border-radius:6px;margin-bottom:2px;}
.section-row.dragging{opacity:.35;background:#d9c8ff !important;}
.section-row.drag-over-top{border-top:2px solid #8b3dff !important;margin-top:-1px;}
.section-row.drag-over-bot{border-bottom:2px solid #8b3dff !important;margin-bottom:-1px;}
.section-drag-handle{width:8px;height:8px;border-radius:50%;cursor:grab;margin:0 8px 0 8px;flex-shrink:0;}
.section-row:active .section-drag-handle{cursor:grabbing;}
.section-row:nth-child(7n+1) .section-drag-handle{background:#f59e0b;}
.section-row:nth-child(7n+2) .section-drag-handle{background:#3b82f6;}
.section-row:nth-child(7n+3) .section-drag-handle{background:#ec4899;}
.section-row:nth-child(7n+4) .section-drag-handle{background:#10b981;}
.section-row:nth-child(7n+5) .section-drag-handle{background:#8b5cf6;}
.section-row:nth-child(7n+6) .section-drag-handle{background:#ef4444;}
.section-row:nth-child(7n+7) .section-drag-handle{background:#06b6d4;}

.section-row:hover{background:#e0d8cc;}.section-row.active{background:#d9c8ff;}

.section-name{flex:1;padding:9px 10px;font-size:14px;border:none;background:transparent;cursor:pointer;text-align:left;outline:none;color:#333;}

.section-row.active .section-name{font-weight:600;color:#5b21b6;}

.sec-actions{display:none;padding-right:6px;gap:2px;}.section-row:hover .sec-actions{display:flex;}

.sec-icon{background:none;border:none;cursor:pointer;font-size:13px;padding:2px 3px;border-radius:4px;color:#999;}

.sec-icon:hover{color:#5b21b6;background:#c4b5fd44;}

.add-section-btn{margin:8px;padding:8px 12px;background:none;border:2px dashed #c4b5fd;border-radius:6px;color:#8b3dff;font-size:13px;cursor:pointer;width:calc(100% - 16px);text-align:left;}

.add-section-btn:hover{background:#f0ebff;}

.pages{width:230px;background:#fafafa;border-left:1px solid #ddd;border-right:1px solid #ddd;display:flex;flex-direction:column;flex-shrink:0;}

.pages-top{padding:12px;border-bottom:1px solid #eee;}

.split-wrap{display:flex;position:relative;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(109,40,217,.4);}

.add-page-btn{flex:1;padding:10px 16px;background:#7c3aed;color:white;border:none;cursor:pointer;font-size:13px;font-weight:700;letter-spacing:.2px;white-space:nowrap;transition:background .12s;}

.add-page-btn:hover{background:#6d28d9;}

.split-arrow{width:38px;flex-shrink:0;background:#6d28d9;color:white;border:none;border-left:1.5px solid rgba(255,255,255,.25);cursor:pointer;font-size:11px;display:flex;align-items:center;justify-content:center;transition:background .12s;}

.split-arrow:hover{background:#5b21b6;}

.sc-dropdown{display:none;position:fixed;min-width:260px;background:white;border:1px solid #e0d9f0;border-radius:12px;box-shadow:0 12px 36px rgba(0,0,0,.18);z-index:9999;overflow:hidden;}

.sc-dd-section{background:linear-gradient(135deg,#f5f3ff,#f0ebff);padding:8px 14px 6px;display:flex;align-items:center;gap:6px;}

.sc-dd-section-dot{width:6px;height:6px;border-radius:50%;flex-shrink:0;}

.sc-dd-section-label{font-size:10px;font-weight:800;color:#6d28d9;text-transform:uppercase;letter-spacing:.8px;}

.sc-dd-divider{height:1px;background:linear-gradient(90deg,#ddd6fe,transparent);margin:0;}

.sc-dd-item{padding:9px 14px 9px 28px;font-size:13px;color:#374151;cursor:pointer;border-bottom:1px solid #f9f7ff;display:flex;align-items:center;gap:8px;transition:.12s;position:relative;}

.sc-dd-item::before{content:"";position:absolute;left:16px;top:50%;transform:translateY(-50%);width:5px;height:5px;border-radius:50%;background:#c4b5fd;}

.sc-dd-item:last-child{border-bottom:none;}

.sc-dd-item:hover{background:#f5f0ff;color:#5b21b6;padding-left:32px;}

.sc-dd-item.featured{color:#5b21b6;font-weight:600;}.sc-dd-item.featured::before{background:#7c3aed;}

.sc-dd-item-icon{font-size:14px;flex-shrink:0;}.sc-dd-item-text{flex:1;}

.sc-dd-item-badge{font-size:9px;font-weight:700;padding:2px 6px;border-radius:10px;background:#ede9fe;color:#6d28d9;letter-spacing:.3px;}

.pages-list{flex:1;overflow-y:auto;padding:8px;}

.page-link{display:flex;align-items:center;justify-content:space-between;padding:9px 10px;margin-bottom:3px;text-decoration:none;color:#333;border-radius:6px;font-size:14px;}

.page-link:hover{background:#eee;}.page-link.active{background:#e5d7ff;font-weight:600;border-left:3px solid #8b3dff;padding-left:7px;}

.drag-handle{width:8px;height:8px;border-radius:50%;cursor:grab;margin:0 8px 0 0;flex-shrink:0;}
.page-link:nth-child(7n+1) .drag-handle{background:#f59e0b;}
.page-link:nth-child(7n+2) .drag-handle{background:#3b82f6;}
.page-link:nth-child(7n+3) .drag-handle{background:#ec4899;}
.page-link:nth-child(7n+4) .drag-handle{background:#10b981;}
.page-link:nth-child(7n+5) .drag-handle{background:#8b5cf6;}
.page-link:nth-child(7n+6) .drag-handle{background:#ef4444;}
.page-link:nth-child(7n+7) .drag-handle{background:#06b6d4;}

.page-link:active .drag-handle{cursor:grabbing;}

.page-title{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}

.page-link.dragging{opacity:.35;background:#e5d7ff !important;}

.page-link.drag-over-top{border-top:2px solid #8b3dff !important;margin-top:-1px;}

.page-link.drag-over-bot{border-bottom:2px solid #8b3dff !important;margin-bottom:-1px;}

.del-page-btn{background:none;border:none;color:#bbb;cursor:pointer;font-size:15px;padding:0 2px;display:none;}

.page-link:hover .del-page-btn{display:inline;}.del-page-btn:hover{color:#e53e3e;}

.editor{flex:1;display:flex;flex-direction:column;background:white;overflow:hidden;}

.toolbar{display:flex;align-items:center;gap:6px;padding:8px 16px;border-bottom:1px solid #e2e8f0;flex-shrink:0;flex-wrap:wrap;background:#fff;}

.tb-group{display:inline-flex;align-items:center;border-radius:8px;overflow:hidden;border:1.5px solid #e8e8f0;}

.tb-group button{padding:6px 11px;border:none;border-right:1px solid #e8e8f0;cursor:pointer;font-size:12px;font-weight:700;transition:all .13s;white-space:nowrap;line-height:1;}

.tb-group button:last-child{border-right:none;}

.tb-fmt-grp{border-color:#ddd6fe;}.tb-fmt-grp button{background:#faf8ff;color:#6d28d9;border-right-color:#ddd6fe;}.tb-fmt-grp button:hover{background:#ede9fe;color:#5b21b6;}

.tb-list-grp{border-color:#bfdbfe;}.tb-list-grp button{background:#f0f7ff;color:#1d4ed8;border-right-color:#bfdbfe;}.tb-list-grp button:hover{background:#dbeafe;color:#1e40af;}

.tb-align-grp{border-color:#e2e8f0;}.tb-align-grp button{background:#f8fafc;color:#475569;border-right-color:#e2e8f0;font-size:14px;}.tb-align-grp button:hover{background:#e2e8f0;color:#1e293b;}

.tb-media-grp{border-color:#bae6fd;}.tb-media-grp button{background:#f0f9ff;color:#0369a1;border-right-color:#bae6fd;}.tb-media-grp button:hover{background:#e0f2fe;color:#075985;}

.tb-insert-grp{border-color:#99f6e4;}.tb-insert-grp button{background:#f0fdfa;color:#0f766e;border-right-color:#99f6e4;}.tb-insert-grp button:hover{background:#ccfbf1;color:#115e59;}

.tb-chart{padding:6px 14px;border:none;border-radius:8px;background:linear-gradient(135deg,#059669,#0d9488);color:white;cursor:pointer;font-size:12px;font-weight:700;box-shadow:0 2px 6px rgba(5,150,105,.3);transition:.15s;letter-spacing:.2px;}

.tb-chart:hover{background:linear-gradient(135deg,#047857,#0f766e);}

.tb-btn-ppt{padding:6px 16px;border:none;border-radius:8px;background:linear-gradient(135deg,#dc2626,#b91c1c);cursor:pointer;font-size:12px;color:white;font-weight:700;margin-left:auto;box-shadow:0 2px 8px rgba(220,38,38,.3);transition:.15s;letter-spacing:.2px;}

.tb-btn-ppt:hover{background:linear-gradient(135deg,#b91c1c,#991b1b);}

.tb-sep{width:1px;height:20px;background:#e2e8f0;margin:0 2px;flex-shrink:0;}

.note-title-bar{display:flex;align-items:center;justify-content:space-between;padding:0 20px 0 0;flex-shrink:0;border-bottom:1px solid #f5f5f5;}

#title.hidden-title{display:none!important;}
#content:has(.ty-wrap) ~ * #title,
.note-title-bar:has(+ #content .ty-wrap) #title{display:none!important;}

#title{font-size:28px;font-weight:700;border:none;outline:none;padding:20px 16px 10px 40px;color:#1e293b;flex:1;background:transparent;}

#title::placeholder{color:#cbd5e1;}

.note-title-logo{height:26px;width:auto;opacity:1;margin-right:10px;}

#content{flex:1;padding:0 40px 20px;outline:none;font-size:15px;line-height:1.8;color:#334155;overflow-y:auto;word-wrap:break-word;}

#content:empty:before{content:attr(data-placeholder);color:#cbd5e1;pointer-events:none;}

#content img{max-width:100%;border-radius:6px;margin:8px 0;cursor:pointer;border:2px solid transparent;display:block;}

#content img:hover{border-color:#c4b5fd;}

#content img.img-selected{border-color:#8b3dff!important;outline:none;}

/* ── Image Resize Overlay ── */

#imgResizeOverlay{display:none;position:fixed;z-index:9100;border:2px solid #8b3dff;box-sizing:border-box;pointer-events:none;}

.img-rh{position:absolute;width:11px;height:11px;background:#8b3dff;border:2px solid white;border-radius:2px;pointer-events:all;box-sizing:border-box;}

.img-rh-tl{top:-6px;left:-6px;cursor:nw-resize;}

.img-rh-tc{top:-6px;left:50%;transform:translateX(-50%);cursor:n-resize;}

.img-rh-tr{top:-6px;right:-6px;cursor:ne-resize;}

.img-rh-ml{top:50%;left:-6px;transform:translateY(-50%);cursor:w-resize;}

.img-rh-mr{top:50%;right:-6px;transform:translateY(-50%);cursor:e-resize;}

.img-rh-bl{bottom:-6px;left:-6px;cursor:sw-resize;}

.img-rh-bc{bottom:-6px;left:50%;transform:translateX(-50%);cursor:s-resize;}

.img-rh-br{bottom:-6px;right:-6px;cursor:se-resize;}

#imgResizeToolbar{position:absolute;top:-46px;left:0;background:#1e293b;border-radius:8px;padding:4px 8px;display:flex;align-items:center;gap:3px;pointer-events:all;white-space:nowrap;box-shadow:0 4px 16px rgba(0,0,0,.35);}

.img-tb-btn{background:none;border:none;color:white;cursor:pointer;font-size:12px;font-weight:600;padding:4px 9px;border-radius:5px;transition:.12s;line-height:1.3;}

.img-tb-btn:hover{background:rgba(255,255,255,.18);}

.img-tb-sep{width:1px;height:16px;background:rgba(255,255,255,.25);margin:0 3px;flex-shrink:0;}

#imgSizeLabel{font-size:11px;color:#94a3b8;padding:0 4px;min-width:70px;}

#content table{border-collapse:separate;border-spacing:0;width:100%;margin:14px 0;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.10);}

#content table thead tr{background:linear-gradient(135deg,#0d152b 0%,#1a2d5a 70%);}

#content table th{color:white;padding:12px 16px;text-align:left;font-size:12px;font-weight:700;border-right:1px solid rgba(255,255,255,.1);text-transform:uppercase;letter-spacing:.4px;}

#content table th:last-child{border-right:none;}

#content table td{padding:10px 16px;border-bottom:1px solid #e5e7eb;border-right:1px solid #f0f0f0;color:#1e293b;vertical-align:middle;font-size:14px;}

#content table td:last-child{border-right:none;}

#content table tbody tr:last-child td{border-bottom:none;}

#content table tbody tr:hover td{background:#f8f5ff;}

.sc-chart-del-btn{position:absolute;top:10px;right:10px;z-index:30;background:rgba(220,38,38,.75);border:none;color:white;border-radius:8px;padding:5px 13px;font-size:11px;font-weight:700;cursor:pointer;letter-spacing:.3px;font-family:Segoe UI,sans-serif;transition:background .15s;}
.sc-chart-del-btn:hover{background:rgba(185,28,28,.95);}
.sc-table{border-collapse:separate !important;border-spacing:0 !important;width:100%;margin:16px 0 !important;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.13) !important;}

.sc-table thead tr{background:linear-gradient(135deg,#0d152b 0%,#1e3868 60%,#0a2a5e 100%) !important;}

.sc-table th{color:white !important;padding:14px 18px !important;font-size:12px !important;font-weight:700 !important;text-transform:uppercase !important;letter-spacing:.6px !important;border-right:1px solid rgba(255,255,255,.12) !important;}

.sc-table th:last-child{border-right:none !important;}

.sc-table th.sc-th-name{min-width:220px;background:rgba(3,141,189,.2) !important;}

.sc-table td{padding:0 !important;border-bottom:1px solid #eaecf0 !important;vertical-align:middle !important;}

.sc-table td:last-child{border-right:none !important;}

.sc-table tbody tr:last-child td{border-bottom:none !important;}

.sc-table tbody tr:hover td{filter:brightness(.97);}

.sc-name-td{padding:12px 18px !important;font-weight:700 !important;font-size:14px;background:#f8f9ff !important;border-right:2px solid #ddd6fe !important;}

.sc-week-td{min-width:140px;}

.sc-cell-inner{display:flex;align-items:center;gap:8px;padding:10px 16px;min-height:48px;}

.sc-val{min-width:60px;max-width:90px;width:70px;padding:5px 10px;background:#fff;border:1.5px solid #d1d5db;border-radius:7px;font-size:14px;font-weight:700;color:#1e293b;text-align:center;outline:none;cursor:text;display:inline-block;min-height:32px;line-height:22px;}

.sc-val:empty::before{content:"0";color:#cbd5e1;font-weight:400;pointer-events:none;}

.sc-val:focus{border-color:#8b3dff;box-shadow:0 0 0 3px rgba(139,61,255,.12);background:#faf7ff;}

.sc-pct{font-size:15px;font-weight:700;color:#64748b;}

.sc-table tbody tr:nth-child(1) .sc-week-td{background:#faf5ff;}

.sc-table tbody tr:nth-child(2) .sc-week-td{background:#f0fdf4;}

.sc-table tbody tr:nth-child(3) .sc-week-td{background:#fffbeb;}

.sc-table tbody tr:nth-child(4) .sc-week-td{background:#eff6ff;}

.sc-table tbody tr:nth-child(5) .sc-week-td{background:#fdf2f8;}

.sc-table tbody tr:nth-child(n+6) .sc-week-td{background:#f0fdfa;}

.sc-month-wrap{display:inline-flex;align-items:center;gap:6px;margin-bottom:12px;background:#f5f3ff;border:1.5px solid #ddd6fe;border-radius:20px;padding:6px 16px;}

.sc-month-val{font-size:14px;font-weight:700;color:#5b21b6;outline:none;min-width:80px;cursor:text;border-bottom:1.5px dashed #a78bfa;padding-bottom:1px;}

.sc-month-val:empty::before{content:"Click to set month";color:#a78bfa;font-weight:400;}

.ob-wrap{font-family:'Segoe UI',system-ui,sans-serif;margin:8px 0 20px;border-radius:14px;box-shadow:0 4px 24px rgba(0,0,0,.12);}
.ob-wrap .ob-header-customer,.ob-wrap .ob-header-vendor,.ob-wrap .ob-header-mah{border-radius:14px 14px 0 0;overflow:hidden;}

display:flex;align-items:center;justify-content:space-between;gap:24px;position:relative;overflow:hidden;}

.ob-header-customer{background:linear-gradient(135deg,#1e3a8a 0%,#2563eb 60%,#3b82f6 100%);padding:10px 22px;display:flex;align-items:center;gap:16px;position:relative;overflow:hidden;}

.ob-header-vendor{background:linear-gradient(135deg,#064e3b 0%,#059669 60%,#10b981 100%);padding:12px 22px;display:flex;align-items:center;justify-content:space-between;gap:16px;position:relative;overflow:hidden;}

.ob-header-mah{background:linear-gradient(135deg,#7c2d12 0%,#c2410c 60%,#ea580c 100%);padding:12px 22px;display:flex;align-items:center;justify-content:space-between;gap:16px;position:relative;overflow:hidden;}

.ob-header-text{flex:1;z-index:1;}

.ob-header-title{font-size:20px;font-weight:900;line-height:1.1;color:white;text-transform:uppercase;letter-spacing:.5px;}

.ob-header-sub{font-size:13px;font-weight:500;opacity:.75;margin-top:8px;color:white;}

.ob-month-wrap{display:inline-flex;align-items:center;gap:6px;margin-top:12px;background:rgba(255,255,255,.15);border:1.5px solid rgba(255,255,255,.3);border-radius:20px;padding:5px 14px;}

.ob-month-val{font-size:13px;font-weight:700;outline:none;min-width:70px;cursor:text;color:rgba(255,255,255,.95);border-bottom:1.5px dashed rgba(255,255,255,.5);}

.ob-month-val:empty::before{content:"Set month";color:rgba(255,255,255,.5);font-weight:400;}

.ob-header-illus{display:none!important;}

.ob-table-wrap{overflow:hidden;}

.ob-table{width:100%;border-collapse:collapse;font-size:14px;}

.ob-table td{border-right:1px solid #e2e8f0!important;border-bottom:1px solid #e2e8f0!important;}

.ob-table td:last-child{border-right:none!important;}

.ob-table td[colspan]{padding:0!important;height:auto!important;border-top:none!important;border-bottom:1px solid #e2e8f0!important;}

.ob-table tr:last-child td{border-bottom:none!important;}

.ob-wrap /* removed wrap border - using cell borders + outline on table instead */

.ob-table,.ob-table th,.ob-table td{border:none!important;}
.ob-table-wrap{border:none!important;border-radius:0!important;}
.ob-wrap{border-radius:10px!important;}

.ob-table th{color:white;padding:13px 20px;text-align:left;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.8px;border-right:1px solid rgba(255,255,255,.15);}

.ob-table th:last-child{border-right:none;}

.ob-table td{padding:12px 20px;border-bottom:1px solid #f0f4f8;border-right:1px solid #f0f4f8;color:#1e293b;vertical-align:middle;font-size:14px;}

.ob-table td:last-child{border-right:none;}
.ob-comment-cell:empty:before{content:"Add comment...";color:#cbd5e1;font-style:italic;pointer-events:none;}

.ob-status-btn{display:inline-block;padding:4px 14px;border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;user-select:none;border:none;transition:all .15s;letter-spacing:.3px;}

.ob-status-btn.pending{background:#fef08a;color:#713f12;}

.ob-status-btn.inprogress{background:#fde68a;color:#92400e;box-shadow:0 0 0 2px #f59e0b;}

.ob-status-btn.hold{background:#fee2e2;color:#991b1b;box-shadow:0 0 0 2px #ef4444;}

.ob-status-btn.completed{background:#dcfce7;color:#166534;box-shadow:0 0 0 2px #22c55e;font-weight:800;}

.ob-progress-bar-wrap{background:#f8fafc;border-bottom:1px solid #e2e8f0;}


/* ── Date Picker ── */

.date-cell-wrap{display:flex;align-items:center;justify-content:center;gap:4px;}

.date-val{min-width:72px;font-size:12px;color:#1e293b;outline:none;cursor:text;min-height:18px;display:inline-block;border-bottom:1.5px dashed #cbd5e1;padding-bottom:1px;text-align:center;font-weight:700;}

.date-val:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;font-size:11px;pointer-events:none;}

.date-val:focus{border-bottom-color:#038dbd;}

.date-cal-btn{background:none;border:none;cursor:pointer;padding:2px 3px;font-size:14px;line-height:1;border-radius:4px;transition:background .12s;flex-shrink:0;color:#94a3b8;}

.date-cal-btn:hover{background:#e0f2fe;color:#038dbd;}

.si-date-wrap{display:flex;align-items:center;gap:4px;padding:5px 9px;background:#f8fafc;border-radius:6px;border:1.5px solid #e2e8f0;min-height:32px;}

.si-date-wrap:focus-within{border-color:var(--si-col,#038dbd);background:#fff;}

.si-date-wrap .date-val{flex:1;border-bottom:none;min-height:16px;font-size:12px;}

.si-date-wrap .date-cal-btn{font-size:13px;}

#datePickerPopup{display:none;position:fixed;z-index:99999;background:white;border-radius:10px;

  box-shadow:0 8px 32px rgba(0,0,0,.22);border:1px solid #e2e8f0;width:252px;overflow:hidden;}

/* calendar internals */

.dp-hdr{background:#0d152b;padding:10px 14px;display:flex;align-items:center;justify-content:space-between;border-radius:10px 10px 0 0;}

.dp-nav{background:rgba(255,255,255,.12);border:none;color:white;cursor:pointer;border-radius:6px;width:28px;height:28px;font-size:20px;display:flex;align-items:center;justify-content:center;line-height:1;}

.dp-nav:hover{background:rgba(255,255,255,.28);}

.dp-hdr-lbl{text-align:center;}

.dp-month-lbl{color:white;font-size:13px;font-weight:700;}

.dp-year-lbl{color:#4db2de;font-size:11px;font-weight:600;}

.dp-dow-row{display:grid;grid-template-columns:repeat(7,1fr);background:#f8fafc;padding:6px 10px 2px;}

.dp-dow{text-align:center;font-size:10px;font-weight:700;color:#94a3b8;}

.dp-dow.dp-wknd{color:#ef4444;}

.dp-grid{display:grid;grid-template-columns:repeat(7,1fr);padding:4px 10px 10px;gap:1px;}

.dp-day{border:none;background:transparent;font-size:12px;border-radius:6px;padding:5px 2px;cursor:pointer;font-weight:400;color:#1e293b;width:100%;}

.dp-day:hover{background:#e0f2fe;}

.dp-day.dp-today{background:#038dbd;color:white;font-weight:700;}

.dp-day.dp-today:hover{background:#0272a0;}

.dp-day.dp-wknd{color:#ef4444;}

.dp-day.dp-today.dp-wknd{color:white;}

.dp-footer{padding:0 10px 10px;text-align:center;}

.dp-today-btn{font-size:11px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:6px;padding:4px 14px;cursor:pointer;color:#475569;font-weight:600;}

.dp-today-btn:hover{background:#e2e8f0;}

/* clickable month / year in header */

.dp-hdr-month-btn{background:none;border:none;color:white;cursor:pointer;font-size:13px;font-weight:700;border-radius:4px;padding:2px 6px;line-height:1.4;}

.dp-hdr-year-btn{background:none;border:none;color:#4db2de;cursor:pointer;font-size:11px;font-weight:600;border-radius:4px;padding:1px 5px;line-height:1.4;}

.dp-hdr-month-btn:hover,.dp-hdr-year-btn:hover{background:rgba(255,255,255,.18);}

/* month grid */

.dp-month-grid{display:grid;grid-template-columns:repeat(3,1fr);padding:10px;gap:5px;}

.dp-month-btn{border:none;background:#f1f5f9;font-size:12px;border-radius:7px;padding:9px 4px;cursor:pointer;color:#334155;font-weight:500;}

.dp-month-btn:hover{background:#bae6fd;color:#0369a1;}

.dp-month-btn.dp-sel{background:#038dbd;color:white;font-weight:700;}

/* year grid */

.dp-year-grid{display:grid;grid-template-columns:repeat(3,1fr);padding:10px;gap:5px;}

.dp-year-btn{border:none;background:#f1f5f9;font-size:12px;border-radius:7px;padding:9px 4px;cursor:pointer;color:#334155;font-weight:500;}

.dp-year-btn:hover{background:#bae6fd;color:#0369a1;}

.dp-year-btn.dp-sel{background:#038dbd;color:white;font-weight:700;}

.dp-year-range{color:#4db2de;font-size:12px;font-weight:600;}

.dp-back-btn{font-size:11px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:6px;padding:4px 14px;cursor:pointer;color:#475569;font-weight:600;}

.dp-back-btn:hover{background:#e2e8f0;}

.si-wrap{font-family:'Segoe UI',system-ui,sans-serif;margin:8px 0 20px;border-radius:14px;box-shadow:0 4px 24px rgba(0,0,0,.12);overflow:hidden;}
.si-banner{background:linear-gradient(135deg,#0d152b 0%,#1e3868 50%,#0a2a5e 100%);padding:10px 22px;position:relative;overflow:hidden;text-align:left;border-radius:14px 14px 0 0;}

.si-banner-title{font-size:20px;font-weight:900;color:white;position:relative;z-index:1;text-transform:uppercase;letter-spacing:.5px;line-height:1.1;}

.si-banner-subtitle{font-size:13px;color:rgba(77,178,222,.85);margin-top:6px;position:relative;z-index:1;font-weight:500;}

.si-banner-meta{display:flex;align-items:center;justify-content:center;gap:16px;margin-top:14px;position:relative;z-index:1;}

.si-month-wrap{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.12);border:1.5px solid rgba(255,255,255,.25);border-radius:20px;padding:5px 14px;}

.si-month-val{font-size:13px;font-weight:700;color:white;outline:none;min-width:70px;cursor:text;border-bottom:1.5px dashed rgba(255,255,255,.4);}

.si-cards-area{background:#f1f5f9;padding:20px 20px 8px;border-radius:0 0 14px 14px;}

.si-cards-grid{display:flex;flex-wrap:wrap;align-items:flex-start;gap:16px;margin-bottom:12px;}

.si-card{background:white;border-radius:14px;padding:0;width:calc(50% - 10px);min-width:300px;box-shadow:0 3px 20px rgba(0,0,0,.12);border:none;display:flex;flex-direction:column;overflow:hidden;transition:.25s;}

.si-card:hover{box-shadow:0 6px 24px rgba(0,0,0,.16);transform:translateY(-3px);}

.si-card-head{padding:12px 22px;background:linear-gradient(135deg,#1e3a8a 0%,#2563eb 60%,#3b82f6 100%);position:relative;overflow:hidden;}

.si-card-title{font-size:14px;font-weight:800;color:white;line-height:1.3;outline:none;cursor:text;min-height:20px;}

.si-card-title:empty::before{content:"Issue Title";color:rgba(255,255,255,.5);font-weight:400;}

.si-card-body{padding:12px 16px;flex:1;display:flex;flex-direction:column;gap:10px;background:white;}

.si-field{display:flex;flex-direction:column;gap:3px;}

.si-field-label{font-size:9px;font-weight:800;color:var(--si-col,#038dbd);text-transform:uppercase;letter-spacing:.8px;}

.si-field-val{font-size:12px;color:#334155;outline:none;cursor:text;padding:5px 9px;background:#f8fafc;border-radius:6px;border:1.5px solid #e2e8f0;min-height:26px;line-height:1.5;transition:.15s;}

.si-field-val:focus{border-color:var(--si-col,#038dbd);background:#fff;}

.si-field-val:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;}

.si-card-foot{padding:10px 16px;border-top:2px solid;border-color:var(--si-col-light,#e0f2fe);display:flex;align-items:center;justify-content:space-between;background:var(--si-col-bg,#f0f9ff);}
.si-status-label{font-size:9px;font-weight:800;color:var(--si-col,#038dbd);text-transform:uppercase;letter-spacing:.8px;}
.si-status-group{display:flex;align-items:center;gap:8px;}

.si-del-card{background:#fee2e2;border:1.5px solid #fecaca;color:#dc2626;cursor:pointer;font-size:13px;padding:4px 10px;border-radius:7px;font-weight:700;transition:.15s;white-space:nowrap;}

.si-del-card:hover{background:#dc2626;color:white;}

.si-multi-rows{display:flex;flex-direction:column;gap:4px;}

.si-multi-row{display:flex;align-items:flex-start;gap:6px;position:relative;}

.si-multi-row .si-field-val{flex:1;min-height:32px;}

.si-field-del-row{flex-shrink:0;width:24px;height:24px;background:none;border:1.5px solid #fecaca;color:#f87171;border-radius:50%;cursor:pointer;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:4px;transition:.15s;line-height:1;padding:0;opacity:0;}

.si-multi-row:hover .si-field-del-row{opacity:1;}

.si-field-del-row:hover{background:#dc2626;border-color:#dc2626;color:white;transform:scale(1.1);}

.si-field-add-row{align-self:flex-start;margin-top:4px;padding:3px 12px;background:none;border:1.5px dashed rgba(3,141,189,.4);border-radius:20px;color:#038dbd;font-size:11px;font-weight:700;cursor:pointer;transition:.15s;display:flex;align-items:center;gap:4px;letter-spacing:.3px;}

.si-field-add-row:hover{border-color:#038dbd;background:rgba(3,141,189,.08);}.si-multi-rows{display:flex;flex-direction:column;gap:4px;}.si-multi-row{display:flex;align-items:flex-start;gap:6px;position:relative;}.si-multi-row .si-field-val{flex:1;min-height:32px;transition:.15s;}.si-field-del-row{flex-shrink:0;width:24px;height:24px;background:none;border:1.5px solid #fecaca;color:#f87171;border-radius:50%;cursor:pointer;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:4px;transition:.15s;line-height:1;padding:0;opacity:0;}.si-multi-row:hover .si-field-del-row{opacity:1;}.si-field-del-row:hover{background:#dc2626;border-color:#dc2626;color:white;transform:scale(1.1);}.si-field-add-row{align-self:flex-start;margin-top:4px;padding:3px 12px;background:none;border:1.5px dashed rgba(3,141,189,.4);border-radius:20px;color:#038dbd;font-size:11px;font-weight:700;cursor:pointer;transition:.15s;display:flex;align-items:center;gap:4px;letter-spacing:.3px;}.si-field-add-row:hover{border-color:#038dbd;background:rgba(3,141,189,.08);transform:none;}

.si-add-btn{width:calc(50% - 10px);min-width:300px;min-height:100px;border:2px dashed #038dbd;border-radius:14px;background:rgba(3,141,189,.04);color:#038dbd;font-size:14px;font-weight:700;cursor:pointer;transition:.15s;display:flex;align-items:center;justify-content:center;gap:6px;}

.si-add-btn:hover{background:rgba(3,141,189,.1);}
.si-tl-card{border-radius:12px!important;}
.si-tl-issue-title{font-size:15px;font-weight:800;color:#fff;padding:12px 16px;margin:-12px -16px 10px -16px;background:linear-gradient(135deg,rgba(255,255,255,.20),rgba(255,255,255,0) 60%),var(--si-col,#038dbd);border-radius:12px 12px 0 0;outline:none;cursor:text;text-shadow:0 1px 2px rgba(0,0,0,.18);}
.si-tl-issue-title:empty::before{content:attr(data-ph);color:rgba(255,255,255,.65);font-weight:700;}
.si-tl-issue-title:focus{box-shadow:inset 0 0 0 2px rgba(255,255,255,.45);}
.si-tl-sublabel{font-size:9px;font-weight:800;color:#94a3b8;text-transform:uppercase;letter-spacing:.5px;margin:2px 0 4px;}

.si-tl{position:relative;padding-left:2px;}
.si-tl-step{position:relative;display:flex;gap:12px;padding-bottom:14px;}
.si-tl-step:last-child{padding-bottom:2px;}
.si-tl-node{width:30px;height:30px;border-radius:50%;background:var(--step-col,#3b82f6);color:#fff;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0;position:relative;z-index:1;box-shadow:0 2px 6px rgba(0,0,0,.16);}
.si-tl-line{position:absolute;left:14px;top:30px;bottom:-14px;width:2px;background:#e2e8f0;z-index:0;}
.si-tl-step-last .si-tl-line,.si-tl-step:last-child .si-tl-line{display:none;}
.si-tl-content{flex:1;min-width:0;}
.si-tl-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:4px;flex-wrap:wrap;}
.si-tl-title{font-size:11px;font-weight:800;color:#1e293b;text-transform:uppercase;letter-spacing:.4px;}
.si-tl-date-wrap{display:flex;align-items:center;gap:3px;font-size:11px;color:#94a3b8;font-weight:600;}
.si-tl-date-wrap .date-val{font-size:11px;color:#64748b;}
.si-tl-body{font-size:12px;color:#475569;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;min-height:38px;outline:none;cursor:text;line-height:1.5;}
.si-tl-body:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;}
.si-tl-body:focus{border-color:#038dbd;background:#fff;}
.si-action-list{display:flex;flex-direction:column;gap:5px;margin-bottom:6px;}
.si-action-row{display:flex;align-items:center;gap:7px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:7px;padding:5px 8px;}
.si-action-check{width:15px;height:15px;cursor:pointer;accent-color:#22c55e;flex-shrink:0;}
.si-action-text{flex:1;font-size:12px;color:#334155;outline:none;cursor:text;}
.si-action-text:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;}
.si-tl-body ol,.si-tl-body ul,.si-action-text ol,.si-action-text ul,.exec-content ol,.exec-content ul{list-style-position:inside;margin:0;padding-left:0;}
.si-action-row.si-action-done .si-action-text{text-decoration:line-through;color:#94a3b8;}
.si-action-del{background:none;border:none;color:#cbd5e1;cursor:pointer;font-size:14px;padding:0 2px;flex-shrink:0;transition:.15s;}
.si-action-del:hover{color:#ef4444;}


/* ── Executive Summary Template ── */

.exec-wrap{font-family:'Segoe UI',system-ui,sans-serif;background:white;border-radius:14px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.12);margin:8px 0 20px;}

.exec-header{background:linear-gradient(135deg,#1e3a8a 0%,#2563eb 55%,#3b82f6 100%);padding:12px 24px;display:flex;align-items:center;position:relative;overflow:hidden;width:100%;box-sizing:border-box;}

.exec-header::before{content:'';position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;background:rgba(255,255,255,.06);}

.exec-header::after{content:'';position:absolute;right:60px;bottom:-60px;width:160px;height:160px;border-radius:50%;background:rgba(255,255,255,.04);}

.exec-header-text{z-index:1;}

.exec-title{font-size:18px;font-weight:900;color:white;margin:0;text-transform:uppercase;letter-spacing:1px;}

.exec-subtitle{font-size:13px;color:rgba(255,255,255,.75);margin-top:6px;font-weight:500;}

.exec-footer,.agenda-footer{display:none!important;}

.exec-vc{position:absolute;top:12px;right:20px;width:46px;height:36px;background:#ffffff;border-radius:6px;box-shadow:0 2px 10px rgba(0,0,0,.18);display:flex;align-items:center;justify-content:center;z-index:2;padding:4px;}
.exec-vc img{width:100%;height:100%;object-fit:contain;}

.exec-body{padding:14px 32px 18px;}

.exec-section{margin-bottom:14px;}

.exec-section-hdr{display:flex;align-items:center;gap:10px;margin-bottom:6px;}

.exec-dot{width:13px;height:13px;border-radius:50%;flex-shrink:0;}

.exec-dot.red{background:#dc2626;}

.exec-dot.green{background:#16a34a;}

.exec-dot.amber{background:#d97706;}

.exec-section-label{font-size:15px;font-weight:700;color:#1e293b;}

/* ── Exec rows (add/delete/handle) ── */

.exec-row{display:flex;align-items:flex-start;gap:8px;margin-bottom:4px;background:#fafafa;border:1.5px solid #e2e8f0;border-radius:8px;padding:2px 6px;transition:.15s;}

.exec-row:hover{border-color:#cbd5e1;background:white;}

.exec-row-handle{cursor:grab;color:#cbd5e1;font-size:16px;padding:6px 2px;flex-shrink:0;user-select:none;line-height:1;transition:.15s;}

.exec-row-handle:hover{color:#94a3b8;}

.exec-row-content{flex:1;font-size:14px;color:#334155;outline:none;min-height:24px;padding:4px 6px;line-height:1.4;background:transparent;}

.exec-row-content:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;pointer-events:none;}

.exec-del-row{flex-shrink:0;background:none;border:1px solid #e2e8f0;color:#94a3b8;font-size:12px;cursor:pointer;padding:4px 7px;border-radius:6px;line-height:1;transition:.15s;align-self:center;}

.exec-del-row:hover{color:#ef4444;background:#fee2e2;border-color:#fecaca;}

.exec-add-row{display:flex;align-items:center;gap:6px;background:none;border:1.5px dashed #cbd5e1;border-radius:8px;color:#64748b;font-size:12px;font-weight:600;padding:7px 14px;cursor:pointer;width:100%;transition:.15s;margin-top:4px;}

.exec-add-row:hover{border-color:#038dbd;color:#038dbd;background:rgba(3,141,189,.04);}

.exec-content{font-size:14px;color:#334155;outline:none;min-height:72px;padding:10px 14px;border-radius:6px;border:1.5px solid #e2e8f0;line-height:1.8;background:#fafafa;display:block;width:calc(100% - 23px);margin-left:23px;box-sizing:border-box;}

.exec-content:focus{border-color:#2563eb;background:#fff;}

.exec-content:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;pointer-events:none;}

.exec-row{display:flex;align-items:flex-start;gap:8px;margin-bottom:4px;background:#fafafa;border:1.5px solid #e2e8f0;border-radius:8px;padding:2px 6px;transition:.15s;}

.exec-row:hover{border-color:#cbd5e1;background:white;}

.exec-row-handle{cursor:grab;color:#cbd5e1;font-size:16px;padding:6px 2px;flex-shrink:0;user-select:none;line-height:1;transition:.15s;}

.exec-row-handle:hover{color:#94a3b8;}

.exec-row-content{flex:1;font-size:14px;color:#334155;outline:none;min-height:24px;padding:4px 6px;line-height:1.4;background:transparent;}

.exec-row-content:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;pointer-events:none;}

.exec-del-row{flex-shrink:0;background:none;border:1px solid #e2e8f0;color:#94a3b8;font-size:12px;cursor:pointer;padding:4px 7px;border-radius:6px;line-height:1;transition:.15s;align-self:center;}

.exec-del-row:hover{color:#ef4444;background:#fee2e2;border-color:#fecaca;}

.exec-add-row{display:flex;align-items:center;gap:6px;background:none;border:1.5px dashed #cbd5e1;border-radius:8px;color:#64748b;font-size:12px;font-weight:600;padding:7px 14px;cursor:pointer;width:100%;transition:.15s;margin-top:4px;}

.exec-add-row:hover{border-color:#038dbd;color:#038dbd;background:rgba(3,141,189,.04);}

.exec-footer-url{font-size:11px;color:rgba(255,255,255,.75);}


/* ── Agenda Template ── */

.agenda-wrap{font-family:'Segoe UI',system-ui,sans-serif;background:white;border-radius:14px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.12);margin:8px 0 20px;min-height:calc(100vh - 220px);display:flex;flex-direction:column;}

.agenda-header{background:linear-gradient(135deg,#0d152b 0%,#1e3868 55%,#038dbd 100%);padding:12px 24px;display:flex;align-items:center;justify-content:space-between;position:relative;overflow:hidden;}

.agenda-header::before{content:'';position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;background:rgba(255,255,255,.05);}

.agenda-header::after{content:'';position:absolute;right:60px;bottom:-60px;width:160px;height:160px;border-radius:50%;background:rgba(77,178,222,.12);}

.agenda-title{font-size:22px;font-weight:900;color:white;margin:0;position:relative;z-index:1;text-transform:uppercase;letter-spacing:1px;}

.agenda-subtitle{font-size:14px;color:rgba(255,255,255,.7);margin-top:6px;position:relative;z-index:1;}

.agenda-vc{position:absolute;top:12px;right:20px;width:46px;height:36px;background:#ffffff;border-radius:6px;box-shadow:0 2px 10px rgba(0,0,0,.18);display:flex;align-items:center;justify-content:center;z-index:2;padding:4px;}
.agenda-vc img{width:100%;height:100%;object-fit:contain;}

.exec-footer,.agenda-footer{display:none!important;}

.agenda-body{display:flex;align-items:stretch;justify-content:center;gap:18px;padding:30px 24px;flex:1;background:white;flex-wrap:wrap;}

.agenda-item{display:flex;align-items:center;gap:14px;flex:0 1 29%;min-width:220px;border-radius:16px;padding:18px 18px 18px 16px;background:linear-gradient(160deg,var(--ag-bg,#eff6ff) 0%,#ffffff 100%);border:1.5px solid var(--ag-border,#bfdbfe);box-shadow:0 3px 14px rgba(0,0,0,.05);transition:.2s;}

.agenda-num{font-size:62px;font-weight:900;line-height:1;color:transparent;-webkit-text-stroke:3px var(--ag-col,#2563eb);letter-spacing:-2px;flex-shrink:0;width:64px;text-align:center;}

.agenda-label{font-size:16px;font-weight:700;color:#1e293b;outline:none;cursor:text;min-height:24px;line-height:1.4;word-wrap:break-word;}

.agenda-label:empty::before{content:"Item label";color:#cbd5e1;font-style:italic;}
.agenda-item{position:relative;}
.agenda-del-item{position:absolute;top:-8px;right:-8px;width:22px;height:22px;background:#fff;border:1.5px solid #fecaca;color:#f87171;border-radius:50%;cursor:pointer;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;line-height:1;padding:0;opacity:0;transition:.15s;box-shadow:0 2px 6px rgba(0,0,0,.08);}
.agenda-item:hover .agenda-del-item{opacity:1;}
.agenda-del-item:hover{background:#dc2626;border-color:#dc2626;color:white;transform:scale(1.1);}
.agenda-item:nth-child(7n+1){--ag-col:#8b5cf6;--ag-bg:#f5f3ff;--ag-border:#ddd6fe;}
.agenda-item:nth-child(7n+2){--ag-col:#059669;--ag-bg:#f0fdf4;--ag-border:#bbf7d0;}
.agenda-item:nth-child(7n+3){--ag-col:#d97706;--ag-bg:#fffbeb;--ag-border:#fde68a;}
.agenda-item:nth-child(7n+4){--ag-col:#2563eb;--ag-bg:#eff6ff;--ag-border:#bfdbfe;}
.agenda-item:nth-child(7n+5){--ag-col:#db2777;--ag-bg:#fdf2f8;--ag-border:#fbcfe8;}
.agenda-item:nth-child(7n+6){--ag-col:#0891b2;--ag-bg:#f0fdfa;--ag-border:#a5f3fc;}
.agenda-item:nth-child(7n+7){--ag-col:#7c3aed;--ag-bg:#f5f3ff;--ag-border:#ddd6fe;}
.agenda-item::before{content:'';position:absolute;top:0;left:18px;right:18px;height:4px;background:var(--ag-col,#2563eb);border-radius:0 0 6px 6px;}
.agenda-item:hover{transform:translateY(-3px);box-shadow:0 8px 22px rgba(0,0,0,.10);}
.agenda-num{cursor:text;}
.agenda-add-row{padding:0 16px 28px;text-align:center;}
.agenda-add-item{display:inline-flex;align-items:center;gap:8px;background:none;border:1.5px dashed #93c5fd;border-radius:10px;color:#2563eb;font-size:13px;font-weight:700;padding:9px 22px;cursor:pointer;transition:.15s;}
.agenda-add-item:hover{border-color:#2563eb;background:rgba(37,99,235,.06);}

.agenda-footer-url{font-size:11px;color:rgba(255,255,255,.75);}

/* ── Cover Slide ── */

.cover-wrap{font-family:'Segoe UI',system-ui,sans-serif;position:relative;display:flex;flex-direction:column;height:calc(100vh - 190px);min-height:420px;max-height:600px;border-radius:16px;overflow:hidden;box-shadow:0 8px 40px rgba(0,0,0,.22);margin:8px 0 20px;background:#000000;padding:32px 40px 0;}

.cover-bg-art{position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:0;}

.cover-top{position:relative;z-index:1;display:flex;justify-content:space-between;align-items:flex-start;gap:24px;}

.cover-title{font-size:26px;font-weight:900;color:white;line-height:1.25;outline:none;cursor:text;margin:0;max-width:520px;}

.cover-top-left{display:flex;flex-direction:column;gap:14px;flex:1;min-width:0;}

.cover-title:empty::before{content:"Monthly Operations & Support Review";color:rgba(255,255,255,.4);}

.cover-badge-wrap{align-self:flex-start;}

.cover-badge{display:inline-flex;align-items:center;border:2px solid #38bdf8;border-radius:50px;padding:6px 22px;background:rgba(56,189,248,.08);cursor:text;}

.cover-badge-name{font-size:14px;font-weight:800;color:#7dd3fc;letter-spacing:2px;outline:none;cursor:text;min-width:80px;text-align:center;text-transform:uppercase;}

.cover-badge-name:empty::before{content:"Company Name";color:rgba(125,211,252,.5);text-transform:none;letter-spacing:0;}

.cover-logo-box{margin-top:10px;width:180px;height:90px;border-radius:12px;background:rgba(255,255,255,.95);display:flex;align-items:center;justify-content:center;cursor:pointer;overflow:hidden;border:2px dashed rgba(255,255,255,.3);}
.cover-logo-box:hover{border-color:#38bdf8;}
.cover-logo-box img{max-width:100%;max-height:100%;object-fit:contain;}
.cover-logo-box .cover-logo-placeholder{font-size:12px;color:#94a3b8;text-align:center;padding:8px;font-family:'Segoe UI',sans-serif;}
.cover-logo-input{display:none;}
.cover-vc-logo{align-self:flex-start;font-family:'Arial',Arial,sans-serif;font-size:22px;font-weight:800;color:white;letter-spacing:6px;text-align:right;white-space:nowrap;}

.cover-vc-logo em{font-style:normal;font-weight:300;font-size:16px;color:#cbd5e1;letter-spacing:2px;margin-left:2px;}

.cover-center{flex:1;display:flex;align-items:center;justify-content:center;gap:48px;position:relative;z-index:1;padding:20px 0;}

.cover-illustration{width:240px;height:auto;flex-shrink:0;}

.cover-tagline-area{display:flex;flex-direction:column;align-items:flex-start;gap:4px;}

.cover-tagline-main{font-size:28px;font-weight:800;color:white;letter-spacing:.5px;}

.cover-tagline-main.accent{color:#38bdf8;}

.cover-tagline-sub{font-size:16px;font-style:italic;color:rgba(255,255,255,.5);margin:2px 0;}

.cover-bottom{position:relative;z-index:1;padding-bottom:18px;}

.cover-presented{display:flex;flex-direction:column;gap:8px;}

.cover-presented-by{font-size:13px;color:#7dd3fc;font-weight:700;outline:none;cursor:text;}

.cover-presented-by:empty::before{content:"Presented by: VariTec Consulting";color:rgba(125,211,252,.5);}

.cover-date-val{font-size:12px;color:rgba(255,255,255,.55);font-weight:600;outline:none;cursor:text;letter-spacing:1px;text-transform:uppercase;}

.cover-date-val:empty::before{content:"Month-Year";color:rgba(255,255,255,.35);}

.cover-footer-bar{background:linear-gradient(90deg,#038dbd,#0284c7);margin:0 -40px;padding:11px 40px;font-size:11px;color:white;font-weight:700;text-align:right;letter-spacing:1px;position:relative;z-index:1;}


/* ── Thank You Slide ── */

.ty-wrap{font-family:'Segoe UI',system-ui,sans-serif;position:relative;display:flex;flex-direction:column;height:calc(100vh - 190px);min-height:420px;max-height:600px;border-radius:16px;overflow:hidden;box-shadow:0 8px 40px rgba(0,0,0,.22);margin:8px 0 20px;background:white;}

.ty-left{position:absolute;top:0;left:0;width:100%;height:100%;z-index:0;}

.ty-left-art{display:block;width:100%;height:100%;}

.ty-right{position:relative;z-index:1;flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:40px;}

.ty-vc-logo{position:absolute;top:32px;right:40px;font-family:'Arial',Arial,sans-serif;font-size:20px;font-weight:800;color:white;letter-spacing:5px;}

.ty-vc-logo em{font-style:normal;font-weight:300;font-size:14px;color:#cbd5e1;letter-spacing:2px;margin-left:2px;}

.ty-title{font-size:48px;font-weight:900;color:white;letter-spacing:2px;margin-bottom:18px;}

.ty-divider{width:60px;height:4px;background:#38bdf8;border-radius:2px;margin-bottom:22px;}

.ty-message{font-size:15px;color:rgba(255,255,255,.75);line-height:1.8;text-align:center;max-width:420px;outline:none;cursor:text;font-weight:400;}

.ty-message:empty::before{content:"We remain committed to ensuring seamless serialization compliance through operational excellence and strategic partnership.";color:rgba(255,255,255,.4);}

.ty-footer-bar{position:absolute;bottom:0;left:0;right:0;background:linear-gradient(90deg,#038dbd,#0284c7);padding:11px 40px;font-size:11px;color:white;font-weight:700;text-align:right;letter-spacing:1px;}



.rr-wrap{font-family:'Segoe UI',system-ui,sans-serif;margin:8px 0 20px;border-radius:14px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.12);}
.rr-banner{background:linear-gradient(135deg,#0d152b 0%,#1e3868 55%,#0a2a5e 100%);padding:16px 24px 16px 28px;position:relative;overflow:hidden;}
.rr-banner::before{content:'';position:absolute;left:0;top:0;bottom:0;width:5px;background:#0099CC;}
.rr-banner::after{content:'';position:absolute;right:-30px;top:-40px;width:160px;height:160px;border-radius:50%;background:rgba(255,255,255,.05);pointer-events:none;}
.rr-banner-title{font-size:20px;font-weight:900;color:white;letter-spacing:1px;text-transform:uppercase;line-height:1.1;position:relative;z-index:1;}
.rr-banner-sub{font-size:12px;color:rgba(77,178,222,.7);margin-top:5px;font-weight:500;position:relative;z-index:1;}
.rr-cards-area{background:linear-gradient(180deg,#eef2f7 0%,#e8edf5 100%);padding:16px;border-radius:0 0 14px 14px;}
.rr-cards-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;align-items:stretch;}
.rr-card{background:white;border-radius:10px;padding:12px 16px;box-shadow:0 1px 6px rgba(0,0,0,.06);display:flex;align-items:center;gap:12px;position:relative;transition:.15s;border-left:4px solid var(--rr-accent,#2563eb);}
.rr-card:hover{box-shadow:0 4px 16px rgba(0,0,0,.12);}
.rr-num{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;color:white;flex-shrink:0;letter-spacing:.3px;background:linear-gradient(135deg,var(--rr-accent,#2563eb),var(--rr-accent-light,#60a5fa));box-shadow:0 2px 6px rgba(0,0,0,.18);}
.rr-card-body{flex:1;min-width:0;}
.rr-card-title{font-size:15px;font-weight:600;color:#1e293b;outline:none;cursor:text;min-height:18px;line-height:1.45;}
.rr-card-title.rr-bold{font-weight:600;color:#1e293b;font-size:15px;}
.rr-card-title:empty::before{content:attr(data-ph);color:#cbd5e1;font-weight:400;pointer-events:none;}
.rr-card-desc{font-size:13px;color:#64748b;outline:none;cursor:text;min-height:14px;line-height:1.5;margin-top:2px;}
.rr-card-desc:empty{display:none;}
.rr-card-desc:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;pointer-events:none;display:none;}
.rr-card.rr-empty-desc .rr-card-desc:empty::before{display:inline;}
.rr-del-card{position:absolute;top:6px;right:6px;background:#fee2e2;border:1.5px solid #fecaca;color:#dc2626;cursor:pointer;font-size:9px;font-weight:700;padding:2px 6px;border-radius:6px;transition:.15s;display:none;font-family:Segoe UI,sans-serif;}
.rr-card:hover .rr-del-card{display:block;}
.rr-del-card:hover{background:#dc2626;color:white;border-color:#dc2626;}
.rr-add-btn{grid-column:span 2;border:2px dashed #0099CC;border-radius:10px;background:rgba(0,153,204,.05);color:#0099CC;font-size:13px;font-weight:700;cursor:pointer;transition:.15s;display:flex;align-items:center;justify-content:center;padding:12px;gap:6px;margin-top:2px;}
.rr-add-btn:hover{background:rgba(0,153,204,.12);border-color:#0099CC;}
@media print{
.rr-wrap{box-shadow:none!important;break-inside:avoid!important;}
.rr-cards-grid{display:grid!important;grid-template-columns:repeat(2,1fr)!important;gap:6px!important;}
.rr-card{break-inside:avoid!important;box-shadow:none!important;}
.rr-del-card,.rr-add-btn{display:none!important;}
.rr-num{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}
}


.pi-wrap{font-family:'Segoe UI',system-ui,sans-serif;margin:8px 0 20px;border-radius:14px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.12);position:relative;}
.pi-wrap::before{content:'';position:absolute;left:0;top:0;bottom:0;width:5px;background:#10b981;z-index:2;}
.pi-banner{background:linear-gradient(135deg,#0d152b 0%,#1e3868 55%,#0a2a5e 100%);padding:14px 24px 14px 30px;position:relative;overflow:hidden;display:flex;align-items:center;gap:12px;}
.pi-banner-icon{display:none;}
.pi-banner-title{font-size:18px;font-weight:900;color:white;letter-spacing:1px;text-transform:uppercase;line-height:1.15;position:relative;z-index:1;}
.pi-banner::after{content:'';position:absolute;right:0;top:0;bottom:0;width:220px;pointer-events:none;background-image:radial-gradient(circle,rgba(255,255,255,.16) 1.3px,transparent 1.6px);background-size:11px 11px;-webkit-mask-image:linear-gradient(to left,rgba(0,0,0,1),rgba(0,0,0,0));mask-image:linear-gradient(to left,rgba(0,0,0,1),rgba(0,0,0,0));}
.pi-body{background:linear-gradient(180deg,#eef2f7 0%,#e8edf5 100%);padding:16px 16px 16px 19px;border-radius:0 0 14px 14px;display:flex;flex-direction:column;gap:10px;}
.pi-item{background:white;border-radius:10px;padding:12px 16px;box-shadow:0 1px 6px rgba(0,0,0,.06);position:relative;transition:.15s;}
.pi-item:hover{box-shadow:0 4px 16px rgba(0,0,0,.12);}
.pi-item-row{display:flex;align-items:center;gap:12px;}
.pi-main-icon{width:11px;height:11px;border-radius:50%;flex-shrink:0;background:#10b981;box-shadow:0 0 0 3px rgba(16,185,129,.15);}
.pi-main-icon.pi-ghost{background:transparent;border:2px dashed #cbd5e1;}
.pi-item-text{flex:1;font-size:15.5px;font-weight:700;color:#0d152b;outline:none;cursor:text;line-height:1.5;min-height:18px;min-width:0;}
.pi-item-text:empty::before{content:attr(data-ph);color:#cbd5e1;font-weight:400;font-style:italic;pointer-events:none;}
.pi-chevron{display:none;}
.pi-del-item{position:absolute;top:8px;right:6px;background:#fee2e2;border:1.5px solid #fecaca;color:#dc2626;cursor:pointer;font-size:9px;font-weight:700;padding:2px 6px;border-radius:6px;transition:.15s;display:none;font-family:Segoe UI,sans-serif;}
.pi-item:hover .pi-del-item{display:block;}
.pi-del-item:hover{background:#dc2626;color:white;border-color:#dc2626;}
.pi-status-cols{display:flex;align-items:center;gap:14px;flex-shrink:0;}
.pi-status-col{display:flex;flex-direction:column;align-items:center;gap:3px;}
.pi-status-label{font-size:9.5px;font-weight:800;color:#94a3b8;text-transform:uppercase;letter-spacing:.5px;}
.pi-pill{display:inline-flex;align-items:center;gap:4px;padding:4px 12px;border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;transition:.15s;border:1.5px solid transparent;}
.pi-pill:hover{filter:brightness(.97);}
.pi-pill-completed{background:#dcfce7;color:#15803d;}
.pi-pill-pending{background:#fef3c7;color:#92400e;}
.pi-pill-inprogress{background:#fde68a;color:#92400e;}
.pi-pill-hold{background:#fee2e2;color:#991b1b;}
.pi-sub-list{margin:10px 0 4px 34px;display:flex;flex-direction:column;gap:6px;}
.pi-sub-row{display:flex;align-items:center;gap:12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:9px;padding:8px 12px;position:relative;}
.pi-sub-icon{width:8px;height:8px;border-radius:50%;flex-shrink:0;background:#94a3b8;}
.pi-sub-text{flex:1;font-size:14px;font-weight:600;color:#334155;outline:none;cursor:text;line-height:1.5;min-width:0;}
.pi-sub-text:empty::before{content:attr(data-ph);color:#cbd5e1;font-style:italic;font-weight:400;pointer-events:none;}
.pi-del-sub{flex-shrink:0;background:none;border:1px solid #fecaca;color:#f87171;border-radius:50%;width:18px;height:18px;cursor:pointer;font-size:10px;display:flex;align-items:center;justify-content:center;opacity:0;transition:.15s;padding:0;line-height:1;}
.pi-sub-row:hover .pi-del-sub{opacity:1;}
.pi-del-sub:hover{background:#dc2626;border-color:#dc2626;color:white;}
.pi-add-sub-wrap{display:flex;justify-content:center;margin-top:8px;}
.pi-add-sub{padding:5px 16px;background:#ecfdf5;border:1.5px dashed #6ee7b7;border-radius:20px;color:#0f9b6e;font-size:11.5px;font-weight:700;cursor:pointer;transition:.15s;}
.pi-add-sub:hover{border-color:#10b981;background:#d1fae5;}
.pi-add-item{border:2px dashed #93c5fd;border-radius:10px;background:#eff6ff;color:#2563eb;font-size:13px;font-weight:700;cursor:pointer;transition:.15s;display:flex;align-items:center;justify-content:center;padding:12px;gap:6px;}
.pi-add-item:hover{background:#dbeafe;border-color:#60a5fa;}
@media print{
.pi-wrap{box-shadow:none!important;break-inside:avoid!important;}
.pi-item{break-inside:avoid!important;box-shadow:none!important;}
.pi-del-item,.pi-del-sub,.pi-add-sub-wrap,.pi-add-item,.pi-chevron{display:none!important;}
.pi-main-icon,.pi-sub-icon,.pi-pill,.pi-wrap::before{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}
}

.insert-table-modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);align-items:center;justify-content:center;z-index:9999;}

.insert-table-modal.show{display:flex;}

.itm-box{background:white;border-radius:12px;padding:24px;width:320px;box-shadow:0 8px 32px rgba(0,0,0,.2);}

.itm-box h4{margin-bottom:14px;color:#1e293b;font-size:15px;}

.itm-row{display:flex;gap:12px;margin-bottom:14px;}

.itm-row label{flex:1;font-size:12px;color:#555;font-weight:600;}

.itm-row input{width:100%;padding:7px 10px;border:1px solid #ddd;border-radius:6px;font-size:14px;margin-top:4px;outline:none;}

.itm-row input:focus{border-color:#8b3dff;}

.itm-btns{display:flex;gap:8px;margin-top:6px;}

.itm-btns button{flex:1;padding:9px;border:none;border-radius:7px;font-size:13px;cursor:pointer;font-weight:600;}

.itm-ok{background:#8b3dff;color:white;}.itm-ok:hover{background:#7230d4;}

.itm-cancel{background:#f0f0f0;color:#555;}.itm-cancel:hover{background:#e0e0e0;}

.paste-hint{font-size:12px;color:#bbb;padding:4px 40px;flex-shrink:0;}

.statusbar{height:28px;font-size:12px;color:#aaa;padding:0 40px;line-height:28px;border-top:1px solid #f0f0f0;flex-shrink:0;}

.row-action-bar{display:none;position:fixed;z-index:9200;background:#1e293b;border-radius:8px;padding:4px 5px;gap:4px;box-shadow:0 4px 16px rgba(0,0,0,.4);align-items:center;}

.row-action-bar button{padding:5px 12px;border:none;border-radius:6px;cursor:pointer;font-size:11px;font-weight:700;color:white;white-space:nowrap;transition:.12s;}

.rab-clone{background:#7c3aed;}.rab-clone:hover{background:#6d28d9;}

.rab-delete{background:#dc2626;}.rab-delete:hover{background:#b91c1c;}

.rab-close{background:#475569;padding:4px 8px!important;}.rab-close:hover{background:#334155;}


.note-footer{height:36px;background:linear-gradient(90deg,#038dbd,#0d152b);display:flex;align-items:center;padding:0 20px 0 16px;flex-shrink:0;}

.note-footer-url{font-size:12px;color:rgba(255,255,255,.8);letter-spacing:.3px;text-decoration:none;font-weight:500;}

.note-footer-url:hover{color:white;text-decoration:underline;}

.saved{color:#38a169;}.saving{color:#d69e2e;}.error{color:#e53e3e;}

.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);align-items:center;justify-content:center;z-index:999;}

.modal-overlay.show{display:flex;}

.modal{background:white;border-radius:14px;padding:28px;width:440px;box-shadow:0 12px 40px rgba(0,0,0,.2);}

.modal h3{margin-bottom:6px;color:#1e293b;font-size:17px;}

.modal p{font-size:13px;color:#64748b;margin-bottom:16px;line-height:1.5;}

.modal input[type=text]{width:100%;padding:10px 12px;border:1.5px solid #ddd;border-radius:8px;font-size:15px;margin-bottom:14px;outline:none;}

.modal input[type=text]:focus{border-color:#8b3dff;}

.modal-btns{display:flex;gap:10px;margin-top:4px;}

.modal-btns button{flex:1;padding:10px;border:none;border-radius:8px;font-size:14px;cursor:pointer;font-weight:600;}

.btn-confirm{background:#8b3dff;color:white;}.btn-confirm:hover{background:#7230d4;}

.btn-cancel{background:#f0f0f0;color:#555;}.btn-cancel:hover{background:#e0e0e0;}

/* ── Hide subtitle/date from all template headers ── */

.ob-header-sub,.ob-month-wrap,.sc-month-wrap,.si-banner-subtitle,

.si-banner-meta,.si-month-wrap,.exec-subtitle,.agenda-subtitle{display:none!important;}

@media print{body{overflow:visible!important;}#sc-inline-wrap canvas{display:none!important;}#sc-chart-pdf-img{display:block!important;}.sc-chart-del-btn,.topbar,.sidebar,.pages,.toolbar,.paste-hint,.statusbar,.tb-btn-ppt,.si-del-card,.si-add-btn,.si-action-add-btn,.si-action-del,.date-cal-btn,.note-footer,.fa-drag-handle,.si-drag-handle{display:none!important;} /* HIDE_DRAG_HANDLES_PRINT_V1 */.fa-card{break-inside:avoid!important;page-break-inside:avoid!important;} /* FA_CARD_BREAK_AVOID_V1 */.main{display:block!important;}.editor{overflow:visible!important;height:auto!important;}#content{overflow:visible!important;height:auto!important;padding:20px 40px!important;}.si-cards-grid{display:grid!important;grid-template-columns:repeat(2,1fr)!important;}@page{margin:1.5cm;size:A4 landscape;}}

</style>
<style>
/* Hide + New Page from sidebar tree */
.stree-add-page { display: none !important; }
</style>

<style>
/* ===== Sidebar clean reset — overrides all previous sidebar CSS ===== */
.sidebar {
  width: 240px !important;
  background: #f5f4fb !important;
  border-right: 1px solid #e2e0f0 !important;
  display: flex !important;
  flex-direction: column !important;
  flex-shrink: 0 !important;
  overflow: hidden !important;
}
.sidebar-header {
  padding: 12px 14px 8px !important;
  font-size: 10px !important;
  font-weight: 800 !important;
  color: #9896be !important;
  text-transform: uppercase !important;
  letter-spacing: 1.5px !important;
  background: #f5f4fb !important;
  border-bottom: 1px solid #e8e6f4 !important;
}
.sections-list {
  flex: 1 !important;
  overflow-y: auto !important;
  padding: 4px 0 !important;
  background: #f5f4fb !important;
}

/* CLIENT ROW */
.stree-client-row {
  display: flex !important;
  align-items: center !important;
  gap: 6px !important;
  padding: 7px 8px 7px 8px !important;
  cursor: pointer !important;
  border-radius: 6px !important;
  margin: 1px 4px !important;
  background: transparent !important;
  transition: background .12s !important;
  color: #2d2b55 !important;
}
.stree-client-row:hover { background: #eceaf6 !important; }
.stree-client-row.stree-open { background: #e8e6f4 !important; }
.stree-arr {
  font-size: 9px !important;
  color: #9896be !important;
  width: 12px !important;
  flex-shrink: 0 !important;
  display: inline-block !important;
}
.stree-client-name {
  flex: 1 !important;
  font-size: 13px !important;
  font-weight: 700 !important;
  color: #2d2b55 !important;
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
}
.stree-client-row.stree-open .stree-client-name { color: #4338ca !important; }

/* MONTH ROW */
.stree-month-row {
  display: flex !important;
  align-items: center !important;
  gap: 6px !important;
  padding: 6px 8px 6px 20px !important;
  cursor: pointer !important;
  border-radius: 6px !important;
  margin: 1px 4px !important;
  background: transparent !important;
  transition: background .12s !important;
}
.stree-month-row:hover { background: #eceaf6 !important; }
.stree-month-row.stree-open { background: #e4e2f4 !important; }
.stree-month-arr {
  font-size: 9px !important;
  color: #9896be !important;
  width: 12px !important;
  flex-shrink: 0 !important;
  display: inline-block !important;
}
.stree-month-name {
  flex: 1 !important;
  font-size: 12.5px !important;
  font-weight: 600 !important;
  color: #3a3870 !important;
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
}
.stree-month-row.stree-open .stree-month-name { color: #4338ca !important; }
.stree-folder-icon { flex-shrink: 0 !important; }

/* MONTH PAGES CONTAINER */
.stree-month-pages {
  padding: 2px 0 4px 0 !important;
  margin: 0 !important;
  position: static !important;
}
.stree-month-pages::before { display: none !important; }

/* PAGE ITEMS */
.stree-page,
.stree-client-body > .stree-page {
  display: flex !important;
  align-items: center !important;
  gap: 8px !important;
  padding: 6px 10px 6px 36px !important;
  text-decoration: none !important;
  color: #4a4878 !important;
  font-size: 12.5px !important;
  border-radius: 6px !important;
  margin: 1px 4px !important;
  transition: background .1s !important;
  position: relative !important;
}
.stree-page::before { display: none !important; }
.stree-page:hover {
  background: #eceaf6 !important;
  color: #1e1c40 !important;
}
.stree-page-active {
  background: #ddd9f5 !important;
  color: #3b2fa0 !important;
  font-weight: 600 !important;
}
.stree-dot {
  width: 8px !important;
  height: 8px !important;
  border-radius: 50% !important;
  flex-shrink: 0 !important;
}
.stree-page-title {
  flex: 1 !important;
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
}

/* DOT COLORS */
.stree-page:nth-child(1)  .stree-dot { background: #8b5cf6 !important; }
.stree-page:nth-child(2)  .stree-dot { background: #3b82f6 !important; }
.stree-page:nth-child(3)  .stree-dot { background: #ec4899 !important; }
.stree-page:nth-child(4)  .stree-dot { background: #10b981 !important; }
.stree-page:nth-child(5)  .stree-dot { background: #8b5cf6 !important; }
.stree-page:nth-child(6)  .stree-dot { background: #ef4444 !important; }
.stree-page:nth-child(7)  .stree-dot { background: #06b6d4 !important; }
.stree-page:nth-child(8)  .stree-dot { background: #f59e0b !important; }
.stree-page:nth-child(9)  .stree-dot { background: #3b82f6 !important; }
.stree-page:nth-child(10) .stree-dot { background: #ec4899 !important; }
.stree-page:nth-child(11) .stree-dot { background: #10b981 !important; }
.stree-page:nth-child(12) .stree-dot { background: #8b5cf6 !important; }

/* PAGE BUTTONS */
.stree-page-btns { display: none !important; }
.stree-page:hover .stree-page-btns {
  display: flex !important;
  gap: 2px !important;
  flex-shrink: 0 !important;
  align-items: center !important;
}
.stree-page .clone-page-btn,
.stree-page .del-page-btn {
  background: none !important;
  border: none !important;
  cursor: pointer !important;
  font-size: 11px !important;
  padding: 1px 3px !important;
  color: #b0aed0 !important;
  line-height: 1 !important;
  opacity: .8 !important;
}
.stree-page .clone-page-btn:hover,
.stree-page .del-page-btn:hover { color: #ef4444 !important; opacity: 1 !important; }

/* CLIENT/MONTH ACTION BUTTONS */
.stree-client-actions,
.stree-month-actions { display: none !important; }
.stree-client-row:hover .stree-client-actions,
.stree-month-row:hover .stree-month-actions {
  display: flex !important;
  gap: 2px !important;
  flex-shrink: 0 !important;
  align-items: center !important;
}
.sec-icon {
  background: none !important; border: none !important;
  cursor: pointer !important; font-size: 11px !important;
  padding: 2px !important; color: #b0aed0 !important;
  line-height: 1 !important;
}
.sec-icon:hover { color: #4338ca !important; }
.sec-clone-btn {
  background: none !important; border: none !important;
  cursor: pointer !important; padding: 2px 3px !important;
  color: #b0aed0 !important; display: flex !important;
  align-items: center !important;
}
.sec-clone-btn:hover { color: #4338ca !important; }
.sec-clone-btn svg { stroke: currentColor !important; }

/* ADD BUTTONS */
.stree-add-month {
  display: block !important;
  margin: 3px 8px 3px 36px !important;
  padding: 4px 10px !important;
  background: none !important;
  border: 1.5px dashed #d0cce8 !important;
  color: #9896be !important;
  font-size: 11px !important;
  font-weight: 600 !important;
  border-radius: 5px !important;
  cursor: pointer !important;
  text-align: left !important;
  transition: .12s !important;
}
.stree-add-month:hover {
  border-color: #4338ca !important;
  color: #4338ca !important;
  background: #f0eeff !important;
}
.stree-add-page {
  display: block !important;
  padding: 4px 10px 4px 44px !important;
  color: #b0aed0 !important;
  font-size: 11px !important;
  text-decoration: none !important;
}
.stree-add-page:hover { color: #4338ca !important; }
.add-section-btn {
  margin: 6px 8px 10px !important;
  padding: 8px !important;
  background: none !important;
  border: 1.5px dashed #d0cce8 !important;
  color: #9896be !important;
  font-size: 12px !important;
  font-weight: 600 !important;
  border-radius: 6px !important;
  cursor: pointer !important;
  transition: .15s !important;
  text-align: center !important;
  display: block !important;
}
.add-section-btn:hover {
  border-color: #4338ca !important;
  color: #4338ca !important;
  background: #f0eeff !important;
}
</style>

<style>
/* ===== Fix page item scratched look ===== */
/* Remove the horizontal tick that overlaps text */
.stree-page::before { display: none !important; }

/* Redesign pages to be clean with dot + text */
.stree-month-pages {
  position: relative;
  padding: 2px 0 4px 0;
  margin-left: 28px;
}
/* Vertical line */
.stree-month-pages::before {
  content: '';
  position: absolute;
  left: 10px;
  top: 4px; bottom: 8px;
  width: 1.5px;
  background: #d0cce8;
  border-radius: 1px;
}
.stree-page {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 5px 10px 5px 28px;
  text-decoration: none;
  color: #4a4870;
  font-size: 12.5px;
  transition: .1s;
  border-radius: 6px;
  margin: 1px 6px 1px 0;
  position: relative;
}
.stree-page:hover { background: #eeecf8; color: #1e1c40; }
.stree-page-active {
  background: #e8e4ff !important;
  color: #3b2fa0 !important;
  font-weight: 600;
}
.stree-dot {
  width: 8px; height: 8px;
  border-radius: 50%;
  flex-shrink: 0;
  position: relative;
  z-index: 1;
}
.stree-page-title {
  flex: 1;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

/* Pages directly under client (no month) */
.stree-client-body > .stree-page {
  padding-left: 36px;
}
</style>

<style>
</style>

<style>
/* ===== Sidebar tree with connecting lines (matches reference) ===== */
.sidebar{
  width:240px;background:#f8f7fd;
  display:flex;flex-direction:column;flex-shrink:0;overflow:hidden;
  border-right:1px solid #e5e2f0;
}
.sidebar-header{
  padding:12px 16px 8px;font-size:10px;font-weight:900;
  color:#9896c4;text-transform:uppercase;letter-spacing:1.5px;
  border-bottom:1px solid #eceaf5;
}
.sections-list{flex:1;overflow-y:auto;padding:6px 0 12px;}
.sections-list::-webkit-scrollbar{width:3px;}
.sections-list::-webkit-scrollbar-thumb{background:#cccae0;border-radius:2px;}

/* ── CLIENT row (top level, e.g. Natco) ── */
.stree-client-row{
  display:flex;align-items:center;gap:6px;
  padding:7px 8px 7px 10px;cursor:pointer;
  transition:.12s;border-radius:7px;margin:1px 6px;
}
.stree-client-row:hover{background:#eeecf8;}
.stree-client-row.stree-open{background:#eceaf7;}
.stree-arr{
  font-size:9px;color:#a09ec8;width:12px;flex-shrink:0;
  user-select:none;transition:transform .15s;display:inline-block;
}
.stree-client-name{
  flex:1;font-size:13px;font-weight:700;color:#1e1c40;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
}
.stree-client-row.stree-open .stree-client-name{color:#4338c9;}
.stree-client-actions{display:none;gap:1px;flex-shrink:0;align-items:center;}
.stree-client-row:hover .stree-client-actions{display:flex;}
.sec-icon{
  background:none;border:none;cursor:pointer;font-size:11px;
  padding:2px;color:#b8b6d8;line-height:1;
}
.sec-icon:hover{color:#4338c9;}
.sec-clone-btn{
  background:none;border:none;cursor:pointer;
  padding:2px 3px;color:#b8b6d8;line-height:1;display:flex;
}
.sec-clone-btn:hover{color:#4338c9;}

/* ── MONTH row (mid level, e.g. July 2025) ── */
.stree-month{margin:2px 0;}
.stree-month-row{
  display:flex;align-items:center;gap:6px;
  padding:6px 8px 6px 22px;cursor:pointer;
  transition:.12s;border-radius:6px;margin:0 6px;
}
.stree-month-row:hover{background:#eeecf8;}
.stree-month-row.stree-open{background:#e8e5f7;}
.stree-month-arr{
  font-size:9px;color:#a09ec8;width:12px;flex-shrink:0;
  user-select:none;display:inline-block;
}

/* Purple folder SVG for month */
.stree-folder-icon{flex-shrink:0;}

.stree-month-name{
  flex:1;font-size:12.5px;font-weight:700;color:#2a2860;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
}
.stree-month-row.stree-open .stree-month-name{color:#4338c9;}
.stree-month-actions{display:none;gap:1px;align-items:center;}
.stree-month-row:hover .stree-month-actions{display:flex;}

/* ── PAGES with vertical tree line ── */
.stree-month-pages{
  position:relative;
  padding:2px 0 4px 0;
  margin-left:30px;         /* indent from left */
}
/* Vertical connecting line */
.stree-month-pages::before{
  content:'';
  position:absolute;
  left:14px;                /* align with dots */
  top:0; bottom:8px;
  width:1.5px;
  background:#d8d5ee;
  border-radius:1px;
}

.stree-page{
  display:flex;align-items:center;gap:8px;
  padding:5px 10px 5px 32px;   /* 32px = space for line + dot */
  text-decoration:none;color:#4a4870;font-size:12.5px;
  transition:.1s;border-radius:6px;margin:1px 6px 1px 0;
  position:relative;
}
.stree-page:hover{background:#eeecf8;color:#1e1c40;}
.stree-page-active{
  background:#e8e4ff!important;
  color:#4338c9!important;font-weight:600;
}

/* Horizontal tick from vertical line to dot */
.stree-page::before{
  content:'';
  position:absolute;
  left:14px; top:50%;
  width:10px; height:1.5px;
  background:#d8d5ee;
  transform:translateY(-50%);
}

/* Colored dot */
.stree-dot{
  width:8px;height:8px;border-radius:50%;
  flex-shrink:0;position:relative;z-index:1;
  background:#6366f1;   /* fallback */
}
.stree-page-active .stree-dot{background:#4338c9!important;}

/* Dot colors by position */
.stree-page:nth-child(1)  .stree-dot{background:#8b5cf6;}
.stree-page:nth-child(2)  .stree-dot{background:#3b82f6;}
.stree-page:nth-child(3)  .stree-dot{background:#ec4899;}
.stree-page:nth-child(4)  .stree-dot{background:#10b981;}
.stree-page:nth-child(5)  .stree-dot{background:#8b5cf6;}
.stree-page:nth-child(6)  .stree-dot{background:#ef4444;}
.stree-page:nth-child(7)  .stree-dot{background:#06b6d4;}
.stree-page:nth-child(8)  .stree-dot{background:#f59e0b;}
.stree-page:nth-child(9)  .stree-dot{background:#3b82f6;}
.stree-page:nth-child(10) .stree-dot{background:#ec4899;}
.stree-page:nth-child(11) .stree-dot{background:#10b981;}
.stree-page:nth-child(12) .stree-dot{background:#8b5cf6;}

.stree-page-title{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-page-btns{display:none;gap:1px;flex-shrink:0;align-items:center;}
.stree-page:hover .stree-page-btns{display:flex;}
.stree-page .clone-page-btn,
.stree-page .del-page-btn{
  background:none;border:none;cursor:pointer;
  font-size:11px;padding:1px 3px;color:#c0bede;line-height:1;
}
.stree-page .clone-page-btn:hover,
.stree-page .del-page-btn:hover{color:#ef4444;}

/* ── Add buttons ── */
.stree-add-month{
  display:block;margin:4px 6px 4px 52px;
  padding:4px 10px;background:none;
  border:1.5px dashed #d0cce8;color:#a09ec8;
  font-size:11px;font-weight:600;border-radius:5px;
  cursor:pointer;text-align:left;transition:.12s;
}
.stree-add-month:hover{border-color:#4338c9;color:#4338c9;background:#f0eeff;}
.stree-add-page{
  display:block;padding:5px 10px 5px 52px;
  color:#b8b6d8;font-size:11.5px;text-decoration:none;
}
.stree-add-page:hover{color:#4338c9;}
.add-section-btn{
  display:block;margin:8px 10px 10px;padding:9px;
  background:none;border:1.5px dashed #d0cce8;color:#a09ec8;
  font-size:12px;font-weight:600;border-radius:7px;
  cursor:pointer;transition:.15s;text-align:center;
}
.add-section-btn:hover{border-color:#4338c9;color:#4338c9;background:#f0eeff;}
</style>

<style>
/* ===== Sidebar final design ===== */
.sidebar{width:235px;background:#f7f7fb;display:flex;flex-direction:column;flex-shrink:0;overflow:hidden;border-right:1px solid #e0dff0;}
.sidebar-header{padding:14px 14px 8px;font-size:10px;font-weight:900;color:#7b7aaa;text-transform:uppercase;letter-spacing:1.8px;border-bottom:1px solid #ebebf5;}
.sections-list{flex:1;overflow-y:auto;padding:6px 0;}
.sections-list::-webkit-scrollbar{width:3px;}
.sections-list::-webkit-scrollbar-thumb{background:#d0ceee;border-radius:2px;}

/* ── CLIENT (top level) ── */
.stree-client-row{display:flex;align-items:center;gap:6px;padding:7px 8px 7px 8px;cursor:pointer;transition:.12s;border-radius:7px;margin:1px 5px;}
.stree-client-row:hover{background:#eeecf8;}
.stree-client-row.stree-open{background:#e8e6f7;}
.stree-arr{font-size:8px;color:#8886b8;width:10px;flex-shrink:0;user-select:none;display:inline-block;}
.stree-client-name{flex:1;font-size:13px;font-weight:700;color:#1e1d40;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-client-row.stree-open .stree-client-name{color:#3b2fa0;}
.stree-client-actions{display:none;gap:1px;flex-shrink:0;align-items:center;}
.stree-client-row:hover .stree-client-actions{display:flex;}
.sec-icon{background:none;border:none;cursor:pointer;font-size:11px;padding:1px 2px;color:#b0aed0;line-height:1;}
.sec-icon:hover{color:#3b2fa0;}
.sec-clone-btn{background:none;border:none;cursor:pointer;padding:2px 3px;color:#b0aed0;line-height:1;display:flex;align-items:center;}
.sec-clone-btn:hover svg{stroke:#3b2fa0;}

/* ── MONTH (mid level) ── */
.stree-month-row{display:flex;align-items:center;gap:6px;padding:6px 8px 6px 24px;cursor:pointer;transition:.12s;border-radius:6px;margin:1px 5px;}
.stree-month-row:hover{background:#eeecf8;}
.stree-month-row.stree-open{background:#e4e1f5;}
.stree-month-arr{font-size:8px;color:#8886b8;width:10px;flex-shrink:0;user-select:none;display:inline-block;}
.stree-month-name{flex:1;font-size:12.5px;font-weight:600;color:#2d2b60;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-month-row.stree-open .stree-month-name{color:#3b2fa0;}
.stree-month-actions{display:none;gap:1px;align-items:center;}
.stree-month-row:hover .stree-month-actions{display:flex;}

/* ── PAGES ── */
.stree-page{display:flex;align-items:center;gap:8px;padding:6px 10px 6px 44px;text-decoration:none;color:#4a4870;font-size:12.5px;transition:.1s;border-radius:6px;margin:1px 5px;}
.stree-page:hover{background:#eeecf8;color:#1e1d40;}
.stree-page-active{background:#ddd9f7!important;color:#3b2fa0!important;font-weight:600;}

/* Colored bullet dots */
.stree-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.stree-page:nth-child(7n+1) .stree-dot{background:#f59e0b;}
.stree-page:nth-child(7n+2) .stree-dot{background:#3b82f6;}
.stree-page:nth-child(7n+3) .stree-dot{background:#ec4899;}
.stree-page:nth-child(7n+4) .stree-dot{background:#10b981;}
.stree-page:nth-child(7n+5) .stree-dot{background:#8b5cf6;}
.stree-page:nth-child(7n+6) .stree-dot{background:#ef4444;}
.stree-page:nth-child(7n+7) .stree-dot{background:#06b6d4;}

.stree-page-title{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-page-btns{display:none;gap:1px;flex-shrink:0;align-items:center;}
.stree-page:hover .stree-page-btns{display:flex;}
.stree-page .clone-page-btn,.stree-page .del-page-btn{background:none;border:none;cursor:pointer;font-size:11px;padding:1px 3px;color:#b0aed0;opacity:.8;line-height:1;}
.stree-page .clone-page-btn:hover,.stree-page .del-page-btn:hover{color:#ef4444;opacity:1;}

/* ── BUTTONS ── */
.stree-add-month{display:block;width:calc(100% - 60px);margin:3px 5px 3px 50px;padding:4px 8px;background:none;border:1.5px dashed #c5c2e8;color:#8886b8;font-size:11px;font-weight:600;border-radius:5px;cursor:pointer;text-align:left;transition:.12s;}
.stree-add-month:hover{border-color:#5b4fc9;color:#5b4fc9;background:#f0eef9;}
.stree-add-page{display:block;padding:5px 10px 5px 52px;color:#a5a3cc;font-size:11.5px;text-decoration:none;transition:.1s;}
.stree-add-page:hover{color:#3b2fa0;}
.add-section-btn{display:block;margin:8px 10px 12px;padding:9px;background:none;border:1.5px dashed #c5c2e8;color:#8886b8;font-size:12px;font-weight:600;border-radius:7px;cursor:pointer;transition:.15s;text-align:center;width:calc(100%-20px);}
.add-section-btn:hover{border-color:#5b4fc9;color:#5b4fc9;background:#f0eef9;}
</style>

<style>
/* ===== Sidebar style v2 — matches reference image ===== */
.sidebar{width:230px;background:#f0eff5;display:flex;flex-direction:column;flex-shrink:0;overflow:hidden;border-right:1px solid #e2e0ea;}
.sidebar-header{padding:12px 14px 6px;font-size:10px;font-weight:800;color:#7c6fcd;text-transform:uppercase;letter-spacing:1.5px;}
.sections-list{flex:1;overflow-y:auto;padding:4px 0 8px;}
.sections-list::-webkit-scrollbar{width:4px;}
.sections-list::-webkit-scrollbar-thumb{background:rgba(0,0,0,.12);border-radius:2px;}

/* CLIENT row */
.stree-client-row{display:flex;align-items:center;gap:4px;padding:6px 8px 6px 6px;cursor:pointer;transition:.12s;border-radius:6px;margin:1px 4px;}
.stree-client-row:hover{background:rgba(0,0,0,.05);}
.stree-arr{font-size:9px;color:#5b5b8a;width:12px;flex-shrink:0;display:inline-block;}
.stree-folder{font-size:15px;flex-shrink:0;}
.stree-folder-sm{font-size:13px;flex-shrink:0;}
.stree-client-name{flex:1;font-size:13px;font-weight:600;color:#2d2b55;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-client-row.stree-open .stree-client-name{color:#4f3cc9;}
.stree-client-actions{display:none;gap:1px;flex-shrink:0;}
.stree-client-row:hover .stree-client-actions{display:flex;}
.sec-icon{background:none;border:none;cursor:pointer;font-size:11px;padding:1px 2px;color:#9ca3af;opacity:.6;}
.sec-icon:hover{opacity:1;}
.sec-clone-btn{background:none;border:none;cursor:pointer;padding:1px 2px;color:#9ca3af;opacity:.6;}
.sec-clone-btn:hover{opacity:1;}

/* CLIENT BODY */
.stree-client-body{padding-left:0;}

/* MONTH row */
.stree-month-row{display:flex;align-items:center;gap:4px;padding:5px 8px 5px 22px;cursor:pointer;transition:.12s;border-radius:5px;margin:1px 4px;}
.stree-month-row:hover{background:rgba(0,0,0,.05);}
.stree-month-row.stree-open .stree-month-name{color:#4f3cc9;font-weight:700;}
.stree-month-arr{font-size:9px;color:#5b5b8a;width:12px;flex-shrink:0;display:inline-block;transition:transform .15s;}
.stree-month-row.stree-open .stree-month-arr{transform:rotate(0deg);}
.stree-month-name{flex:1;font-size:12.5px;font-weight:600;color:#3d3b6e;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-month-actions{display:none;gap:1px;}
.stree-month-row:hover .stree-month-actions{display:flex;}

/* PAGE items */
.stree-page{display:flex;align-items:center;gap:7px;padding:5px 8px 5px 42px;text-decoration:none;color:#555;font-size:12.5px;transition:.1s;border-radius:5px;margin:1px 4px;}
.stree-page:hover{background:rgba(79,60,201,.08);color:#2d2b55;}
.stree-page-active{background:#ede9ff!important;color:#4f3cc9!important;font-weight:600;}
.stree-dot{width:7px;height:7px;border-radius:50%;background:#6366f1;flex-shrink:0;}
.stree-page-active .stree-dot{background:#4f3cc9;}
.stree-page:nth-child(7n+1) .stree-dot{background:#f59e0b;}
.stree-page:nth-child(7n+2) .stree-dot{background:#3b82f6;}
.stree-page:nth-child(7n+3) .stree-dot{background:#ec4899;}
.stree-page:nth-child(7n+4) .stree-dot{background:#10b981;}
.stree-page:nth-child(7n+5) .stree-dot{background:#8b5cf6;}
.stree-page:nth-child(7n+6) .stree-dot{background:#ef4444;}
.stree-page:nth-child(7n+7) .stree-dot{background:#06b6d4;}
.stree-page-title{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-page-btns{display:none;gap:1px;flex-shrink:0;}
.stree-page:hover .stree-page-btns{display:flex;}
.stree-page .clone-page-btn,.stree-page .del-page-btn{background:none;border:none;cursor:pointer;font-size:10px;padding:1px 2px;color:#aaa;}
.stree-page .clone-page-btn:hover,.stree-page .del-page-btn:hover{color:#ef4444;}

/* Buttons */
.stree-add-month{display:block;width:calc(100% - 52px);margin:3px 4px 3px 44px;padding:4px 8px;background:none;border:1px dashed rgba(79,60,201,.3);color:#7c6fcd;font-size:11px;border-radius:4px;cursor:pointer;text-align:left;}
.stree-add-month:hover{border-color:#4f3cc9;color:#4f3cc9;background:rgba(79,60,201,.05);}
.stree-add-page{display:block;padding:4px 8px 4px 48px;color:#9ca3af;font-size:11px;text-decoration:none;}
.stree-add-page:hover{color:#4f3cc9;}
.add-section-btn{margin:6px 10px 10px;padding:8px;background:none;border:1px dashed rgba(79,60,201,.3);color:#7c6fcd;font-size:12px;font-weight:600;border-radius:6px;cursor:pointer;transition:.15s;text-align:center;}
.add-section-btn:hover{border-color:#4f3cc9;color:#4f3cc9;background:rgba(79,60,201,.05);}
</style>

<style>
/* ===== 3-level sidebar tree ===== */
.sidebar{width:230px;background:#1a1726;display:flex;flex-direction:column;flex-shrink:0;overflow:hidden;}
.sidebar-header{padding:12px 14px 6px;font-size:10px;font-weight:800;color:#6d5dab;text-transform:uppercase;letter-spacing:1.5px;}
.sections-list{flex:1;overflow-y:auto;padding:4px 0 8px;}
.sections-list::-webkit-scrollbar{width:3px;}
.sections-list::-webkit-scrollbar-thumb{background:rgba(255,255,255,.1);border-radius:2px;}

/* Client (top level) */
.stree-client-row{display:flex;align-items:center;gap:6px;padding:8px 10px 8px 10px;cursor:pointer;transition:.12s;border-radius:6px;margin:1px 4px;}
.stree-client-row:hover{background:rgba(255,255,255,.06);}
.stree-client-row.stree-open{background:rgba(139,61,255,.15);}
.stree-arr{font-size:8px;color:#6d5dab;width:10px;flex-shrink:0;}
.stree-client-name{flex:1;font-size:13px;font-weight:700;color:#c4b5fd;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-client-actions{display:none;gap:1px;flex-shrink:0;}
.stree-client-row:hover .stree-client-actions{display:flex;}
.sec-icon{background:none;border:none;cursor:pointer;font-size:11px;padding:1px 2px;color:#9ca3af;opacity:.7;}
.sec-icon:hover{opacity:1;}
.sec-clone-btn{background:none;border:none;cursor:pointer;padding:1px 3px;color:#a78bfa;opacity:.7;}
.sec-clone-btn:hover{opacity:1;}

/* Month (mid level) */
.stree-month-row{display:flex;align-items:center;gap:6px;padding:6px 8px 6px 22px;cursor:pointer;transition:.12s;}
.stree-month-row:hover{background:rgba(255,255,255,.05);}
.stree-month-row.stree-open{background:rgba(99,102,241,.12);}
.stree-month-icon{font-size:11px;flex-shrink:0;}
.stree-month-name{flex:1;font-size:12.5px;font-weight:600;color:#a5b4fc;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-month-actions{display:none;gap:1px;}
.stree-month-row:hover .stree-month-actions{display:flex;}

/* Pages (leaf level) */
.stree-month-pages,.stree-client-body .stree-page{padding-left:0;}
.stree-page{display:flex;align-items:center;gap:7px;padding:5px 8px 5px 36px;text-decoration:none;color:#9ca3af;font-size:12px;transition:.1s;}
.stree-page:hover{background:rgba(255,255,255,.04);color:#e2e8f0;}
.stree-page-active{background:rgba(139,61,255,.18)!important;color:#c4b5fd!important;border-left:2px solid #8b3dff;}
.stree-dot{width:6px;height:6px;border-radius:50%;background:#6366f1;flex-shrink:0;}
.stree-page-active .stree-dot{background:#a78bfa;}
.stree-page-title{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stree-page-btns{display:none;gap:1px;flex-shrink:0;}
.stree-page:hover .stree-page-btns{display:flex;}
.stree-page .clone-page-btn,.stree-page .del-page-btn{background:none;border:none;cursor:pointer;font-size:10px;padding:1px 2px;color:#9ca3af;}
.stree-page .clone-page-btn:hover,.stree-page .del-page-btn:hover{color:#f87171;}

/* Colored dots */
.stree-page:nth-child(7n+1) .stree-dot{background:#f59e0b;}
.stree-page:nth-child(7n+2) .stree-dot{background:#3b82f6;}
.stree-page:nth-child(7n+3) .stree-dot{background:#ec4899;}
.stree-page:nth-child(7n+4) .stree-dot{background:#10b981;}
.stree-page:nth-child(7n+5) .stree-dot{background:#8b5cf6;}
.stree-page:nth-child(7n+6) .stree-dot{background:#ef4444;}
.stree-page:nth-child(7n+7) .stree-dot{background:#06b6d4;}

/* Add buttons */
.stree-add-month{display:block;width:calc(100% - 28px);margin:4px 14px;padding:5px 8px;background:none;border:1px dashed rgba(99,102,241,.35);color:#6d5dab;font-size:11px;font-weight:600;border-radius:5px;cursor:pointer;text-align:left;}
.stree-add-month:hover{border-color:#8b5cf6;color:#a78bfa;}
.stree-add-page{display:block;padding:4px 8px 4px 42px;color:#6d5dab;font-size:11px;text-decoration:none;}
.stree-add-page:hover{color:#a78bfa;}
.add-section-btn{margin:6px 10px 10px;padding:8px;background:none;border:1px dashed rgba(139,61,255,.35);color:#6d5dab;font-size:12px;font-weight:600;border-radius:6px;cursor:pointer;transition:.15s;}
.add-section-btn:hover{border-color:#8b3dff;color:#c4b5fd;background:rgba(139,61,255,.08);}
</style>

<style>
/* FORCE SI header colors - no green no red */
html body div.si-card div.si-card-head,
html body .si-card .si-card-head {
  background: linear-gradient(135deg,#1e3a8a 0%,#2563eb 60%,#3b82f6 100%) !important;
}
html body .si-card[data-sis="completed"] .si-card-head {
  background: linear-gradient(135deg,#134e4a 0%,#0d9488 60%,#2dd4bf 100%) !important;
}
html body .si-card[data-sis="inprogress"] .si-card-head {
  background: linear-gradient(135deg,#78350f 0%,#c2410c 60%,#fb923c 100%) !important;
}
html body .si-card[data-sis="hold"] .si-card-head {
  background: linear-gradient(135deg,#1e1b4b 0%,#4338ca 60%,#818cf8 100%) !important;
}
</style>

<style>
/* SI header: stylesheet !important beats inline style saved in DB */
.si-card-head{background:linear-gradient(135deg,#1e3a8a 0%,#2563eb 60%,#3b82f6 100%)!important;}
.si-card[data-si-status="completed"] .si-card-head{background:linear-gradient(135deg,#134e4a 0%,#0d9488 60%,#2dd4bf 100%)!important;}
.si-card[data-si-status="inprogress"] .si-card-head{background:linear-gradient(135deg,#78350f 0%,#c2410c 60%,#fb923c 100%)!important;}
.si-card[data-si-status="hold"] .si-card-head{background:linear-gradient(135deg,#1e1b4b 0%,#4338ca 60%,#818cf8 100%)!important;}
.si-card[data-si-status="pending"] .si-card-head{background:linear-gradient(135deg,#1e3a8a 0%,#2563eb 60%,#3b82f6 100%)!important;}
</style>

<style>
/* SI card soft colors */
</style>

<style>
/* SI: bold dates + status */
.date-val{font-weight:800!important;color:#0f172a!important;}
.si-status-label{font-size:11px!important;font-weight:800!important;color:#0f172a!important;letter-spacing:.5px!important;}
.ob-status-btn{font-weight:800!important;font-size:13px!important;padding:5px 16px!important;}
</style>

<style>
/* SI date bold */
.date-val,
.si-tl-date,
.date-cell-wrap .date-val,
.si-card .date-val {
  font-weight: 800 !important;
  color: #0f172a !important;
}
</style>

<style>
/* ===== PI sub-row alignment fix ===== */
/* Remove first-child special treatment that caused misalignment */
.pi-sub-list>.pi-sub-row:first-child{align-items:center!important;padding-bottom:5px!important;}
.pi-sub-list>.pi-sub-row:first-child .pi-sub-icon{margin-top:0!important;}
.pi-sub-list>.pi-sub-row:first-child .pi-sub-text{padding-top:0!important;}
.pi-sub-list>.pi-sub-row:first-child .pi-del-sub{margin-top:0!important;}

/* All sub-rows: consistent center alignment */
.pi-sub-row{align-items:center!important;}

/* Show TEST/PRODUCTION labels on ALL rows as small inline labels */
.pi-sub-row .pi-status-label{
  display:block!important;
  font-size:8px!important;font-weight:800!important;
  color:#94a3b8!important;letter-spacing:.8px!important;
  text-transform:uppercase!important;margin-bottom:2px!important;
}

/* Hide labels on 2nd+ rows to avoid repetition — only show on first */
.pi-sub-row:not(:first-child) .pi-status-label{
  display:none!important;
  height:0!important;overflow:hidden!important;
}
</style>

<style>
/* ===== SI card header colour by status — header only ===== */
.si-card:has(.ob-status-btn.completed) .si-card-head{
  background:linear-gradient(135deg,#134e4a 0%,#0d9488 60%,#2dd4bf 100%)!important;
}
.si-card:has(.ob-status-btn.inprogress) .si-card-head{
  background:linear-gradient(135deg,#7c2d12 0%,#ea580c 60%,#fb923c 100%)!important;
}
.si-card:has(.ob-status-btn.hold) .si-card-head{
  background:linear-gradient(135deg,#312e81 0%,#4f46e5 60%,#818cf8 100%)!important;
}
</style>

<style>
</style>

<style>
/* ===== Clone icon SVG styling ===== */
.clone-page-btn svg, .sec-clone-btn svg{
  display:inline-block;vertical-align:middle;
  width:13px;height:13px;stroke:#8b3dff;
}
.clone-page-btn:hover svg, .sec-clone-btn:hover svg{stroke:#6d28d9;}
.clone-page-btn, .sec-clone-btn{line-height:0;}
</style>

<style>
/* ===== Uniform banner title across all templates ===== */
.agenda-title,
.exec-title,
.si-banner-title,
.pi-banner-title,
.rr-banner-title,
.ob-header-title,
.fa-banner-title {
  font-size: 20px !important;
  font-weight: 900 !important;
  color: white !important;
  text-transform: uppercase !important;
  letter-spacing: 1px !important;
  line-height: 1.1 !important;
}
</style>

<style>
/* ===== Section clone button ===== */
.sec-clone-btn{
  background:none;border:none;cursor:pointer;
  font-size:13px;padding:2px 4px;border-radius:4px;
  color:#8b3dff;line-height:1;transition:.15s;
}
.sec-clone-btn:hover{background:#ede9fe;}
</style>

<style>
/* ===== Clone button force visible ===== */
.clone-page-btn{
  display:inline-flex!important;visibility:visible!important;
  opacity:1!important;color:#8b3dff!important;
  background:none!important;border:none!important;
  cursor:pointer!important;font-size:14px!important;
  padding:1px 4px!important;border-radius:4px!important;
  flex-shrink:0!important;line-height:1!important;
}
.clone-page-btn:hover{background:#ede9fe!important;}
</style>

<style>
/* ===== Clone button always visible ===== */
.clone-page-btn{
  opacity:0.45 !important;
  color:#8b3dff !important;
  background:none;border:none;cursor:pointer;
  font-size:14px;padding:2px 5px;border-radius:4px;
  flex-shrink:0;line-height:1;transition:.15s;
}
.clone-page-btn:hover{opacity:1 !important;background:#ede9fe !important;}
.del-page-btn{opacity:0.45;transition:.15s;}
.del-page-btn:hover{opacity:1;}
</style>

<style>
/* ===== Clone page button ===== */
.clone-page-btn{
  background:none;border:none;cursor:pointer;
  font-size:13px;padding:2px 4px;border-radius:4px;
  color:#94a3b8;opacity:0;transition:.15s;flex-shrink:0;line-height:1;
}
.page-link:hover .clone-page-btn{opacity:1;}
.clone-page-btn:hover{color:#8b3dff;background:#ede9fe;}
</style>

<style>
/* ===== Remove duplicate ::before dot from impact badge ===== */
.fa-impact-badge::before{content:none!important;display:none!important;width:0!important;height:0!important;}
</style>

<style>
/* ===== Focus Areas: impact badge with explicit dot colours ===== */
.fa-impact-badge{
  display:inline-flex !important;align-items:center !important;gap:6px;
  padding:4px 12px 4px 8px !important;border-radius:20px;
  cursor:pointer;user-select:none;
  font-size:11px;font-weight:800;letter-spacing:.2px;
  transition:opacity .15s;border:1.5px solid transparent;
}
.fa-impact-badge:hover{opacity:.75;}
/* dot via span.fa-impact-icon */
.fa-impact-icon{
  display:inline-block !important;
  width:8px !important;height:8px !important;
  border-radius:50% !important;
  flex-shrink:0 !important;
}
.fa-impact-badge[data-level="High"]  {color:#b91c1c;background:#fef2f2;border-color:#fca5a5;}
.fa-impact-badge[data-level="Medium"]{color:#92400e;background:#fffbeb;border-color:#fcd34d;}
.fa-impact-badge[data-level="Low"]   {color:#065f46;background:#f0fdf4;border-color:#6ee7b7;}
.fa-impact-badge[data-level="High"]   .fa-impact-icon{background:#ef4444;}
.fa-impact-badge[data-level="Medium"] .fa-impact-icon{background:#f59e0b;}
.fa-impact-badge[data-level="Low"]    .fa-impact-icon{background:#10b981;}
/* hide legacy elements */
.fa-impact-dot,.fa-impact-text{display:none!important;}
</style>

<style>
/* ===== Focus Areas: impact toggle badge ===== */
.fa-impact-row{display:flex;align-items:center;gap:0;margin-bottom:10px;margin-top:6px;}
.fa-impact-badge{
  display:inline-flex;align-items:center;gap:5px;
  padding:3px 10px 3px 7px;border-radius:20px;
  cursor:pointer;user-select:none;
  font-size:11px;font-weight:800;
  transition:background .15s,color .15s;
  border:1.5px solid currentColor;
}
.fa-impact-badge::before{content:'';width:7px;height:7px;border-radius:50%;background:currentColor;flex-shrink:0;}
.fa-impact-badge[data-level="High"]  {color:#ef4444;background:#fef2f2;}
.fa-impact-badge[data-level="Medium"]{color:#f59e0b;background:#fffbeb;}
.fa-impact-badge[data-level="Low"]   {color:#10b981;background:#f0fdf4;}
.fa-impact-badge:hover{opacity:.8;}
.fa-impact-dot,.fa-impact-text{display:none!important;}
</style>

<style>
/* ===== Focus Areas v2: banner header + direct card colours ===== */
.fa-banner{background:linear-gradient(120deg,#0d152b 0%,#1a3464 55%,#0a2550 100%);padding:14px 22px;border-radius:10px 10px 0 0;margin-bottom:0;}
.fa-banner-title{color:white;font-size:13px;font-weight:900;letter-spacing:2px;text-transform:uppercase;}
.fa-wrap{border-radius:10px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1);}
.fa-grid{padding:14px;background:#f8fafc;gap:14px;}
.fa-add-btn{margin:0 14px 14px;width:calc(100% - 28px);background:#f8fafc;}
.fa-main-title{display:none;}
</style>

<style>
/* ===== Current Focus Areas & Challenges Template ===== */
.fa-wrap{font-family:'Segoe UI',system-ui,sans-serif;margin:8px 0 20px;}
.fa-main-title{font-size:22px;font-weight:900;color:#0f172a;margin-bottom:18px;padding:2px 4px;outline:none;cursor:text;line-height:1.3;}
.fa-main-title:empty::before{content:attr(data-ph);color:#94a3b8;font-weight:400;}
.fa-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;}
/* FA_DRAG_REORDER_V1 */
.fa-drag-handle{
  position:absolute;top:2px;left:50%;transform:translateX(-50%);
  cursor:grab;color:rgba(255,255,255,.85);font-size:18px;
  user-select:none;padding:1px 12px 5px;border-radius:0 0 7px 7px;
  letter-spacing:3px;line-height:1;z-index:20;
  background:rgba(255,255,255,.16);
}
.fa-drag-handle:hover{color:#fff;background:rgba(255,255,255,.28);}
.fa-drag-handle:active{cursor:grabbing;}
.fa-card.fa-dragging{opacity:.35;}
.fa-card{background:white;border-radius:10px;border:1.5px solid #e2e8f0;position:relative;overflow:visible;}
/* Chevron header */
.fa-card-head{
  position:relative;
  background:var(--fa-color,#1e40af);
  clip-path:polygon(0 0,calc(100% - 20px) 0,100% 50%,calc(100% - 20px) 100%,0 100%);
  padding:13px 44px 13px 16px;
  border-radius:8px 8px 0 0;
  min-height:48px;
  display:flex;align-items:center;
}
.fa-card-title{color:white;font-size:14px;font-weight:800;outline:none;cursor:text;flex:1;line-height:1.3;}
.fa-card-title:empty::before{content:'Title';color:rgba(255,255,255,.45);}
/* Colour cycle button — sits at chevron tip */
.fa-color-btn{
  position:absolute;top:50%;right:6px;transform:translateY(-50%);
  width:16px;height:16px;border-radius:50%;
  background:rgba(255,255,255,.35);border:2px solid rgba(255,255,255,.7);
  cursor:pointer;z-index:5;transition:.15s;flex-shrink:0;
}
.fa-color-btn:hover{background:rgba(255,255,255,.65);}
/* Card body */
.fa-card-inner{padding:10px 14px 14px;}
.fa-impact-row{display:flex;align-items:center;gap:6px;margin-bottom:10px;margin-top:4px;}
.fa-impact-dot{width:8px;height:8px;border-radius:50%;background:#ef4444;flex-shrink:0;}
.fa-impact-text{font-size:12px;font-weight:700;color:#ef4444;outline:none;cursor:text;}
.fa-impact-text:empty::before{content:'High Impact';color:#fca5a5;}
.fa-section{margin-bottom:8px;}
.fa-section-label{font-size:11px;font-weight:800;color:#1e293b;letter-spacing:.2px;margin-bottom:3px;text-transform:none;}
.fa-section-body{font-size:12px;color:#374151;line-height:1.55;outline:none;min-height:16px;cursor:text;}
.fa-section-body:empty::before{content:attr(data-ph);color:#cbd5e1;}
/* Delete button */
.fa-del-card{
  position:absolute;top:-9px;right:-9px;
  width:22px;height:22px;background:#ef4444;color:white;
  border:none;border-radius:50%;cursor:pointer;
  font-size:13px;font-weight:700;
  display:none;align-items:center;justify-content:center;
  z-index:10;line-height:1;padding:0;
}
.fa-card:hover .fa-del-card{display:flex;}
/* Add button */
.fa-add-btn{
  display:block;width:100%;margin-top:14px;
  background:none;border:2px dashed #cbd5e1;color:#94a3b8;
  font-size:12px;font-weight:700;padding:10px;
  border-radius:8px;cursor:pointer;box-sizing:border-box;transition:.15s;
}
.fa-add-btn:hover{border-color:#3b82f6;color:#2563eb;background:#eff6ff;}
</style>

<style>
/* ===== SI card drag-and-drop ===== */
/* SI_DRAG_UX_V1 */
/* SI_DRAG_HITAREA_V1: larger fixed hit-box, whole badge clickable */
.si-drag-handle{
  position:absolute;top:2px;left:50%;transform:translateX(-50%);
  box-sizing:border-box;width:44px;height:22px;
  display:flex;align-items:center;justify-content:center;
  cursor:grab;color:rgba(255,255,255,.85);font-size:20px;
  user-select:none;border-radius:0 0 8px 8px;
  letter-spacing:3px;line-height:1;z-index:50;pointer-events:auto;
  background:rgba(255,255,255,.14);
}
.si-drag-handle:hover{color:#fff;background:rgba(255,255,255,.26);}
.si-drag-handle:active{cursor:grabbing;}

/* Card being dragged */
.si-card.si-dragging{opacity:.35;box-shadow:0 0 0 2px #3b82f6;}

/* Drop insertion line */
.si-drop-line{
  height:3px;background:#3b82f6;border-radius:2px;
  margin:3px 0;display:none;box-shadow:0 0 8px rgba(59,130,246,.5);
}
.si-drop-line.visible{display:block;}
</style>

<style>
/* ===== SI: compact list indentation in text areas ===== */
.si-tl-body ol, .si-tl-body ul,
.si-tl-body [contenteditable] ol,
.si-tl-body [contenteditable] ul {
  padding-left: 20px !important;
  margin: 2px 0 !important;
}
.si-tl-body li { padding-left: 2px !important; margin: 1px 0 !important; }
</style>

<style>
/* ===== SI: keep action bullets/counters inside the card box ===== */
.si-action-list{overflow:hidden;padding-left:0!important;margin-left:0!important;list-style:none!important;}
.si-action-row{position:relative;display:flex;align-items:flex-start;gap:8px;padding:6px 8px;border-radius:4px;margin:0!important;}
.si-action-row::before{content:none!important;}
.si-action-text{flex:1;min-height:18px;outline:none;word-break:break-word;}
.si-action-list ol,.si-action-list ul{padding-left:20px!important;margin:0!important;overflow:hidden!important;}
.si-tl-content [contenteditable] ul,.si-tl-content [contenteditable] ol{padding-left:20px!important;margin:0!important;}
</style>

<style>
/* ===== PI: hide status on main item, keep only on sub-rows ===== */
.pi-item-row > .pi-status-cols { display: none !important; }
.pi-item > .pi-status-cols     { display: none !important; }
.pi-sub-row   .pi-status-cols  { display: flex !important; }
</style>

<style>
/* ===== PI Status: TEST + PRODUCTION on one line ===== */
.pi-status-cols{flex-direction:row!important;display:flex!important;align-items:flex-start;gap:12px;}
.pi-status-col{display:flex!important;flex-direction:column;align-items:flex-start;gap:3px;min-width:90px;}
.pi-item>.pi-status-cols{padding:4px 14px 12px 70px;flex-wrap:wrap;}
.pi-item-row>.pi-status-cols{flex-wrap:nowrap;align-self:center;gap:10px;}
.pi-sub-row .pi-status-cols{flex-wrap:nowrap;gap:8px;}
.pi-sub-row .pi-status-col{min-width:88px;}
</style>

<style>
/* ===== PI Template — Modern Fluent UI Style ===== */

/* Wrapper */
.pi-wrap{border-radius:14px;overflow:hidden;box-shadow:0 2px 16px rgba(0,0,0,.1);background:#f1f5f9;font-family:'Segoe UI',sans-serif;}
.pi-wrap::before{display:none!important;}

/* Banner */
.pi-banner{background:linear-gradient(120deg,#0d152b 0%,#1a3464 55%,#0a2550 100%);padding:16px 22px 14px;position:relative;overflow:hidden;border-bottom:2px solid rgba(16,185,129,.4);}
.pi-banner::before{content:'';position:absolute;right:-30px;top:-30px;width:130px;height:130px;border-radius:50%;background:rgba(255,255,255,.04);pointer-events:none;}
.pi-banner-icon{display:none!important;}
.pi-banner-title{color:white;font-size:13px;font-weight:900;letter-spacing:2.5px;text-transform:uppercase;position:relative;z-index:1;}
.pi-chevron{display:none!important;}

/* Body — sets up CSS counter for auto-numbering */
.pi-body{padding:12px;background:#f1f5f9;counter-reset:pi-item-num;}

/* Item card */
.pi-item{background:white;border-radius:12px;box-shadow:0 1px 6px rgba(0,0,0,.07),0 4px 16px rgba(0,0,0,.04);margin-bottom:10px;border:1px solid #eaeff5;border-left:none!important;position:relative;counter-increment:pi-item-num;}

/* Item header row */
.pi-item-row{display:flex;align-items:center;gap:14px;padding:14px 48px 12px 14px;}

/* Main icon → dark rounded square with auto-number */
.pi-main-icon,.pi-main-icon.pi-ghost{width:42px!important;height:42px!important;border-radius:10px!important;background:#0d152b!important;flex-shrink:0!important;display:flex!important;align-items:center!important;justify-content:center!important;font-size:0!important;box-shadow:0 2px 8px rgba(13,21,43,.25)!important;}
.pi-main-icon::after,.pi-main-icon.pi-ghost::after{content:counter(pi-item-num,decimal-leading-zero);font-size:13px;font-weight:900;color:white;font-family:'Segoe UI',sans-serif;letter-spacing:.5px;}

/* Item text */
.pi-item-text{flex:1;font-size:15px;font-weight:700;color:#0f172a;line-height:1.4;min-height:24px;}

/* Status cols INSIDE header row (standalone items — no sub-points) */
.pi-item-row>.pi-status-cols{padding:0;flex-shrink:0;align-self:center;flex-direction:row;gap:8px;}
/* Status cols as direct child of item (between header and sub-list) */
.pi-item>.pi-status-cols{padding:2px 14px 10px 70px;}

/* Status base */
.pi-status-cols{display:flex;gap:8px;}
.pi-status-col{display:flex;flex-direction:column;align-items:flex-start;gap:2px;min-width:100px;}
.pi-status-label{font-size:9px;font-weight:800;color:#94a3b8;letter-spacing:1px;text-transform:uppercase;}

/* Pills */
.pi-pill{font-size:11px;font-weight:700;padding:4px 13px;border-radius:20px;cursor:pointer;white-space:nowrap;user-select:none;transition:opacity .1s,transform .08s;}
.pi-pill:active{transform:scale(.96);}
.pi-pill-completed{background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;}
.pi-pill-inprogress{background:#dbeafe;color:#1e40af;border:1px solid #93c5fd;}
.pi-pill-pending{background:#fef3c7;color:#92400e;border:1px solid #fcd34d;}
.pi-pill-hold{background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;}

/* Sub-list */
.pi-sub-list{padding:0 12px 6px 70px;}

/* ── Column-header trick: first sub-row shows TEST/PROD labels once ── */
/* Align first row to top so labels sit above the pills as headers */
.pi-sub-list>.pi-sub-row:first-child{align-items:flex-start;padding-bottom:6px;}
/* Push icon, text, del button down to align with pill (not label) */
.pi-sub-list>.pi-sub-row:first-child .pi-sub-icon{margin-top:16px;}
.pi-sub-list>.pi-sub-row:first-child .pi-sub-text{padding-top:15px;}
.pi-sub-list>.pi-sub-row:first-child .pi-del-sub{margin-top:15px;}
/* Hide labels on 2nd, 3rd, ... sub-rows — only pills show */
.pi-sub-list>.pi-sub-row:not(:first-child) .pi-status-label{display:none;}

/* Sub-row */
.pi-sub-row{display:flex;align-items:center;gap:10px;padding:5px 6px;border-radius:7px;background:transparent;border:none!important;margin-bottom:2px;transition:background .1s;}
.pi-sub-row:hover{background:#f0f5fb;}

/* Sub-row status */
.pi-sub-row .pi-status-cols{padding:0;flex-shrink:0;gap:8px;flex-direction:row;}
.pi-sub-row .pi-status-col{min-width:100px;}

/* Sub icon → small blue-tinted circle */
.pi-sub-icon{width:22px;height:22px;border-radius:50%;background:#eff6ff;border:2px solid #93c5fd;flex-shrink:0;}

/* Sub text */
.pi-sub-text{flex:1;font-size:13px;color:#374151;min-height:18px;font-weight:500;line-height:1.35;}

/* Delete sub button */
.pi-del-sub{background:none;border:none;color:#e2e8f0;cursor:pointer;font-size:13px;padding:2px 5px;border-radius:4px;flex-shrink:0;line-height:1;transition:.1s;}
.pi-del-sub:hover{color:#ef4444;background:#fef2f2;}

/* Add sub-point button */
.pi-add-sub-wrap{padding:4px 12px 12px 70px;}
.pi-add-sub{background:none;border:1.5px dashed #34d399;color:#059669;font-size:11px;font-weight:700;padding:5px 16px;border-radius:20px;cursor:pointer;letter-spacing:.2px;transition:.1s;}
.pi-add-sub:hover{background:#ecfdf5;border-color:#10b981;}

/* Delete item button */
.pi-del-item{position:absolute;top:12px;right:12px;background:none;border:none;color:#e2e8f0;cursor:pointer;font-size:14px;padding:2px 6px;border-radius:4px;line-height:1;z-index:1;transition:.1s;}
.pi-del-item:hover{color:#ef4444;background:#fef2f2;}

/* Add point button */
.pi-add-item{display:block;width:100%;background:none;border:2px dashed #cbd5e1;color:#94a3b8;font-size:12px;font-weight:700;padding:10px;border-radius:8px;cursor:pointer;text-align:center;margin-top:2px;box-sizing:border-box;transition:.15s;letter-spacing:.2px;}
.pi-add-item:hover{border-color:#10b981;color:#059669;background:#f0fdf4;}
</style>

<style>
/* ===== PI Template Redesign ===== */
.pi-wrap{border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(13,21,43,.13);background:#f1f5f9;font-family:'Segoe UI',sans-serif;}
.pi-wrap::before{display:none!important;}

/* Banner */
.pi-banner{background:linear-gradient(120deg,#0d152b 0%,#1a3464 55%,#0a2550 100%);padding:18px 24px 16px;position:relative;overflow:hidden;border-bottom:3px solid rgba(16,185,129,.45);}
.pi-banner::before{content:'';position:absolute;right:-30px;top:-30px;width:140px;height:140px;border-radius:50%;background:rgba(255,255,255,.04);pointer-events:none;}
.pi-banner::after{content:'';position:absolute;left:-20px;bottom:-30px;width:100px;height:100px;border-radius:50%;background:rgba(16,185,129,.06);pointer-events:none;}
.pi-banner-icon{display:none!important;}
.pi-banner-title{color:white;font-size:15px;font-weight:900;letter-spacing:2px;text-transform:uppercase;position:relative;z-index:1;}
.pi-chevron{display:none!important;}

/* Body: set up CSS counter for auto-numbering */
.pi-body{padding:14px;background:#f1f5f9;counter-reset:pi-item-num;}

/* Item cards */
.pi-item{background:white;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.07);margin-bottom:10px;border:1px solid #e2eaf2;border-left:4px solid #10b981;position:relative;counter-increment:pi-item-num;}

/* Item header row */
.pi-item-row{display:flex;align-items:flex-start;gap:12px;padding:14px 44px 10px 14px;}

/* Transform the main icon dot into an auto-numbered green badge */
.pi-main-icon,.pi-main-icon.pi-ghost{
  width:28px!important;height:28px!important;border-radius:50%!important;
  background:linear-gradient(135deg,#10b981,#059669)!important;
  flex-shrink:0!important;margin-top:2px!important;
  display:flex!important;align-items:center!important;justify-content:center!important;
  font-size:0!important;box-shadow:0 2px 6px rgba(16,185,129,.3)!important;
}
.pi-main-icon::after,.pi-main-icon.pi-ghost::after{
  content:counter(pi-item-num,decimal-leading-zero);
  font-size:10px;font-weight:900;color:white;font-family:'Segoe UI',sans-serif;
}

/* Item text */
.pi-item-text{flex:1;font-size:14px;font-weight:600;color:#1e293b;line-height:1.55;min-height:28px;}

/* Status cols: inside item-row (standalone items — no sub-points) */
.pi-item-row>.pi-status-cols{padding:0;flex-shrink:0;align-self:center;flex-direction:column;gap:5px;}
/* Status cols: direct child of item (between item-row and sub-list) */
.pi-item>.pi-status-cols{padding:0 14px 10px 54px;}
/* Base */
.pi-status-cols{display:flex;gap:10px;}
.pi-status-col{display:flex;flex-direction:column;align-items:flex-start;gap:3px;min-width:88px;}
.pi-status-label{font-size:9px;font-weight:800;color:#94a3b8;letter-spacing:1px;text-transform:uppercase;}

/* Status pills */
.pi-pill{font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;cursor:pointer;white-space:nowrap;user-select:none;transition:opacity .1s,transform .08s;}
.pi-pill:active{transform:scale(.96);}
.pi-pill-completed{background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;}
.pi-pill-inprogress{background:#dbeafe;color:#1e40af;border:1px solid #93c5fd;}
.pi-pill-pending{background:#fef3c7;color:#92400e;border:1px solid #fcd34d;}
.pi-pill-hold{background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;}

/* Sub-points list */
.pi-sub-list{padding:0 14px 4px 54px;}
.pi-sub-row{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:7px;background:#f8fafc;margin-bottom:5px;border:1px solid #e8edf5;transition:background .1s;}
.pi-sub-row:hover{background:#f0f5fc;}
.pi-sub-row .pi-status-cols{padding:0;flex-shrink:0;gap:8px;flex-direction:row;}
.pi-sub-row .pi-status-col{min-width:80px;}
.pi-sub-icon{width:6px;height:6px;border-radius:50%;background:#94a3b8;flex-shrink:0;}
.pi-sub-text{flex:1;font-size:13px;color:#374151;min-height:18px;line-height:1.4;}

/* Sub-point delete button */
.pi-del-sub{background:none;border:none;color:#e2e8f0;cursor:pointer;font-size:13px;padding:2px 5px;border-radius:4px;flex-shrink:0;line-height:1;transition:.1s;}
.pi-del-sub:hover{color:#ef4444;background:#fef2f2;}

/* Add sub-point button */
.pi-add-sub-wrap{padding:2px 14px 12px 54px;}
.pi-add-sub{background:none;border:1.5px dashed #34d399;color:#059669;font-size:11px;font-weight:700;padding:5px 16px;border-radius:20px;cursor:pointer;transition:.1s;letter-spacing:.2px;}
.pi-add-sub:hover{background:#ecfdf5;border-color:#10b981;}

/* Item delete button */
.pi-del-item{position:absolute;top:10px;right:10px;background:none;border:none;color:#e2e8f0;cursor:pointer;font-size:14px;padding:2px 6px;border-radius:4px;line-height:1;z-index:1;transition:.1s;}
.pi-del-item:hover{color:#ef4444;background:#fef2f2;}

/* Add point button */
.pi-add-item{display:block;width:100%;background:none;border:2px dashed #cbd5e1;color:#94a3b8;font-size:12px;font-weight:700;padding:10px;border-radius:8px;cursor:pointer;text-align:center;margin-top:2px;box-sizing:border-box;transition:.15s;letter-spacing:.2px;}
.pi-add-item:hover{border-color:#10b981;color:#059669;background:#f0fdf4;}
</style>



</head><body>

<div class="topbar">
  <div style="display:flex;align-items:center;padding-left:4px;">
    <img src="{{ LOGO_SRC }}?v=2" alt="VariTec Consulting" style="height:40px;width:auto;object-fit:contain;display:block;">
  </div>

  

  <div class="topbar-right">

    <div class="profile-wrap">

      <button class="profile-btn" onclick="toggleProfileMenu(event)" type="button">

        <div class="profile-avatar">{{ email[0] }}</div>

        <span class="profile-name-text">{{ email.split('@')[0] }}</span>

        <span class="profile-caret">▾</span>

      </button>

      <div class="profile-menu" id="profileMenu">

        <div class="profile-menu-header">

          <div class="profile-avatar-lg">{{ email[0] }}</div>

          <div>

            <div class="profile-menu-username">{{ email.split('@')[0] }}</div>

            <div class="profile-menu-email">{{ email }}</div>

            <span class="profile-role-pill">{{ role.upper() if role else 'USER' }}</span>

          </div>

        </div>

        {% if role == 'admin' %}

        <a href="/admin/users" class="profile-menu-item">

          <span class="profile-menu-item-icon">👥</span> Manage Users

        </a>

        {% endif %}

        <a href="/logout" class="profile-menu-item danger">

          <span class="profile-menu-item-icon">⏻</span> Logout

        </a>

      </div>

    </div>

  </div>

</div>

<div class="main">

  <div class="sidebar">
    <div class="sidebar-header">Sections</div>
    <div class="sections-list" id="sectionsList">

      {% for sec in top_sections %}
      {% set sub_secs = sub_map.get(sec.id, []) %}
      {% set sec_notes = notes_map.get(sec.name, []) %}
      {% set sec_active = (sec.name == active_section) %}
      {% set any_child_active = (active_section in sub_map.get(sec.id, [])|map(attribute='name')|list) %}

      <div class="stree-client" data-id="{{ sec.id }}">

        <!-- CLIENT HEADER -->
        <div class="stree-client-row {% if sec_active or any_child_active %}stree-open{% endif %}"
             onclick="streeClient({{ sec.id }})">
          <span class="stree-arr-icon">{% if sec_active or any_child_active or sub_secs %}<span class="stree-arr">▼</span>{% else %}<span class="stree-arr">▶</span>{% endif %}</span><span class="stree-folder"><svg width="16" height="14" viewBox="0 0 16 14" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M1 2.5C1 1.67 1.67 1 2.5 1H6L7.5 3H13.5C14.33 3 15 3.67 15 4.5V11.5C15 12.33 14.33 13 13.5 13H2.5C1.67 13 1 12.33 1 11.5V2.5Z" fill="#2d3a6b" stroke="#2d3a6b" stroke-width="0.5"/></svg></span>
          <span class="stree-client-name">{{ sec.name }}</span>
          <span class="stree-client-actions" onclick="event.stopPropagation()">
            <button class="sec-clone-btn" onclick="cloneSection({{ sec.id }},'{{ sec.name }}')" title="Clone">
              <svg width="12" height="12" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="8" height="8" rx="1.5"/><path d="M1 9V2.5A1.5 1.5 0 0 1 2.5 1H9"/></svg>
            </button>
            <button class="sec-icon" onclick="openRename({{ sec.id }},'{{ sec.name }}')">&#9999;&#65039;</button>
            <button class="sec-icon" onclick="deleteSection({{ sec.id }},'{{ sec.name }}')">&#128465;</button>
          </span>
        </div>

        <!-- CLIENT BODY (sub-sections + direct pages) -->
        <div class="stree-client-body" id="cb-{{ sec.id }}"
             style="{% if not sec_active and not any_child_active %}display:none{% endif %}">

          {% if sub_secs %}
            <!-- MONTH SUB-SECTIONS -->
            {% for sub in sub_secs %}
            {% set sub_notes = notes_map.get(sub.name, []) %}
            {% set sub_active = (sub.name == active_section) %}
            <div class="stree-month" data-id="{{ sub.id }}">
              <div class="stree-month-row {% if sub_active %}stree-open{% endif %}"
                   data-fullname="{{ sub.name | e }}"
                   onclick="streeMonth({{ sub.id }})">
                <span class="stree-arr stree-month-arr">▶</span><span class="stree-folder stree-folder-sm"><svg width="14" height="12" viewBox="0 0 16 14" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M1 2.5C1 1.67 1.67 1 2.5 1H6L7.5 3H13.5C14.33 3 15 3.67 15 4.5V11.5C15 12.33 14.33 13 13.5 13H2.5C1.67 13 1 12.33 1 11.5V2.5Z" fill="#3b4fa8" stroke="#3b4fa8" stroke-width="0.5"/></svg></span>
                <span class="stree-month-name">{{ sub.name.split(" > ", 1)[1] if " > " in sub.name else sub.name }}</span>
                <span class="stree-month-actions" onclick="event.stopPropagation()">
                  <button class="sec-clone-btn" onclick="cloneMonthSection({{ sub.id }},'{{ sub.name }}')" title="Clone month"><svg width="12" height="12" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="8" height="8" rx="1.5"/><path d="M1 9V2.5A1.5 1.5 0 0 1 2.5 1H9"/></svg></button>
                  <button class="sec-icon" onclick="openRename({{ sub.id }},'{{ sub.name }}')">&#9999;&#65039;</button>
                  <button class="sec-icon" onclick="deleteSection({{ sub.id }},'{{ sub.name }}')">&#128465;</button>
                </span>
              </div>
              <div class="stree-month-pages" id="mp-{{ sub.id }}"
                   style="{% if not sub_active %}display:none{% endif %}">
                {% for note in sub_notes %}
                <a class="stree-page {% if note.id == current.id %}stree-page-active{% endif %}"
                   href="/note/{{ note.id }}">
                  <span class="stree-dot"></span>
                  <span class="stree-page-title">{{ note.title or 'Untitled' }}</span>
                  <span class="stree-page-btns" onclick="event.stopPropagation()">
                    <button class="clone-page-btn" onclick="clonePage(event,{{ note.id }})" title="Clone">&#128203;</button>
                    <button class="del-page-btn" onclick="deletePage(event,{{ note.id }})">&#128465;</button>
                  </span>
                </a>
                {% endfor %}
              </div>
            </div>
            {% endfor %}

            <!-- ADD MONTH BUTTON -->
            {% if can_write %}
            <button class="stree-add-month" onclick="streeAddMonth({{ sec.id }},'{{ sec.name }}')">
              + Add Month
            </button>
            {% endif %}

          {% else %}
            <!-- NO SUB-SECTIONS: show direct pages -->
            {% for note in sec_notes %}
            <a class="stree-page {% if note.id == current.id %}stree-page-active{% endif %}"
               href="/note/{{ note.id }}">
              <span class="stree-dot"></span>
              <span class="stree-page-title">{{ note.title or 'Untitled' }}</span>
              <span class="stree-page-btns" onclick="event.stopPropagation()">
                <button class="clone-page-btn" onclick="clonePage(event,{{ note.id }})" title="Clone">&#128203;</button>
                <button class="del-page-btn" onclick="deletePage(event,{{ note.id }})">&#128465;</button>
              </span>
            </a>
            {% endfor %}
            {% if can_write %}
            <button class="stree-add-month" onclick="streeAddMonth({{ sec.id }},'{{ sec.name }}')">
              + Add Month
            </button>
            {% endif %}
          {% endif %}

        </div>
      </div>
      {% endfor %}

    </div>
    <button class="add-section-btn" onclick="openAddSection()">&#65291; Add Section</button>
  </div>
<div class="pages">

    <!-- CURRENT_SECTION_INDICATOR_V1 -->
    <style>
.current-section-indicator{padding:10px 14px;background:linear-gradient(135deg,#4c1d95 0%,#6d28d9 60%,#7c3aed 100%);color:white;border-bottom:1px solid #5b21b6;}
.csi-label{font-size:9px;font-weight:800;letter-spacing:1px;opacity:.75;margin-bottom:4px;text-transform:uppercase;}
.csi-client{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:800;}
.csi-client svg{flex-shrink:0;}
.csi-month{display:flex;align-items:center;gap:6px;font-size:11px;font-weight:600;opacity:.85;margin-top:3px;}
.csi-month svg{flex-shrink:0;}
    </style>
    {% set _csi_parts = (active_section or '').split(' > ') %}
    <div class="current-section-indicator">
      <div class="csi-label">Current Section</div>
      <div class="csi-client">
        <svg width="13" height="11" viewBox="0 0 16 14" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M1 3a1 1 0 0 1 1-1h4l1.5 2H14a1 1 0 0 1 1 1v7a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1V3z" fill="currentColor"/></svg>
        <span>{{ _csi_parts[0] }}</span>
      </div>
      {% if _csi_parts|length > 1 %}
      <div class="csi-month">
        <svg width="12" height="12" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="1.5" y="2.5" width="13" height="12" rx="1.5" stroke="currentColor" stroke-width="1.3"/><path d="M1.5 6h13" stroke="currentColor" stroke-width="1.3"/></svg>
        <span>{{ _csi_parts[1] }}</span>
      </div>
      {% endif %}
    </div>

    <div class="pages-top">

      {% if can_write %}

      <div class="split-wrap">

        <form action="/new/{{ active_section }}" method="post" style="flex:1;margin:0;padding:0;display:flex;">

          <button type="submit" class="add-page-btn" style="width:100%;border-radius:0;">&#65291; New Page</button>

        </form>

        <button type="button" class="split-arrow" id="scArrowBtn" onclick="toggleScDropdown(event)">&#9660;</button>

      </div>

      {% else %}

      <div style="padding:10px 14px;font-size:12px;color:#94a3b8;text-align:center;background:#f8fafc;border-bottom:1px solid #e2e8f0;">👁 Read-only access</div>

      {% endif %}

    <div class="sc-dropdown" id="scDropdown">

        <div class="sc-dd-section" style="background:linear-gradient(135deg,#fff7ed,#fef3c7);"><div class="sc-dd-section-dot" style="background:#d97706;"></div><span class="sc-dd-section-label" style="color:#d97706;">Presentation</span></div>

        <div class="sc-dd-divider"></div>

        <div class="sc-dd-item" onclick="createCoverSlide()"><span class="sc-dd-item-icon">&#127970;</span><span class="sc-dd-item-text">Cover</span></div>

        <div class="sc-dd-item" onclick="createAgenda()"><span class="sc-dd-item-icon">&#128203;</span><span class="sc-dd-item-text">Agenda</span></div>

        <div class="sc-dd-item" onclick="createExecutiveSummary()"><span class="sc-dd-item-icon">&#128196;</span><span class="sc-dd-item-text">Executive Summary</span></div>

        <div class="sc-dd-item" onclick="createThankYou()"><span class="sc-dd-item-icon">&#128591;</span><span class="sc-dd-item-text">Thank You</span></div>
        <div class="sc-dd-item" onclick="createRolesResponsibilities()"><span class="sc-dd-item-icon">&#128101;</span><span class="sc-dd-item-text">Roles &amp; Responsibilities</span></div>
        <div class="sc-dd-item" onclick="createProcessImprovements()"><span class="sc-dd-item-icon">&#128200;</span><span class="sc-dd-item-text">Process Improvements</span></div>
        <div class="sc-dd-item" onclick="createFocusAreas()"><span class="sc-dd-item-icon">&#127919;</span><span class="sc-dd-item-text">Focus Areas &amp; Challenges</span></div>

        <div class="sc-dd-section" style="background:linear-gradient(135deg,#f5f3ff,#f0ebff);margin-top:4px;"><div class="sc-dd-section-dot" style="background:#7c3aed;"></div><span class="sc-dd-section-label">Scorecard</span></div>

        <div class="sc-dd-divider"></div>

         <div class="sc-dd-item" onclick="createSerializationIssue()"><span class="sc-dd-item-icon">&#128203;</span><span class="sc-dd-item-text">Serialization Issues</span></div>

        <div class="sc-dd-section" style="background:linear-gradient(135deg,#f0f9ff,#e0f2fe);margin-top:4px;"><div class="sc-dd-section-dot" style="background:#0369a1;"></div><span class="sc-dd-section-label" style="color:#0369a1;">Onboarding</span></div>

        <div class="sc-dd-divider"></div>

        <div class="sc-dd-item" onclick="createOnboarding('customer')"><span class="sc-dd-item-icon">&#127970;</span><span class="sc-dd-item-text">Customer Onboarding</span></div>

        <div class="sc-dd-item" onclick="createOnboarding('vendor')"><span class="sc-dd-item-icon">&#128203;</span><span class="sc-dd-item-text">Vendor/CMO Onboarding</span></div>

        <div class="sc-dd-item" onclick="createOnboarding('mah')"><span class="sc-dd-item-icon">&#127970;</span><span class="sc-dd-item-text">MAH Onboarding</span></div>

        <div class="sc-dd-section" style="background:linear-gradient(135deg,#f0f9ff,#e0f2fe);margin-top:4px;"><div class="sc-dd-section-dot" style="background:#0369a1;"></div><span class="sc-dd-section-label" style="color:#0369a1;">Serialization Issues</span></div>

        <div class="sc-dd-divider"></div>

       
        <div class="sc-dd-item featured" onclick="createScorecard('all')"><span class="sc-dd-item-icon">&#128202;</span><span class="sc-dd-item-text">All Customer Scorecards</span></div>
      </div>

    </div>

    <div class="pages-list" id="pagesList">

      {% for note in notes %}

      <a class="page-link {% if note.id == current.id %}active{% endif %}"

         href="/note/{{ note.id }}"

         draggable="true"

         data-note-id="{{ note.id }}">

        <span class="drag-handle" title="Drag to reorder"></span>

        <span class="page-title">{{ note.title or 'Untitled' }}</span>

        <button class="clone-page-btn" onclick="clonePage(event,{{ note.id }})" title="Clone page"><svg width="13" height="13" viewBox="0 0 13 13" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="8" height="8" rx="1.5"/><path d="M1 9V2.5A1.5 1.5 0 0 1 2.5 1H9"/></svg></button><button class="del-page-btn" onclick="deletePage(event,{{ note.id }})">&#128465;</button>

      </a>

      {% endfor %}

    </div>

  </div>

  <div class="editor">

    <div class="toolbar">

      <div class="tb-group tb-fmt-grp">

        <button onclick="fmt('bold')"><b>B</b></button>

        <button onclick="fmt('italic')"><em>I</em></button>

        <button onclick="fmt('underline')"><u>U</u></button>

        <button onclick="fmt('strikeThrough')"><s>S</s></button>

      </div>

      <div class="tb-group tb-list-grp">

        <button onclick="fmt('insertUnorderedList')">&#8226; List</button>

        <button onclick="fmt('insertOrderedList')">1. List</button>

      </div>

      <div class="tb-group tb-align-grp">

        <button onclick="fmt('justifyLeft')">&#8676;</button>

        <button onclick="fmt('justifyCenter')">&#9632;</button>

        <button onclick="fmt('justifyRight')">&#8677;</button>

      </div>

      <div class="tb-group tb-media-grp">

        <button onclick="insertImageFile()">&#128247; Image</button>

        <button onclick="fmt('removeFormat')" style="color:#dc2626;background:#fff5f5;">&#10006; Clear</button>

      </div>

      <div class="tb-group tb-insert-grp">

        <button onclick="openTableModal()">&#9776; Table</button>

        <button onclick="insertBox()">&#9634; Box</button>

      </div>

      <button class="tb-chart" onclick="generateScorecardChartInline()">&#128202; Chart</button>

      <button id="addTextOverlayBtn" onclick="addOverlayTextBox()" style="display:none;padding:6px 14px;border:none;border-radius:8px;background:linear-gradient(135deg,#7c3aed,#6d28d9);color:white;cursor:pointer;font-size:12px;font-weight:700;box-shadow:0 2px 6px rgba(124,58,237,.3);transition:.15s;letter-spacing:.2px;margin-left:4px;">&#9707; Add Text</button>
      <button class="tb-btn-ppt" onclick="exportAsPDF()">&#128196; Export PDF</button>

    </div>

    <div class="note-title-bar">

      <input id="title" placeholder="Page title…" value="{{ current.title | e }}">

      

    </div>

    <div id="content" contenteditable="true" data-placeholder="Start typing…">{{ current.content | safe }}</div>
<script>document.querySelectorAll('.ty-wrap,.cover-wrap,.si-wrap,.ob-wrap,.sc-table,.exec-wrap,.agenda-wrap,.proc-wrap,.pcard-wrap,.si-banner,.rr-wrap,.pi-wrap,.fa-wrap').length&&(document.getElementById('title').style.display='none');</script>

    <div class="paste-hint">&#128247; Paste Ctrl+V &middot; Drag &amp; drop &middot; Ctrl+S to save</div>

    <div class="statusbar"><span id="status">Ready</span></div>

    <div class="note-footer">

      <a class="note-footer-url" href="https://www.varitecconsulting.com" target="_blank" rel="noopener">www.varitecconsulting.com</a>

    </div>

  </div>

</div>

<input type="file" id="imageFileInput" accept="image/*" style="display:none" onchange="handleImageFile(event)">

<div class="insert-table-modal" id="tableModal">

  <div class="itm-box">

    <h4>&#9776; Insert Table</h4>

    <div class="itm-row">

      <label>Rows<input type="number" id="tblRows" value="3" min="1" max="20"></label>

      <label>Columns<input type="number" id="tblCols" value="3" min="1" max="10"></label>

    </div>

    <div style="font-size:11px;font-weight:700;color:#555;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px;">Column Headers (optional)</div>

    <div id="tblHeaderInputs" style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px;"></div>

    <div class="itm-btns">

      <button class="itm-cancel" onclick="closeTableModal()">Cancel</button>

      <button class="itm-ok" onclick="doInsertTable()">Insert</button>

    </div>

  </div>

</div>


<div class="modal-overlay" id="chartModal">

  <div class="modal" style="width:min(1400px,98vw);max-height:92vh;overflow-y:auto;">

    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px;">

      <div><h3>&#128202; Scorecard Performance Chart</h3><p style="margin-bottom:0;">Average weekly scores from your scorecard table</p></div>

      <button onclick="closeChartModal()" style="background:#f1f5f9;border:none;font-size:20px;cursor:pointer;color:#64748b;width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;">&#215;</button>

    </div>

    <div style="height:3px;background:linear-gradient(90deg,#8b3dff,#059669,#d97706,#2563eb);border-radius:3px;margin:12px 0 16px;"></div>

    <div id="chartNoData" style="display:none;text-align:center;padding:30px;color:#94a3b8;font-size:13px;">&#9888; No percentage values found.</div>

    <canvas id="scorecardChart" style="height:360px;width:100%;display:block;"></canvas>

    <div id="chartDataTable" style="margin-top:14px;font-size:12px;"></div>

    <div style="display:flex;gap:10px;margin-top:18px;justify-content:flex-end;">

      <button onclick="closeChartModal()" style="padding:9px 20px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:8px;font-size:13px;cursor:pointer;color:#475569;">Close</button>

      <button onclick="insertChartIntoNote()" style="padding:9px 22px;background:linear-gradient(135deg,#059669,#0d9488);color:white;border:none;border-radius:8px;font-size:13px;cursor:pointer;font-weight:700;">&#11015; Insert into Note</button>

    </div>

  </div>

</div>


<!-- Row Action Bar — clone/delete for any template row -->

<div class="row-action-bar" id="rowActionBar">

  <button class="rab-clone" onclick="rabClone()">&#128203; Clone Row</button>

  <button class="rab-delete" onclick="rabDelete()">&#128465; Delete Row</button>

  <button class="rab-del-cust" onclick="rabDeleteCustomer()" style="display:none;background:#7f1d1d;padding:5px 12px;border:none;border-radius:6px;cursor:pointer;font-size:11px;font-weight:700;color:white;white-space:nowrap;">&#128465; Delete Customer</button>

  <button class="rab-close" onclick="rabHide()">&#215;</button>

</div>

<div class="modal-overlay" id="sectionModal">

  <div class="modal">

    <h3 id="secModalTitle">Add Section</h3><p></p>

    <input id="sectionInput" type="text" placeholder="Section name…" maxlength="50">

    <div class="modal-btns">

      <button class="btn-cancel" onclick="closeSecModal()">Cancel</button>

      <button class="btn-confirm" onclick="confirmSecModal()">Save</button>

    </div>

  </div>

</div>


<!-- Image resize overlay -->

<div id="imgResizeOverlay">

  <div id="imgResizeToolbar">

    <button class="img-tb-btn" onclick="imgTbAlign('left')" title="Float left">&#8676; Left</button>

    <button class="img-tb-btn" onclick="imgTbAlign('center')" title="Center">&#9632; Center</button>

    <button class="img-tb-btn" onclick="imgTbAlign('right')" title="Float right">&#8677; Right</button>

    <div class="img-tb-sep"></div>

    <button class="img-tb-btn" onclick="imgTbSize(25)">25%</button>

    <button class="img-tb-btn" onclick="imgTbSize(50)">50%</button>

    <button class="img-tb-btn" onclick="imgTbSize(75)">75%</button>

    <button class="img-tb-btn" onclick="imgTbSize(100)">100%</button>

    <div class="img-tb-sep"></div>

    <span id="imgSizeLabel"></span>

    <div class="img-tb-sep"></div>

    <button class="img-tb-btn" onclick="imgTbDelete()" style="color:#f87171;" title="Delete image">&#128465;</button>

  </div>

  <div class="img-rh img-rh-tl" data-d="tl"></div>

  <div class="img-rh img-rh-tc" data-d="tc"></div>

  <div class="img-rh img-rh-tr" data-d="tr"></div>

  <div class="img-rh img-rh-ml" data-d="ml"></div>

  <div class="img-rh img-rh-mr" data-d="mr"></div>

  <div class="img-rh img-rh-bl" data-d="bl"></div>

  <div class="img-rh img-rh-bc" data-d="bc"></div>

  <div class="img-rh img-rh-br" data-d="br"></div>

</div>


<!-- Date Picker Popup — single shared instance -->

<div id="datePickerPopup"></div>


<script>

const NOTE_ID = {{ current.id }};

const ACTIVE_SECTION = {{ active_section | tojson }};

const SECTION_NOTE_IDS = {{ notes | map(attribute='id') | list | tojson }};

const ROW_ACCENTS = ['#8b5cf6','#059669','#d97706','#2563eb','#db2777','#0891b2','#7c3aed'];


let saveTimer=null, changed=false, secModal=null, renameId=null;
var _chartImages={};   /* {customerName: dataURL} */
var _chartImageMode='none'; /* none|fill|overlay|replace|top|pictograph */

let exportFile=null, tplFile=null, exportScope='this';

let _tblActiveCell=null;


/* ── Autosave ── */

document.getElementById('title').addEventListener('input', scheduleSave);

document.getElementById('content').addEventListener('input', scheduleSave);

function scheduleSave(){changed=true;setStatus('saving','Saving…');clearTimeout(saveTimer);saveTimer=setTimeout(saveNote,1500);}

document.addEventListener('keydown',function(e){

  if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();saveNote();}

  if(e.key==='Escape'){closeSecModal();closeChartModal();}

});

function saveNote(){

  if(!changed)return;

  fetch('/save/'+NOTE_ID,{method:'POST',headers:{'Content-Type':'application/json'},

    body:JSON.stringify({title:document.getElementById('title').value,content:document.getElementById('content').innerHTML})})

  .then(r=>r.json()).then(d=>{

    if(d.status==='ok'){changed=false;setStatus('saved','Saved '+new Date().toLocaleTimeString());

      const a=document.querySelector('.page-link.active .page-title');if(a)a.innerText=document.getElementById('title').value||'Untitled';}

  }).catch(()=>setStatus('error','Save failed'));

}

setInterval(()=>{if(changed)saveNote();},30000);

window.addEventListener('beforeunload',e=>{if(changed){saveNote();e.preventDefault();}});

function setStatus(cls,msg){const el=document.getElementById('status');el.className=cls;el.innerText=msg;}


/* ── Images ── */

document.getElementById('content').addEventListener('paste',function(e){
  var overlay=document.querySelector('.cover-drop-layer');
  if(overlay && overlay.closest('.cover-overlay-wrap')){
    var items=(e.clipboardData||e.originalEvent.clipboardData).items;
    for(const item of items){
      if(item.type.startsWith('image/')){
        e.preventDefault();
        const r=new FileReader();
        r.onload=ev=>{placeOverlayImg(ev.target.result,'10','10');scheduleSave();};
        r.readAsDataURL(item.getAsFile());
        return;
      }
    }
  }
  // original paste handler below
  if(false){

  const items=(e.clipboardData||e.originalEvent.clipboardData).items;

  for(const item of items){if(item.type.startsWith('image/')){e.preventDefault();const r=new FileReader();r.onload=ev=>{insertImgAtCursor(ev.target.result);scheduleSave();};r.readAsDataURL(item.getAsFile());return;}}

  if(e.clipboardData.types.includes('text/plain')){e.preventDefault();document.execCommand('insertText',false,e.clipboardData.getData('text/plain'));}

}});

const editor=document.getElementById('content');

editor.addEventListener('dragover',e=>e.preventDefault());

editor.addEventListener('drop',function(e){
  e.preventDefault();
  var overlay=document.querySelector('.cover-drop-layer');
  if(overlay){var wrap=overlay.closest('.cover-overlay-wrap');if(wrap){var rect=wrap.getBoundingClientRect();if(e.clientX>=rect.left&&e.clientX<=rect.right&&e.clientY>=rect.top&&e.clientY<=rect.bottom){var xPct=((e.clientX-rect.left)/rect.width*100).toFixed(1);var yPct=((e.clientY-rect.top)/rect.height*100).toFixed(1);for(const f of e.dataTransfer.files){if(f.type.startsWith('image/')){const r=new FileReader();r.onload=ev=>{placeOverlayImg(ev.target.result,xPct,yPct);scheduleSave();};r.readAsDataURL(f);return;}}}}}
  for(const f of e.dataTransfer.files){if(f.type.startsWith('image/')){const r=new FileReader();r.onload=ev=>{insertImgAtCursor(ev.target.result);scheduleSave();};r.readAsDataURL(f);}}
});

function placeOverlayImg(src,xPct,yPct){
  var overlay=document.querySelector('.cover-drop-layer');
  if(!overlay)return;
  overlay.style.pointerEvents='all';
  var hint=document.querySelector('.cover-overlay-hint');
  if(hint)hint.style.display='none';

  var wrapper=document.createElement('div');
  wrapper.style.cssText='position:absolute;left:'+xPct+'%;top:'+yPct+'%;transform:translate(-50%,-50%);width:15%;z-index:15;cursor:move;';wrapper.setAttribute('data-cov','img');

  var img=document.createElement('img');
  img.src=src;
  img.style.cssText='width:100%;height:auto;display:block;border:none;border-radius:0;pointer-events:none;';
  img.draggable=false;

  // Delete button
  var del=document.createElement('button');
  del.innerHTML='&#10005;';
  del.style.cssText='position:absolute;top:-10px;right:-10px;width:22px;height:22px;background:#dc2626;color:white;border:none;border-radius:50%;cursor:pointer;font-size:12px;font-weight:700;display:none;align-items:center;justify-content:center;z-index:20;line-height:1;padding:0;';
  del.addEventListener('mousedown',function(e){e.stopPropagation();});
  del.addEventListener('click',function(e){e.stopPropagation();if(confirm('Delete this logo?')){wrapper.remove();scheduleSave();}});

  // Resize handle (bottom-right)
  var rh=document.createElement('div');
  rh.style.cssText='position:absolute;bottom:-6px;right:-6px;width:14px;height:14px;background:#8b3dff;border:2px solid white;border-radius:3px;cursor:se-resize;z-index:20;display:none;';
  rh.addEventListener('mousedown',function(e){
    e.preventDefault();e.stopPropagation();
    var startX=e.clientX,startW=wrapper.offsetWidth;
    function mm(e){
      var nw=Math.max(40,startW+(e.clientX-startX));
      wrapper.style.width=nw+'px';
    }
    function mu(){document.removeEventListener('mousemove',mm);document.removeEventListener('mouseup',mu);scheduleSave();}
    document.addEventListener('mousemove',mm);
    document.addEventListener('mouseup',mu);
  });

  // Drag to move
  wrapper.addEventListener('mousedown',function(e){
    if(e.target===del||e.target===rh)return;
    e.preventDefault();e.stopPropagation();
    var _ox=e.clientX,_oy=e.clientY;
    function mm(e){
      var wr=wrapper.closest('.cover-overlay-wrap');
      if(!wr)return;
      var rect=wr.getBoundingClientRect();
      wrapper.style.left=(parseFloat(wrapper.style.left)+(e.clientX-_ox)/rect.width*100)+'%';
      wrapper.style.top=(parseFloat(wrapper.style.top)+(e.clientY-_oy)/rect.height*100)+'%';
      _ox=e.clientX;_oy=e.clientY;
    }
    function mu(){document.removeEventListener('mousemove',mm);document.removeEventListener('mouseup',mu);scheduleSave();}
    document.addEventListener('mousemove',mm);
    document.addEventListener('mouseup',mu);
  });

  wrapper.addEventListener('click',function(e){e.stopPropagation();del.style.display='flex';rh.style.display='block';});
  document.addEventListener('click',function(e){if(!wrapper.contains(e.target)){del.style.display='none';rh.style.display='none';}});
  rh.addEventListener('mouseup',function(){scheduleSave();setTimeout(function(){del.style.display='none';rh.style.display='none';},300);});

  wrapper.appendChild(img);
  wrapper.appendChild(del);
  wrapper.appendChild(rh);
  overlay.appendChild(wrapper);
  scheduleSave();
}

/* COVER_OVERLAY_IMG_REHYDRATE_V1 */
function coverWireOverlayImgWrapper(wrapper){
  if (wrapper.dataset.covWired === '1') return;
  var img = wrapper.querySelector('img');
  var del = wrapper.querySelector('button');
  var rh = null;
  for (var i=0;i<wrapper.children.length;i++){
    var el = wrapper.children[i];
    if (el.tagName === 'DIV' && el.style && el.style.cursor === 'se-resize') { rh = el; break; }
  }
  if (!img || !del || !rh) return;
  wrapper.dataset.covWired = '1';

  del.addEventListener('mousedown', function(e){ e.stopPropagation(); });
  del.addEventListener('click', function(e){
    e.stopPropagation();
    if (confirm('Delete this logo?')) { wrapper.remove(); if (typeof scheduleSave === 'function') scheduleSave(); }
  });

  rh.addEventListener('mousedown', function(e){
    e.preventDefault(); e.stopPropagation();
    var startX = e.clientX, startW = wrapper.offsetWidth;
    function mm(e){ var nw = Math.max(40, startW + (e.clientX - startX)); wrapper.style.width = nw + 'px'; }
    function mu(){
      document.removeEventListener('mousemove', mm);
      document.removeEventListener('mouseup', mu);
      if (typeof scheduleSave === 'function') scheduleSave();
    }
    document.addEventListener('mousemove', mm);
    document.addEventListener('mouseup', mu);
  });

  wrapper.addEventListener('mousedown', function(e){
    if (e.target === del || e.target === rh) return;
    e.preventDefault(); e.stopPropagation();
    var _ox = e.clientX, _oy = e.clientY;
    function mm(e){
      var wr = wrapper.closest('.cover-overlay-wrap');
      if (!wr) return;
      var rect = wr.getBoundingClientRect();
      wrapper.style.left = (parseFloat(wrapper.style.left) + (e.clientX - _ox) / rect.width * 100) + '%';
      wrapper.style.top = (parseFloat(wrapper.style.top) + (e.clientY - _oy) / rect.height * 100) + '%';
      _ox = e.clientX; _oy = e.clientY;
    }
    function mu(){
      document.removeEventListener('mousemove', mm);
      document.removeEventListener('mouseup', mu);
      if (typeof scheduleSave === 'function') scheduleSave();
    }
    document.addEventListener('mousemove', mm);
    document.addEventListener('mouseup', mu);
  });

  wrapper.addEventListener('click', function(e){
    e.stopPropagation(); del.style.display='flex'; rh.style.display='block';
  });
  document.addEventListener('click', function(e){
    if (!wrapper.contains(e.target)) { del.style.display='none'; rh.style.display='none'; }
  });
  rh.addEventListener('mouseup', function(){
    if (typeof scheduleSave === 'function') scheduleSave();
    setTimeout(function(){ del.style.display='none'; rh.style.display='none'; }, 300);
  });
}
function coverInitOverlayImgHandlers(){
  document.querySelectorAll('.cover-overlay-wrap [data-cov="img"]').forEach(function(wrapper){
    coverWireOverlayImgWrapper(wrapper);
  });
}
if (document.readyState !== 'loading') coverInitOverlayImgHandlers();
else document.addEventListener('DOMContentLoaded', coverInitOverlayImgHandlers);

function insertImgAtCursor(src){const img=document.createElement('img');img.src=src;img.style.maxWidth='100%';img.setAttribute('draggable','true');const contentEl=document.getElementById('content');const sel=window.getSelection();if(sel&&sel.rangeCount){const rng=sel.getRangeAt(0);if(contentEl.contains(rng.commonAncestorContainer)){rng.collapse(false);rng.insertNode(img);rng.setStartAfter(img);rng.collapse(true);sel.removeAllRanges();sel.addRange(rng);return;}}contentEl.focus();contentEl.appendChild(img);}

/* ── Drag-and-drop repositioning of images within content ── */
(function(){
  var _dragImg=null;
  document.getElementById('content').addEventListener('dragstart',function(e){
    if(e.target.tagName==='IMG'){_dragImg=e.target;e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain','img-move');}
  });
  document.getElementById('content').addEventListener('dragover',function(e){
    if(_dragImg)e.preventDefault();
  });
  document.getElementById('content').addEventListener('drop',function(e){
    if(!_dragImg)return;
    e.preventDefault();
    var range=document.caretRangeFromPoint?document.caretRangeFromPoint(e.clientX,e.clientY):null;
    if(range){range.insertNode(_dragImg);}
    _dragImg=null; scheduleSave();
  });
})();

function insertImageFile(){document.getElementById('imageFileInput').click();}

function handleImageFile(e){const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{insertImgAtCursor(ev.target.result);scheduleSave();};r.readAsDataURL(f);e.target.value='';}
function handleCoverLogoFile(e){const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{const box=e.target.previousElementSibling.previousElementSibling;const img=box.querySelector('img');const ph=box.querySelector('.cover-logo-placeholder');img.src=ev.target.result;img.style.display='block';ph.style.display='none';scheduleSave();};r.readAsDataURL(f);e.target.value='';}

function fmt(cmd){document.getElementById('content').focus();document.execCommand(cmd,false,null);scheduleSave();}

function deletePage(e,noteId){e.preventDefault();e.stopPropagation();if(!confirm('Delete this page?'))return;fetch('/delete/'+noteId,{method:'POST'}).then(()=>location.href='/section/'+ACTIVE_SECTION);}


/* ── Scorecard dropdown ── */

function toggleScDropdown(ev){ev.stopPropagation();var dd=document.getElementById('scDropdown');if(dd.style.display==='block'){dd.style.display='none';return;}var r=document.getElementById('scArrowBtn').getBoundingClientRect();dd.style.top=(r.bottom+4)+'px';dd.style.left=Math.max(4,r.right-240)+'px';dd.style.display='block';}

document.addEventListener('click',function(e){var dd=document.getElementById('scDropdown');if(dd&&!dd.contains(e.target)&&e.target.id!=='scArrowBtn')dd.style.display='none';});

function createScorecard(tpl){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';var fd=new FormData();fd.append('tpl',tpl);fetch('/new-scorecard/'+ACTIVE_SECTION,{method:'POST',body:fd}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function createSerializationIssue(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-serialization/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function createOnboarding(type){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';var fd=new FormData();fd.append('type',type);fetch('/new-onboarding/'+ACTIVE_SECTION,{method:'POST',body:fd}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function createExecutiveSummary(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-exec-summary/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function createAgenda(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-agenda/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}
function agendaAddItem(btn){var wrap=btn.closest('.agenda-wrap');var body=wrap?wrap.querySelector('.agenda-body'):null;if(!body)return;var count=body.querySelectorAll('.agenda-item').length;var div=document.createElement('div');div.className='agenda-item';div.innerHTML='<button class="agenda-del-item" type="button" onclick="agendaDelItem(this)" title="Remove item">&#10005;</button><div class="agenda-num" contenteditable="true">'+(count+1)+'</div><div class="agenda-label" contenteditable="true"></div>';body.appendChild(div);agendaRenumber(body);var lbl=div.querySelector('.agenda-label');if(lbl)lbl.focus();scheduleSave();}
/* AGENDA_RENUMBER_V1 */
function agendaRenumber(body){
  if(!body)return;
  var items=body.querySelectorAll('.agenda-item');
  items.forEach(function(el,idx){
    var numEl=el.querySelector('.agenda-num');
    if(numEl)numEl.textContent=String(idx+1);
  });
}
function agendaInitRenumber(){
  document.querySelectorAll('.agenda-wrap .agenda-body').forEach(function(body){
    agendaRenumber(body);
  });
}
if(document.readyState!=='loading')agendaInitRenumber();
else document.addEventListener('DOMContentLoaded',agendaInitRenumber);
function agendaDelItem(btn){var item=btn.closest('.agenda-item');var body=item?item.closest('.agenda-body'):null;if(!body)return;if(body.querySelectorAll('.agenda-item').length<=1){alert('Cannot delete the last item.');return;}item.remove();agendaRenumber(body);scheduleSave();}

function createCoverSlide(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-cover-slide/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function createThankYou(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-thank-you/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function createRolesResponsibilities(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-roles-responsibilities/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function createFocusAreas(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-focus-areas/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}
function createProcessImprovements(){var dd=document.getElementById('scDropdown');if(dd)dd.style.display='none';fetch('/new-process-improvements/'+ACTIVE_SECTION,{method:'POST'}).then(r=>{if(r.redirected)location.href=r.url;else location.reload();}).catch(()=>location.reload());}

function piBuildStatusColsHTML(){
  return '<div class="pi-status-cols">'
    +'<div class="pi-status-col"><span class="pi-status-label">Test</span>'
    +'<span class="pi-pill pi-pill-pending" data-status="pending" data-label="Pending">Pending</span></div>'
    +'<div class="pi-status-col"><span class="pi-status-label">Production</span>'
    +'<span class="pi-pill pi-pill-pending" data-status="pending" data-label="Pending">Pending</span></div>'
    +'</div>';
}

function piAddItemRow(body,beforeBtn){
  var div=document.createElement('div');
  div.className='pi-item';
  div.innerHTML='<button class="pi-del-item" type="button" title="Delete point">&#10005;</button>'
    +'<div class="pi-item-row">'
    +'<div class="pi-main-icon pi-ghost"></div>'
    +'<div class="pi-item-text" contenteditable="true" data-ph="Process improvement point..."></div>'
    +piBuildStatusColsHTML()
    +'</div>'
    +'<div class="pi-sub-list"></div>'
    +'<div class="pi-add-sub-wrap"><button class="pi-add-sub" type="button" '
    +'onclick="event.stopPropagation();var w=this.closest(\\'.pi-add-sub-wrap\\');var l=w?w.previousElementSibling:null;if(l)piAddSubRow(l,this);">'
    +'&#43; Add Sub-point</button></div>';
  body.insertBefore(div,beforeBtn);
  var t=div.querySelector('.pi-item-text');if(t)t.focus();
  scheduleSave();
}

function piAddSubRow(list,beforeBtn){
  var row=document.createElement('div');
  row.className='pi-sub-row';
  row.innerHTML='<div class="pi-sub-icon"></div>'
    +'<div class="pi-sub-text" contenteditable="true" data-ph="Sub-point..."></div>'
    +piBuildStatusColsHTML()
    +'<button class="pi-del-sub" type="button" title="Remove">&#10005;</button>';
  list.appendChild(row);
  var t=row.querySelector('.pi-sub-text');if(t)t.focus();
  scheduleSave();
}

var PI_PILL_CYCLE=['pending','inprogress','hold','completed'];
var PI_PILL_LABELS={pending:'Pending',inprogress:'In Progress',hold:'Hold',completed:'Completed'};

document.getElementById('content').addEventListener('click', function(e){
  var pillBtn = e.target.closest('.pi-pill');
  if(pillBtn){
    e.preventDefault(); e.stopPropagation();
    var cur=pillBtn.dataset.status||'pending';
    var next=PI_PILL_CYCLE[(PI_PILL_CYCLE.indexOf(cur)+1)%PI_PILL_CYCLE.length];
    pillBtn.className='pi-pill pi-pill-'+next;
    pillBtn.dataset.status=next;
    pillBtn.dataset.label=PI_PILL_LABELS[next];
    pillBtn.textContent=PI_PILL_LABELS[next];
    scheduleSave();
    return;
  }
  var addItemBtn = e.target.closest('.pi-add-item');
  if(addItemBtn){
    e.preventDefault(); e.stopPropagation();
    var body = addItemBtn.closest('.pi-body');
    if(body) piAddItemRow(body, addItemBtn);
    return;
  }
  var addSubBtn = e.target.closest('.pi-add-sub');
  if(addSubBtn){
    e.preventDefault(); e.stopPropagation();
    var wrap = addSubBtn.closest('.pi-add-sub-wrap');
    var list = wrap ? wrap.previousElementSibling : null;
    if(list && list.classList.contains('pi-sub-list')) piAddSubRow(list, addSubBtn);
    return;
  }
  var delItemBtn = e.target.closest('.pi-del-item');
  if(delItemBtn){
    e.preventDefault(); e.stopPropagation();
    var item = delItemBtn.closest('.pi-item');
    var body2 = item ? item.closest('.pi-body') : null;
    if(!body2) return;
    if(body2.querySelectorAll('.pi-item').length<=1){
      alert('Cannot delete the last point.'); return;
    }
    if(!confirm('Delete this point?')) return;
    item.remove(); scheduleSave();
    return;
  }
  var delSubBtn = e.target.closest('.pi-del-sub');
  if(delSubBtn){
    e.preventDefault(); e.stopPropagation();
    var row = delSubBtn.closest('.pi-sub-row');
    if(row){ row.remove(); scheduleSave(); }
    return;
  }
});



var RR_COLORS=[['#3b82f6','#93c5fd'],['#10b981','#6ee7b7'],['#f59e0b','#fcd34d'],['#06b6d4','#67e8f9'],['#8b5cf6','#c4b5fd'],['#ef4444','#fca5a5'],['#0ea5e9','#7dd3fc'],['#f97316','#fdba74']];

function rrNormalizeWrap(wrap){
  var grid=wrap.querySelector('.rr-cards-grid');
  if(!grid)return false;
  var cards=Array.prototype.slice.call(grid.querySelectorAll('.rr-card'));
  var changed=false;
  cards.forEach(function(card,i){
    var pair=RR_COLORS[i%RR_COLORS.length];
    var col=pair[0],colLight=pair[1];
    var curCol=card.style.getPropertyValue('--rr-accent');
    if(curCol!==col){card.style.setProperty('--rr-accent',col);changed=true;}
    if(card.style.getPropertyValue('--rr-accent-light')!==colLight){
      card.style.setProperty('--rr-accent-light',colLight);changed=true;
    }
    if(card.getAttribute('data-rr-card')!==String(i)){
      card.setAttribute('data-rr-card',i);changed=true;
    }
    var num=card.querySelector('.rr-num');
    if(!num){
      num=document.createElement('div');
      num.className='rr-num';
      var body=card.querySelector('.rr-card-body');
      var delBtn=card.querySelector('.rr-del-card');
      if(body)card.insertBefore(num,body);
      else card.appendChild(num);
      changed=true;
    }
    var expectedNum=String(i+1).padStart(2,'0');
    if(num.textContent!==expectedNum){num.textContent=expectedNum;changed=true;}
    var title=card.querySelector('.rr-card-title');
    if(title&&title.classList.contains('rr-bold')){
      title.classList.remove('rr-bold');changed=true;
    }
  });
  return changed;
}
function rrNormalizeAll(save){
  var any=false;
  document.querySelectorAll('.rr-wrap').forEach(function(wrap){
    if(rrNormalizeWrap(wrap))any=true;
  });
  if(any&&save!==false)scheduleSave();
  return any;
}
document.addEventListener('DOMContentLoaded',function(){rrNormalizeAll(true);});
if(document.readyState!=='loading')rrNormalizeAll(true);


function rrAddCard(btn){
  var grid=btn.closest('.rr-cards-grid');if(!grid)return;
  var div=document.createElement('div');
  div.className='rr-card rr-empty-desc';
  div.innerHTML='<button class="rr-del-card" type="button" onclick="rrDeleteCard(this)" title="Delete">&#10005;</button>'+'<div class="rr-num"></div>'+'<div class="rr-card-body">'+'<div class="rr-card-title" contenteditable="true" data-ph="Role or responsibility..."></div>'+'<div class="rr-card-desc" contenteditable="true" data-ph="Add a short note..."></div>'+'</div>';
  grid.insertBefore(div,btn);
  rrNormalizeAll(false);
  var t=div.querySelector('.rr-card-title');if(t)t.focus();
  scheduleSave();
}

function rrDeleteCard(btn){
  var card=btn.closest('.rr-card');
  var grid=card?card.closest('.rr-cards-grid'):null;
  if(!grid)return;
  if(grid.querySelectorAll('.rr-card').length<=1){alert('Cannot delete the last card.');return;}
  if(!confirm('Delete this card?'))return;
  card.remove();
  rrNormalizeAll(false);
  scheduleSave();
}



/* ── Export PDF ── */

function toggleProfileMenu(e){

  e.stopPropagation();

  var m=document.getElementById('profileMenu');

  m.classList.toggle('show');

}

document.addEventListener('click',function(e){

  var m=document.getElementById('profileMenu');

  if(m&&!m.closest('.profile-wrap').contains(e.target)) m.classList.remove('show');

});

function exportAsPDF(){
  /* Step 1: capture canvas → img BEFORE saving */
  var hasChart=document.getElementById('sc-inline-wrap');
  if(hasChart){
    var c2=hasChart.querySelector('canvas');
    var img=hasChart.querySelector('#sc-chart-pdf-img');
    if(c2&&img&&c2.width>0){img.src=c2.toDataURL('image/png');}
  }
  var _openPDF=function(){
    var url='/export/pdf/'+encodeURIComponent(ACTIVE_SECTION);
    var win=window.open(url,'_blank');
    if(!win){alert('Please allow popups to export PDF.');}
  };
  /* Step 2: save (innerHTML now includes chart img with base64 src) */
  var _title=document.getElementById('title').value;
  var _content=document.getElementById('content').innerHTML;
  fetch('/save/'+NOTE_ID,{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({title:_title,content:_content})
  }).then(function(r){return r.json();})
  .then(function(){
    changed=false;
    setStatus('saved','Saved '+new Date().toLocaleTimeString());
    /* Step 3: open PDF after save confirmed */
    setTimeout(_openPDF,300);
  }).catch(function(){
    setTimeout(_openPDF,300);
  });
}


/* ── Section modal ── */

function openAddSection(){secModal='add';renameId=null;document.getElementById('secModalTitle').innerText='Add New Section';document.getElementById('sectionInput').value='';document.getElementById('sectionModal').classList.add('show');document.getElementById('sectionInput').focus();}

function openRename(id,name){secModal='rename';renameId=id;document.getElementById('secModalTitle').innerText='Rename Section';document.getElementById('sectionInput').value=name;document.getElementById('sectionModal').classList.add('show');document.getElementById('sectionInput').focus();}

function closeSecModal(){document.getElementById('sectionModal').classList.remove('show');}

document.getElementById('sectionInput').addEventListener('keydown',e=>{if(e.key==='Enter')confirmSecModal();});

function confirmSecModal(){const name=document.getElementById('sectionInput').value.trim();if(!name)return;const url=secModal==='add'?'/section/add':'/section/rename/'+renameId;fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})}).then(r=>r.json()).then(d=>{if(d.status==='ok')location.href=secModal==='add'?'/section/'+name:location.href;else alert(d.message||'Error.');});closeSecModal();}

function deleteSection(id,name){if(!confirm('Delete section "'+name+'" and ALL its pages?'))return;fetch('/section/delete/'+id,{method:'POST'}).then(()=>location.href='/notes');}


/* ── Insert Table / Box ── */

function openTableModal(){document.getElementById('tableModal').classList.add('show');renderHeaderInputs();document.getElementById('tblCols').addEventListener('input',renderHeaderInputs);}

function closeTableModal(){document.getElementById('tableModal').classList.remove('show');}

function renderHeaderInputs(){var cols=parseInt(document.getElementById('tblCols').value)||3,div=document.getElementById('tblHeaderInputs'),existing=Array.from(div.querySelectorAll('input')).map(i=>i.value);div.innerHTML='';for(var c=0;c<Math.min(cols,8);c++){var inp=document.createElement('input');inp.placeholder='Col '+(c+1);inp.value=existing[c]||'';inp.style.cssText='flex:1;min-width:70px;padding:5px 7px;border:1px solid #ddd;border-radius:5px;font-size:12px;outline:none;';div.appendChild(inp);}}

function doInsertTable(){var rows=parseInt(document.getElementById('tblRows').value)||3,cols=parseInt(document.getElementById('tblCols').value)||3,headers=Array.from(document.getElementById('tblHeaderInputs').querySelectorAll('input')).map((i,idx)=>i.value||('Column '+(idx+1)));var html='<table><thead><tr>';for(var c=0;c<cols;c++)html+='<th>'+(headers[c]||('Column '+(c+1)))+'</th>';html+='</tr></thead><tbody>';for(var r=0;r<rows;r++){html+='<tr>';for(var c=0;c<cols;c++)html+='<td>&nbsp;</td>';html+='</tr>';}html+='</tbody></table><p><br></p>';document.getElementById('content').focus();document.execCommand('insertHTML',false,html);scheduleSave();closeTableModal();}

function insertBox(){var html='<div style="border:2px solid #8b3dff;border-radius:8px;padding:14px 18px;margin:12px 0;background:#f5f0ff;"><div style="font-weight:700;color:#5b21b6;margin-bottom:6px;font-size:14px;">Box Title</div><div style="color:#333;font-size:14px;">Enter your content here...</div></div><p><br></p>';document.getElementById('content').focus();document.execCommand('insertHTML',false,html);scheduleSave();}


/* ── Scorecard init ── */

function initScorecard(){document.querySelectorAll('#content .sc-val').forEach(function(el){el.setAttribute('contenteditable','true');if(el._scBound)return;el._scBound=true;el.addEventListener('input',scheduleSave);});}

document.addEventListener('keydown',function(e){
  var el=e.target;
  if(el.nodeType===3)el=el.parentElement;
  if(!el||!el.classList||!el.classList.contains('sc-val'))return;
  if(!/^[0-9.]$/.test(e.key)&&!['Backspace','Delete','ArrowLeft','ArrowRight','Tab','Enter'].includes(e.key)){e.preventDefault();return;}
  if(e.key==='Tab'||e.key==='Enter'){
    e.preventDefault();
    e.stopPropagation();
    var table=el.closest('.sc-table');
    var cells=table?Array.prototype.slice.call(table.querySelectorAll('.sc-val')):[];
    var idx2=cells.indexOf(el);
    var next=(e.shiftKey&&e.key==='Tab')?cells[idx2-1]:cells[idx2+1];
    if(next){
      next.focus();
      var range=document.createRange();
      range.selectNodeContents(next);
      range.collapse(false);
      var sel=window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }else{
      el.blur();
    }
  }
}, true);

document.addEventListener('DOMContentLoaded',function(){initScorecard();updateTitleVisibility();siInitDeleteBtns();});

(function(){if(document.readyState!=='loading')initScorecard();})();


/* ── Bar Chart ── */

function generateScorecardChart(){

  var content=document.getElementById('content');

  var table=content.querySelector('.sc-table')||content.querySelector('table');

  if(!table){alert('No scorecard table found.');return;}

  var headers=Array.from(table.querySelectorAll('thead th')).map(th=>th.textContent.trim());

  var customerCol=0,weekCols=[];

  headers.forEach(function(h,i){if(h.toLowerCase().indexOf("week")>=0)weekCols.push(i);if(h.toLowerCase().indexOf("customer")>=0)customerCol=i;});

  if(weekCols.length===0){alert('No Week columns found.');return;}

  var customers=[];

  Array.from(table.querySelectorAll('tbody tr')).forEach(function(row){

    var cells=Array.from(row.querySelectorAll('td'));if(!cells.length)return;

    var name=(cells[customerCol]||{}).textContent.replace(/[\u25cf\u2022\u2713]/g,'').trim();if(!name)return;

    var scores=[];

    weekCols.forEach(function(ci){if(ci<cells.length){var valEl=cells[ci].querySelector('.sc-val');var raw=(valEl?valEl.textContent:cells[ci].textContent).replace(/[^0-9.]/g,'').trim();var val=parseFloat(raw);if(raw!==''&&!isNaN(val)&&val>0)scores.push(val);}});

    if(scores.length>0){var avg=scores.reduce((a,b)=>a+b,0)/scores.length;customers.push({name,avg:Math.round(avg*100)/100,scores});}

  });

  document.getElementById('chartModal').classList.add('show');

  var canvas=document.getElementById('scorecardChart'),noData=document.getElementById('chartNoData'),tbl=document.getElementById('chartDataTable');

  if(customers.length===0){canvas.style.display='none';noData.style.display='block';tbl.innerHTML='';return;}

  noData.style.display='none';canvas.style.display='block';

  if(_chartInstance){_chartInstance.destroy();;}

  var ctx=canvas.getContext('2d');

  function scoreColor(v){
  if(v>=95)return '#4ade80';
  if(v>=90)return '#60a5fa';
  if(v>=85)return '#c084fc';
  if(v>=75)return '#fbbf24';
  return '#f87171';
}
var palette=customers.map(c=>scoreColor(c.avg));

  var grads=customers.map(function(c,i){var col=palette[i];var g=ctx.createLinearGradient(0,0,canvas.width,0);g.addColorStop(0,col);g.addColorStop(1,col+'55');return g;});

  var colors=customers.map((c,i)=>palette[i]);

  _chartInstance=new Chart(canvas,{type:'bar',data:{labels:customers.map(c=>c.name),datasets:[{label:'Average Score (%)',data:customers.map(c=>c.avg),backgroundColor:grads,borderColor:colors,borderWidth:0,borderRadius:0,borderSkipped:false,barPercentage:0.92,categoryPercentage:0.98}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,layout:{padding:{top:8,right:90,bottom:22,left:90}},plugins:{legend:{display:false},tooltip:{backgroundColor:'#1e293b',callbacks:{label:ctx=>'  Score: '+ctx.parsed.x+'%'}}},scales:{x:{beginAtZero:false,min:Math.max(0,Math.floor(Math.min(...customers.map(c=>c.avg))-15)),max:100,ticks:{callback:v=>v+'%',color:'#000000',font:{weight:'900',size:13}},grid:{color:'#f1f5f9'}},y:{grid:{display:true,color:'#f1f5f9'},ticks:{font:{size:14,weight:'900'},color:'#000000'}}}},plugins:[{id:'bar3d',beforeDatasetsDraw(chart){const ctx2=chart.ctx;const depth=14;chart.data.datasets.forEach((ds,i)=>{chart.getDatasetMeta(i).data.forEach((bar,idx)=>{const col=colors[idx]||'#3b82f6';const x0=bar.base,x1=bar.x,y0=bar.y-bar.height/2,y1=bar.y+bar.height/2;
  function shade(hex,p){var n=parseInt(hex.slice(1),16);var r=(n>>16)+p,g=((n>>8)&255)+p,b=(n&255)+p;r=Math.min(255,Math.max(0,r));g=Math.min(255,Math.max(0,g));b=Math.min(255,Math.max(0,b));return 'rgb('+r+','+g+','+b+')';}
  ctx2.save();
  // top face cap (lighter) - only near the end
  ctx2.fillStyle=shade(col,45);
  ctx2.beginPath();
  ctx2.moveTo(x1-depth,y0);ctx2.lineTo(x1,y0);ctx2.lineTo(x1+depth,y0-depth);ctx2.lineTo(x1,y0-depth);ctx2.closePath();ctx2.fill();
  // right face (darker)
  ctx2.fillStyle=shade(col,-35);
  ctx2.beginPath();
  ctx2.moveTo(x1,y0);ctx2.lineTo(x1+depth,y0-depth);ctx2.lineTo(x1+depth,y1-depth);ctx2.lineTo(x1,y1);ctx2.closePath();ctx2.fill();
  ctx2.restore();
});});}},
{id:'dl',afterDatasetsDraw(chart){const ctx2=chart.ctx;const depth=14;chart.data.datasets.forEach((ds,i)=>{chart.getDatasetMeta(i).data.forEach((bar,idx)=>{const v=ds.data[idx];const tx=bar.x+depth+8,ty=bar.y;ctx2.save();ctx2.font='900 16px Segoe UI';const txt=v+'%';const tw=ctx2.measureText(txt).width;ctx2.fillStyle='rgba(255,255,255,.9)';ctx2.fillRect(tx-3,ty-11,tw+6,22);ctx2.fillStyle='#000000';ctx2.textAlign='left';ctx2.textBaseline='middle';ctx2.fillText(txt,tx,ty);ctx2.restore();});});}}]});

  var wh=Array(weekCols.length).fill(0).map((_,i)=>'<th style="padding:7px 10px;text-align:center;border:1px solid #e2e8f0;">Week '+(i+1)+'</th>').join('');

  var cr=customers.map(function(c){var col=c.avg>=95?'#16a34a':c.avg>=85?'#d97706':'#dc2626';var wc=Array(weekCols.length).fill(0).map((_,i)=>'<td style="padding:6px 10px;text-align:center;border:1px solid #e2e8f0;">'+(c.scores[i]!=null?c.scores[i]+'%':'—')+'</td>').join('');return'<tr><td style="padding:6px 10px;border:1px solid #e2e8f0;font-weight:600;">'+c.name+'</td>'+wc+'<td style="padding:6px 10px;text-align:center;border:1px solid #e2e8f0;font-weight:700;color:'+col+';">'+c.avg+'%</td></tr>';}).join('');

  tbl.innerHTML='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-top:8px;"><thead><tr style="background:#f1f5f9;"><th style="padding:7px 10px;text-align:left;border:1px solid #e2e8f0;">Customer</th>'+wh+'<th style="padding:7px 10px;text-align:center;border:1px solid #e2e8f0;background:#0d152b;color:white;">Avg</th></tr></thead><tbody>'+cr+'</tbody></table>';

}

function closeChartModal(){document.getElementById('chartModal').classList.remove('show');}

function generateScorecardChartInline(){
  var contentEl=document.getElementById('content');
  var table=contentEl.querySelector('.sc-table')||contentEl.querySelector('table');
  if(!table){alert('No scorecard table found.');return;}
  var existing=document.getElementById('sc-inline-wrap');
  if(existing)existing.remove();

  /* Extract data */
  var headers=Array.from(table.querySelectorAll('thead th')).map(function(th){return th.textContent.trim();});
  var customerCol=0,weekCols=[];
  headers.forEach(function(h,i){
    if(h.toLowerCase().indexOf('week')>=0)weekCols.push(i);
    if(h.toLowerCase().indexOf('customer')>=0)customerCol=i;
  });
  if(!weekCols.length){alert('No Week columns found.');return;}

  var customers=[];
  Array.from(table.querySelectorAll('tbody tr')).forEach(function(row){
    var cells=Array.from(row.querySelectorAll('td'));
    if(!cells.length)return;
    var name=(cells[customerCol]||{}).textContent.trim().replace(/[^a-zA-Z0-9 ]/g,'').trim();
    if(!name)return;
    var scores=[];
    weekCols.forEach(function(ci){
      if(ci<cells.length){
        var valEl=cells[ci].querySelector('.sc-val');
        var raw=(valEl?valEl.textContent:cells[ci].textContent).replace(/[^0-9.]/g,'').trim();
        var val=parseFloat(raw);
        if(raw!==''&&!isNaN(val)&&val>0)scores.push(val);
      }
    });
    if(scores.length>0){
      var avg=scores.reduce(function(a,b){return a+b;},0)/scores.length;
      customers.push({name:name,avg:Math.round(avg*10)/10});
    }
  });
  if(!customers.length){alert('No data found.');return;}

  /* Build wrapper */
  var wrap=document.createElement('div');
  wrap.id='sc-inline-wrap';
  wrap.setAttribute('contenteditable','false');
  wrap.style.cssText='background:#ffffff;border-radius:16px;padding:16px 16px 10px;margin-top:16px;box-shadow:0 4px 20px rgba(0,0,0,.12);border:1px solid #e2e8f0;position:relative;overflow:hidden;';

  /* Title */
  var titleEl=document.createElement('div');
  titleEl.style.cssText='font-size:11px;font-weight:800;color:#334155;text-transform:uppercase;letter-spacing:1.5px;margin-bottom:10px;';
  titleEl.textContent='MONTHLY AVERAGE SCORECARD PERFORMANCE';
  wrap.appendChild(titleEl);

  /* Remove button */
  var delBtn=document.createElement('button');
  delBtn.className='sc-chart-del-btn';
  delBtn.setAttribute('contenteditable','false');
  delBtn.innerHTML='&#10005; Remove Chart';
  delBtn.style.cssText='position:absolute;top:10px;right:10px;z-index:30;background:rgba(220,38,38,.75);border:none;color:white;border-radius:8px;padding:5px 13px;font-size:11px;font-weight:700;cursor:pointer;letter-spacing:.3px;font-family:Segoe UI,sans-serif;';
  delBtn.addEventListener('click',function(e){
    e.preventDefault();e.stopPropagation();
    var w=document.getElementById('sc-inline-wrap');
    if(w){w.remove();scheduleSave();}
  });
  wrap.appendChild(delBtn);

  /* Canvas */
  var canvas=document.createElement('canvas');
  canvas.style.cssText='width:100%;display:block;border-radius:8px;';
  wrap.appendChild(canvas);

  /* PDF image snapshot */
  var pdfImg=document.createElement('img');
  pdfImg.id='sc-chart-pdf-img';
  pdfImg.style.cssText='width:100%;height:auto;display:none;';
  wrap.appendChild(pdfImg);

  table.parentNode.insertBefore(wrap,table.nextSibling);scheduleSave();

  /* Draw after layout */
  setTimeout(function(){
    var W=Math.max(300, wrap.clientWidth-32);
    var CHART_H=Math.max(480, Math.round(W*0.52));
    var dpr=Math.max(window.devicePixelRatio||1,4);
    canvas.width=Math.round(W*dpr);
    canvas.height=Math.round(CHART_H*dpr);
    canvas.style.width=W+'px';
    canvas.style.height=CHART_H+'px';
    try{
      _draw3DPie(canvas,customers,W,CHART_H,dpr);
    }catch(err){
      console.error('Chart error:',err);
    }
    setTimeout(function(){
      pdfImg.src=canvas.toDataURL('image/png');scheduleSave();
    },400);
  },150);
}

function _draw3DPie(canvas,customers,W,H,dpr){
  var ctx=canvas.getContext('2d');
  dpr=dpr||Math.max(window.devicePixelRatio||1,4);
  ctx.scale(dpr,dpr);
  /* White background */
  ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,W,H);
  ctx.imageSmoothingEnabled=true; ctx.imageSmoothingQuality='high';

  var PAL=[
    [[250,204,21],[160,120,0]],
    [[251,146,60],[180,70,8]],
    [[34,211,238],[10,140,160]],
    [[244,114,182],[180,40,110]],
    [[56,189,248],[10,120,200]],
    [[0,210,140],[0,140,90]],
    [[168,85,247],[100,30,190]],
    [[239,68,68],[160,20,20]],
  ];

  var total=customers.reduce(function(s,c){return s+c.avg;},0);
  if(!total)return;
  var maxVal=Math.max.apply(null,customers.map(function(c){return c.avg;}));

  /* ── TOP LEGEND ─────────────────────────────────────────────────── */
  var LNM=22, LVAL=20, LDOT=10, LGAP=32, LROWH=52, LTOP=12;
  ctx.font='800 '+LNM+'px Segoe UI';
  var items=customers.map(function(c,i){
    var dup=customers.slice(0,i).filter(function(x){return x.avg===c.avg;}).length;
    var off=dup===0?0:dup%2===1?22:-22;
    var cl=function(v){return Math.min(255,Math.max(0,v+off));};
    var col=PAL[i%PAL.length];
    var top=[cl(col[0][0]),cl(col[0][1]),cl(col[0][2])];
    var sid=[cl(col[1][0]),cl(col[1][1]),cl(col[1][2])];
    var nm=c.name;
    var nmW=ctx.measureText(nm).width;
    var valW=ctx.measureText(c.avg+'%').width;
    return {name:nm,val:c.avg,top:top,sid:sid,itemW:LDOT*2+8+Math.max(nmW,valW)+LGAP};
  });

  /* Pack into rows */
  var rows=[[]],rowW=[0];
  items.forEach(function(it){
    var ri=rows.length-1;
    if(rowW[ri]+it.itemW>W-10&&rows[ri].length>0){rows.push([]);rowW.push(0);ri++;}
    rows[ri].push(it); rowW[ri]+=it.itemW;
  });
  var legendH=rows.length*LROWH+LTOP+8;

  /* Draw legend */
  rows.forEach(function(row,ri){
    var usedW=row.reduce(function(s,it){return s+it.itemW;},0);
    var x=(W-usedW)/2;
    var y=LTOP+ri*LROWH;
    row.forEach(function(it){
      var r=it.top[0],g=it.top[1],b=it.top[2];
      /* Dot */
      ctx.save();
      ctx.beginPath();ctx.arc(x+LDOT,y+LROWH/2,LDOT,0,Math.PI*2);
      ctx.fillStyle='rgb('+r+','+g+','+b+')';
      ctx.shadowColor='rgba('+r+','+g+','+b+',.35)';ctx.shadowBlur=6;
      ctx.fill();ctx.restore();
      /* Name — bold black */
      ctx.save();
      ctx.shadowBlur=0;ctx.shadowColor='transparent';
      ctx.fillStyle='#000000';
      ctx.textAlign='left';ctx.textBaseline='bottom';
      ctx.font='800 '+LNM+'px Segoe UI';
      ctx.fillText(it.name, x+LDOT*2+8, y+LROWH/2+1);
      /* Value — bold black */
      ctx.font='800 '+LVAL+'px Segoe UI';
      ctx.fillStyle='#111111';
      ctx.textBaseline='top';
      ctx.fillText(it.val+'%', x+LDOT*2+8, y+LROWH/2+3);
      ctx.restore();
      x+=it.itemW;
    });
  });

  /* ── PIE GEOMETRY ───────────────────────────────────────────────── */
  var PIE_X=W*0.50;
  var PIE_Y=legendH+(H-legendH)*0.54;
  var RX=W*0.26, RY=RX*0.38;
  var MAX_H=55, MIN_H=14;

  var ang=-Math.PI*0.65;
  var slices=items.map(function(it,i){
    var sweep=(it.val/total)*Math.PI*2;
    var h=MIN_H+(it.val/maxVal)*(MAX_H-MIN_H);
    var s={name:it.name,val:it.val,sa:ang,ea:ang+sweep,ma:ang+sweep*0.5,
           top:it.top,sid:it.sid,h:h};
    ang+=sweep; return s;
  });

  var order=slices.slice().sort(function(a,b){return Math.sin(a.ma)-Math.sin(b.ma);});
  var ep=function(a,dy){return{x:PIE_X+Math.cos(a)*RX,y:PIE_Y+Math.sin(a)*RY+(dy||0)};};
  var arc=function(sa,ea,dy){
    var steps=Math.max(16,Math.ceil(Math.abs(ea-sa)*16)),pts=[];
    for(var t=0;t<=steps;t++){var a=sa+(ea-sa)*t/steps;pts.push(ep(a,dy));}
    return pts;
  };

  /* Shadow */
  ctx.save();
  var shd=ctx.createRadialGradient(PIE_X,PIE_Y+8,0,PIE_X,PIE_Y+8,RX*.9);
  shd.addColorStop(0,'rgba(0,0,0,.12)');shd.addColorStop(1,'rgba(0,0,0,0)');
  ctx.fillStyle=shd;
  ctx.beginPath();ctx.ellipse(PIE_X,PIE_Y+10,RX*.85,RY*.5,0,0,Math.PI*2);
  ctx.fill();ctx.restore();

  /* Side walls */
  order.forEach(function(s){
    [s.sa,s.ea].forEach(function(a){
      var po=ep(a,0),pt=ep(a,-s.h);
      ctx.save();ctx.beginPath();
      ctx.moveTo(PIE_X,PIE_Y-s.h);ctx.lineTo(pt.x,pt.y);
      ctx.lineTo(po.x,po.y);ctx.lineTo(PIE_X,PIE_Y);ctx.closePath();
      var r=s.sid[0],g=s.sid[1],b=s.sid[2];
      var wg=ctx.createLinearGradient(PIE_X,PIE_Y-s.h,PIE_X,PIE_Y);
      wg.addColorStop(0,'rgba('+Math.min(r+20,255)+','+Math.min(g+20,255)+','+Math.min(b+20,255)+',.92)');
      wg.addColorStop(1,'rgba('+r+','+g+','+b+',.80)');
      ctx.fillStyle=wg;ctx.fill();
      ctx.strokeStyle='rgba(0,0,0,.1)';ctx.lineWidth=.5;ctx.stroke();
      ctx.restore();
    });
  });

  /* Outer curved side */
  order.forEach(function(s){
    var r=s.sid[0],g=s.sid[1],b=s.sid[2];
    var bot=arc(s.sa,s.ea,0),top=arc(s.sa,s.ea,-s.h);
    ctx.save();ctx.beginPath();
    bot.forEach(function(p,i){i===0?ctx.moveTo(p.x,p.y):ctx.lineTo(p.x,p.y);});
    for(var t=top.length-1;t>=0;t--)ctx.lineTo(top[t].x,top[t].y);
    ctx.closePath();
    var sg=ctx.createLinearGradient(PIE_X,PIE_Y-s.h,PIE_X,PIE_Y+RY);
    sg.addColorStop(0,'rgba('+Math.min(r+35,255)+','+Math.min(g+35,255)+','+Math.min(b+35,255)+',.95)');
    sg.addColorStop(1,'rgba('+Math.max(r-15,0)+','+Math.max(g-15,0)+','+Math.max(b-15,0)+',.85)');
    ctx.fillStyle=sg;ctx.fill();
    ctx.strokeStyle='rgba(0,0,0,.06)';ctx.lineWidth=.5;ctx.stroke();
    ctx.restore();
  });

  /* Top face */
  order.forEach(function(s){
    var r=s.top[0],g=s.top[1],b=s.top[2];
    var top=arc(s.sa,s.ea,-s.h);
    ctx.save();ctx.beginPath();
    ctx.moveTo(PIE_X,PIE_Y-s.h);
    top.forEach(function(p){ctx.lineTo(p.x,p.y);});
    ctx.closePath();
    var tg=ctx.createRadialGradient(PIE_X,PIE_Y-s.h,0,PIE_X,PIE_Y-s.h,RX);
    tg.addColorStop(0,'rgba('+Math.min(r+75,255)+','+Math.min(g+75,255)+','+Math.min(b+75,255)+',.98)');
    tg.addColorStop(.5,'rgba('+r+','+g+','+b+',.95)');
    tg.addColorStop(1,'rgba('+Math.max(r-35,0)+','+Math.max(g-35,0)+','+Math.max(b-35,0)+',.9)');
    ctx.fillStyle=tg;ctx.fill();
    var hl=ctx.createLinearGradient(PIE_X-RX*.4,PIE_Y-s.h-RY,PIE_X,PIE_Y-s.h);
    hl.addColorStop(0,'rgba(255,255,255,.28)');hl.addColorStop(1,'rgba(255,255,255,0)');
    ctx.fillStyle=hl;ctx.fill();
    ctx.strokeStyle='rgba(255,255,255,.22)';ctx.lineWidth=1;ctx.stroke();
    ctx.restore();
  });

  /* % labels — drawn in their own pass AFTER all top faces, so a
     neighboring slice can never paint over another slice's label
     (this was why thin slices like a 20% wedge looked cut off) */
  order.forEach(function(s){
    var sweepDeg=(s.ea-s.sa)*180/Math.PI;
    var fsize=sweepDeg<30?16:22;
    var lx=PIE_X+Math.cos(s.ma)*RX*.6;
    var ly=(PIE_Y-s.h)+Math.sin(s.ma)*RY*.55;
    ctx.save();
    ctx.font='bold '+fsize+'px Segoe UI';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.shadowColor='rgba(0,0,0,.8)';ctx.shadowBlur=3;
    ctx.fillStyle='white';ctx.fillText(s.val+'%',lx,ly);
    ctx.restore();
  });
}


function insertChartIntoNote(){var canvas=document.getElementById('scorecardChart');if(!canvas||canvas.style.display==='none')return;var img=document.createElement('img');img.src=canvas.toDataURL('image/png');img.style.cssText='max-width:100%;max-height:140mm;width:auto;height:auto;border-radius:8px;margin:12px 0;display:block;';var cont=document.getElementById('content');cont.appendChild(img);cont.appendChild(document.createElement('p'));scheduleSave();closeChartModal();}


/* ── Onboarding status cycling ── */

var OB_STATUSES=['pending','inprogress','hold','completed'];

var OB_LABELS={pending:'Pending',inprogress:'In Progress',hold:'Hold',completed:'Completed'};

function obUpdateProgress(wrap){
  if(!wrap)return;
  var tbody=wrap.querySelector('tbody');
  if(!tbody)return;

  // Get all unique customer groups
  var allCgs=Array.from(new Set(Array.from(tbody.querySelectorAll('[data-cg]')).map(function(r){return r.getAttribute('data-cg');})));

  // Update per-customer progress bars (each group's bar above its table)
  allCgs.forEach(function(cg){
    var cgRows=Array.from(tbody.querySelectorAll('[data-cg="'+cg+'"]'));
    var cgBtns=[];
    cgRows.forEach(function(r){var b=r.querySelector('.ob-status-btn');if(b)cgBtns.push(b);});
    var cgDone=cgBtns.filter(function(b){return b.classList.contains('completed');}).length;
    var cgPct=cgBtns.length?Math.round(cgDone/cgBtns.length*100):0;
    var firstRow=cgRows[0];
    var prev=firstRow?firstRow.previousElementSibling:null;
    while(prev){
      var fill=prev.querySelector('.ob-progress-fill');
      var lbl=prev.querySelector('.ob-progress-pct');
      if(fill){fill.style.width=cgPct+'%';if(lbl)lbl.textContent=cgPct+'%';break;}
      prev=prev.previousElementSibling;
    }
  });

  // Update main top progress bar based on FIRST customer group only (cg=0)
  var firstCgRows=Array.from(tbody.querySelectorAll('[data-cg="0"]'));
  var firstBtns=[];
  firstCgRows.forEach(function(r){var b=r.querySelector('.ob-status-btn');if(b)firstBtns.push(b);});
  var firstDone=firstBtns.filter(function(b){return b.classList.contains('completed');}).length;
  var firstPct=firstBtns.length?Math.round(firstDone/firstBtns.length*100):0;
  var mainFill=wrap.querySelector('.ob-progress-bar-wrap .ob-progress-fill');
  var mainLbl=wrap.querySelector('.ob-progress-bar-wrap .ob-progress-pct');
  if(mainFill)mainFill.style.width=firstPct+'%';
  if(mainLbl)mainLbl.textContent=firstPct+'%';
}

document.getElementById('content').addEventListener('click',function(e){var badge=e.target.closest('.ob-status-btn');if(!badge)return;e.preventDefault();e.stopPropagation();var cur=badge.dataset.status||'pending';var next=OB_STATUSES[(OB_STATUSES.indexOf(cur)+1)%OB_STATUSES.length];badge.dataset.status=next;badge.className='ob-status-btn '+next;badge.textContent=OB_LABELS[next];obUpdateProgress(badge.closest('.ob-wrap'));scheduleSave();});


function obAddCustomerGroup(btn){var thColor=btn.getAttribute('data-thcolor'),rowEven=btn.getAttribute('data-roweven');var steps=JSON.parse(btn.getAttribute('data-steps').replace(/&quot;/g,'"'));var stepCount=steps.length,tbody=btn.closest('tr').parentNode,addRow=btn.closest('tr');var allCgIdxSet=new Set(Array.from(tbody.querySelectorAll('[data-cg]')).map(r=>r.getAttribute('data-cg')));var cgIdx=allCgIdxSet.size;var newRows=[];var thS='background:'+thColor+'dd;color:white;padding:9px 16px;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.8px;';for(var si=0;si<stepCount;si++){var bg=(si%2===1)?'background:'+rowEven+';':'background:#fff;';var tr=document.createElement('tr');tr.setAttribute('data-cg',cgIdx);if(si===0){var ns='background:'+thColor+'22;border-left:4px solid '+thColor+';font-weight:700;font-size:14px;color:'+thColor+';padding:12px 18px;vertical-align:middle;text-align:center;min-width:160px;';var nameTd=document.createElement('td');nameTd.rowSpan=stepCount;nameTd.style.cssText=ns;nameTd.setAttribute('contenteditable','true');nameTd.textContent='';tr.appendChild(nameTd);}var stepTd=document.createElement('td');stepTd.style.cssText=bg+'padding:11px 16px;font-size:13px;';var stepName=(steps&&steps[si])?steps[si]:'Step '+(si+1);stepTd.innerHTML='<span style="display:inline-flex;align-items:center;gap:10px;"><span style="width:10px;height:10px;border-radius:50%;flex-shrink:0;background:'+thColor+';display:inline-block;"></span><span style="color:#1e293b;font-weight:500;">'+stepName+'</span></span>';tr.appendChild(stepTd);var statusTd=document.createElement('td');statusTd.style.cssText=bg+'padding:11px 14px;text-align:center;';statusTd.innerHTML='<span class="ob-status-btn pending" data-status="pending">Pending</span>';tr.appendChild(statusTd);

/* date cells with calendar picker */

['Target Date','Completion Date'].forEach(function(){var dtTd=document.createElement('td');dtTd.style.cssText=bg+'padding:8px 10px;text-align:center;';dtTd.innerHTML='<div class="date-cell-wrap"><span class="date-val" contenteditable="true" data-ph="DD-Mon-YYYY"></span><button class="date-cal-btn" onclick="openDatePicker(event,this)" type="button" title="Pick date">📅</button></div>';tr.appendChild(dtTd);});

var briefTd=document.createElement('td');briefTd.className='ob-comment-cell';briefTd.style.cssText=bg+'padding:11px 14px;';briefTd.setAttribute('contenteditable','true');tr.appendChild(briefTd);newRows.push(tr);}var sep=document.createElement('tr');sep.innerHTML='<td colspan="6" style="height:5px;background:linear-gradient(90deg,'+thColor+'22,transparent);border-bottom:2px solid '+thColor+'33;"></td>';newRows.push(sep);var repeatHdr=document.createElement('tr');repeatHdr.setAttribute('data-repeat-hdr','1');repeatHdr.innerHTML='<th style="'+thS+'min-width:160px;">Customer Name</th><th style="'+thS+'">Check Point</th><th style="'+thS+'text-align:center;width:140px;">Status</th><th style="'+thS+'min-width:120px;text-align:center;">Target Date</th><th style="'+thS+'min-width:130px;text-align:center;">Completion Date</th><th style="'+thS+'min-width:160px;">Comment</th>';var spacerRow=document.createElement('tr');spacerRow.innerHTML='<td colspan="6" class="ob-spacer-row" style="height:32px;background:#ffffff;border:none;padding:0;"></td>';
var progressRow=document.createElement('tr');
progressRow.className='ob-progress-row';
progressRow.innerHTML='<td colspan="6" style="padding:0;background:#f8fafc;border-bottom:1px solid #e2e8f0;"><div style="display:flex;justify-content:space-between;align-items:center;padding:8px 18px 4px;"><span style="font-size:12px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.5px;">Overall Progress</span><span class="ob-progress-pct" style="font-size:13px;font-weight:700;color:'+thColor+';">0%</span></div><div style="margin:0 18px 10px;height:8px;background:#e2e8f0;border-radius:10px;overflow:hidden;"><div class="ob-progress-fill" style="height:100%;width:0%;background:'+thColor+';border-radius:10px;transition:width .4s;"></div></div></td>';
tbody.insertBefore(spacerRow,addRow);tbody.insertBefore(progressRow,addRow);tbody.insertBefore(repeatHdr,addRow);newRows.forEach(function(r){tbody.insertBefore(r,addRow);});if(newRows[0])newRows[0].querySelector('[contenteditable]').focus();scheduleSave();}


/* ── Serialization cards ── */

var SI_COLORS=['#038dbd','#0d152b','#6366f1','#059669','#d97706','#dc2626','#7c3aed','#0891b2'];


var SI_COLORS=['#038dbd','#0d152b','#6366f1','#059669','#d97706','#dc2626','#7c3aed','#0891b2'];

function siTimelineCardHTML(idx,col){
  var siColorMap={'#038dbd':['rgba(3,141,189,.08)','#bae6fd'],'#0d152b':['rgba(13,21,43,.06)','#c7d2fe'],'#6366f1':['rgba(99,102,241,.08)','#c4b5fd'],'#059669':['rgba(5,150,105,.08)','#a7f3d0'],'#d97706':['rgba(217,119,6,.08)','#fde68a'],'#dc2626':['rgba(220,38,38,.08)','#fecaca'],'#7c3aed':['rgba(124,58,237,.08)','#ddd6fe'],'#0891b2':['rgba(8,145,178,.08)','#a5f3fc']};
  var cv=siColorMap[col]||['rgba(3,141,189,.08)','#bae6fd'];
  var dateInline='<span class="si-tl-date-wrap"><span class="date-val" contenteditable="true" data-ph="DD-Mon-YYYY"></span><button class="date-cal-btn" onclick="openDatePicker(event,this)" type="button" title="Pick date">📅</button></span>';
  var titleBox='<div class="si-tl-issue-title" contenteditable="true" data-ph="Issue Title"></div>';
  var tl='<div class="si-tl">'
    +'<div class="si-tl-step" data-step="created"><div class="si-tl-node" style="--step-col:#3b82f6;">📝</div><div class="si-tl-line"></div><div class="si-tl-content"><div class="si-tl-head"><span class="si-tl-title">Issue Created</span>'+dateInline+'</div><div class="si-tl-sublabel">Issue Description</div><div class="si-tl-body" contenteditable="true" data-ph="Describe the issue..."></div></div></div>'
    +'<div class="si-tl-step" data-step="rootcause"><div class="si-tl-node" style="--step-col:#f59e0b;">\u26A0</div><div class="si-tl-line"></div><div class="si-tl-content"><div class="si-tl-head"><span class="si-tl-title">Root Cause</span></div><div class="si-tl-body" contenteditable="true" data-ph="Identify root cause..."></div></div></div>'
    +'<div class="si-tl-step" data-step="actions"><div class="si-tl-node" style="--step-col:#22c55e;">\u2705</div><div class="si-tl-line"></div><div class="si-tl-content"><div class="si-tl-head"><span class="si-tl-title">Action Items / Comments</span></div><div class="si-action-list">'+siActionRowHTML()+'</div></div></div>'
    +'<div class="si-tl-step si-tl-step-last" data-step="resolution"><div class="si-tl-node" style="--step-col:#8b5cf6;">🎯</div><div class="si-tl-content"><div class="si-tl-head"><span class="si-tl-title">Resolution Date:</span>'+dateInline+'</div></div></div>'
    +'</div>';
  return '<div class="si-card si-tl-card" style="--si-col:'+col+';--si-col-bg:'+cv[0]+';--si-col-light:'+cv[1]+';" data-si-card="'+idx+'">'
    +'<div class="si-card-body">'+titleBox+tl+'</div>'
    +'<div class="si-card-foot"><div class="si-status-group"><span class="si-status-label">Status</span><span class="ob-status-btn pending" data-status="pending" title="Click to change">Pending</span></div>'
    +'<button class="si-del-card" onclick="siDeleteCard(this)" type="button">&#128465; Delete</button></div></div>';
}

function siActionRowHTML(){
  return '<div class="si-action-row"><span class="si-action-text" contenteditable="true" data-ph="Action item..."></span></div>';
}
function siActionInsertAfter(afterRow){
  var list=afterRow?afterRow.parentElement:null;
  if(!list)return null;
  var tmp=document.createElement('div');
  tmp.innerHTML=siActionRowHTML();
  var newRow=tmp.firstElementChild;
  if(afterRow.nextSibling)list.insertBefore(newRow,afterRow.nextSibling);
  else list.appendChild(newRow);
  return newRow;
}
function siTimelineDelAction(btn){var row=btn.closest('.si-action-row');if(!row)return;var list=row.parentElement;row.remove();if(list&&list.querySelectorAll('.si-action-row').length===0){var tmp=document.createElement('div');tmp.innerHTML=siActionRowHTML();list.appendChild(tmp.firstElementChild);}scheduleSave();}
function siTimelineToggleAction(checkbox){
  var row=checkbox.closest('.si-action-row');
  if(row){row.classList.toggle('si-action-done',checkbox.checked);scheduleSave();}
}

function siAddCard(btn){var grid=btn.closest('.si-cards-grid');if(!grid)return;var existing=grid.querySelectorAll('.si-card').length;var col=SI_COLORS[existing%SI_COLORS.length];var tmp=document.createElement('div');tmp.innerHTML=siTimelineCardHTML(existing,col);var card=tmp.firstElementChild;grid.insertBefore(card,btn);scheduleSave();}

function siDeleteCard(btn){var card=btn.closest('.si-card');var grid=card?card.closest('.si-cards-grid'):null;if(!grid)return;if(grid.querySelectorAll('.si-card').length<=1){alert('Cannot delete the last card.');return;}if(!confirm('Delete this issue card?'))return;card.remove();scheduleSave();}
/* SI_DRAG_RUNTIME_INJECT_V1 */
function siCardHasHandle(card){
  for (var i=0;i<card.children.length;i++){
    if (card.children[i].classList && card.children[i].classList.contains('si-drag-handle')) return true;
  }
  return false;
}
function siInitDragHandles(){
  document.querySelectorAll('.si-cards-grid .si-card').forEach(function(card){
    if (siCardHasHandle(card)) return;
    var handle = document.createElement('div');
    handle.className = 'si-drag-handle';
    handle.title = 'Drag to reorder';
    handle.textContent = '\u22ee\u22ee';
    card.insertBefore(handle, card.firstChild);
  });
}
function siGetDragAfterElement(grid, x, y){
  var cards = Array.prototype.slice.call(grid.querySelectorAll('.si-card:not(.si-dragging)'));
  for (var i=0;i<cards.length;i++){
    var box = cards[i].getBoundingClientRect();
    var inRow = (y >= box.top && y <= box.bottom);
    if (inRow){
      if (x < box.left + box.width/2) return cards[i];
    } else if (y < box.top) {
      return cards[i];
    }
  }
  return null;
}
(function(){
  if (document.readyState !== 'loading') siInitDragHandles();
  else document.addEventListener('DOMContentLoaded', siInitDragHandles);

  var _siMo = new MutationObserver(function(){ siInitDragHandles(); });
  function _siStartObserver(){ _siMo.observe(document.body, {childList:true, subtree:true}); }
  if (document.body) _siStartObserver();
  else document.addEventListener('DOMContentLoaded', _siStartObserver);

  document.addEventListener('mousedown', function(e){
    var handle = e.target.closest('.si-drag-handle');
    if(!handle) return;
    var card = handle.closest('.si-card');
    if(card) card.setAttribute('draggable','true');
  });
  document.addEventListener('mouseup', function(){
    document.querySelectorAll('.si-card[draggable="true"]').forEach(function(c){
      if(!c.classList.contains('si-dragging')) c.removeAttribute('draggable');
    });
  });
  document.addEventListener('dragstart', function(e){
    var card = e.target.closest('.si-card');
    if(!card || card.getAttribute('draggable')!=='true'){ return; }
    try{ e.dataTransfer.setData('text/plain','si-card'); }catch(err){}
    e.dataTransfer.effectAllowed='move';
    card.classList.add('si-dragging');
    window._siDragCard = card;
  });
  document.addEventListener('dragend', function(e){
    var card = e.target.closest('.si-card');
    if(card){ card.classList.remove('si-dragging'); card.removeAttribute('draggable'); }
    window._siDragCard = null;
    if (typeof scheduleSave === 'function') scheduleSave();
  });
  var _siDragPending = false;
  var _siLastEvt = null;
  function _siApplyDragMove(){
    _siDragPending = false;
    var e = _siLastEvt;
    if(!e) return;
    var grid = e.target.closest('.si-cards-grid');
    if(!grid || !window._siDragCard) return;
    var afterEl = siGetDragAfterElement(grid, e.clientX, e.clientY);
    var addBtn = grid.querySelector('.si-add-btn');
    if(afterEl === window._siDragCard || (afterEl && afterEl.previousElementSibling === window._siDragCard)) return;
    if(afterEl){ grid.insertBefore(window._siDragCard, afterEl); }
    else if(addBtn){ grid.insertBefore(window._siDragCard, addBtn); }
    else { grid.appendChild(window._siDragCard); }
  }
  /* SI_DRAG_UX_V1: throttled with requestAnimationFrame for smoother movement */
  document.addEventListener('dragover', function(e){
    var grid = e.target.closest('.si-cards-grid');
    if(!grid || !window._siDragCard) return;
    e.preventDefault();
    _siLastEvt = e;
    if(!_siDragPending){
      _siDragPending = true;
      requestAnimationFrame(_siApplyDragMove);
    }
  });
})();

function siAddFieldRow(btn){

  var rows=btn.closest('.si-multi-rows'); if(!rows) return;

  var ph=(rows.querySelector('.si-field-val')||{getAttribute:function(){return '';}}).getAttribute('data-ph')||'';

  var row=document.createElement('div'); row.className='si-multi-row';

  var val=document.createElement('div'); val.className='si-field-val';

  val.setAttribute('contenteditable','true'); val.setAttribute('data-ph',ph);

  var del=document.createElement('button'); del.className='si-field-del-row';

  del.type='button'; del.title='Remove row'; del.innerHTML='&#10005;';

  // onclick handled by delegation

  row.appendChild(val); row.appendChild(del);

  rows.insertBefore(row,btn); val.focus(); scheduleSave();

}

function siDelFieldRow(btn){

  var row=btn.closest('.si-multi-row'); if(!row) return;

  var rows=row.parentNode;

  if(rows.querySelectorAll('.si-multi-row').length<=1) return;

  row.remove(); scheduleSave();

}

function siInitDeleteBtns(){

  /* Only wrap existing single si-field-val into si-multi-rows structure.

     Delete button already exists in .si-card-foot — don't add a second one. */

  document.querySelectorAll('.si-field').forEach(function(field){

    var val=field.querySelector('.si-field-val');

    if(!val||val.closest('.si-multi-row')) return;

    var wrap=document.createElement('div'); wrap.className='si-multi-rows';

    var row=document.createElement('div'); row.className='si-multi-row';

    val.parentNode.insertBefore(wrap,val);

    row.appendChild(val); wrap.appendChild(row);

    var addBtn=document.createElement('button');

    addBtn.className='si-field-add-row'; addBtn.type='button';

    addBtn.innerHTML='<span style="font-size:14px;line-height:1;">&#43;</span> Add';

    // onclick handled by delegation

    wrap.appendChild(addBtn);

  });

}


/* ── Row Action Bar (Clone / Delete) for all template rows ── */

var _rabRow = null, _rabTbody = null, _rabType = null;

var _rabHideTimer = null;


function rabShow(row, tbody, type, anchorRect){

  _rabRow = row; _rabTbody = tbody; _rabType = type;

  var bar = document.getElementById('rowActionBar');

  bar.style.display = 'flex';

  var barW = 240;

  var centre = anchorRect.left + anchorRect.width / 2;

  var leftPos = Math.min(Math.max(centre - barW / 2, 8), window.innerWidth - barW - 8);

  bar.style.top  = (anchorRect.bottom + 4) + 'px';

  bar.style.left = leftPos + 'px';

  bar.querySelector('.rab-clone').style.display = (type === 'si') ? 'none' : '';

  bar.querySelector('.rab-del-cust').style.display = (type === 'ob') ? '' : 'none';

}

function rabHide(){ 

  _rabRow = null; _rabTbody = null; _rabType = null;

  var bar = document.getElementById('rowActionBar');

  if(bar) bar.style.display = 'none';

}


function rabClone(){

  if(!_rabRow) return;

  if(_rabType==='exec'){

    var clone=_rabRow.cloneNode(true);

    var cc=clone.querySelector('.exec-row-content'); if(cc) cc.textContent='';

    var addBtn=_rabRow.closest('.exec-section').querySelector('.exec-add-row');

    var ns=_rabRow.nextElementSibling;

    _rabRow.parentNode.insertBefore(clone,(ns&&ns!==addBtn)?ns:addBtn);

    if(cc) cc.focus(); scheduleSave(); rabHide(); return;

  }

  if(!_rabTbody) return;

  if(_rabType === 'sc'){

    var cloneTr = _rabRow.cloneNode(true);

    cloneTr.className = 'sc-row';

    Array.from(cloneTr.querySelectorAll('.sc-val')).forEach(function(v){ v.textContent=''; });

    var next = _rabRow.nextElementSibling;

    if(next) _rabTbody.insertBefore(cloneTr, next);

    else _rabTbody.appendChild(cloneTr);

    cloneTr.querySelectorAll('.sc-val').forEach(function(el){

      el.setAttribute('contenteditable','true');

      el.addEventListener('keydown',function(e){

        if(!/^[0-9.]$/.test(e.key)&&!['Backspace','Delete','ArrowLeft','ArrowRight','Tab'].includes(e.key)){if(e.key!=='Enter')e.preventDefault();}

        if(e.key==='Enter'){e.preventDefault();el.blur();}

      });

      el.addEventListener('input', scheduleSave);

    });

    rabRecolorRows(_rabTbody);

    var nameTd = cloneTr.querySelector('.sc-name-td');

    if(nameTd){ nameTd.setAttribute('contenteditable','true'); nameTd.focus(); }

    scheduleSave();

  } else if(_rabType === 'ob'){

    var cgIdx = _rabRow.getAttribute('data-cg');

    if(!cgIdx){ rabHide(); return; }

    var cloneTrOb = document.createElement('tr');

    cloneTrOb.setAttribute('data-cg', cgIdx);

    Array.from(_rabRow.cells).forEach(function(srcTd){

      if(srcTd.classList.contains('sc-name-td')) return;

      cloneTrOb.appendChild(srcTd.cloneNode(true));

    });

    var next2 = _rabRow.nextElementSibling;

    if(next2) _rabTbody.insertBefore(cloneTrOb, next2);

    else _rabTbody.appendChild(cloneTrOb);

    var fg = _rabTbody.querySelector('[data-cg="'+cgIdx+'"]');

    if(fg){

      var nt = fg.querySelector('.sc-name-td') || Array.from(fg.cells).find(function(c){ return c.rowSpan>1; });

      if(nt) nt.rowSpan = (nt.rowSpan||1)+1;

    }

    obUpdateProgress(_rabRow.closest('.ob-wrap'));

    scheduleSave();

  }

  rabHide();

}


function rabDelete(){

  if(!_rabRow) return;

  if(_rabType==='exec'){

    var sec=_rabRow.closest('.exec-section');

    if(sec&&sec.querySelectorAll('.exec-row').length<=1){ rabHide(); return; }

    _rabRow.remove(); scheduleSave(); rabHide(); return;

  }

  if(!_rabTbody) return;

  if(_rabType === 'sc'){

    if(_rabTbody.querySelectorAll('tr').length <= 1){ alert('Cannot delete the last row.'); return; }

    if(!confirm('Delete this row?')) return;

    _rabTbody.removeChild(_rabRow);

    rabRecolorRows(_rabTbody);

    scheduleSave();

  } else if(_rabType === 'ob'){

    var obWrap = _rabRow.closest('.ob-wrap');

    var cgIdx = _rabRow.getAttribute('data-cg');

    if(!cgIdx){ rabHide(); return; }

    var groupRows = Array.from(_rabTbody.querySelectorAll('[data-cg="'+cgIdx+'"]'));

    var allGroups = new Set(Array.from(_rabTbody.querySelectorAll('[data-cg]')).map(function(r){ return r.getAttribute('data-cg'); }));

    if(groupRows.length > 1){

      if(!confirm('Delete this row?')) return;

      var hasNameTd = _rabRow.querySelector('.sc-name-td') || Array.from(_rabRow.cells).find(function(td){ return td.rowSpan>1; });

      if(hasNameTd){

        var nextGRow = _rabRow.nextElementSibling;

        while(nextGRow && nextGRow.getAttribute('data-cg') !== cgIdx) nextGRow = nextGRow.nextElementSibling;

        if(nextGRow){ var mv = hasNameTd.cloneNode(true); mv.rowSpan = Math.max(1,(hasNameTd.rowSpan||1)-1); nextGRow.insertBefore(mv, nextGRow.firstChild); }

      } else {

        var fg2 = _rabTbody.querySelector('[data-cg="'+cgIdx+'"]');

        if(fg2){ var nt2 = fg2.querySelector('.sc-name-td') || Array.from(fg2.cells).find(function(td){ return td.rowSpan>1; }); if(nt2 && nt2.rowSpan>1) nt2.rowSpan--; }

      }

      _rabTbody.removeChild(_rabRow);

    } else {

      if(allGroups.size <= 1){ alert('Cannot delete the last customer.'); return; }

      if(!confirm('Delete entire customer and all their rows?')) return;

      var lastRow = groupRows[groupRows.length-1];

      var sep = lastRow.nextElementSibling;

      if(sep && !sep.getAttribute('data-cg') && sep.querySelector('td[colspan]')) _rabTbody.removeChild(sep);

      var prevRow = groupRows[0].previousElementSibling;

      if(prevRow && prevRow.getAttribute('data-repeat-hdr')) _rabTbody.removeChild(prevRow);

      groupRows.forEach(function(r){ _rabTbody.removeChild(r); });

    }

    if(obWrap) obUpdateProgress(obWrap);

    scheduleSave();

  }

  rabHide();

}


function rabRecolorRows(tbody){

  var ACCENTS = ['#8b5cf6','#059669','#d97706','#2563eb','#db2777','#0891b2','#7c3aed'];

  Array.from(tbody.rows).forEach(function(row, ri){

    var acc = ACCENTS[ri % ACCENTS.length];

    var nm = row.querySelector('.sc-name-td');

    if(nm){ nm.style.borderLeft = '4px solid '+acc; nm.style.color = acc; }

  });

}


document.getElementById('content').addEventListener('click', function(e){

  /* ── Exec row actions (delegation — more reliable than inline onclick in contenteditable) ── */

  var execDel = e.target.closest('.exec-del-row');

  var execAdd = e.target.closest('.exec-add-row');

  var execHnd = e.target.closest('.exec-row-handle');

  if(execDel){ e.preventDefault(); e.stopPropagation(); execDelRow(execDel); return; }

  if(execAdd){ e.preventDefault(); e.stopPropagation(); execAddRow(execAdd); return; }

  if(execHnd){ e.preventDefault(); e.stopPropagation(); execHandleRab(execHnd); return; }

  /* ── SI field row actions ── */

  var siFieldDel = e.target.closest('.si-field-del-row');

  var siFieldAdd = e.target.closest('.si-field-add-row');

  if(siFieldDel){ e.preventDefault(); e.stopPropagation(); siDelFieldRow(siFieldDel); return; }

  if(siFieldAdd){ e.preventDefault(); e.stopPropagation(); siAddFieldRow(siFieldAdd); return; }

  if(e.target.closest('button,.ob-status-btn,.sc-val,.si-card-title,.si-field-val')) return;

  var row = e.target.closest('tr');

  if(!row) { rabHide(); return; }

  var scRow = row.closest('.sc-table') ? row : null;

  var obRow = (!scRow && row.getAttribute('data-cg')) ? row : null;

  if(scRow && scRow.closest('tbody')){

    var rect = scRow.getBoundingClientRect();

    rabShow(scRow, scRow.closest('tbody'), 'sc', rect);

  } else if(obRow && obRow.closest('tbody')){

    var rect2 = obRow.getBoundingClientRect();

    rabShow(obRow, obRow.closest('tbody'), 'ob', rect2);

  } else {

    rabHide();

  }

});


document.addEventListener('click', function(e){

  var bar = document.getElementById('rowActionBar');

  if(bar && bar.style.display==='flex' && !bar.contains(e.target)){

    if(!e.target.closest('#content')) rabHide();

  }

});


function rabDeleteCustomer(){

  if(!_rabRow || !_rabTbody) return;

  var obWrap = _rabRow.closest('.ob-wrap');

  var cgIdx = _rabRow.getAttribute('data-cg');

  if(!cgIdx){ rabHide(); return; }

  var allGroups = new Set(Array.from(_rabTbody.querySelectorAll('[data-cg]')).map(function(r){ return r.getAttribute('data-cg'); }));

  if(allGroups.size <= 1){ alert('Cannot delete the last customer.'); return; }

  if(!confirm('Delete entire Customer and all their checklist rows?')) return;

  var groupRows = Array.from(_rabTbody.querySelectorAll('[data-cg="'+cgIdx+'"]'));

  var lastRow = groupRows[groupRows.length-1];

  var sep = lastRow.nextElementSibling;

  if(sep && !sep.getAttribute('data-cg') && sep.querySelector('td[colspan]')) _rabTbody.removeChild(sep);

  var prevRow = groupRows[0].previousElementSibling;

  if(prevRow && prevRow.getAttribute('data-repeat-hdr')) _rabTbody.removeChild(prevRow);

  groupRows.forEach(function(r){ _rabTbody.removeChild(r); });

  if(obWrap) obUpdateProgress(obWrap);

  scheduleSave();

  rabHide();

}


/* ── Hide page title when content has a template banner ── */

/* ── Date Picker ────────────────────────────────────────────────────────── */

var _dpTarget = null, _dpYear = new Date().getFullYear(), _dpMonth = new Date().getMonth();

var _dpView = 'days';   // 'days' | 'months' | 'years'

var _dpYearPage = Math.floor(new Date().getFullYear() / 12);

var DP_MONTHS_S = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

var DP_MONTHS_L = ['January','February','March','April','May','June','July','August','September','October','November','December'];


function openDatePicker(e, btn){

  e.preventDefault(); e.stopPropagation();

  var wrap = btn.closest('.date-cell-wrap') || btn.closest('.si-date-wrap') || btn.closest('.si-tl-date-wrap');

  _dpTarget = wrap ? wrap.querySelector('.date-val') : null;

  if(!_dpTarget) return;

  _dpView = 'days';

  var existing = (_dpTarget.textContent||'').trim();

  if(existing){

    var p = existing.split('-');

    if(p.length===3){ var mi=DP_MONTHS_S.indexOf(p[1]); if(mi>=0){_dpMonth=mi; _dpYear=parseInt(p[2])||_dpYear;} }

  } else { _dpMonth=new Date().getMonth(); _dpYear=new Date().getFullYear(); }

  _dpYearPage = Math.floor(_dpYear / 12);

  _dpRender();

  var dp = document.getElementById('datePickerPopup');

  dp.style.display = 'block';

  var r = btn.getBoundingClientRect();

  var left = r.left, top = r.bottom + 5;

  if(left + 256 > window.innerWidth) left = window.innerWidth - 260;

  if(top + 320 > window.innerHeight) top = r.top - 325;

  dp.style.left = Math.max(4,left) + 'px';

  dp.style.top  = Math.max(4,top)  + 'px';

}


function _dpRender(){

  if(_dpView==='months') { _dpRenderMonths(); return; }

  if(_dpView==='years')  { _dpRenderYears();  return; }

  _dpRenderDays();

}


/* ── Day view ── */

function _dpRenderDays(){

  var dp=document.getElementById('datePickerPopup');

  var first=new Date(_dpYear,_dpMonth,1).getDay();

  var days=new Date(_dpYear,_dpMonth+1,0).getDate();

  var today=new Date();

  var h='<div class="dp-hdr">';

  h+='<button class="dp-nav" data-dir="-1">&#8249;</button>';

  h+='<div class="dp-hdr-lbl">';

  h+='<button class="dp-hdr-month-btn">'+DP_MONTHS_L[_dpMonth]+'</button>';

  h+='<button class="dp-hdr-year-btn">'+_dpYear+'</button>';

  h+='</div>';

  h+='<button class="dp-nav" data-dir="1">&#8250;</button></div>';

  h+='<div class="dp-dow-row">';

  ['S','M','T','W','T','F','S'].forEach(function(d,i){

    h+='<div class="dp-dow'+(i===0||i===6?' dp-wknd':'')+'">'+d+'</div>';

  });

  h+='</div><div class="dp-grid">';

  for(var i=0;i<first;i++) h+='<div></div>';

  for(var d=1;d<=days;d++){

    var isTd=(d===today.getDate()&&_dpMonth===today.getMonth()&&_dpYear===today.getFullYear());

    var dow=(first+d-1)%7;

    var cls='dp-day'+(isTd?' dp-today':'')+(dow===0||dow===6?' dp-wknd':'');

    h+='<button class="'+cls+'" data-day="'+d+'">'+d+'</button>';

  }

  h+='</div><div class="dp-footer"><button class="dp-today-btn">Today</button></div>';

  dp.innerHTML=h;

  dp.querySelectorAll('.dp-nav').forEach(function(b){

    b.addEventListener('click',function(e){e.stopPropagation();dpNav(parseInt(this.getAttribute('data-dir')));});

  });

  dp.querySelectorAll('.dp-day').forEach(function(b){

    b.addEventListener('click',function(e){e.stopPropagation();dpSelect(parseInt(this.getAttribute('data-day')));});

  });

  dp.querySelector('.dp-hdr-month-btn').addEventListener('click',function(e){

    e.stopPropagation(); _dpView='months'; _dpRender();

  });

  dp.querySelector('.dp-hdr-year-btn').addEventListener('click',function(e){

    e.stopPropagation(); _dpYearPage=Math.floor(_dpYear/12); _dpView='years'; _dpRender();

  });

  var tb=dp.querySelector('.dp-today-btn');

  if(tb) tb.addEventListener('click',function(e){e.stopPropagation();dpSelectToday();});

}


/* ── Month view ── */

function _dpRenderMonths(){

  var dp=document.getElementById('datePickerPopup');

  var h='<div class="dp-hdr">';

  h+='<button class="dp-nav" data-dir="-1">&#8249;</button>';

  h+='<div class="dp-hdr-lbl"><button class="dp-hdr-year-btn">'+_dpYear+'</button></div>';

  h+='<button class="dp-nav" data-dir="1">&#8250;</button></div>';

  h+='<div class="dp-month-grid">';

  DP_MONTHS_S.forEach(function(m,i){

    h+='<button class="dp-month-btn'+(i===_dpMonth?' dp-sel':'')+'" data-mi="'+i+'">'+m+'</button>';

  });

  h+='</div>';

  h+='<div class="dp-footer"><button class="dp-back-btn">&#8592; Back</button></div>';

  dp.innerHTML=h;

  dp.querySelectorAll('.dp-nav').forEach(function(b){

    b.addEventListener('click',function(e){

      e.stopPropagation();

      _dpYear+=parseInt(this.getAttribute('data-dir'));

      _dpRender();

    });

  });

  dp.querySelectorAll('.dp-month-btn').forEach(function(b){

    b.addEventListener('click',function(e){

      e.stopPropagation();

      _dpMonth=parseInt(this.getAttribute('data-mi'));

      _dpView='days'; _dpRender();

    });

  });

  dp.querySelector('.dp-hdr-year-btn').addEventListener('click',function(e){

    e.stopPropagation(); _dpYearPage=Math.floor(_dpYear/12); _dpView='years'; _dpRender();

  });

  dp.querySelector('.dp-back-btn').addEventListener('click',function(e){

    e.stopPropagation(); _dpView='days'; _dpRender();

  });

}


/* ── Year view ── */

function _dpRenderYears(){

  var dp=document.getElementById('datePickerPopup');

  var startY=_dpYearPage*12;

  var endY=startY+11;

  var h='<div class="dp-hdr">';

  h+='<button class="dp-nav" data-dir="-1">&#8249;</button>';

  h+='<div class="dp-hdr-lbl"><span class="dp-year-range">'+startY+' – '+endY+'</span></div>';

  h+='<button class="dp-nav" data-dir="1">&#8250;</button></div>';

  h+='<div class="dp-year-grid">';

  for(var y=startY;y<=endY;y++){

    h+='<button class="dp-year-btn'+(y===_dpYear?' dp-sel':'')+'" data-yr="'+y+'">'+y+'</button>';

  }

  h+='</div>';

  h+='<div class="dp-footer"><button class="dp-back-btn">&#8592; Back</button></div>';

  dp.innerHTML=h;

  dp.querySelectorAll('.dp-nav').forEach(function(b){

    b.addEventListener('click',function(e){

      e.stopPropagation();

      _dpYearPage+=parseInt(this.getAttribute('data-dir'));

      _dpRender();

    });

  });

  dp.querySelectorAll('.dp-year-btn').forEach(function(b){

    b.addEventListener('click',function(e){

      e.stopPropagation();

      _dpYear=parseInt(this.getAttribute('data-yr'));

      _dpView='months'; _dpRender();

    });

  });

  dp.querySelector('.dp-back-btn').addEventListener('click',function(e){

    e.stopPropagation(); _dpView='days'; _dpRender();

  });

}


function dpNav(dir){

  _dpMonth+=dir;

  if(_dpMonth<0){_dpMonth=11;_dpYear--;}

  if(_dpMonth>11){_dpMonth=0;_dpYear++;}

  _dpRender();

}


function dpSelect(day){

  if(!_dpTarget) return;

  var formatted=String(day).padStart(2,'0')+'-'+DP_MONTHS_S[_dpMonth]+'-'+_dpYear;

  _dpTarget.textContent=formatted;

  _dpTarget.dispatchEvent(new Event('input',{bubbles:true}));

  dpClose(); scheduleSave();

}


function dpSelectToday(){

  var t=new Date();

  _dpMonth=t.getMonth(); _dpYear=t.getFullYear();

  dpSelect(t.getDate());

}


function dpClose(){

  var dp=document.getElementById('datePickerPopup');

  if(dp) dp.style.display='none';

  _dpTarget=null;

}


document.addEventListener('click',function(e){

  var dp=document.getElementById('datePickerPopup');

  if(dp&&dp.style.display==='block'&&!dp.contains(e.target)&&!e.target.classList.contains('date-cal-btn')){

    dpClose();

  }

});


/* ── Page Drag & Drop Reorder ──────────────────────────────────────────── */

(function(){

  var _dragSrc    = null;

  var _canDrag    = false;   // true only when mousedown was on .drag-handle


  function getList(){ return document.getElementById('pagesList'); }


  function clearDragStyles(){

    var list = getList(); if(!list) return;

    list.querySelectorAll('.page-link').forEach(function(el){

      el.classList.remove('dragging','drag-over-top','drag-over-bot');

    });

  }


  function initPageDrag(){

    var list = getList(); if(!list) return;


    // Track whether the drag started from the handle

    list.addEventListener('mousedown', function(e){

      _canDrag = !!e.target.closest('.drag-handle');

    });

    // Reset on mouseup (in case user didn't drag)

    document.addEventListener('mouseup', function(){ _canDrag = false; });


    list.addEventListener('dragstart', function(e){

      if(!_canDrag){ e.preventDefault(); return; }

      var link = e.target.closest('.page-link');

      if(!link){ e.preventDefault(); return; }

      _dragSrc = link;

      link.classList.add('dragging');

      e.dataTransfer.effectAllowed = 'move';

      e.dataTransfer.setData('text/plain', link.getAttribute('data-note-id'));

    });


    list.addEventListener('dragend', function(){

      clearDragStyles();

      _dragSrc = null;

      _canDrag = false;

    });


    list.addEventListener('dragover', function(e){

      e.preventDefault();

      e.dataTransfer.dropEffect = 'move';

      var target = e.target.closest('.page-link');

      clearDragStyles();

      if(_dragSrc) _dragSrc.classList.add('dragging');

      if(!target || target === _dragSrc) return;

      var rect = target.getBoundingClientRect();

      var mid  = rect.top + rect.height / 2;

      if(e.clientY < mid) target.classList.add('drag-over-top');

      else                target.classList.add('drag-over-bot');

    });


    list.addEventListener('dragleave', function(e){

      if(!list.contains(e.relatedTarget)) clearDragStyles();

    });


    list.addEventListener('drop', function(e){

      e.preventDefault();

      var target = e.target.closest('.page-link');

      if(!target || !_dragSrc || target === _dragSrc){ clearDragStyles(); return; }


      var rect = target.getBoundingClientRect();

      var insertBefore = (e.clientY < rect.top + rect.height / 2);


      if(insertBefore) list.insertBefore(_dragSrc, target);

      else {

        var next = target.nextElementSibling;

        if(next) list.insertBefore(_dragSrc, next);

        else     list.appendChild(_dragSrc);

      }

      clearDragStyles();


      // Persist new order

      var newOrder = Array.from(list.querySelectorAll('.page-link'))

        .map(function(el){ return parseInt(el.getAttribute('data-note-id')); });

      fetch('/reorder-notes',{

        method:'POST',

        headers:{'Content-Type':'application/json'},

        body:JSON.stringify({note_ids:newOrder})

      });

    });

  }


  document.addEventListener('DOMContentLoaded', initPageDrag);

  if(document.readyState !== 'loading') initPageDrag();

})();

(function(){
  var _dragSrc = null;
  var _canDrag = false;
  function getList(){ return document.getElementById('sectionsList'); }
  function clearDragStyles(){
    var list = getList(); if(!list) return;
    list.querySelectorAll('.section-row').forEach(function(el){
      el.classList.remove('dragging','drag-over-top','drag-over-bot');
    });
  }
  function initSectionDrag(){
    var list = getList(); if(!list) return;
    list.addEventListener('mousedown', function(e){
      _canDrag = !!e.target.closest('.section-drag-handle');
    });
    document.addEventListener('mouseup', function(){ _canDrag = false; });
    list.addEventListener('dragstart', function(e){
      var row = e.target.closest('.section-row');
      if(!row || !_canDrag){ e.preventDefault(); return; }
      _dragSrc = row;
      row.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', row.getAttribute('data-section-id'));
    });
    list.addEventListener('dragend', function(){ clearDragStyles(); _dragSrc = null; _canDrag = false; });
    list.addEventListener('dragover', function(e){
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      var target = e.target.closest('.section-row');
      clearDragStyles();
      if(_dragSrc) _dragSrc.classList.add('dragging');
      if(!target || target === _dragSrc) return;
      var rect = target.getBoundingClientRect();
      var mid  = rect.top + rect.height / 2;
      if(e.clientY < mid) target.classList.add('drag-over-top');
      else                target.classList.add('drag-over-bot');
    });
    list.addEventListener('dragleave', function(e){
      if(!list.contains(e.relatedTarget)) clearDragStyles();
    });
    list.addEventListener('drop', function(e){
      e.preventDefault();
      var target = e.target.closest('.section-row');
      if(!target || !_dragSrc || target === _dragSrc){ clearDragStyles(); return; }
      var rect = target.getBoundingClientRect();
      var insertBefore = (e.clientY < rect.top + rect.height / 2);
      if(insertBefore) list.insertBefore(_dragSrc, target);
      else {
        var next = target.nextElementSibling;
        if(next) list.insertBefore(_dragSrc, next);
        else     list.appendChild(_dragSrc);
      }
      clearDragStyles();
      var newOrder = Array.from(list.querySelectorAll('.section-row'))
        .map(function(el){ return parseInt(el.getAttribute('data-section-id')); });
      fetch('/reorder-sections',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({section_ids:newOrder})
      });
    });
  }
  document.addEventListener('DOMContentLoaded', initSectionDrag);
  if(document.readyState !== 'loading') initSectionDrag();
})();


/* ── Image Resize ────────────────────────────────────────────────────────── */

(function(){

  var _selImg  = null;

  var _dragDir = null;

  var _startX=0, _startY=0, _startW=0, _startH=0;


  function overlay(){ return document.getElementById('imgResizeOverlay'); }

  function sizeLabel(){ return document.getElementById('imgSizeLabel'); }


  function positionOverlay(){

    if(!_selImg) return;

    var o = overlay(); if(!o) return;

    var r = _selImg.getBoundingClientRect();

    o.style.left   = r.left   + 'px';

    o.style.top    = r.top    + 'px';

    o.style.width  = r.width  + 'px';

    o.style.height = r.height + 'px';

    o.style.display = 'block';

    var sl = sizeLabel();

    if(sl) sl.textContent = Math.round(r.width) + ' × ' + Math.round(r.height) + 'px';

  }


  function selectImg(img){

    if(_selImg) _selImg.classList.remove('img-selected');

    _selImg = img;

    img.classList.add('img-selected');

    positionOverlay();

  }


  function deselect(){

    if(_selImg) _selImg.classList.remove('img-selected');

    _selImg = null;

    var o = overlay(); if(o) o.style.display = 'none';

  }


  // Resize drag

  function onHandleDown(e){

    if(!_selImg) return;

    e.preventDefault(); e.stopPropagation();

    _dragDir = this.getAttribute('data-d');

    _startX  = e.clientX; _startY = e.clientY;

    _startW  = _selImg.offsetWidth; _startH = _selImg.offsetHeight;

    document.addEventListener('mousemove', onDrag);

    document.addEventListener('mouseup',   onDragEnd);

  }


  function onDrag(e){

    if(!_selImg || !_dragDir) return;

    var dx = e.clientX - _startX;

    var dy = e.clientY - _startY;

    var ar = _startH / _startW;

    var nw = _startW, nh = _startH;


    if(_dragDir.indexOf('r') !== -1) nw = Math.max(40, _startW + dx);

    if(_dragDir.indexOf('l') !== -1) nw = Math.max(40, _startW - dx);

    if(_dragDir.indexOf('b') !== -1) nh = Math.max(30, _startH + dy);

    if(_dragDir.indexOf('t') !== -1 && _dragDir.length > 1) nh = Math.max(30, _startH - dy);


    // Lock aspect ratio on corner drags

    if(_dragDir.length === 2 && _dragDir !== 'tc' && _dragDir !== 'bc'){

      nh = Math.round(nw * ar);

    }

    _selImg.style.width  = Math.round(nw) + 'px';

    _selImg.style.height = _dragDir === 'ml' || _dragDir === 'mr' ? 'auto' : Math.round(nh) + 'px';

    positionOverlay();

  }


  function onDragEnd(){

    _dragDir = null;

    document.removeEventListener('mousemove', onDrag);

    document.removeEventListener('mouseup',   onDragEnd);

    scheduleSave();

  }


  // Toolbar actions (exposed globally)

  window.imgTbAlign = function(align){

    if(!_selImg) return;

    _selImg.style.float   = align === 'left' ? 'left' : align === 'right' ? 'right' : '';

    _selImg.style.display = align === 'center' ? 'block' : '';

    _selImg.style.margin  = align === 'center' ? '12px auto'

                           : align === 'left'   ? '8px 16px 8px 0'

                           :                      '8px 0 8px 16px';

    positionOverlay(); scheduleSave();

  };

  window.imgTbSize = function(pct){

    if(!_selImg) return;

    _selImg.style.width  = pct + '%';

    _selImg.style.height = 'auto';

    positionOverlay(); scheduleSave();

  };

  window.imgTbDelete = function(){

    if(!_selImg) return;

    if(!confirm('Delete this image?')) return;

    _selImg.parentNode.removeChild(_selImg);

    deselect(); scheduleSave();

  };


  // Wire up everything after DOM ready

  function init(){

    var content = document.getElementById('content');

    var o = overlay(); if(!content || !o) return;


    // Attach handle mousedown

    o.querySelectorAll('.img-rh').forEach(function(h){

      h.addEventListener('mousedown', onHandleDown);

    });


    // Click image to select

    content.addEventListener('click', function(e){

      if(e.target.tagName === 'IMG'){

        if(e.target.closest('.cover-overlay-wrap')){return;}

        selectImg(e.target);

        e.stopPropagation();

      }

    });


    // Click outside → deselect

    document.addEventListener('click', function(e){

      if(!_selImg) return;

      var o2 = overlay();

      if(e.target.tagName !== 'IMG' && !(o2 && o2.contains(e.target))){

        deselect();

      }

    });


    // Reposition overlay on content scroll

    content.addEventListener('scroll', function(){

      if(_selImg) positionOverlay();

    });


    // Reposition on window resize/scroll

    window.addEventListener('scroll', function(){ if(_selImg) positionOverlay(); });

    window.addEventListener('resize', function(){ if(_selImg) positionOverlay(); });

  }


  document.addEventListener('DOMContentLoaded', init);

  if(document.readyState !== 'loading') init();

})();


/* ── Executive Summary row actions ── */

function execHandleRab(btn){

  var row=btn.closest('.exec-row'); if(!row) return;

  rabShow(row, null, 'exec', row.getBoundingClientRect());

}

function execDelRow(btn){

  var section=btn.closest('.exec-section');

  var rows=section?section.querySelectorAll('.exec-row'):[];

  if(rows.length<=1) return;

  btn.closest('.exec-row').remove(); scheduleSave();

}

function execAddRow(btn){

  var section=btn.closest('.exec-section');

  var ph=section.querySelector('.exec-row-content')?section.querySelector('.exec-row-content').getAttribute('data-ph')||'Enter item...':'Enter item...';

  var row=document.createElement('div'); row.className='exec-row';

  row.innerHTML='<span class="exec-row-handle" title="Options">&#8942;</span>'

    +'<div class="exec-row-content" contenteditable="true" data-ph="'+ph+'"></div>'

    +'<button class="exec-del-row" type="button" title="Delete">&#10005;</button>';

  section.insertBefore(row,btn);

  row.querySelector('.exec-row-content').focus(); scheduleSave();

}

/* ── Serialization Issues field rows ── */

function siAddFieldRow(btn){

  var rows=btn.closest('.si-multi-rows'); if(!rows) return;

  var ph=(rows.querySelector('.si-field-val')||{getAttribute:function(){return '';}}).getAttribute('data-ph')||'';

  var row=document.createElement('div'); row.className='si-multi-row';

  var val=document.createElement('div'); val.className='si-field-val';

  val.setAttribute('contenteditable','true'); val.setAttribute('data-ph',ph);

  var del=document.createElement('button'); del.className='si-field-del-row';

  del.type='button'; del.title='Remove row'; del.innerHTML='&#10005;';

  // onclick handled by delegation

  row.appendChild(val); row.appendChild(del);

  rows.insertBefore(row,btn); val.focus(); scheduleSave();

}

function siDelFieldRow(btn){

  var row=btn.closest('.si-multi-row'); if(!row) return;

  var rows=row.parentNode;

  if(rows.querySelectorAll('.si-multi-row').length<=1) return;

  row.remove(); scheduleSave();

}

function siInitDeleteBtns(){

  document.querySelectorAll('.si-field').forEach(function(field){

    var val=field.querySelector('.si-field-val');

    if(!val||val.closest('.si-multi-row')) return;

    var wrap=document.createElement('div'); wrap.className='si-multi-rows';

    var row=document.createElement('div'); row.className='si-multi-row';

    val.parentNode.insertBefore(wrap,val);

    row.appendChild(val); wrap.appendChild(row);

    var addBtn=document.createElement('button');

    addBtn.className='si-field-add-row'; addBtn.type='button';

    addBtn.innerHTML='&#43; Add'; // onclick handled by delegation

    wrap.appendChild(addBtn);

  });

}


(function(){
  function initActionEnterKey(){
    var content=document.getElementById('content');
    if(!content||content._actionEnterWired)return;
    content._actionEnterWired=true;
    content.addEventListener('keydown',function(e){
      if(e.key==='Enter'&&!e.shiftKey&&e.target.classList&&e.target.classList.contains('si-action-text')){
        e.preventDefault();
        var row=e.target.closest('.si-action-row');
        var newRow=siActionInsertAfter(row);
        if(newRow){
          var txt=newRow.querySelector('.si-action-text');
          if(txt)txt.focus();
        }
        scheduleSave();
      }
    });
  }
  document.addEventListener('DOMContentLoaded',initActionEnterKey);
  if(document.readyState!=='loading')initActionEnterKey();
})();
(function(){
  function initSafeEnterBoxes(){
    var content=document.getElementById('content');
    if(!content||content._safeEnterWired)return;
    content._safeEnterWired=true;
    var BOX_SELECTOR='.exec-content,.si-tl-body';
    content.addEventListener('keydown',function(e){
      if(e.key!=='Enter')return;
      var sel=window.getSelection();
      if(!sel||!sel.rangeCount)return;
      var node=sel.getRangeAt(0).startContainer;
      var el=(node.nodeType===3)?node.parentElement:node;
      var box=el&&el.closest?el.closest(BOX_SELECTOR):null;
      if(!box)return;
      var li=el.closest?el.closest('li'):null;
      if(li&&box.contains(li))return;
      e.preventDefault();
      e.stopPropagation();
      document.execCommand('insertLineBreak');
      scheduleSave();
    });
  }
  document.addEventListener('DOMContentLoaded',initSafeEnterBoxes);
  if(document.readyState!=='loading')initSafeEnterBoxes();
})();
function addOverlayTextBox(){
  var overlay=document.querySelector('.cover-drop-layer');
  if(overlay){
    overlay.style.pointerEvents='all';
    placeOverlayTextBox('50','40');
  } else {
    var tw=document.createElement('div');
    tw.style.cssText='position:relative;display:inline-block;margin:8px 0;vertical-align:top;min-width:120px;';
    tw.setAttribute('contenteditable','false');
    var tb2=document.createElement('div');
    tb2.style.cssText='display:none;position:absolute;top:-40px;left:0;background:#1e293b;border-radius:8px;padding:4px 8px;gap:4px;white-space:nowrap;z-index:25;align-items:center;box-shadow:0 4px 12px rgba(0,0,0,.5);';
    tb2.innerHTML='<button title="Bold" style="background:none;border:none;color:white;cursor:pointer;font-weight:900;font-size:14px;padding:2px 7px;" onmousedown="event.preventDefault();oTxtBold(this)">B</button>'
     +'<button title="Italic" style="background:none;border:none;color:white;cursor:pointer;font-style:italic;font-size:14px;padding:2px 7px;" onmousedown="event.preventDefault();oTxtItalic(this)"><i>I</i></button>'
     +'<button title="Bigger" style="background:none;border:none;color:white;cursor:pointer;font-size:13px;padding:2px 7px;" onmousedown="event.preventDefault();oTxtSize(this,2)">A+</button>'
     +'<button title="Smaller" style="background:none;border:none;color:white;cursor:pointer;font-size:13px;padding:2px 7px;" onmousedown="event.preventDefault();oTxtSize(this,-2)">A-</button>'
     +'<input type="color" value="#1e293b" style="width:24px;height:24px;border:none;background:none;cursor:pointer;padding:0;vertical-align:middle;" oninput="oTxtColor(this)" onchange="oTxtColor(this)">'
     +'<button title="Delete" style="background:#dc2626;border:none;color:white;cursor:pointer;font-size:11px;font-weight:700;padding:3px 8px;border-radius:4px;margin-left:4px;" onmousedown="event.preventDefault();oTxtDel(this)">&#10005;</button>';
    var txt2=document.createElement('div');
    txt2.contentEditable='true';
    txt2.style.cssText='min-width:120px;min-height:36px;font-size:18px;font-weight:600;color:#1e293b;border:none;border-radius:8px;padding:10px 14px;outline:none;cursor:text;line-height:1.5;';
    txt2.textContent='Type text here';
    txt2.addEventListener('input',scheduleSave);
    txt2.addEventListener('focus',function(){tb2.style.display='flex';txt2.style.outline='2px dashed #8b3dff';});
    txt2.addEventListener('blur',function(){setTimeout(function(){if(!tw.contains(document.activeElement)){tb2.style.display='none';txt2.style.outline='none';}},200);});
    document.addEventListener('mousedown',function(e){if(!tw.contains(e.target)){tb2.style.display='none';txt2.style.outline='none';}});
    tb2._txtRef=txt2;
    tw.appendChild(tb2);tw.appendChild(txt2);
    var contentEl=document.getElementById('content');
    contentEl.appendChild(tw);
    contentEl.appendChild(document.createElement('p'));
    txt2.focus();
    var r2=document.createRange();r2.selectNodeContents(txt2);
    var s2=window.getSelection();s2.removeAllRanges();s2.addRange(r2);
    scheduleSave();
  }
}
function placeOverlayTextBox(xPct,yPct,initText,initSize,initColor,initBold){
  var overlay=document.querySelector('.cover-drop-layer');
  if(!overlay)return;
  overlay.style.pointerEvents='all';

  var tw=document.createElement('div');
  tw.style.cssText='position:absolute;left:'+xPct+'%;top:'+yPct+'%;z-index:16;min-width:120px;min-height:36px;';tw.setAttribute('data-cov','txt');

  // toolbar
  var tb=document.createElement('div');
  tb.style.cssText='display:none;position:absolute;top:-40px;left:0;background:#1e293b;border-radius:8px;padding:4px 8px;gap:4px;white-space:nowrap;z-index:25;align-items:center;box-shadow:0 4px 12px rgba(0,0,0,.5);';
  tb.innerHTML=
    '<button title="Bold" style="background:none;border:none;color:white;cursor:pointer;font-weight:900;font-size:14px;padding:2px 7px;border-radius:4px;line-height:1;" onmousedown="event.preventDefault();oTxtBold(this)">B</button>'
   +'<button title="Italic" style="background:none;border:none;color:white;cursor:pointer;font-style:italic;font-size:14px;padding:2px 7px;border-radius:4px;line-height:1;" onmousedown="event.preventDefault();oTxtItalic(this)"><i>I</i></button>'
   +'<button title="Bigger" style="background:none;border:none;color:white;cursor:pointer;font-size:13px;padding:2px 7px;" onmousedown="event.preventDefault();oTxtSize(this,2)">A+</button>'
   +'<button title="Smaller" style="background:none;border:none;color:white;cursor:pointer;font-size:13px;padding:2px 7px;" onmousedown="event.preventDefault();oTxtSize(this,-2)">A-</button>'
   +'<input type="color" title="Color" value="'+(initColor||'#ffffff')+'" style="width:24px;height:24px;border:none;background:none;cursor:pointer;padding:0;vertical-align:middle;" oninput="oTxtColor(this)" onchange="oTxtColor(this)">'
   +'<span style="width:1px;height:18px;background:rgba(255,255,255,.3);display:inline-block;margin:0 3px;vertical-align:middle;"></span>'
   +'<button title="Move" style="background:none;border:none;color:#94a3b8;cursor:move;font-size:14px;padding:2px 6px;" class="otxt-move-btn">&#10021;</button>'
   +'<button title="Delete" style="background:#dc2626;border:none;color:white;cursor:pointer;font-size:11px;font-weight:700;padding:3px 8px;border-radius:4px;margin-left:4px;" onmousedown="event.preventDefault();oTxtDel(this)">&#10005;</button>';

  var txt=document.createElement('div');
  txt.contentEditable='true';
  txt.style.cssText='min-width:120px;min-height:36px;color:'+(initColor||'#ffffff')+';font-size:'+(initSize||'20')+'px;font-weight:'+(initBold||'600')+';font-family:Segoe UI,sans-serif;outline:none;padding:6px 10px;cursor:text;text-shadow:1px 1px 6px rgba(0,0,0,.7);line-height:1.4;white-space:pre-wrap;min-width:80px;';
  txt.textContent=initText||'Text';

  txt.setAttribute('onfocus',"var tb=this.parentNode.children[0];if(tb)tb.style.display='flex';this.style.outline='2px dashed rgba(139,61,255,.7)';");
  txt.setAttribute('onblur',"var el=this,p=this.parentNode;setTimeout(function(){if(!p.contains(document.activeElement)){var tb=p.children[0];if(tb)tb.style.display='none';el.style.outline='none';}},220);");
  document.addEventListener('mousedown',function(e){if(!tw.contains(e.target)){tb.style.display='none';txt.style.outline='none';}});
  txt.setAttribute('oninput','scheduleSave()');

  // drag via move button
  tb.addEventListener('mousedown',function(e){
    if(!e.target.classList.contains('otxt-move-btn'))return;
    e.preventDefault();
    var _ox=e.clientX,_oy=e.clientY;
    function mm(e){
      var wr=tw.closest('.cover-overlay-wrap');if(!wr)return;
      var rect=wr.getBoundingClientRect();
      tw.style.left=(parseFloat(tw.style.left)+(e.clientX-_ox)/rect.width*100)+'%';
      tw.style.top=(parseFloat(tw.style.top)+(e.clientY-_oy)/rect.height*100)+'%';
      _ox=e.clientX;_oy=e.clientY;
    }
    function mu(){document.removeEventListener('mousemove',mm);document.removeEventListener('mouseup',mu);scheduleSave();}
    document.addEventListener('mousemove',mm);document.addEventListener('mouseup',mu);
  });

  tb._txtRef=txt;tw.appendChild(tb);tw.appendChild(txt);
  overlay.appendChild(tw);
  setTimeout(function(){txt.focus();var r=document.createRange();r.selectNodeContents(txt);var s=window.getSelection();s.removeAllRanges();s.addRange(r);},50);
  scheduleSave();
}

/* COVER_OVERLAY_TXT_REHYDRATE_V1 */
function coverWireOverlayTextWrapper(tw){
  if (tw.dataset.covTxtWired === '1') return;
  var tb = tw.children[0];
  var txt = tw.children[1];
  if (!tb || !txt) return;
  tw.dataset.covTxtWired = '1';

  tb.addEventListener('mousedown', function(e){
    if(!e.target.classList.contains('otxt-move-btn')) return;
    e.preventDefault();
    var _ox=e.clientX,_oy=e.clientY;
    function mm(e){
      var wr=tw.closest('.cover-overlay-wrap'); if(!wr) return;
      var rect=wr.getBoundingClientRect();
      tw.style.left=(parseFloat(tw.style.left)+(e.clientX-_ox)/rect.width*100)+'%';
      tw.style.top=(parseFloat(tw.style.top)+(e.clientY-_oy)/rect.height*100)+'%';
      _ox=e.clientX;_oy=e.clientY;
    }
    function mu(){
      document.removeEventListener('mousemove',mm);
      document.removeEventListener('mouseup',mu);
      if (typeof scheduleSave === 'function') scheduleSave();
    }
    document.addEventListener('mousemove',mm);
    document.addEventListener('mouseup',mu);
  });

  document.addEventListener('mousedown', function(e){
    if(!tw.contains(e.target)){ tb.style.display='none'; txt.style.outline='none'; }
  });
}
function coverInitOverlayTextHandlers(){
  document.querySelectorAll('.cover-overlay-wrap [data-cov="txt"]').forEach(function(tw){
    coverWireOverlayTextWrapper(tw);
  });
}
if (document.readyState !== 'loading') coverInitOverlayTextHandlers();
else document.addEventListener('DOMContentLoaded', coverInitOverlayTextHandlers);

function oTxtBold(btn){var tb=btn.closest('div');var txt=tb?tb._txtRef:null;if(txt){txt.style.fontWeight=txt.style.fontWeight==='bold'?'normal':'bold';scheduleSave();}}
function oTxtItalic(btn){var tb=btn.closest('div');var txt=tb?tb._txtRef:null;if(txt){txt.style.fontStyle=txt.style.fontStyle==='italic'?'normal':'italic';scheduleSave();}}
function oTxtSize(btn,d){var tb=btn.closest('div');var txt=tb?tb._txtRef:null;if(txt){txt.style.fontSize=Math.max(8,parseInt(txt.style.fontSize||20)+d)+'px';scheduleSave();}}
function oTxtColor(inp){var tb=inp.closest('div');var txt=tb?tb._txtRef:null;if(txt){txt.style.color=inp.value;scheduleSave();}}
function oTxtDel(btn){var tb=btn.closest('div');var tw=tb?tb.parentNode:null;if(tw&&confirm('Delete text box?')){tw.remove();scheduleSave();}}

function updateTitleVisibility(){

  var content = document.getElementById('content');

  var title   = document.getElementById('title');

  if(!content || !title) return;

  var hasTemplate = content.querySelector(

    '.si-wrap, .ob-wrap, .sc-table, .si-banner, ' +

    '.ob-header-customer, .ob-header-vendor, .ob-header-mah, ' +

    '.exec-wrap, .agenda-wrap'

  );

  title.classList.toggle('hidden-title', !!hasTemplate);
  var addTxtBtn=document.getElementById('addTextOverlayBtn');
  if(addTxtBtn) addTxtBtn.style.display='inline-block';

}

document.addEventListener('DOMContentLoaded', function(){

  updateTitleVisibility();

  siInitDeleteBtns();

  /* Lock all date-val to calendar-only */

  document.querySelectorAll('.date-val').forEach(function(el){

    el.setAttribute('contenteditable','false');

    el.style.cursor='pointer';

    el.addEventListener('keydown',function(e){e.preventDefault();e.stopPropagation();});

    el.addEventListener('paste',function(e){e.preventDefault();});

  });

  new MutationObserver(function(muts){

    muts.forEach(function(m){

      m.addedNodes.forEach(function(n){

        if(n.nodeType===1){

          (n.querySelectorAll?n.querySelectorAll('.date-val'):[]).forEach(function(el){

            el.setAttribute('contenteditable','false');

            el.style.cursor='pointer';

            el.addEventListener('keydown',function(e){e.preventDefault();});

            el.addEventListener('paste',function(e){e.preventDefault();});

          });

        }

      });

    });

  }).observe(document.getElementById('content')||document.body,{childList:true,subtree:true});

});

(function(){ if(document.readyState !== 'loading'){ updateTitleVisibility(); siInitDeleteBtns(); } })();

</script></body>
<script>
var FA_IMPACT_LEVELS=['High','Medium','Low'];
function faImpactBadgeHTML(level){
  level=level||'High';
  return '<span class="fa-impact-badge" data-level="'+level+'" onclick="faToggleImpact(this)" title="Click to change impact level">'+level+' Impact</span>';
}
function faToggleImpact(badge){
  var cur=badge.getAttribute('data-level')||'High';
  var next=FA_IMPACT_LEVELS[(FA_IMPACT_LEVELS.indexOf(cur)+1)%FA_IMPACT_LEVELS.length];
  badge.setAttribute('data-level',next);
  badge.textContent=next+' Impact';
  if(typeof scheduleSave==='function')scheduleSave();
}
(function(){
  var _orig=window.faCardHTML;
  window.faCardHTML=function(color,title,impact,issue,imp,action){
    color=color||FA_COLORS[0];
    impact=(impact||'High').replace(' Impact','').trim();
    if(FA_IMPACT_LEVELS.indexOf(impact)===-1)impact='High';
    return '<div class="fa-card" data-fa-color="'+color+'">'
     +'<div class="fa-card-head" style="background:'+color+';">'
     +'<div class="fa-card-title" contenteditable="true">'+(title||'')+'</div>'
     +'<button class="fa-color-btn" onclick="faCycleColor(this)" title="Change colour"></button>'
     +'</div>'
     +'<div class="fa-card-inner">'
     +'<div class="fa-impact-row">'+faImpactBadgeHTML(impact)+'</div>'
     +'<div class="fa-section"><div class="fa-section-label">Issue</div>'
     +'<div class="fa-section-body" contenteditable="true" data-ph="Describe the issue...">'+(issue||'')+'</div></div>'
     +'<div class="fa-section"><div class="fa-section-label">Impact</div>'
     +'<div class="fa-section-body" contenteditable="true" data-ph="Describe the impact...">'+(imp||'')+'</div></div>'
     +'<div class="fa-section"><div class="fa-section-label">Action</div>'
     +'<div class="fa-section-body" contenteditable="true" data-ph="Describe the action...">'+(action||'')+'</div></div>'
     +'</div>'
     +'<button class="fa-del-card" onclick="faDelCard(this)">&#10005;</button>'
     +'</div>';
  };
})();
</script>
<script>
/* Focus Areas: impact badge helpers */
var FA_IMPACT_LEVELS = ['High','Medium','Low'];

function faImpactBadgeHTML(level){
  level = level || 'High';
  level = level.replace(' Impact','').trim();
  if(FA_IMPACT_LEVELS.indexOf(level) === -1) level = 'High';
  return '<span class="fa-impact-badge" data-level="'+level
    +'" onclick="faToggleImpact(this)" title="Click: High / Medium / Low">'
    +'<span class="fa-impact-icon"></span>'
    +level+' Impact</span>';
}

function faToggleImpact(badge){
  var cur  = badge.getAttribute('data-level') || 'High';
  var next = FA_IMPACT_LEVELS[(FA_IMPACT_LEVELS.indexOf(cur)+1) % FA_IMPACT_LEVELS.length];
  badge.setAttribute('data-level', next);
  var icon = badge.querySelector('.fa-impact-icon');
  /* rebuild inner content */
  badge.innerHTML = '<span class="fa-impact-icon"></span>'+next+' Impact';
  if(typeof scheduleSave==='function') scheduleSave();
}

/* Upgrade existing cards that still use old dot+text structure */
function faUpgradeImpactRows(){
  document.querySelectorAll('.fa-impact-row').forEach(function(row){
    /* Already upgraded */
    if(row.querySelector('.fa-impact-badge')) return;
    /* Old structure: .fa-impact-dot + .fa-impact-text */
    var txt = row.querySelector('.fa-impact-text');
    var level = txt ? txt.textContent.replace(' Impact','').trim() : 'High';
    if(FA_IMPACT_LEVELS.indexOf(level) === -1) level = 'High';
    row.innerHTML = faImpactBadgeHTML(level);
  });
}

/* Patch faCardHTML to always use new badge */
(function(){
  var _orig = typeof faCardHTML !== 'undefined' ? faCardHTML : null;
  window.faCardHTML = function(color,title,impact,issue,imp,action){
    color  = color  || (typeof FA_COLORS!=='undefined'?FA_COLORS[0]:'#2563eb');
    impact = (impact||'High').replace(' Impact','').trim();
    if(FA_IMPACT_LEVELS.indexOf(impact)===-1) impact='High';
    return '<div class="fa-card" data-fa-color="'+color+'">'
     +'<div class="fa-card-head" style="background:'+color+';">'
     +'<div class="fa-card-title" contenteditable="true">'+(title||'')+'</div>'
     +'<button class="fa-color-btn" onclick="faCycleColor(this)" title="Change colour"></button>'
     +'</div>'
     +'<div class="fa-card-inner">'
     +'<div class="fa-impact-row">'+faImpactBadgeHTML(impact)+'</div>'
     +'<div class="fa-section"><div class="fa-section-label">Issue</div>'
     +'<div class="fa-section-body" contenteditable="true" data-ph="Describe the issue...">'+(issue||'')+'</div></div>'
     +'<div class="fa-section"><div class="fa-section-label">Impact</div>'
     +'<div class="fa-section-body" contenteditable="true" data-ph="Describe the impact...">'+(imp||'')+'</div></div>'
     +'<div class="fa-section"><div class="fa-section-label">Action</div>'
     +'<div class="fa-section-body" contenteditable="true" data-ph="Describe the action...">'+(action||'')+'</div></div>'
     +'</div>'
     +'<button class="fa-del-card" onclick="faDelCard(this)">&#10005;</button>'
     +'</div>';
  };
})();

if(document.readyState!=='loading') faUpgradeImpactRows();
else document.addEventListener('DOMContentLoaded', faUpgradeImpactRows);
</script>
<script>
function clonePage(e, noteId){
  e.preventDefault();
  e.stopPropagation();
  fetch('/clone/'+noteId, {method:'POST'})
    .then(function(r){ if(r.redirected) location.href=r.url; else location.reload(); })
    .catch(function(){ location.reload(); });
}
</script>
<script>
function cloneSection(secId, secName){
  if(!confirm('Clone section "'+secName+'" and all its pages?')) return;
  fetch('/clone-section/'+secId, {method:'POST'})
    .then(function(r){ if(r.redirected) location.href=r.url; else location.reload(); })
    .catch(function(){ location.reload(); });
}
</script>



<script>
/* SI card header color — single definitive implementation */
(function(){
  var C={
    completed:'linear-gradient(135deg,#0c4a6e,#0369a1,#38bdf8)',
    inprogress:'linear-gradient(135deg,#78350f,#b45309,#fbbf24)',
    hold:'linear-gradient(135deg,#312e81,#4338ca,#a5b4fc)',
    pending:'linear-gradient(135deg,#1e3a8a,#1d4ed8,#60a5fa)'
  };
  function applyCard(card){
    var h=card.querySelector('.si-card-head');
    if(!h)return;
    var btn=card.querySelector('.ob-status-btn');
    var s=btn?(btn.getAttribute('data-status')||'pending').toLowerCase():'pending';
    if(!C[s])s='pending';
    h.removeAttribute('style');
    h.style.setProperty('background',C[s],'important');
  }
  function applyAll(){document.querySelectorAll('.si-card').forEach(applyCard);}
  [100,500,1500].forEach(function(t){setTimeout(applyAll,t);});
  document.addEventListener('click',function(e){
    var b=e.target.closest&&e.target.closest('.ob-status-btn');
    if(!b)return;
    var c=b.closest('.si-card');
    if(c)setTimeout(function(){
      applyCard(c);
      if(typeof scheduleSave==='function')scheduleSave();
    },80);
  },true);
})();
</script>
<script>
/* Restored: SI Enter key fix + duplicate body cleanup */
(function(){
  /* 1. Remove duplicate .si-tl-body on load */
  function removeDuplicateBodies(){
    document.querySelectorAll('.si-tl-content').forEach(function(c){
      var bodies=c.querySelectorAll('.si-tl-body');
      for(var i=1;i<bodies.length;i++) bodies[i].remove();
    });
  }

  /* 2. Enter in SI text areas = line break, not new element */
  function fixSiEnter(){
    var content=document.getElementById('content');
    if(!content||content._siEnterRestored)return;
    content._siEnterRestored=true;
    content.addEventListener('keydown',function(e){
      if(e.key!=='Enter'||e.shiftKey)return;
      var t=e.target;
      if(!t||!t.closest)return;
      if(!t.closest('.si-card'))return;
      if(t.classList&&t.classList.contains('si-action-text'))return;
      e.preventDefault(); e.stopPropagation();
      document.execCommand('insertLineBreak');
    },true);
  }

  if(document.readyState!=='loading'){removeDuplicateBodies();fixSiEnter();}
  else document.addEventListener('DOMContentLoaded',function(){removeDuplicateBodies();fixSiEnter();});
})();
</script>
<script>
/* SI: Enter=linebreak + remove duplicate boxes (restored v2) */
(function(){
  function removeDups(){
    document.querySelectorAll('.si-tl-content').forEach(function(c){
      var b=c.querySelectorAll('.si-tl-body');
      for(var i=1;i<b.length;i++)b[i].remove();
    });
  }
  function siEnter(){
    var el=document.getElementById('content');
    if(!el||el._siEnterV3)return;
    el._siEnterV3=true;
    el.addEventListener('keydown',function(e){
      if(e.key!=='Enter'||e.shiftKey)return;
      var t=e.target;
      if(!t||!t.closest||!t.closest('.si-card'))return;
      if(t.classList&&t.classList.contains('si-action-text'))return;
      e.preventDefault();e.stopPropagation();
      document.execCommand('insertLineBreak');
    },true);
  }
  if(document.readyState!=='loading'){removeDups();siEnter();}
  else document.addEventListener('DOMContentLoaded',function(){removeDups();siEnter();});
})();
/* EXEC_CONTENT_ENTER_FIX_V1 */
/* EXEC_CONTENT_ANTILIST_V1 */
(function(){
  function execUnwrapLists(el){
    if(!el.querySelector('ul,ol,li')) return false;
    var sel=window.getSelection();
    var range=null, startContainer=null, startOffset=0;
    if(sel && sel.rangeCount>0){
      range=sel.getRangeAt(0);
      startContainer=range.startContainer;
      startOffset=range.startOffset;
    }
    el.querySelectorAll('li').forEach(function(li){
      var div=document.createElement('div');
      while(li.firstChild) div.appendChild(li.firstChild);
      if(li.parentNode) li.parentNode.insertBefore(div, li);
      li.remove();
    });
    el.querySelectorAll('ul,ol').forEach(function(listEl){
      while(listEl.firstChild) listEl.parentNode.insertBefore(listEl.firstChild, listEl);
      listEl.remove();
    });
    if(range && startContainer && el.contains(startContainer)){
      try{
        var newRange=document.createRange();
        var maxOffset = startContainer.nodeType===3 ? (startContainer.textContent||'').length : startContainer.childNodes.length;
        newRange.setStart(startContainer, Math.min(startOffset, maxOffset));
        newRange.collapse(true);
        sel.removeAllRanges();
        sel.addRange(newRange);
      }catch(e){}
    }
    return true;
  }
  function execInitAntiList(){
    document.querySelectorAll('.exec-content').forEach(function(el){
      if(el._execAntiList) return;
      el._execAntiList = true;
      var changed = execUnwrapLists(el);
      if(changed && typeof scheduleSave === 'function') scheduleSave();
      el.addEventListener('input', function(){ execUnwrapLists(el); });
    });
  }
  if(document.readyState!=='loading') execInitAntiList();
  else document.addEventListener('DOMContentLoaded', execInitAntiList);
})();
(function(){
  function execEnter(){
    var el=document.getElementById('content');
    if(!el||el._execEnterV1)return;
    el._execEnterV1=true;
    el.addEventListener('keydown',function(e){
      if(e.key!=='Enter'||e.shiftKey)return;
      var t=e.target;
      if(!t||!t.closest||!t.closest('.exec-content'))return;
      e.preventDefault();e.stopPropagation();
      document.execCommand('insertLineBreak');
    },true);
  }
  if(document.readyState!=='loading')execEnter();
  else document.addEventListener('DOMContentLoaded',execEnter);
})();
</script>
<script>
/* 3-level sidebar tree */
function streeClient(id){
  var body=document.getElementById('cb-'+id);
  var row=body?body.previousElementSibling:null;
  if(!body)return;
  var open=body.style.display!=='none';
  body.style.display=open?'none':'block';
  if(row){
    row.classList.toggle('stree-open',!open);
    var arr=row.querySelector('.stree-arr');
    if(arr)arr.textContent=open?'▶':'▼';
  }
}

function streeGetExpandedMonths(){
  try{ return JSON.parse(localStorage.getItem('streeExpandedMonths')||'[]'); }
  catch(e){ return []; }
}
function streeSaveMonthState(id, isOpen){
  var arr = streeGetExpandedMonths();
  var idStr = String(id);
  var idx = arr.indexOf(idStr);
  if(isOpen){ if(idx===-1) arr.push(idStr); }
  else { if(idx!==-1) arr.splice(idx,1); }
  try{ localStorage.setItem('streeExpandedMonths', JSON.stringify(arr)); }catch(e){}
}
function streeRestoreExpandedMonths(){
  var arr = streeGetExpandedMonths();
  arr.forEach(function(idStr){
    var pages=document.getElementById('mp-'+idStr);
    if(!pages) return;
    var row=pages.previousElementSibling;
    pages.style.display='block';
    if(row) row.classList.add('stree-open');
    var clientBody = pages.closest('.stree-client-body');
    if(clientBody){
      clientBody.style.display='block';
      var clientRow = clientBody.previousElementSibling;
      if(clientRow) clientRow.classList.add('stree-open');
    }
  });
}
if(document.readyState!=='loading') streeRestoreExpandedMonths();
else document.addEventListener('DOMContentLoaded', streeRestoreExpandedMonths);
function streeMonth(id){
  var pages=document.getElementById('mp-'+id);
  var row=pages?pages.previousElementSibling:null;
  if(!pages)return;
  var open=pages.style.display!=='none';
  pages.style.display=open?'none':'block';
  if(row) row.classList.toggle('stree-open',!open);
  streeSaveMonthState(id, !open);
  if(!open && row){
    // Use data-fullname for navigation (stored name may differ from display)
    var fullName=row.getAttribute('data-fullname');
    if(fullName) location.href='/section/'+encodeURIComponent(fullName);
  }
}

function streeNewPage(secName){
  var form=document.createElement('form');
  form.method='POST'; form.action='/new/'+encodeURIComponent(secName);
  document.body.appendChild(form); form.submit();
}

function streeAddMonth(parentId, parentName){
  var name=prompt('Month name (e.g. July 2025):','');
  if(!name||!name.trim())return;
  fetch('/section/add-sub',{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({name:name.trim(),parent_id:parentId})
  }).then(r=>r.json()).then(d=>{
    if(d.status==='ok') location.href='/section/'+encodeURIComponent(d.name);
    else alert(d.message||'Error creating month.');
  });
}
</script>
<script>
function cloneMonthSection(secId, secName){
  if(!confirm('Clone this month and all its pages?')) return;
  fetch('/clone-month/'+secId, {method:'POST'})
    .then(r=>r.json())
    .then(d=>{
      if(d.status==='ok') location.href='/section/'+encodeURIComponent(d.name);
      else alert(d.message||'Error');
    });
}
</script>
<!-- FA_ADD_DEL_CYCLE_FUNCTIONS_ADDED_V3 -->
<script>
/* Focus Areas: Add / Delete / Cycle-color card handlers.
   Referenced via onclick= in the template but never defined,
   so "+ Add Focus Area" silently failed (ReferenceError). */
var FA_COLORS = (typeof FA_COLORS !== 'undefined') ? FA_COLORS : ['#2563eb', '#0f172a', '#059669', '#d97706', '#7c3aed', '#dc2626'];

function faBuildCardHTML(color) {
  if (typeof faCardHTML === 'function') {
    return faCardHTML(color, '', 'High', '', '', '');
  }
  var badge = (typeof faImpactBadgeHTML === 'function') ? faImpactBadgeHTML('High') : '';
  return '<div class="fa-card" data-fa-color="' + color + '">'
    + '<div class="fa-card-head" style="background:' + color + ';">'
    + '<div class="fa-card-title" contenteditable="true"></div>'
    + '<button class="fa-color-btn" onclick="faCycleColor(this)" title="Change colour"></button>'
    + '</div><div class="fa-card-inner">'
    + '<div class="fa-impact-row">' + badge + '</div>'
    + '<div class="fa-section"><div class="fa-section-label">Issue</div>'
    + '<div class="fa-section-body" contenteditable="true" data-ph="Describe the issue..."></div></div>'
    + '<div class="fa-section"><div class="fa-section-label">Impact</div>'
    + '<div class="fa-section-body" contenteditable="true" data-ph="Describe the impact..."></div></div>'
    + '<div class="fa-section"><div class="fa-section-label">Action</div>'
    + '<div class="fa-section-body" contenteditable="true" data-ph="Describe the action..."></div></div>'
    + '</div><button class="fa-del-card" onclick="faDelCard(this)">&#10005;</button></div>';
}

function faAddCard(btn) {
  var wrap = btn.closest('.fa-wrap');
  var grid = wrap ? wrap.querySelector('.fa-grid') : btn.closest('.fa-grid');
  if (!grid) return;
  var existing = grid.querySelectorAll('.fa-card').length;
  var color = FA_COLORS[existing % FA_COLORS.length];
  var tmp = document.createElement('div');
  tmp.innerHTML = faBuildCardHTML(color);
  var card = tmp.firstElementChild;
  grid.appendChild(card);
  var title = card.querySelector('.fa-card-title');
  if (title) title.focus();
  if (typeof scheduleSave === 'function') scheduleSave();
}

function faDelCard(btn) {
  var card = btn.closest('.fa-card');
  var grid = card ? card.closest('.fa-grid') : null;
  if (!grid) return;
  if (grid.querySelectorAll('.fa-card').length <= 1) {
    alert('Cannot delete the last focus area.');
    return;
  }
  if (!confirm('Delete this focus area?')) return;
  card.remove();
  if (typeof scheduleSave === 'function') scheduleSave();
}

function faCycleColor(btn) {
  var card = btn.closest('.fa-card');
  if (!card) return;
  var head = card.querySelector('.fa-card-head');
  var cur = card.getAttribute('data-fa-color') || FA_COLORS[0];
  var idx = FA_COLORS.indexOf(cur);
  var next = FA_COLORS[(idx + 1) % FA_COLORS.length];
  card.setAttribute('data-fa-color', next);
  if (head) head.style.setProperty('background', next, 'important');
  if (typeof scheduleSave === 'function') scheduleSave();
}
/* FA_DRAG_REORDER_V1 */
function faCardHasHandle(card){
  for (var i=0;i<card.children.length;i++){
    if (card.children[i].classList && card.children[i].classList.contains('fa-drag-handle')) return true;
  }
  return false;
}
function faInitDragHandles(){
  document.querySelectorAll('.fa-grid .fa-card').forEach(function(card){
    if (faCardHasHandle(card)) return;
    var handle = document.createElement('div');
    handle.className = 'fa-drag-handle';
    handle.title = 'Drag to reorder';
    handle.textContent = '\u22ee\u22ee';
    card.insertBefore(handle, card.firstChild);
  });
}
function faGetDragAfterElement(grid, x, y){
  var cards = Array.prototype.slice.call(grid.querySelectorAll('.fa-card:not(.fa-dragging)'));
  for (var i=0;i<cards.length;i++){
    var box = cards[i].getBoundingClientRect();
    var inRow = (y >= box.top && y <= box.bottom);
    if (inRow){
      if (x < box.left + box.width/2) return cards[i];
    } else if (y < box.top) {
      return cards[i];
    }
  }
  return null;
}
(function(){
  if (document.readyState !== 'loading') faInitDragHandles();
  else document.addEventListener('DOMContentLoaded', faInitDragHandles);

  var _faMo = new MutationObserver(function(){ faInitDragHandles(); });
  function _faStartObserver(){ _faMo.observe(document.body, {childList:true, subtree:true}); }
  if (document.body) _faStartObserver();
  else document.addEventListener('DOMContentLoaded', _faStartObserver);

  document.addEventListener('mousedown', function(e){
    var handle = e.target.closest('.fa-drag-handle');
    if(!handle) return;
    var card = handle.closest('.fa-card');
    if(card) card.setAttribute('draggable','true');
  });
  document.addEventListener('mouseup', function(){
    document.querySelectorAll('.fa-card[draggable="true"]').forEach(function(c){
      if(!c.classList.contains('fa-dragging')) c.removeAttribute('draggable');
    });
  });
  document.addEventListener('dragstart', function(e){
    var card = e.target.closest('.fa-card');
    if(!card || card.getAttribute('draggable')!=='true'){ return; }
    try{ e.dataTransfer.setData('text/plain','fa-card'); }catch(err){}
    e.dataTransfer.effectAllowed='move';
    card.classList.add('fa-dragging');
    window._faDragCard = card;
  });
  document.addEventListener('dragend', function(e){
    var card = e.target.closest('.fa-card');
    if(card){ card.classList.remove('fa-dragging'); card.removeAttribute('draggable'); }
    window._faDragCard = null;
    if (typeof scheduleSave === 'function') scheduleSave();
  });
  var _faDragPending = false;
  var _faLastEvt = null;
  function _faApplyDragMove(){
    _faDragPending = false;
    var e = _faLastEvt;
    if(!e) return;
    var grid = e.target.closest('.fa-grid');
    if(!grid || !window._faDragCard) return;
    var afterEl = faGetDragAfterElement(grid, e.clientX, e.clientY);
    var addBtn = grid.querySelector('.fa-add-btn');
    if(afterEl === window._faDragCard || (afterEl && afterEl.previousElementSibling === window._faDragCard)) return;
    if(afterEl){ grid.insertBefore(window._faDragCard, afterEl); }
    else if(addBtn){ grid.insertBefore(window._faDragCard, addBtn); }
    else { grid.appendChild(window._faDragCard); }
  }
  document.addEventListener('dragover', function(e){
    var grid = e.target.closest('.fa-grid');
    if(!grid || !window._faDragCard) return;
    e.preventDefault();
    _faLastEvt = e;
    if(!_faDragPending){
      _faDragPending = true;
      requestAnimationFrame(_faApplyDragMove);
    }
  });
})();
</script>
</html>"""


CUSTOMER_SVG_GL = '<svg width="180" height="130" viewBox="0 0 180 130" xmlns="http://www.w3.org/2000/svg"><circle cx="55" cy="35" r="15" fill="#93c5fd"/><rect x="38" y="54" width="34" height="44" rx="8" fill="#2563eb"/><rect x="26" y="66" width="13" height="28" rx="6" fill="#93c5fd"/><rect x="59" y="66" width="13" height="28" rx="6" fill="#93c5fd"/><circle cx="125" cy="35" r="15" fill="#fca5a5"/><rect x="108" y="54" width="34" height="44" rx="8" fill="#1d4ed8"/><rect x="96" y="66" width="13" height="28" rx="6" fill="#fca5a5"/><rect x="129" y="66" width="13" height="28" rx="6" fill="#fca5a5"/><ellipse cx="90" cy="86" rx="13" ry="9" fill="#fbbf24" opacity=".9"/><text x="66" y="50" font-size="13" fill="#fbbf24">&#9733;&#9733;&#9733;&#9733;&#9733;</text><text x="146" y="72" font-size="22">&#128077;</text></svg>'

VENDOR_SVG_GL   = '<svg width="200" height="130" viewBox="0 0 200 130" xmlns="http://www.w3.org/2000/svg"><rect x="55" y="8" width="100" height="72" rx="8" fill="white" opacity=".15"/><rect x="58" y="11" width="94" height="64" rx="5" fill="white" opacity=".2"/><rect x="65" y="20" width="40" height="6" rx="3" fill="white" opacity=".6"/><rect x="115" y="50" width="10" height="18" rx="2" fill="#6ee7b7"/><rect x="128" y="44" width="10" height="24" rx="2" fill="#34d399"/><rect x="141" y="38" width="10" height="30" rx="2" fill="#10b981"/><circle cx="38" cy="48" r="12" fill="#fde68a"/><rect x="24" y="63" width="28" height="38" rx="7" fill="#f59e0b"/><circle cx="162" cy="44" r="13" fill="#a7f3d0"/><rect x="147" y="60" width="30" height="40" rx="7" fill="#059669"/><circle cx="172" cy="28" r="14" fill="#fbbf24" opacity=".9"/><text x="166" y="33" font-size="13" font-weight="bold" fill="white">$</text><circle cx="176" cy="68" r="11" fill="#4ade80" opacity=".9"/><text x="171" y="73" font-size="11" font-weight="bold" fill="white">&#10003;</text></svg>'

MAH_SVG_GL      = '<svg width="200" height="130" viewBox="0 0 200 130" xmlns="http://www.w3.org/2000/svg"><rect x="55" y="8" width="100" height="72" rx="8" fill="white" opacity=".15"/><rect x="58" y="11" width="94" height="64" rx="5" fill="white" opacity=".2"/><rect x="65" y="20" width="40" height="6" rx="3" fill="white" opacity=".6"/><rect x="115" y="50" width="10" height="18" rx="2" fill="#fed7aa"/><rect x="128" y="44" width="10" height="24" rx="2" fill="#fb923c"/><rect x="141" y="38" width="10" height="30" rx="2" fill="#f97316"/><circle cx="38" cy="48" r="12" fill="#fde68a"/><rect x="24" y="63" width="28" height="38" rx="7" fill="#f59e0b"/><circle cx="162" cy="44" r="13" fill="#fde68a"/><rect x="147" y="60" width="30" height="40" rx="7" fill="#ea580c"/><circle cx="172" cy="28" r="14" fill="#fbbf24" opacity=".9"/><text x="166" y="33" font-size="13" font-weight="bold" fill="white">$</text><circle cx="176" cy="68" r="11" fill="#fb923c" opacity=".9"/><text x="171" y="73" font-size="11" font-weight="bold" fill="white">&#10003;</text></svg>'

OB_CONFIGS_GL   = {'customer': {'header_cls': 'ob-header-customer', 'title1': 'CUSTOMER', 'title2': 'ONBOARDING', 'subtitle': 'Track and manage your customer onboarding journey', 'th_color': '#1e40af', 'row_even': '#eff6ff', 'steps': ['Customer Contacted?', 'Master Data Shared', 'Configuration completed in Test', 'Send Test File', 'Add Company and Product MD to PROD', 'Promote connection to Production', 'Verify 1st Production File']}, 'vendor': {'header_cls': 'ob-header-vendor', 'title1': 'VENDOR/CMO', 'title2': 'ONBOARDING', 'subtitle': 'Track and manage your Vendor/CMO onboarding process', 'th_color': '#065f46', 'row_even': '#f0fdf4', 'steps': ['Add Company/Location MD to Test', 'Share Product MD with CMO', 'Configuration completed in Test', 'Test EPCIS file from Vendor/CMO received and processed successfully', 'Add Company and Product MD to PROD', 'Promote connection to Production', 'Verify 1st Production File']}, 'mah': {'header_cls': 'ob-header-mah', 'title1': 'MAH', 'title2': 'ONBOARDING', 'subtitle': 'Track and manage your MAH onboarding process', 'th_color': '#7c2d12', 'row_even': '#fff7ed', 'steps': ['Add Company/Location MD to Test', 'Share Product MD with CMO', 'Configuration completed in Test', 'Test EPCIS file from CMO - routing to MAH', 'Add Company and Product MD to PROD', 'Promote connection to Production', 'Verify 1st Production File']}}


def _build_onboarding_html(ob_type):

    """Matrix layout: Customer Name rows x Checklist Step columns."""

    from datetime import datetime

    current_month = datetime.now().strftime("%B %Y")

    cfg   = OB_CONFIGS_GL.get(ob_type, OB_CONFIGS_GL["customer"])

    steps = cfg["steps"]

    th_col= cfg["th_color"]

    row_ev= cfg["row_even"]

    illus = {"customer": CUSTOMER_SVG_GL, "vendor": VENDOR_SVG_GL, "mah": MAH_SVG_GL}.get(ob_type, CUSTOMER_SVG_GL)


    def status_btn():

        return ('<span class="ob-status-btn pending" data-status="pending" '

                'title="Click to cycle status">Pending</span>')


    def customer_rows(steps, th_col, row_even, cust_label, cust_idx):

        n = len(steps)

        rows = ""

        for si, step in enumerate(steps):

            bg = "background:" + row_even + ";" if si % 2 == 1 else "background:#fff;"

            bullet = ('<span style="width:10px;height:10px;border-radius:50%;flex-shrink:0;'

                      'background:' + th_col + ';display:inline-block;margin-right:10px;"></span>')

            step_td = ('<td style="' + bg + 'padding:11px 16px;font-size:13px;">'

                       '<span style="display:inline-flex;align-items:center;">' + bullet +

                       '<span style="color:#1e293b;font-weight:500;">' + step + '</span></span></td>')

            status_td = ('<td style="' + bg + 'padding:11px 14px;text-align:center;">'

                         + status_btn() + '</td>')

            date1_td = ('<td style="' + bg + 'padding:8px 10px;text-align:center;">'

                        '<div class="date-cell-wrap">'

                        '<span class="date-val" contenteditable="true" data-ph="DD-Mon-YYYY"></span>'

                        '<button class="date-cal-btn" onclick="openDatePicker(event,this)" type="button" title="Pick date">📅</button>'

                        '</div></td>')

            date2_td = ('<td style="' + bg + 'padding:8px 10px;text-align:center;">'

                        '<div class="date-cell-wrap">'

                        '<span class="date-val" contenteditable="true" data-ph="DD-Mon-YYYY"></span>'

                        '<button class="date-cal-btn" onclick="openDatePicker(event,this)" type="button" title="Pick date">📅</button>'

                        '</div></td>')

            brief_td = '<td class="ob-comment-cell" style="' + bg + 'padding:11px 14px;" contenteditable="true"></td>'

            if si == 0:

                ns = ("background:" + th_col + "22;border-left:4px solid " + th_col + ";"

                      "font-weight:700;font-size:14px;color:" + th_col + ";"

                      "padding:12px 18px;vertical-align:middle;text-align:center;min-width:160px;")

                name_td = ('<td rowspan="' + str(n) + '" class="sc-name-td" style="' + ns + '" contenteditable="true">'

                           + cust_label + '</td>')

                rows += '<tr data-cg="' + str(cust_idx) + '">' + name_td + step_td + status_td + date1_td + date2_td + brief_td + '</tr>'

            else:

                rows += '<tr data-cg="' + str(cust_idx) + '">' + step_td + status_td + date1_td + date2_td + brief_td + '</tr>'

        rows += ('<tr><td colspan="6" style="height:5px;background:linear-gradient(90deg,'

                 + th_col + '22,transparent);border-bottom:2px solid ' + th_col + '33;"></td></tr>')

        return rows


    rows_html = customer_rows(steps, th_col, row_ev, "Customer 1", 0)

    steps_json = json.dumps(steps).replace('"', '&quot;')

    add_btn_row = (

        '<tr id="ob-add-btn-row"><td colspan="6" style="padding:10px 16px;">'

        '<button onclick="obAddCustomerGroup(this)" data-steps="' + steps_json + '" '

        'data-thcolor="' + th_col + '" data-roweven="' + row_ev + '" '

        'style="padding:6px 18px;background:none;border:1.5px dashed ' + th_col + ';'

        'color:' + th_col + ';border-radius:8px;cursor:pointer;font-size:13px;font-weight:700;">'

        '+ Add Customer</button></td></tr>'

    )


    # Header background colors per type

    hdr_bg = {

        'customer': 'linear-gradient(135deg,#1e3a8a 0%,#2563eb 60%,#3b82f6 100%)',

        'vendor':   'linear-gradient(135deg,#064e3b 0%,#059669 60%,#10b981 100%)',

        'mah':      'linear-gradient(135deg,#7c2d12 0%,#c2410c 60%,#ea580c 100%)',

    }.get(ob_type, 'linear-gradient(135deg,#1e3a8a 0%,#2563eb 60%,#3b82f6 100%)')


    header = (

        '<div style="background:' + hdr_bg + ';padding:14px 24px;display:flex;'

        'align-items:center;gap:16px;position:relative;overflow:hidden;color:white;'

        'border-radius:10px 10px 0 0;">'

        # decorative circles

        '<div style="position:absolute;right:-30px;top:-40px;width:160px;height:160px;'

        'border-radius:50%;background:rgba(255,255,255,0.06);pointer-events:none;"></div>'

        '<div style="position:absolute;right:60px;bottom:-50px;width:120px;height:120px;'

        'border-radius:50%;background:rgba(255,255,255,0.04);pointer-events:none;"></div>'

        '<div style="flex:1;position:relative;z-index:1;">'

        '<div style="font-size:22px;font-weight:900;color:white;letter-spacing:1px;'

        'text-transform:uppercase;line-height:1.1;">'

        + cfg["title1"] + ' ' + cfg["title2"] +

        '</div>'

        '</div>'

        '</div>'

    )

    progress = (

        '<div class="ob-progress-bar-wrap">'

        '<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 18px 6px;">'

        '<span style="font-size:12px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.5px;">Overall Progress</span>'

        '<span class="ob-progress-pct" style="font-size:13px;font-weight:700;color:' + th_col + ';">0%</span></div>'

        '<div style="margin:0 18px 12px;height:8px;background:#e2e8f0;border-radius:10px;overflow:hidden;">'

        '<div class="ob-progress-fill" style="height:100%;width:0%;background:' + th_col + ';border-radius:10px;transition:width .4s;"></div>'

        '</div></div>'

    )

    table = (

        '<div class="ob-table-wrap" style="overflow-x:auto;">'

        '<table class="ob-table">'

        '<thead><tr style="background:' + th_col + ';font-size:12px;">'

        '<th style="min-width:160px;text-align:center;">Customer Name</th>'

        '<th>Check Point</th>'

        '<th style="text-align:center;width:140px;">Status</th>'

        '<th style="min-width:120px;text-align:center;">Target Date</th>'

        '<th style="min-width:130px;text-align:center;">Completion Date</th>'

        '<th style="min-width:160px;">Comment</th></tr></thead>'

        '<tbody>' + rows_html + add_btn_row + '</tbody></table></div>'

    )

    return '<div class="ob-wrap" style="--ob-border:' + th_col + '">' + header + progress + table + '</div>'


def _build_serialization_html():
    """Serialization Issues -- Timeline layout template."""
    COLORS = ['#038dbd','#0d152b','#6366f1','#059669','#d97706','#dc2626','#7c3aed','#0891b2']
    COLOR_MAP = {
        '#038dbd': ('rgba(3,141,189,.08)',  '#bae6fd'),
        '#0d152b': ('rgba(13,21,43,.06)',   '#c7d2fe'),
        '#6366f1': ('rgba(99,102,241,.08)', '#c4b5fd'),
        '#059669': ('rgba(5,150,105,.08)',  '#a7f3d0'),
        '#d97706': ('rgba(217,119,6,.08)',  '#fde68a'),
        '#dc2626': ('rgba(220,38,38,.08)',  '#fecaca'),
        '#7c3aed': ('rgba(124,58,237,.08)', '#ddd6fe'),
        '#0891b2': ('rgba(8,145,178,.08)',  '#a5f3fc'),
    }

    def date_inline():
        return ('<span class="si-tl-date-wrap">'
                '<span class="date-val" contenteditable="true" data-ph="DD-Mon-YYYY"></span>'
                '<button class="date-cal-btn" onclick="openDatePicker(event,this)" type="button" title="Pick date">\U0001F4C5</button>'
                '</span>')

    def action_row():
        return ('<div class="si-action-row">'
                '<span class="si-action-text" contenteditable="true" data-ph="Action item..."></span>'
                '</div>')

    def card(idx, col):
        bg, border = COLOR_MAP.get(col, ('rgba(3,141,189,.08)', '#bae6fd'))
        css_vars = "--si-col:" + col + ";--si-col-bg:" + bg + ";--si-col-light:" + border + ";"
        title_box = '<div class="si-tl-issue-title" contenteditable="true" data-ph="Issue Title"></div>'
        tl = (
            '<div class="si-tl">'
            '<div class="si-tl-step" data-step="created">'
            '<div class="si-tl-node" style="--step-col:#3b82f6;">\U0001F4DD</div>'
            '<div class="si-tl-line"></div>'
            '<div class="si-tl-content">'
            '<div class="si-tl-head"><span class="si-tl-title">Issue Created</span>' + date_inline() + '</div>'
            '<div class="si-tl-sublabel">Issue Description</div>'
            '<div class="si-tl-body" contenteditable="true" data-ph="Describe the issue..."></div>'
            '</div></div>'
            '<div class="si-tl-step" data-step="rootcause">'
            '<div class="si-tl-node" style="--step-col:#f59e0b;">\u26A0</div>'
            '<div class="si-tl-line"></div>'
            '<div class="si-tl-content">'
            '<div class="si-tl-head"><span class="si-tl-title">Root Cause</span></div>'
            '<div class="si-tl-body" contenteditable="true" data-ph="Identify root cause..."></div>'
            '</div></div>'
            '<div class="si-tl-step" data-step="actions">'
            '<div class="si-tl-node" style="--step-col:#22c55e;">\u2705</div>'
            '<div class="si-tl-line"></div>'
            '<div class="si-tl-content">'
            '<div class="si-tl-head"><span class="si-tl-title">Action Items / Comments</span></div>'
            '<div class="si-action-list">' + action_row() + '</div>'
            '</div></div>'
            '<div class="si-tl-step si-tl-step-last" data-step="resolution">'
            '<div class="si-tl-node" style="--step-col:#8b5cf6;">\U0001F3AF</div>'
            '<div class="si-tl-content">'
            '<div class="si-tl-head"><span class="si-tl-title">Resolution Date:</span>' + date_inline() + '</div>'
            '</div></div>'
            '</div>'
        )
        return (
            '<div class="si-card si-tl-card" style="' + css_vars + '" data-si-card="' + str(idx) + '">'
            '<div class="si-card-body">' + title_box + tl + '</div>'
            '<div class="si-card-foot">'
            '<div class="si-status-group">'
            '<span class="si-status-label">Status</span>'
            '<span class="ob-status-btn pending" data-status="pending" title="Click to change">Pending</span>'
            '</div>'
            '<button class="si-del-card" onclick="siDeleteCard(this)" type="button">&#128465; Delete</button>'
            '</div></div>'
        )

    cards_html = "".join(card(i, COLORS[i % len(COLORS)]) for i in range(3))
    add_btn = '<button class="si-add-btn" onclick="siAddCard(this)" type="button">&#43; Add Issue</button>'
    banner = ('<div class="si-banner">'
              '<div class="si-banner-title">SERIALIZATION ISSUES</div>'
              '</div>')
    return ('<div class="si-wrap">' + banner
            + '<div class="si-cards-area"><div class="si-cards-grid">'
            + cards_html + add_btn
            + '</div></div></div>')


@app.route("/thankyou-graphic.png")
def serve_thankyou_graphic():
    from flask import send_file
    path = os.path.join(BASE_DIR, "Support Image", "thankyou_graphic.png")
    if not os.path.exists(path):
        # thankyou_graphic.png isn't shipped in Support Image; fall back to
        # the thank-you background so this route never 404s.
        path = os.path.join(BASE_DIR, "Support Image", "thankyou_bg.png")
    if not os.path.exists(path):
        return "Not found", 404
    with open(path, "rb") as _f:
        _header = _f.read(3)
    _mt = "image/jpeg" if _header[:3] == b"\xff\xd8\xff" else "image/png"
    return send_file(path, mimetype=_mt)

@app.route("/thankyou-bg.png")
def serve_thankyou_bg():
    from flask import send_file
    path = os.path.join(BASE_DIR, "Support Image", "thankyou_bg.png")
    if not os.path.exists(path):
        return "Not found", 404
    with open(path, "rb") as _f:
        _header = _f.read(3)
    _mt = "image/jpeg" if _header[:3] == b"\xff\xd8\xff" else "image/png"
    return send_file(path, mimetype=_mt)

@app.route("/cover-logo.png")
def serve_cover_logo():
    from flask import send_file
    path = os.path.join(BASE_DIR, "Support Image", "cover_logo.png")
    if not os.path.exists(path):
        # cover_logo.png isn't shipped in Support Image; fall back to the
        # full-colour logo (kept in the folder but otherwise unused) so this
        # route never 404s.
        path = os.path.join(BASE_DIR, "Support Image", "Varitec Logo_Full_Black-Blue.png")
    if not os.path.exists(path):
        path = os.path.join(BASE_DIR, "Support Image", "varitec_logo.png")
    if not os.path.exists(path):
        return "Not found", 404
    with open(path, "rb") as _f:
        _header = _f.read(3)
    _mt = "image/jpeg" if _header[:3] == b"\xff\xd8\xff" else "image/png"
    return send_file(path, mimetype=_mt)

@app.route("/cover-slide-bg.png")
def serve_cover_slide_bg():
    from flask import send_file
    path = os.path.join(BASE_DIR, "Support Image", "cover_Slide.png")
    if not os.path.exists(path):
        return "Not found", 404
    with open(path, "rb") as _f:
        _header = _f.read(3)
    _mt = "image/jpeg" if _header[:3] == b"\xff\xd8\xff" else "image/png"
    return send_file(path, mimetype=_mt)

@app.route("/logo.png")
def serve_logo():
    from flask import send_file
    logo_path = os.path.join(BASE_DIR, "Support Image", "varitec_logo.png")
    if not os.path.exists(logo_path):
        return "Logo not found", 404
    with open(logo_path, "rb") as _f:
        _header = _f.read(3)
    _mt = "image/jpeg" if _header[:3] == b"\xff\xd8\xff" else "image/png"
    return send_file(logo_path, mimetype=_mt)


@app.route("/vc-logo.png")
def serve_vc_logo():
    from flask import send_file
    vc_path = os.path.join(BASE_DIR, "Support Image", "vc_logo_bw.png")
    if not os.path.exists(vc_path):
        return "Not found", 404
    with open(vc_path, "rb") as _f:
        _header = _f.read(3)
    _mt = "image/jpeg" if _header[:3] == b"\xff\xd8\xff" else "image/png"
    return send_file(vc_path, mimetype=_mt)

@app.route("/",methods=["GET","POST"])

def login():

    error=None

    MAX_FAILED_ATTEMPTS = 5

    if request.method=="POST":

        email=request.form.get("user_email_field", request.form.get("email","")).strip().lower()

        password=request.form.get("password","")

        generic_error="Access denied. Contact your administrator."

        if not password or not UserAccess.has_access(email):

            error=generic_error

        else:

            ua = UserAccess.query.filter_by(email=email).first()

            if ua is None:

                # DEFAULT_ADMINS with no UserAccess row yet - create one on first login

                ua = UserAccess(email=email, role=UserAccess.get_role(email))

                ua.set_password(password)

                db.session.add(ua)

                db.session.commit()

                session["email"]=email

                session["role"]=UserAccess.get_role(email)

                return redirect("/notes")

            if ua.locked:

                error="Account locked. Contact your administrator."

            elif not ua.password:

                # No password on file yet - first login sets it

                ua.set_password(password)

                db.session.commit()

                session["email"]=email

                session["role"]=UserAccess.get_role(email)

                return redirect("/notes")

            elif ua.check_password(password):

                ua.failed_attempts = 0

                db.session.commit()

                session["email"]=email

                session["role"]=UserAccess.get_role(email)

                return redirect("/notes")

            else:

                ua.failed_attempts = (ua.failed_attempts or 0) + 1

                if ua.failed_attempts >= MAX_FAILED_ATTEMPTS:

                    ua.locked = True

                db.session.commit()

                error=generic_error

    return render_template_string(LOGIN_HTML,error=error)


@app.route("/signup", methods=["GET","POST"])

def signup():

    error=None

    name=email=company=reason=""

    if request.method=="POST":

        name = request.form.get("name","").strip()

        email = request.form.get("email","").strip().lower()

        password = request.form.get("password","")

        confirm_password = request.form.get("confirm_password","")

        company = request.form.get("company","").strip()

        reason = request.form.get("reason","").strip()

        if not name or not email or not password:

            error="Please fill in your name, email, and password."

        elif "@" not in email or "." not in email.split("@")[-1]:

            error="Please enter a valid email address."

        elif password != confirm_password:

            error="Passwords do not match."

        elif len(password) < 6:

            error="Password must be at least 6 characters."

        elif UserAccess.query.filter_by(email=email).first() or email in DEFAULT_ADMINS:

            error="An account with this email already exists. Try signing in instead."

        elif PendingSignup.query.filter_by(email=email, status="pending").first():

            error="A request for this email is already pending review."

        else:

            req = PendingSignup(name=name, email=email, company=company, reason=reason)

            req.password_hash = generate_password_hash(password)

            db.session.add(req)

            db.session.commit()

            return render_template_string(SIGNUP_HTML, submitted=True, name=name)

    return render_template_string(SIGNUP_HTML, submitted=False, error=error,

        name=name, email=email, company=company, reason=reason)


@app.route("/logout")

def logout():

    session.clear(); return redirect("/")


def _require_write():

    email = session.get("email","")

    if not email: return redirect("/")

    if not UserAccess.can_write(email):

        return ("❌ Read-only access. Ask your admin for write permissions.", 403)

    return None


def _require_admin():

    email = session.get("email","")

    if not email: return redirect("/")

    if not UserAccess.is_admin(email):

        return ("❌ Admin access required.", 403)

    return None


# ── Admin: User Management ────────────────────────────────────────────────

@app.route("/admin/users")

def admin_users():

    err = _require_admin()

    if err: return err

    users = UserAccess.query.order_by(UserAccess.role, UserAccess.email).all()

    cur_email = session["email"]

    shown_emails = {u.email for u in users}

    def _pw_form(email):
        return (f'''<form method="post" action="/admin/users/set-password" class="inline-form pw-form" onsubmit="return this.password.value.length>=6 || (alert('Password must be at least 6 characters.'),false);">'''
                f'''<input type="hidden" name="email" value="{email}">'''
                f'''<input type="password" name="password" placeholder="New password" class="pw-input" minlength="6" required>'''
                f'''<button type="submit" class="btn-update">🔑 Set Password</button></form>''')

    default_rows = ""

    for de in sorted(DEFAULT_ADMINS):

        if de not in shown_emails:

            default_rows += f'''<div class="user-card"><div class="user-info"><div class="user-avatar admin">{de[0].upper()}</div><div><div class="user-email">{de}</div><div class="user-meta">Default Administrator</div></div></div><div class="user-actions"><span class="role-badge admin">ADMIN</span><span class="default-lock">🔒 Default</span>{_pw_form(de)}</div></div>'''

    db_rows = ""

    for u in users:

        is_default = u.email in DEFAULT_ADMINS

        role_icon = "🛡️" if u.role=="admin" else "✏️" if u.role=="write" else "👁"

        sel_opts = "".join(f'''<option value="{r}"{" selected" if r==u.role else ""}>{r.title()}</option>''' for r in ROLES)

        lock_note = ''' <span class="default-lock" style="color:#dc2626;">🔒 Locked</span>''' if u.locked else ""

        actions = ((f'''<span class="default-lock">🔒 Default</span>{_pw_form(u.email)}''') if is_default else

            f'''<form method="post" action="/admin/users/{u.id}/role" class="inline-form"><select name="role" class="role-select">{sel_opts}</select><button type="submit" class="btn-update">Save</button></form>{_pw_form(u.email)}<form method="post" action="/admin/users/{u.id}/delete" onsubmit="return confirm('Remove {u.email}?')" class="inline-form"><button type="submit" class="btn-remove">✕ Remove</button></form>''')

        db_rows += f'''<div class="user-card"><div class="user-info"><div class="user-avatar {u.role}">{u.email[0].upper()}</div><div><div class="user-email">{u.email}</div><div class="user-meta">{role_icon} {u.role.title()} Access{lock_note}</div></div></div><div class="user-actions"><span class="role-badge {u.role}">{u.role.upper()}</span>{actions}</div></div>'''

    total = len(users) + len([d for d in DEFAULT_ADMINS if d not in shown_emails])

    managed_label = f'''<div class="section-label" style="margin-top:16px;">Managed Users</div>''' if db_rows else ""

    pending = PendingSignup.query.filter_by(status="pending").order_by(PendingSignup.created_at.desc()).all()

    pending_rows = ""

    for p in pending:

        meta_bits = []

        if p.company: meta_bits.append(p.company)

        if p.reason: meta_bits.append(p.reason)

        meta = " &middot; ".join(meta_bits) if meta_bits else "No additional details provided"

        pending_rows += f'''<div class="user-card"><div class="user-info"><div class="user-avatar read">{p.name[0].upper()}</div><div><div class="user-email">{p.name} &middot; {p.email}</div><div class="user-meta">{meta}</div></div></div><div class="user-actions"><form method="post" action="/admin/signups/{p.id}/approve" class="inline-form"><select name="role" class="role-select"><option value="read">👁 Read</option><option value="write">✏️ Write</option><option value="admin">🛡️ Admin</option></select><button type="submit" class="btn-update">✓ Approve</button></form><form method="post" action="/admin/signups/{p.id}/reject" onsubmit="return confirm('Reject request from {p.email}?')" class="inline-form"><button type="submit" class="btn-remove">✕ Reject</button></form></div></div>'''

    pending_section = f'''<div class="user-list" style="border-top:1px solid #f1f5f9;"><div class="section-label">Pending Sign-Up Requests ({len(pending)})</div>{pending_rows or '<div class="user-meta">No pending requests.</div>'}</div>'''

    panel = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>User Management</title><meta name="viewport" content="width=device-width,initial-scale=1">

<style>*{{box-sizing:border-box;margin:0;padding:0;}}body{{font-family:'Segoe UI',system-ui,sans-serif;background:linear-gradient(135deg,#0d152b,#1e3868);min-height:100vh;padding:28px 16px;}}.page-wrap{{max-width:780px;margin:0 auto;}}.page-header{{display:flex;align-items:center;gap:14px;margin-bottom:24px;}}.logo-mark{{font-family:'Arial Black',Arial,sans-serif;font-size:18px;font-weight:900;color:white;letter-spacing:2px;}}.logo-mark span{{font-weight:300;font-size:12px;color:#4db2de;}}.back-link{{margin-left:auto;color:#4db2de;text-decoration:none;font-size:13px;font-weight:600;border:1px solid rgba(77,178,222,.4);border-radius:20px;padding:5px 14px;}}.back-link:hover{{background:rgba(77,178,222,.15);}}.card{{background:white;border-radius:16px;overflow:hidden;box-shadow:0 12px 48px rgba(0,0,0,.25);}}.card-header{{background:linear-gradient(135deg,#0d152b,#1e3868);padding:24px 28px;display:flex;align-items:center;gap:16px;}}.card-header-icon{{width:46px;height:46px;border-radius:12px;background:rgba(3,141,189,.3);display:flex;align-items:center;justify-content:center;font-size:22px;}}.card-header-text h2{{color:white;font-size:20px;font-weight:800;margin-bottom:3px;}}.card-header-text p{{color:#4db2de;font-size:12px;}}.stat-chips{{margin-left:auto;}}.stat-chip{{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.15);border-radius:20px;padding:4px 12px;font-size:11px;color:white;font-weight:600;}}.user-list{{padding:20px 28px;}}.section-label{{font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px;}}.user-card{{display:flex;align-items:center;justify-content:space-between;gap:10px 16px;flex-wrap:wrap;padding:14px 16px;border-radius:10px;border:1px solid #f1f5f9;margin-bottom:8px;background:#fafafa;}}.user-card:hover{{border-color:#e2e8f0;background:white;box-shadow:0 2px 8px rgba(0,0,0,.06);}}.user-info{{display:flex;align-items:center;gap:12px;min-width:0;}}.user-avatar{{width:38px;height:38px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:800;color:white;flex-shrink:0;}}.user-avatar.admin{{background:linear-gradient(135deg,#dc2626,#ef4444);}}.user-avatar.write{{background:linear-gradient(135deg,#059669,#10b981);}}.user-avatar.read{{background:linear-gradient(135deg,#0369a1,#0ea5e9);}}.user-email{{font-size:14px;font-weight:600;color:#0d152b;word-break:break-all;}}.user-meta{{font-size:11px;color:#94a3b8;margin-top:2px;}}.user-actions{{display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;}}.role-badge{{padding:3px 10px;border-radius:20px;font-size:11px;font-weight:800;}}.role-badge.admin{{background:#fee2e2;color:#dc2626;}}.role-badge.write{{background:#dcfce7;color:#16a34a;}}.role-badge.read{{background:#e0f2fe;color:#0369a1;}}.default-lock{{font-size:11px;color:#94a3b8;font-weight:600;}}.inline-form{{display:inline;}}.role-select{{border:1.5px solid #e2e8f0;border-radius:7px;padding:4px 8px;font-size:12px;color:#374151;cursor:pointer;outline:none;background:white;}}.btn-update{{background:#038dbd;color:white;border:none;border-radius:7px;padding:5px 12px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;}}.btn-remove{{background:none;border:1.5px solid #fecaca;color:#ef4444;border-radius:7px;padding:4px 10px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;}}.pw-form{{display:inline-flex;gap:6px;flex-wrap:nowrap;}}.pw-input{{border:1.5px solid #e2e8f0;border-radius:7px;padding:4px 8px;font-size:12px;color:#374151;outline:none;width:110px;}}.pw-input:focus{{border-color:#038dbd;}}.add-section{{padding:20px 28px 28px;border-top:1px solid #f1f5f9;}}.add-section h3{{font-size:14px;font-weight:700;color:#0d152b;margin-bottom:14px;}}.add-form{{display:flex;gap:10px;flex-wrap:wrap;}}.add-form input{{flex:1;min-width:200px;border:1.5px solid #e2e8f0;border-radius:10px;padding:10px 14px;font-size:13px;color:#374151;outline:none;}}.add-form select{{border:1.5px solid #e2e8f0;border-radius:10px;padding:10px 14px;font-size:13px;color:#374151;outline:none;background:white;cursor:pointer;}}.add-form button{{background:linear-gradient(135deg,#0d152b,#1e3868);color:white;border:none;border-radius:10px;padding:10px 20px;font-size:13px;font-weight:700;cursor:pointer;white-space:nowrap;}}.add-form button:hover{{background:linear-gradient(135deg,#038dbd,#0284c7);}}</style></head><body>

<div class="page-wrap"><div class="page-header"><div class="logo-mark">VARITEC<span>consulting</span></div><a class="back-link" href="/notes">← Back to Notes</a></div>

<div class="card"><div class="card-header"><div class="card-header-icon">👥</div><div class="card-header-text"><h2>User Management</h2><p>Logged in as {cur_email}</p></div><div class="stat-chips"><div class="stat-chip">Total: {total}</div></div></div>

<div class="user-list"><div class="section-label">Default Administrators</div>{default_rows}{managed_label}{db_rows}</div>

{pending_section}

<div class="add-section"><h3>➕ Grant Access to New User</h3><form class="add-form" method="post" action="/admin/users/add">
<input name="email" type="email" placeholder="you@varitecconsulting.com" required autocomplete="username">
<input name="password" type="password" placeholder="Create a password" required minlength="6" autocomplete="new-password">
<input name="confirm_password" type="password" placeholder="Confirm password" required minlength="6" autocomplete="new-password">
<select name="role"><option value="read">👁 Read Only</option><option value="write">✏️ Write</option><option value="admin">🛡️ Admin</option></select>
<button type="submit">+ Add User</button>
</form></div>

</div></div></body></html>"""

    from flask import Response

    return Response(panel, mimetype="text/html")


@app.route("/admin/signups/<int:sid>/approve", methods=["POST"])

def admin_approve_signup(sid):

    err = _require_admin()

    if err: return err

    req = PendingSignup.query.get(sid)

    role = request.form.get("role", ROLE_READ)

    if req and req.status == "pending" and role in ROLES:

        if not UserAccess.query.filter_by(email=req.email).first():

            ua = UserAccess(email=req.email, role=role)

            ua.password = req.password_hash

            db.session.add(ua)

        req.status = "approved"

        db.session.commit()

    return redirect("/admin/users")


@app.route("/admin/signups/<int:sid>/reject", methods=["POST"])

def admin_reject_signup(sid):

    err = _require_admin()

    if err: return err

    req = PendingSignup.query.get(sid)

    if req and req.status == "pending":

        req.status = "rejected"

        db.session.commit()

    return redirect("/admin/users")


@app.route("/admin/users/add", methods=["POST"])

def admin_add_user():

    err = _require_admin()

    if err: return err

    email = request.form.get("email","").strip().lower()

    role  = request.form.get("role", ROLE_READ)

    password = request.form.get("password","")

    confirm_password = request.form.get("confirm_password","")

    if email and role in ROLES and password and password == confirm_password and len(password) >= 6:

        existing = UserAccess.query.filter_by(email=email).first()

        if existing:

            existing.role = role

            existing.set_password(password)

        else:

            ua = UserAccess(email=email, role=role)

            ua.set_password(password)

            db.session.add(ua)

        db.session.commit()

    return redirect("/admin/users")


@app.route("/admin/users/set-password", methods=["POST"])

def admin_set_password():

    err = _require_admin()

    if err: return err

    email = request.form.get("email","").strip().lower()

    password = request.form.get("password","")

    if email and password and len(password) >= 6:

        ua = UserAccess.query.filter_by(email=email).first()

        if ua is None and email in DEFAULT_ADMINS:

            ua = UserAccess(email=email, role=UserAccess.get_role(email))

            db.session.add(ua)

        if ua is not None:

            ua.set_password(password)

            db.session.commit()

    return redirect("/admin/users")


@app.route("/admin/users/<int:uid>/role", methods=["POST"])

def admin_update_role(uid):

    err = _require_admin()

    if err: return err

    u = UserAccess.query.get(uid)

    new_role = request.form.get("role")

    if u and new_role in ROLES:

        u.role = new_role; db.session.commit()

    return redirect("/admin/users")


@app.route("/admin/users/<int:uid>/delete", methods=["POST"])

def admin_delete_user(uid):

    err = _require_admin()

    if err: return err

    u = UserAccess.query.get(uid)

    if u and u.email not in DEFAULT_ADMINS:

        db.session.delete(u); db.session.commit()

    return redirect("/admin/users")


@app.route("/notes")

def notes_home():

    if "email" not in session: return redirect("/")

    first=Section.query.order_by(Section.order).first()

    return redirect(f"/section/{first.name}" if first else "/section/General")


def _render(nl,cur,sec):

    email = session["email"]

    _all_secs = Section.query.order_by(Section.order).all()
    _notes_map = {s.name: Note.query.filter_by(section=s.name).order_by(Note.note_order).all() for s in _all_secs}
    _top_secs = [s for s in _all_secs if s.parent_id is None]
    _sub_map  = {}
    for s in _all_secs:
        if s.parent_id: _sub_map.setdefault(s.parent_id, []).append(s)
    return render_template_string(NOTES_HTML,notes=nl,current=cur,

        sections=_all_secs, top_sections=_top_secs, sub_map=_sub_map, notes_map=_notes_map,

        active_section=sec,active=sec,email=email,

        role=UserAccess.get_role(email),

        can_write=UserAccess.can_write(email),

        tpl_name=get_tpl_name(),

        default_template_name=get_tpl_name(),

        LOGO_SRC=LOGO_SRC)


@app.route("/section/<sec>")

def section_view(sec):

    if "email" not in session: return redirect("/")

    nl=Note.query.filter_by(section=sec).order_by(Note.note_order).all()

    if not nl:

        n=Note(title="Welcome",content="",section=sec); db.session.add(n); db.session.commit(); nl=[n]

    return _render(nl,nl[0],sec)


@app.route("/note/<int:nid>")

def open_note(nid):

    if "email" not in session: return redirect("/")

    cur=db.session.get(Note,nid)

    if not cur: return redirect("/notes")

    return _render(Note.query.filter_by(section=cur.section).order_by(Note.note_order).all(),cur,cur.section)


def _next_order(sec):

    """Return max note_order + 1 for a section (appends new notes at end)."""

    mx = db.session.query(db.func.max(Note.note_order)).filter_by(section=sec).scalar()

    return (mx or 0) + 1


@app.route("/new/<sec>",methods=["POST"])

def new_note(sec):

    g=_require_write();

    if g: return g

    if "email" not in session: return redirect("/")

    n=Note(title="Untitled",content="",section=sec,note_order=_next_order(sec))

    db.session.add(n); db.session.commit()

    return redirect(f"/note/{n.id}")


@app.route("/new-scorecard/<sec>",methods=["POST"])

def new_scorecard(sec):

    if "email" not in session: return redirect("/")

    tpl_key = request.form.get("tpl", "all")

    title, customers = SCORECARD_TEMPLATES.get(tpl_key, SCORECARD_TEMPLATES["all"])

    content = _build_scorecard_html(customers)

    n=Note(title=title, content=content, section=sec, note_order=_next_order(sec))

    db.session.add(n); db.session.commit()

    return redirect(f"/note/{n.id}")


@app.route("/api/config", methods=["GET","POST"])

def api_config():

    if "email" not in session: return jsonify({"status":"error"}),401

    if request.method=="POST":

        data = request.get_json()

        cfg = get_config()

        if "openai_key"    in data: cfg["openai_key"]    = data["openai_key"].strip()

        if "anthropic_key" in data: cfg["anthropic_key"] = data["anthropic_key"].strip()

        save_config(cfg)

        return jsonify({"status":"ok"})

    cfg = get_config()

    def mask(k): return ("sk-..." + k[-4:]) if k and len(k)>8 else ("" if not k else "set")

    return jsonify({

        "openai_set":    bool(cfg.get("openai_key") or os.environ.get("OPENAI_API_KEY","")),

        "anthropic_set": bool(cfg.get("anthropic_key") or os.environ.get("ANTHROPIC_API_KEY","")),

        "openai_masked":    mask(cfg.get("openai_key","") or os.environ.get("OPENAI_API_KEY","")),

        "anthropic_masked": mask(cfg.get("anthropic_key","") or os.environ.get("ANTHROPIC_API_KEY","")),

    })


@app.route("/save/<int:nid>",methods=["POST"])

def save(nid):

    g=_require_write()

    if g: return g

    n=db.session.get(Note,nid)

    if not n: return jsonify({"status":"error"}),404

    d=request.get_json(); n.title=d.get("title",n.title); n.content=d.get("content",n.content)

    db.session.commit(); return jsonify({"status":"ok"})


@app.route("/delete/<int:nid>",methods=["POST"])

def delete_note(nid):

    g=_require_write()

    if g: return g

    n=db.session.get(Note,nid)

    if n: db.session.delete(n); db.session.commit()

    return jsonify({"status":"ok"})


@app.route("/note-content/<int:nid>")

def note_content(nid):

    if "email" not in session: return jsonify({"status":"error"}),401

    note = db.session.get(Note, nid)

    if not note: return jsonify({"status":"error"}),404

    soup  = BeautifulSoup(note.content or "", "html.parser")

    plain = soup.get_text(separator=" ").strip()

    return jsonify({"title": note.title, "plain": plain[:4000]})


@app.route("/store-summary/<int:nid>", methods=["POST"])

def store_summary(nid):

    if "email" not in session: return jsonify({"status":"error"}),401

    note = db.session.get(Note, nid)

    if not note: return jsonify({"status":"error"}),404

    data = request.get_json() or {}

    data["_sig"] = str(hash(note.content))

    note.summary = json.dumps(data)

    db.session.commit()

    return jsonify({"status":"ok"})


def _rule_based_summary(title, plain):

    import re

    text = plain.strip()

    tl = text.lower()

    if any(w in tl for w in ['closed','resolved','completed','complet','done','fixed','rma sent']):

        status = "Completed"

    elif any(w in tl for w in ['in progress','ongoing','pending','open','working','under']):

        status = "In Progress"

    else:

        status = "Pending"

    sm = re.search(r'status\s*[:\-]\s*(.+)', text, re.I)

    if sm:

        sv = sm.group(1).strip().lower().split('\n')[0]

        if any(w in sv for w in ['closed','resolved','complet','done']): status = "Completed"

        elif any(w in sv for w in ['progress','open','pending','ongoing']): status = "In Progress"

    issue = ""

    im = re.search(r'issue\s*[:\-]\s*(.+?)(?=root cause|status|action|$)', text, re.I|re.S)

    if im: issue = im.group(1).strip().split('\n')[0][:300]

    if not issue: issue = text[:250].strip()

    root_cause = "N/A"

    rm = re.search(r'root cause\s*[:\-]\s*(.+?)(?=status|action|issue|$)', text, re.I|re.S)

    if rm: root_cause = rm.group(1).strip().split('\n')[0][:250]

    action = ""

    am = re.search(r'action[s]?\s*[:\-]\s*(.+?)(?=status|issue|root|$)', text, re.I|re.S)

    if am:

        lines = [l.strip().lstrip("•–-▸* ") for l in am.group(1).split('\n') if l.strip().lstrip("•–-▸* ")]

        action = "; ".join(lines[:3])

    if not action:

        bullets = re.findall(r'[•▸\-]\s*(.+)', text)

        action = "; ".join(b.strip() for b in bullets[:3]) if bullets else "Review note for details."

    small_title = title.strip() if title and title != "Untitled" else (issue[:60] if issue else "Note Summary")

    return {

        "small_title": small_title[:80],

        "issue":       issue      or "See note content.",

        "root_cause":  root_cause or "N/A",

        "status":      status,

        "action":      action     or "Review note for action items.",

    }


@app.route("/summarize/<int:nid>",methods=["POST"])

def summarize_note(nid):

    if "email" not in session:

        return jsonify({"status":"error","message":"Not authenticated"}), 401

    note = db.session.get(Note, nid)

    if not note:

        return jsonify({"status":"error","message":"Note not found"}), 404

    sig   = str(hash(note.content))

    force = request.args.get("force","0") == "1"

    if not force and note.summary:

        try:

            c = json.loads(note.summary)

            if c.get("_sig") == sig and c.get("small_title"):

                return jsonify({"status":"ok",

                                "summary":{k:v for k,v in c.items() if k!="_sig"},

                                "page_title":note.title,"cached":True})

        except Exception:

            pass

    soup  = BeautifulSoup(note.content or "", "html.parser")

    plain = soup.get_text(separator=" ").strip()[:4000]

    summary = _rule_based_summary(note.title, plain)

    method  = "rules"

    cfg        = get_config()

    anth_key   = cfg.get("anthropic_key","").strip() or os.environ.get("ANTHROPIC_API_KEY","")

    openai_key = cfg.get("openai_key","").strip()    or os.environ.get("OPENAI_API_KEY","")

    prompt = f"""Analyse this note and return ONLY a valid JSON object (no markdown, no code fences).


Note Title: {note.title}

Content: {plain}


Return exactly:

{{

  "small_title": "Concise title max 10 words",

  "issue": "Brief description of the issue or task",

  "root_cause": "Underlying cause — N/A if not applicable",

  "status": "One of: In Progress | Completed | Pending",

  "action": "Actions taken or recommended next steps"

}}"""


    if anth_key:

        try:

            import anthropic as _sdk

            client = _sdk.Anthropic(api_key=anth_key)

            msg    = client.messages.create(

                model="claude-sonnet-4-20250514", max_tokens=500,

                messages=[{"role":"user","content":prompt}])

            raw = msg.content[0].text.strip()

            raw = re.sub(r'^```[a-z]*','',raw); raw = re.sub(r'```$','',raw)

            ai_result = json.loads(raw.strip())

            summary = ai_result; method = "anthropic"

        except Exception as e:

            print(f"Anthropic failed ({e}), using rule-based result")

    elif openai_key and OPENAI_PKG:

        try:

            client = _OpenAI(api_key=openai_key)

            resp   = client.chat.completions.create(

                model="gpt-4o",

                messages=[{"role":"system","content":"Return only valid JSON."},

                          {"role":"user","content":prompt}],

                max_tokens=500, temperature=0.2)

            raw = resp.choices[0].message.content.strip()

            raw = re.sub(r'^```[a-z]*','',raw); raw = re.sub(r'```$','',raw)

            ai_result = json.loads(raw.strip())

            summary = ai_result; method = "openai"

        except Exception as e:

            print(f"OpenAI failed ({e}), using rule-based result")


    note.summary = json.dumps({**summary, "_sig": sig})

    db.session.commit()

    return jsonify({"status":"ok","summary":summary,

                    "page_title":note.title,"cached":False,"method":method})


@app.route("/export/ppt",methods=["POST"])

def export_ppt():

    if "email" not in session: return jsonify({"status":"error"}),401

    nid=int(request.form.get("note_id",0)); note=db.session.get(Note,nid)

    if not note: return jsonify({"status":"error","message":"Note not found"}),404

    scope=request.form.get("scope","all"); nl=None

    if scope=="all":

        try: ids=json.loads(request.form.get("note_ids","[]")); nl=[db.session.get(Note,i) for i in ids if db.session.get(Note,i)]

        except: nl=Note.query.filter_by(section=note.section).order_by(Note.note_order).all()

    try:

        ppt=generate_ppt(note,nl); safe=re.sub(r'[^\w\s-]','',note.section).strip() or "export"

        return send_file(ppt,as_attachment=True,download_name=f"{safe}.pptx",mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:

        import traceback; traceback.print_exc()

        return jsonify({"status":"error","message":str(e)}),500


@app.route("/template/set-default",methods=["POST"])

def set_default_template():

    if "email" not in session: return jsonify({"status":"error"}),401

    f=request.files.get("template")

    if not f or not f.filename.endswith(".pptx"): return jsonify({"status":"error","message":"Upload a .pptx file."})

    try: f.save(TPL_PATH); save_tpl_name(f.filename); return jsonify({"status":"ok"})

    except Exception as e: return jsonify({"status":"error","message":str(e)})


@app.route("/template/remove-default",methods=["POST"])

def remove_default_template():

    if "email" not in session: return jsonify({"status":"error"}),401

    for p in [TPL_PATH,TPL_NAME_FILE]:

        if os.path.exists(p): os.remove(p)

    return jsonify({"status":"ok"})


@app.route("/section/add",methods=["POST"])

def add_section():

    if "email" not in session: return jsonify({"status":"error"}),401

    data=request.get_json(); name=data.get("name","").strip()

    if not name: return jsonify({"status":"error","message":"Name cannot be empty."})

    if Section.query.filter_by(name=name).first(): return jsonify({"status":"error","message":"Section already exists."})

    mo=db.session.query(db.func.max(Section.order)).scalar() or 0

    db.session.add(Section(name=name,order=mo+1)); db.session.commit(); return jsonify({"status":"ok"})


@app.route("/section/rename/<int:sid>",methods=["POST"])

def rename_section(sid):

    if "email" not in session: return jsonify({"status":"error"}),401

    sec=db.session.get(Section,sid); data=request.get_json(); nn=data.get("name","").strip()

    if not nn: return jsonify({"status":"error","message":"Name cannot be empty."})

    if Section.query.filter_by(name=nn).first(): return jsonify({"status":"error","message":"Name exists."})

    old=sec.name; sec.name=nn; Note.query.filter_by(section=old).update({"section":nn})

    db.session.commit(); return jsonify({"status":"ok"})


@app.route("/section/delete/<int:sid>",methods=["POST"])

def delete_section(sid):

    if "email" not in session: return jsonify({"status":"error"}),401

    sec=db.session.get(Section,sid)

    if sec: Note.query.filter_by(section=sec.name).delete(); db.session.delete(sec); db.session.commit()

    return jsonify({"status":"ok"})


@app.route("/reorder-sections", methods=["POST"])
def reorder_sections():
    if "email" not in session: return jsonify({"status":"error"}), 401
    data = request.get_json()
    section_ids = data.get("section_ids", [])
    for i, sid in enumerate(section_ids):
        sec = db.session.get(Section, int(sid))
        if sec: sec.order = i
    db.session.commit()
    return jsonify({"status": "ok"})

@app.route("/reorder-notes", methods=["POST"])

def reorder_notes():

    if "email" not in session: return jsonify({"status":"error"}), 401

    data = request.get_json()

    note_ids = data.get("note_ids", [])

    for i, nid in enumerate(note_ids):

        note = db.session.get(Note, int(nid))

        if note: note.note_order = i

    db.session.commit()

    return jsonify({"status": "ok"})


@app.route("/new-onboarding/<sec>", methods=["POST"])

def new_onboarding(sec):

    if "email" not in session: return redirect("/")

    ob_type = request.form.get("type","customer")

    titles = {"customer":"Customer Onboarding","vendor":"Vendor/CMO Onboarding","mah":"MAH Onboarding"}

    note = Note(title=titles.get(ob_type,"Onboarding"), content=_build_onboarding_html(ob_type), section=sec, note_order=_next_order(sec))

    db.session.add(note); db.session.commit()

    return redirect(f"/note/{note.id}")


@app.route("/new-serialization/<sec>", methods=["POST"])

def new_serialization(sec):

    if "email" not in session: return redirect("/")

    note = Note(title="Serialization Issues", content=_build_serialization_html(), section=sec, note_order=_next_order(sec))

    db.session.add(note); db.session.commit()

    return redirect(f"/note/{note.id}")


def _build_executive_summary_html():

    """Executive Summary slide with one free-form box per section."""

    from datetime import datetime

    current_month = datetime.now().strftime("%B %Y")

    sections = [

        ("red",   "Key Challenges",   "Enter key challenge..."),

        ("green", "Key Achievements", "Enter key achievement..."),

        ("amber", "Focus Next Month", "Enter focus area..."),

    ]

    body = ""

    for colour, label, ph in sections:

        body += (

            '<div class="exec-section">'

            '<div class="exec-section-hdr">'

            '<span class="exec-dot ' + colour + '"></span>'

            '<span class="exec-section-label">' + label + '</span>'

            '</div>'

            '<div class="exec-content" contenteditable="true" data-ph="' + ph + '"></div>'

            '</div>'

        )

    return (

        '<div class="exec-wrap">'

        '<div class="exec-header">'

        '<div class="exec-header-text">'

        '<h1 class="exec-title">Executive Summary</h1>'

        '</div>'

        '</div>'

        '<div class="exec-body">' + body + '</div>'

        '</div>'

    )


def _build_agenda_html():

    """Agenda slide template with 3 numbered items and colorful header."""

    from datetime import datetime

    current_month = datetime.now().strftime("%B %Y")

    items = [

        ("1", "Summary"),

        ("2", "Focus Areas &amp; Issues"),

        ("3", "Scorecard Performance"),

    ]

    agenda_items = ""

    for num, label in items:

        agenda_items += (

            '<div class="agenda-item">'

            '<button class="agenda-del-item" type="button" onclick="agendaDelItem(this)" title="Remove item">&#10005;</button>'
            '<div class="agenda-num" contenteditable="true">' + num + '</div>'

            '<div class="agenda-label" contenteditable="true">' + label + '</div>'

            '</div>'

        )

    return (

        '<div class="agenda-wrap">'

        '<div class="agenda-header">'

        '<div style="z-index:1;">'

        '<h1 class="agenda-title">Agenda</h1>'

        '</div>'

        '</div>'

        '<div class="agenda-body">' + agenda_items + '</div>'
        '<div class="agenda-add-row">'
        '<button class="agenda-add-item" type="button" onclick="agendaAddItem(this)">&#43; Add Item</button>'
        '</div>'

        '</div>'

    )


@app.route("/new-exec-summary/<sec>", methods=["POST"])

def new_exec_summary(sec):

    if "email" not in session: return redirect("/")

    note = Note(title="Executive Summary", content=_build_executive_summary_html(),

                section=sec, note_order=_next_order(sec))

    db.session.add(note); db.session.commit()

    return redirect(f"/note/{note.id}")


@app.route("/new-agenda/<sec>", methods=["POST"])

def new_agenda(sec):

    if "email" not in session: return redirect("/")

    g = _require_write()

    if g: return g

    note = Note(title="Agenda", content=_build_agenda_html(),

                section=sec, note_order=_next_order(sec))

    db.session.add(note); db.session.commit()

    return redirect(f"/note/{note.id}")


def _build_cover_slide_html():
    return (
        '''<div class="cover-overlay-wrap" style="position:relative;margin:8px 0 20px;border-radius:16px;overflow:hidden;box-shadow:0 8px 40px rgba(0,0,0,.22);display:inline-block;width:100%;line-height:0;font-size:0;">
  <img src="/cover-slide-bg.png?v=2" style="width:100%;height:auto;display:block;border-radius:16px;pointer-events:none;" draggable="false">
  <div class="cover-drop-layer" contenteditable="false" style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:10;">
  </div>
  <div class="cover-overlay-hint" style="position:absolute;bottom:10px;left:50%;transform:translateX(-50%);background:rgba(0,0,0,.6);color:white;font-size:11px;padding:4px 12px;border-radius:20px;pointer-events:none;white-space:nowrap;font-family:Segoe UI,sans-serif;z-index:20;">Paste or drag an image here to place logo</div>
  <div class="cover-url-label" style="position:absolute;bottom:14px;right:22px;color:white;font-size:13px;font-weight:700;font-family:Segoe UI,sans-serif;z-index:20;pointer-events:none;letter-spacing:.4px;text-shadow:0 1px 4px rgba(0,0,0,.8);">www.varitecconsulting.com</div>
</div>'''
    )


@app.route("/new-cover-slide/<sec>", methods=["POST"])

def new_cover_slide(sec):

    if "email" not in session: return redirect("/")

    g = _require_write()

    if g: return g

    note = Note(title="Cover Slide", content=_build_cover_slide_html(),

                section=sec, note_order=_next_order(sec))

    db.session.add(note); db.session.commit()

    return redirect(f"/note/{note.id}")


def _build_thank_you_html():
    return (
        '''<div class="ty-wrap" style="margin:8px 0 20px !important;border-radius:16px !important;overflow:hidden !important;box-shadow:0 8px 40px rgba(0,0,0,.22) !important;width:100% !important;max-width:100% !important;max-height:none !important;min-height:0 !important;height:auto !important;display:block !important;background:#040b18 !important;padding:0 !important;margin-bottom:0 !important;box-sizing:border-box !important;line-height:0 !important;font-size:0 !important;"><img src="/thankyou-bg.png?v=12" style="width:100%;max-width:100%;height:auto;display:block;border-radius:16px;vertical-align:top;pointer-events:none;"></div>'''
    )


@app.route("/new-thank-you/<sec>", methods=["POST"])

def new_thank_you(sec):

    if "email" not in session: return redirect("/")

    g = _require_write()

    if g: return g

    note = Note(title="Thank You", content=_build_thank_you_html(),

                section=sec, note_order=_next_order(sec))

    db.session.add(note); db.session.commit()

    return redirect(f"/note/{note.id}")


@app.route("/export/pdf/<section>")

def export_pdf_section(section):

    if "email" not in session: return redirect("/")

    notes_list = Note.query.filter_by(section=section).order_by(Note.note_order).all()

    def _has_content(note):

        c = (note.content or '').strip()

        if not c: return False

        from bs4 import BeautifulSoup as _BS

        soup = _BS(c, 'html.parser')

        if soup.find(class_=['si-wrap','ob-wrap','sc-table','si-banner','ob-header-customer',

                              'ob-header-vendor','ob-header-mah','exec-wrap','agenda-wrap',

                              'ty-wrap','cover-wrap','cover-overlay-wrap']): return True

        return len(soup.get_text(strip=True)) > 10

    notes_list = [n for n in notes_list if _has_content(n)]

    if not notes_list:

        from flask import Response

        return Response("<html><body><p>No content to export.</p></body></html>", mimetype='text/html')

    css_start = NOTES_HTML.find('<style>') + 7

    css_end   = NOTES_HTML.find('</style>')

    import re as _re2
    app_css = '\n'.join(_re2.findall(r'<style>(.*?)</style>', NOTES_HTML, _re2.DOTALL))
    app_css = _re2.sub(r'min-height\s*:\s*calc\([^)]*\)', 'min-height:0', app_css)
    app_css = _re2.sub(r'height\s*:\s*calc\(100vh[^)]*\)', 'height:auto', app_css)

    def _split_ob(content):

        # If progress bars exist, skip splitting — browser handles page breaks
        # naturally via break-inside:avoid, and splitting would drop the bars
        if 'ob-progress-fill' in (content or ''): return content

        from bs4 import BeautifulSoup

        soup = BeautifulSoup(content, 'html.parser')

        ob_table = soup.find('table', class_='ob-table')

        if not ob_table: return content

        thead = ob_table.find('thead'); tbody = ob_table.find('tbody')

        if not thead or not tbody: return content

        thead_str = str(thead)

        rows = list(tbody.find_all('tr', recursive=False))

        groups, cur, pending_rows = [], [], []

        for row in rows:

            row_cls = ' '.join(row.get('class') or [])

            is_progress = 'ob-progress-row' in row_cls

            is_sep = (not is_progress and (row.get('data-repeat-hdr') or (not row.get('data-cg') and any(td.get('colspan') for td in row.find_all('td')))))

            if is_progress:

                pending_rows.append(row)  # hold for next group

            elif is_sep:

                if cur: groups.append(cur); cur = []; pending_rows = []

            else:

                if pending_rows: cur.extend(pending_rows); pending_rows = []

                cur.append(row)

        if cur: groups.append(cur)

        if len(groups) <= 1: return content

        blocks = ''

        for gi, group in enumerate(groups):

            pb = ''  # browser decides if group fits on current page

            rows_html = ''.join(str(r) for r in group)

            blocks += (f'<div class="ob-cust-block" style="{pb}">'

                       f'<table class="ob-table" style="width:100%;border-collapse:collapse;">'

                       f'{thead_str}<tbody>{rows_html}</tbody></table></div>')

        from bs4 import BeautifulSoup as BS2

        ob_table.replace_with(BS2(blocks, 'html.parser'))

        return str(soup)

    def _si_paginate_for_pdf(content, cards_per_page=1):
        """Split a Serialization Issues note's cards into fixed-size
        groups -- each group becomes its own PDF page. The
        SERIALIZATION ISSUES banner is shown only on the first page;
        later pages show just the card(s), continuing the flow."""
        from bs4 import BeautifulSoup as _SIBS
        soup2 = _SIBS(content, 'html.parser')
        banner_el = soup2.find(class_='si-banner')
        banner_html = str(banner_el) if banner_el else ''
        grid = soup2.find(class_='si-cards-grid')
        if not grid:
            return [content]
        cards = grid.find_all(class_='si-card', recursive=False)
        if not cards:
            return [content]
        groups = [cards[i:i+cards_per_page] for i in range(0, len(cards), cards_per_page)]
        out = []
        for gi, group in enumerate(groups):
            cols = min(len(group), 2)
            cards_html = ''.join(str(c) for c in group)
            head = banner_html if gi == 0 else ''
            out.append('<div class="si-wrap">' + head +
                        '<div class="si-cards-area"><div class="si-cards-grid" '
                        'style="grid-template-columns:repeat(' + str(cols) + ',1fr)!important;">' +
                        cards_html + '</div></div></div>')
        return out

    pages_html = ""

    for note in notes_list:

        is_tpl = any(x in (note.content or '') for x in

                     ['si-wrap','ob-wrap','sc-table','ob-header-customer',

                      'ob-header-vendor','ob-header-mah','exec-wrap','agenda-wrap',

                      'ty-wrap','cover-wrap','cover-overlay-wrap','fa-wrap','rr-wrap','pi-wrap'])

        title_html = ('' if is_tpl else f'<h2 class="pdf-note-title">{note.title or "Untitled"}</h2>')

        content = note.content or ''
        # Cover slide: explicit 170mm height (A4 landscape content area)
        if 'cover-overlay-wrap' in content:
            # 1. Container: explicit 170mm, no border-radius in PDF
            content = content.replace(
                'position:relative;margin:8px 0 20px;border-radius:16px;overflow:hidden;'
                'box-shadow:0 8px 40px rgba(0,0,0,.22);display:inline-block;width:100%;'
                'line-height:0;font-size:0;',
                'position:relative;margin-top:4mm;margin-bottom:0;overflow:hidden;display:block;'
                'width:100%;height:166mm;')
            # 2. Background image: fill 170mm with crop
            content = content.replace(
                'style="width:100%;height:auto;display:block;'
                'border-radius:16px;pointer-events:none;"',
                'style="width:100%;height:auto;display:block;"')
            # 3. Drop layer: keep logos visible
            content = content.replace(
                'style="position:absolute;top:0;left:0;width:100%;'
                'height:100%;pointer-events:none;z-index:10;"',
                'style="position:absolute;top:0;left:0;width:100%;'
                'height:100%;z-index:10;"')
        # Cover slide: explicit 166mm height + 4mm top gap
        if 'cover-overlay-wrap' in content:
            content = content.replace(
                'position:relative;margin:8px 0 20px;border-radius:16px;overflow:hidden;'
                'box-shadow:0 8px 40px rgba(0,0,0,.22);display:inline-block;width:100%;'
                'line-height:0;font-size:0;',
                'position:relative;margin-top:7mm;margin-bottom:7mm;overflow:hidden;'
                'display:block;width:100%;')
            content = content.replace(
                'style="width:100%;height:auto;display:block;'
                'border-radius:16px;pointer-events:none;"',
                'style="width:100%;height:auto;display:block;"')
            content = content.replace(
                'style="position:absolute;top:0;left:0;width:100%;'
                'height:100%;pointer-events:none;z-index:10;"',
                'style="position:absolute;top:0;left:0;width:100%;'
                'height:100%;z-index:10;"')
        if 'ob-table' in content: content = _split_ob(content)

        # Force full-width on exec/agenda headers via inline styles (beats CSS specificity)

        # Zero min-height on RR/PI wraps via inline style (beats all CSS)
        if 'rr-wrap' in content:
            content = content.replace(
                'class="rr-wrap"',
                'class="rr-wrap" style="min-height:0!important;height:auto!important;"',
                1)
        if 'pi-wrap' in content:
            content = content.replace(
                'class="pi-wrap"',
                'class="pi-wrap" style="min-height:0!important;height:auto!important;"',
                1)
        if 'fa-wrap' in content:
            content = content.replace(
                'class="fa-wrap"',
                'class="fa-wrap" style="min-height:0!important;height:auto!important;"',
                1)

        content = content.replace(

            'class="exec-wrap"',

            'class="exec-wrap" style="width:100%!important;max-width:100%!important;box-sizing:border-box!important;display:block!important;overflow:hidden!important;"'

        )

        content = content.replace(

            'class="exec-header"',

            'class="exec-header" style="width:100%!important;min-width:100%!important;box-sizing:border-box!important;display:flex!important;"'

        )

        content = content.replace(

            'class="agenda-wrap"',

            'class="agenda-wrap" style="width:100%!important;max-width:100%!important;box-sizing:border-box!important;display:block!important;overflow:hidden!important;"'

        )

        content = content.replace(

            'class="agenda-header"',

            'class="agenda-header" style="width:100%!important;min-width:100%!important;box-sizing:border-box!important;display:flex!important;"'

        )

        _spacer_div = '<div class="pdf-page-spacer"></div>'
        _full = any(x in (note.content or "") for x in ["ty-wrap","cover-wrap","cover-overlay-wrap"])
        if 'si-wrap' in content:
            for _si_chunk in _si_paginate_for_pdf(content, cards_per_page=1):
                # No pdf-note-body — break-inside:avoid on it causes blank pages (same fix as rr/pi/fa) [SI_PDF_NO_NOTE_BODY_V1]
                pages_html += '<div class="pdf-note">' + _spacer_div + _si_chunk + '</div>'
        elif any(x in content for x in ['rr-wrap','pi-wrap','fa-wrap']):
            # No pdf-note-body — break-inside:avoid on it causes blank pages for these templates
            pages_html += f'<div class="pdf-note">{title_html}{_spacer_div}{content}</div>'
        else:
            pages_html += f'<div class="pdf-note{" full-slide" if _full else ""}">{title_html}<div class="pdf-note-body">{"" if _full else _spacer_div}{content}</div></div>'

    html = f"""<!DOCTYPE html>

<html><head><meta charset="utf-8"><title>{section} — VariTec Export</title>

<style>

{app_css}

*,*::before,*::after{{box-sizing:border-box;}}

html,body{{margin:0;padding:0;font-family:'Segoe UI',system-ui,sans-serif;background:white;-webkit-print-color-adjust:exact;print-color-adjust:exact;}}

/* Each note/template starts on its own page */
.pdf-note{{break-before:page;page-break-before:always;padding:4px 0 16px;overflow:visible;}}
.pdf-note:first-child{{break-before:auto;page-break-before:auto;}}
/* Content within a template flows naturally (e.g. multiple ob customers) */
.pdf-note-body{{break-inside:avoid;page-break-inside:avoid;}}
/* PI PDF: fix banner/card overlap */
.pi-wrap{{display:flex!important;flex-direction:column!important;overflow:visible!important;}}
.pi-banner{{position:relative!important;z-index:2!important;flex-shrink:0!important;margin-bottom:0!important;}}
.pi-body{{position:relative!important;z-index:1!important;margin-top:0!important;padding-top:12px!important;flex:1!important;}}
.pi-chevron{{display:none!important;}}
.pi-item{{margin-bottom:8px!important;page-break-inside:avoid;break-inside:avoid;}}

/* RR large size for PDF — near-original, no blank page risk */
.rr-wrap{{margin:0!important;border-radius:12px!important;}}
.rr-banner{{padding:15px 22px!important;}}
.rr-banner-title{{font-size:19px!important;letter-spacing:1px!important;}}
.rr-banner-sub{{font-size:12px!important;margin-top:4px!important;}}
.rr-cards-area{{padding:14px!important;border-radius:0 0 12px 12px!important;}}
.rr-cards-grid{{gap:10px!important;}}
.rr-card{{padding:11px 14px!important;gap:11px!important;border-radius:9px!important;border-left-width:4px!important;}}
.rr-num{{width:30px!important;height:30px!important;font-size:13px!important;flex-shrink:0!important;}}
.rr-card-title{{font-size:14px!important;line-height:1.45!important;}}
.rr-card-desc{{font-size:12.5px!important;line-height:1.45!important;margin-top:3px!important;}}
.rr-del-card,.rr-add-card{{display:none!important;}}

.fa-add-btn,.fa-del-card,.fa-color-btn{{display:none!important;}}

/* Zero template min-heights — prevent blank pages */
.rr-wrap,.pi-wrap,.exec-wrap,.agenda-wrap,.si-wrap,.sc-table,.fa-wrap,.ty-wrap,.rr-table-wrap,.pi-body,.pi-banner{{min-height:0!important;height:auto!important;}}
.rr-wrap *,.pi-wrap *{{max-height:none!important;}}


/* Cover & Thank You — full page, no padding, no header/footer margin boxes */

.pdf-note.full-slide{{
  break-before:page!important;
  page-break-before:always!important;
  padding:0!important;
  display:block!important;
  overflow:visible!important;
}}

.pdf-note.full-slide .pdf-note-body{{
  width:100%!important;
  padding:0!important;
  margin:0!important;
  display:block!important;
}}

.pdf-note.full-slide .cover-wrap,

.pdf-note.full-slide .ty-wrap{{

  height:calc(100vh - 40mm)!important;

  max-height:calc(100vh - 40mm)!important;

  min-height:0!important;

  width:100%!important;

  margin:0!important;

  border-radius:10px!important;

  box-shadow:0 4px 24px rgba(0,0,0,.18)!important;

  overflow:hidden!important;

}}


.pdf-page-spacer{{height:5mm;display:block;line-height:0;font-size:0;}}


.pdf-note-title{{font-size:18px;font-weight:800;color:#0d152b;border-bottom:3px solid #038dbd;padding-bottom:8px;margin:0 0 14px;}}
/* Hide 'Cover Slide' title text for cover-overlay-wrap notes */
.pdf-note:has(.cover-overlay-wrap) h2.pdf-note-title{{display:none!important;}}
/* Hide duplicate page title for templates that already show their own banner title (Roles & Responsibilities, Process Improvements) */
.pdf-note:has(.rr-wrap) h2.pdf-note-title{{display:none!important;}}
/* RR/PI: remove min-height so 100vh doesn't create blank pages */
.rr-wrap,.pi-wrap,.exec-wrap,.agenda-wrap,.ty-wrap{{min-height:0!important;height:auto!important;}}
.rr-table-wrap{{min-height:0!important;height:auto!important;}}
.pi-body{{min-height:0!important;}}


.pdf-note:has(.pi-wrap) h2.pdf-note-title{{display:none!important;}}
.pdf-note:has(.fa-wrap) h2.pdf-note-title{{display:none!important;}}
.cover-overlay-hint{{display:none!important;}}
/* Placed logos inside cover-drop-layer keep their position */
.pdf-note.full-slide .cover-overlay-wrap .cover-drop-layer img{{
  position:static!important;
  width:100%!important;height:auto!important;
  object-fit:contain!important;
  top:auto!important;left:auto!important;
}}
.pdf-note.full-slide .cover-overlay-wrap .cover-drop-layer > div{{
  position:absolute!important;
}}

.cover-overlay-wrap{{width:100%!important;margin:0!important;display:block!important;border-radius:4px!important;}}
.cover-overlay-wrap img{{width:100%!important;height:auto!important;display:block!important;}}
.pdf-note.full-slide .cover-overlay-wrap > img{{width:100%!important;height:auto!important;display:block!important;}}
/* Cover slide: fill full page height like Thank You slide */
.pdf-note.full-slide .cover-overlay-wrap{{
  position:relative!important;display:block!important;
  height:auto!important;}}
.pdf-note.full-slide h2.pdf-note-title{{display:none!important;}}


.pdf-note-body{{font-size:13px;line-height:1.7;color:#334155;overflow:visible;}}
/* ── Cover slide: fill page height with header gap ── */
.pdf-note.full-slide .cover-overlay-wrap{{
  height:156mm!important;
  margin-top:4mm!important;
  background:#040b18!important;
  margin-bottom:0!important;
  margin-left:0!important;
  margin-right:0!important;
  width:100%!important;
  position:relative!important;
  display:block!important;
  overflow:hidden!important;
  border-radius:4px!important;
}}
.pdf-note.full-slide .cover-overlay-wrap .cover-drop-layer{{
  position:absolute!important;
  top:0!important;left:0!important;
  width:100%!important;height:100%!important;
}}
.cover-overlay-hint{{display:none!important;}}
/* Placed logos inside cover-drop-layer keep their position */
.pdf-note.full-slide .cover-overlay-wrap .cover-drop-layer img{{
  position:static!important;
  width:100%!important;height:auto!important;
  object-fit:contain!important;
  top:auto!important;left:auto!important;
}}
.pdf-note.full-slide .cover-overlay-wrap .cover-drop-layer > div{{
  position:absolute!important;
}}


/* Chart: show img snapshot, hide canvas in PDF */
#sc-inline-wrap canvas{{display:none!important;}}
#sc-chart-pdf-img{{display:block!important;width:100%!important;height:auto!important;max-height:145mm!important;object-fit:contain!important;border-radius:6px!important;margin:0!important;}}
.sc-chart-del-btn{{display:none!important;}}
/* Hide chart title in PDF — full canvas visible */
#sc-inline-wrap>div:not(#sc-img-config-panel):not(canvas):not(img){{
  display:none!important;
}}
/* Hide scorecard banner in PDF — give full page to chart */
.sc-table+div,.pdf-note-body>div:first-child .si-banner{{display:none!important;}}
.sc-table{{display:none!important;}}
/* Hide scorecard table wrapper too */
#content .sc-table{{display:none!important;}}
#sc-inline-wrap{{width:100%!important;margin:0!important;padding:0!important;border-radius:8px!important;overflow:hidden!important;break-inside:avoid!important;page-break-inside:avoid!important;break-before:auto!important;page-break-before:auto!important;display:block!important;margin-top:3mm!important;}}
.note-footer,.topbar,.sidebar,.pages,.toolbar,.paste-hint,.statusbar,.tb-btn-ppt,

.row-action-bar,#imgResizeOverlay,#datePickerPopup,.date-cal-btn,#ob-add-btn-row,

.exec-del-row,.exec-add-row,.exec-row-handle,

.si-del-card,.si-add-btn,.si-field-del-row,.si-field-add-row,

.ob-header-sub,.ob-month-wrap,.sc-month-wrap,.si-banner-subtitle,

.si-banner-meta,.si-month-wrap,.exec-subtitle,.agenda-subtitle,.agenda-add-row,.agenda-del-item,.agenda-vc,.exec-vc{{display:none!important;}}
[data-repeat-hdr]{{display:table-row!important;break-before:avoid!important;page-break-before:avoid!important;break-after:avoid!important;page-break-after:avoid!important;}}
[data-repeat-hdr] th{{font-size:9px!important;padding:6px 8px!important;font-weight:800!important;text-transform:uppercase!important;letter-spacing:.5px!important;color:white!important;}}

.agenda-wrap,.exec-wrap{{min-height:0!important;width:100%!important;max-width:100%!important;border-radius:6px!important;overflow:hidden!important;box-shadow:none!important;margin:0 0 12px!important;position:relative!important;display:block!important;}}
.agenda-body{{padding:48px 32px!important;gap:26px!important;}}
.agenda-item{{padding:36px 26px!important;min-height:160px!important;}}
.agenda-num{{font-size:100px!important;width:96px!important;}}
.agenda-label{{font-size:21px!important;}}

.exec-header{{width:100%!important;min-width:100%!important;box-sizing:border-box!important;border-radius:0!important;display:flex!important;position:relative!important;float:none!important;}}

.agenda-header{{width:100%!important;min-width:100%!important;box-sizing:border-box!important;border-radius:0!important;display:flex!important;position:relative!important;float:none!important;}}

.agenda-wrap{{border-radius:6px!important;overflow:hidden!important;}}

.exec-body{{width:100%!important;box-sizing:border-box!important;}}
/* Strip the screen drop-shadow from the Serialization Issues card
   in print — it was rendering as a gray smudge below the card */
.si-wrap,.si-wrap *{{
  box-shadow:none!important;
}}


.exec-row-content{{border:none!important;background:transparent!important;padding:2px 0!important;}}

.ob-header-illus{{display:none!important;}}

.ob-header-customer,.ob-header-vendor,.ob-header-mah{{padding:10px 18px!important;break-after:avoid!important;}}

.ob-table-wrap{{overflow:visible!important;width:100%!important;}}

.ob-table{{table-layout:fixed!important;width:100%!important;font-size:9px!important;border-collapse:collapse!important;}}

.ob-table thead{{display:table-header-group;padding-top:4mm;}}

.ob-table tbody tr{{break-inside:avoid!important;}}

.ob-table th,.ob-table td{{padding:4px 5px!important;font-size:9px!important;word-wrap:break-word!important;white-space:normal!important;}}
/* Progress bar row: preserve its padding and height */
.ob-table .ob-progress-row td{{padding:0!important;height:auto!important;border:none!important;}}
.ob-table .ob-progress-row .ob-progress-fill{{transition:none!important;}}
.ob-table .ob-progress-row{{break-inside:avoid;page-break-inside:avoid;}}
.ob-table .ob-spacer-row{{height:32px!important;padding:0!important;display:table-row!important;break-inside:avoid!important;break-after:avoid!important;page-break-after:avoid!important;}}
.ob-table .ob-progress-row{{break-after:avoid!important;page-break-after:avoid!important;break-before:avoid!important;page-break-before:avoid!important;}}
.ob-table tr:has(.ob-progress-fill){{break-after:avoid!important;page-break-after:avoid!important;break-before:avoid!important;page-break-before:avoid!important;}}
.ob-table tr:has(td[colspan][style*='height:5px']){{break-after:avoid!important;page-break-after:avoid!important;}}
.ob-table .ob-spacer-row td{{height:32px!important;padding:0!important;border:none!important;background:#ffffff!important;}}
/* ob-cust-block: stay together when possible, move to next page only if needed */
.ob-cust-block{{break-inside:avoid;page-break-inside:avoid;margin-top:6mm;}}
.ob-cust-block:first-child{{margin-top:0;}}
.si-cards-area{{padding-top:6mm!important;}}


.exec-row{{break-inside:avoid!important;page-break-inside:avoid!important;}}
.exec-section{{break-inside:avoid!important;}}
.exec-row{{margin-top:2mm!important;}}
.exec-body>.exec-section:first-child>.exec-row:first-of-type{{margin-top:0!important;}}

.ob-table th:nth-child(1),.ob-table td:nth-child(1){{width:10%!important;}}

.ob-table th:nth-child(2),.ob-table td:nth-child(2){{width:26%!important;}}

.ob-table th:nth-child(3),.ob-table td:nth-child(3){{width:14%!important;}}

.ob-table th:nth-child(4),.ob-table td:nth-child(4){{width:13%!important;}}

.ob-table th:nth-child(5),.ob-table td:nth-child(5){{width:13%!important;}}

.ob-table th:nth-child(6),.ob-table td:nth-child(6){{width:24%!important;}}

.ob-table .ob-status-btn{{font-size:8px!important;padding:3px 6px!important;white-space:nowrap!important;display:inline-block!important;}}

.ob-table td:nth-child(3){{padding:3px 2px!important;text-align:center!important;}}

.ob-table{{border:2px solid var(--ob-border,#1e3a8a)!important;}}
.ob-table td:last-child,.ob-table th:last-child,.ob-table tr:last-child td{{border-color:var(--ob-border,#1e3a8a)!important;}}
.ob-table td[rowspan]{{border-color:#e2e8f0!important;border-bottom-color:var(--ob-border,#1e3a8a)!important;}}
.ob-table td[colspan]{{padding:0!important;height:auto!important;border-top:none!important;border-bottom:1px solid #e2e8f0!important;}}
.sc-table{{table-layout:fixed!important;width:100%!important;font-size:9px!important;border-collapse:collapse!important;break-inside:avoid;page-break-inside:avoid;}}

.sc-table thead{{display:table-header-group;}}

.sc-table tbody tr{{break-inside:avoid!important;}}

.si-cards-grid{{display:grid!important;grid-template-columns:repeat(2,1fr)!important;gap:6px!important;align-items:start!important;}}
.si-card{{min-width:0!important;width:100%!important;}}
.si-cards-area{{padding:10px 4px 4px!important;}}
.si-card-body{{padding:8px 10px!important;}}
.si-field-val{{font-size:10px!important;padding:4px 7px!important;line-height:1.4!important;}}
.si-field-label{{font-size:8px!important;}}
.si-card-title{{font-size:12px!important;}}

.si-card{{break-inside:avoid!important;margin-top:0!important;}}
.si-tl-step{{break-inside:avoid!important;}}
.si-action-add-btn,.si-action-del{{display:none!important;}}

.si-banner{{break-after:avoid!important;}}

.si-cards-area{{break-before:avoid!important;}}

.pdf-note:has(#sc-inline-wrap){{
  padding:0!important;
  margin:0!important;
}}
@media print{{

  @page{{size:A4 landscape;margin:23mm 10mm 20mm 10mm/*PDF_PAGE_HEADER_GAP_V1*/;

    @top-left{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;content:"";background-image:url('data:image/png;base64,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AAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0BAAAAAGiG0BsAAAAAgGYIvQEAAAAAaIbQGwAAAACAZgi9AQAAAABohtAbAAAAAIBmCL0B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AAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAAAAAAAAAAAIAOGfYGAAAAAAAAAAAAAOiQYW8AAAAAAAAAAAAAgA4Z9gYAAAAAAAAAAAAA6JBhbwAAAAAAAAAAAACADhn2BgAAAAAAAAAAAADokGFvAAAA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AIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAAAAAAAAAAABAj9AYAAAAAAAAAAAAAiBF6AwAAAAAAAAAAAADECL0BAAAAAAAAAAAAAGKE3gAAAAAAAAAAAAAAMUJvAAAAAAAAAAAAAIAYoTcAAAAAAAAAAAAAQIzQGwAAAAAAAAAAAAAgRugNAAAA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AAAAAAAAAACDU28AAAAAAAAALjt3IAAAAAAgyN96hAUKJAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAAAAAAAAAAAAAAAAAAACG1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhtQbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIbUGwAAAAAAAAAAIPbu9yqqPHv78N6z5j10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AAAAAAAAAAAAFjDqDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAALGPUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAWMegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAAka9AQAAAAAAAAAAAAAAAAAAAAAAAAAAAGABo94AAAAAAAAAAAAAAAAAAAAAAAAAAAAAsIBRbwAAAAAAAAAAAAAAAAAAAAAAAAAAAABYwKg3AAAAAAAAAAAAAAAAAAAAAAAAAAAAACxg1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFjDqDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAALGPUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAWMegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAAka9AQAAAAAAAAAAAAAAAAAAAAAAAAAAAGABo94AAAAAAAAAAAAAAAAAAAAAAAAAAAAAsIBRbwAAAAAAAAAAAAAAAAAAAAAAAAAAAABYwKg3AAAAAAAAAAAAAAAAAAAAAAAAAAAAACxg1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFjDqDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAALGPUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAWMegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAAka9AQAAAAAAAAAAAAAAAAAAAAAAAAAAAGABo94AAAAAAAAAAAAAAAAAAAAAAAAAAAAAsIBRbwAAAAAAAAAAAAAAAAAAAAAAAAAAAABYwKg3AAAAAAAAAAAAAAAAAAAAAAAAAAAAACxg1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFjDqDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAALGPUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAWMegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAAka9AQAAAAAAAAAAAAAAAAAAAAAAAAAAAGABo94AAAAAAAAAAAAAAAAAAAAAAAAAAAAAsIBRbwAAAAAAAAAAAAAAAAAAAAAAAAAAAABYwKg3AAAAAAAAAAAAAAAAAAAAAAAAAAAAACxg1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFjDqDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAALGPUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAWMegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAAka9AQAAAAAAAAAAAAAAAAAAAAAAAAAAAGABo94AAAAAAAAAAAAAAAAAAAAAAAAAAAAAsIBRbwAAAAAAAAAAAAAAAAAAAAAAAAAAAABYwKg3AAAAAAAAAAAAAAAAAAAAAAAAAAAAACxg1BsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFjDqDQAAAAAAAAAAAAAAAAAAAAAAAAAAAAALGPUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAWMegMAAAAAAAAAAAAAwP+zcwcCAAAAAIL8rUdYoEACAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAABuxs0UAAAsQSURBVAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYUm8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAGFJvAAAAAAAAAAAAAAAAAAAAAAAAAAAAABhSbwAAAAAAAAAAAAAAAABq5w4EAAAAAAT5W4+wQIEEAAAAAAAAAADAkHoDAAAAAAAAAAAAAAAAAAAAAAAAAAAAwJB6AwAAAAAAAAAAAAAAAAAAAAAAAAAAAMCQegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAkHoDAAAAAAAAAAAAAAAAAAAAAAAAAAAAwJB6AwAAAAAAAAAAAAAAAAAAAAAAAAAAAMCQegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAkHoDAAAAAAAAAAAAAAAAAAAAAAAAAAAAwJB6AwAAAAAAAAAAAAAAAAAAAAAAAAAAAMCQegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAkHoDAAAAAAAAAAAAAAAAAAAAAAAAAAAAwJB6AwAAAAAAAAAAAAAAAAAAAAAAAAAAAMCQegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAkHoDAAAAAAAAAAAAAAAAAAAAAAAAAAAAwJB6AwAAAAAAAAAAAAAAAAAAAAAAAAAAAMCQegMAAAAAAAAAAAAAAAAAAAAAAAAAAADAkHoDAAAAAAAAAAAAAAAAAAAAAAAAAAAAwJB6AwAAAAAAAAAAAAAAAAAAAAAAAAAAAMAIwZBwQo9f/BMAAAAASUVORK5CYII=');background-repeat:no-repeat;background-position:left bottom;background-size:auto 34px;font-family:"Arial Black",Arial,sans-serif;font-size:11pt;font-weight:900;color:#0d152b;letter-spacing:1px;border-bottom:3px solid #038dbd;padding-bottom:4pt;vertical-align:bottom;}}

    @top-center{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;content:"";border-bottom:3px solid #038dbd;}}

    @top-right{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;content:"";background-image:url('data:image/png;base64,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');background-repeat:no-repeat;background-position:right bottom;background-size:auto 34px;padding:0 12px 5px 0;box-sizing:border-box;border-bottom:3px solid #038dbd;}}

    @bottom-left{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;content:"www.varitecconsulting.com";font-family:"Segoe UI",Arial,sans-serif;font-size:9pt;color:#334155;font-weight:600;border-top:3px solid #038dbd;padding-top:4pt;vertical-align:top;}}

    @bottom-center{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;content:"";border-top:3px solid #038dbd;}}

    @bottom-right{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;content:"Page " counter(page) " of " counter(pages);font-family:"Segoe UI",Arial,sans-serif;font-size:10pt;font-weight:800;color:#0d152b;border-top:3px solid #038dbd;padding-top:3pt;vertical-align:top;white-space:nowrap;}}

  }}

  /* Full-bleed page for cover/thank-you — override @page margins */

  /* full-bleed override removed */

  html,body{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;overflow:visible!important;}}

  .pdf-note{{overflow:visible!important;}}

  .ob-table thead,.sc-table thead{{display:table-row-group!important;}}

  .ob-table tbody tr,.sc-table tbody tr{{break-inside:avoid!important;}}

  .si-card{{break-inside:avoid!important;margin-top:0!important;}}

  .agenda-wrap,.exec-wrap{{min-height:0!important;width:100%!important;max-width:100%!important;box-shadow:none!important;display:block!important;}}

  .exec-header,.agenda-header{{width:100%!important;min-width:100%!important;box-sizing:border-box!important;display:flex!important;float:none!important;}}

  .ob-table,.ob-table th,.ob-table td{{border:none!important;}}
  .ob-table-wrap{{border:none!important;}}

}}

</style></head><body>


<div id="pdf-body">{pages_html}</div>

<script>
var _printDone=false;
function _doPrint(){{
  if(_printDone)return;_printDone=true;
  /* Zero min-heights before printing — prevents blank pages */
  var _sel='.rr-wrap,.rr-body,.rr-table-wrap,.pi-wrap,.pi-body,.pi-banner,.fa-wrap,.exec-wrap,.agenda-wrap,.si-wrap,.sc-table,.agenda-body';
  document.querySelectorAll(_sel).forEach(function(el){{
    el.style.setProperty('min-height','0','important');
    el.style.setProperty('height','auto','important');
    el.style.setProperty('max-height','none','important');
  }});
  /* PDF_DISABLE_SPELLCHECK_V1: suppress red spellcheck squiggles in the printed/exported PDF */
  var _pdfBody=document.getElementById('pdf-body');
  if(_pdfBody){{
    _pdfBody.setAttribute('spellcheck','false');
    _pdfBody.querySelectorAll('[contenteditable]').forEach(function(el){{el.setAttribute('spellcheck','false');}});
  }}
  window.print();
}}
setTimeout(_doPrint,4000);
window.onload=function(){{setTimeout(_doPrint,800);}};
</script>
</body></html>"""

    from flask import Response

    return Response(html, mimetype='text/html')



@app.route("/new-roles-responsibilities/<sec>", methods=["POST"])
def new_roles_responsibilities(sec):
    if "email" not in session: return redirect("/")
    g = _require_write()
    if g: return g
    note = Note(
        title="Roles & Responsibilities",
        content=_build_roles_responsibilities_html(),
        section=sec,
        note_order=_next_order(sec),
    )
    db.session.add(note); db.session.commit()
    return redirect(f"/note/{note.id}")



@app.route("/new-process-improvements/<sec>", methods=["POST"])
def new_process_improvements(sec):
    if "email" not in session: return redirect("/")
    g = _require_write()
    if g: return g
    note = Note(
        title="Ongoing / Completed Process Improvements",
        content=_build_process_improvements_html(),
        section=sec,
        note_order=_next_order(sec),
    )
    db.session.add(note); db.session.commit()
    return redirect(f"/note/{note.id}")




def _build_focus_areas_html():
    def _card(color, title='', impact='High Impact', issue='', imp='', action=''):
        return (
            f'''<div class="fa-card" data-fa-color="{color}">'''
            + f'''<div class="fa-card-head" style="background:{color};">'''
            + f'''<div class="fa-card-title" contenteditable="true">{title}</div>'''
            + '''<button class="fa-color-btn" onclick="faCycleColor(this)" title="Change colour"></button>'''
            + '''</div>'''
            + '''<div class="fa-card-inner">'''
            + '''<div class="fa-impact-row"><span class="fa-impact-dot"></span>'''
            + f'''<span class="fa-impact-text" contenteditable="true">{impact}</span></div>'''
            + '''<div class="fa-section"><div class="fa-section-label">Issue</div>'''
            + f'''<div class="fa-section-body" contenteditable="true" data-ph="Describe the issue...">{issue}</div></div>'''
            + '''<div class="fa-section"><div class="fa-section-label">Impact</div>'''
            + f'''<div class="fa-section-body" contenteditable="true" data-ph="Describe the impact...">{imp}</div></div>'''
            + '''<div class="fa-section"><div class="fa-section-label">Action</div>'''
            + f'''<div class="fa-section-body" contenteditable="true" data-ph="Describe the action...">{action}</div></div>'''
            + '''</div>'''
            + '''<button class="fa-del-card" onclick="faDelCard(this)">&#10005;</button>'''
            + '''</div>'''
        )
    cards = _card('#2563eb') + _card('#0f172a')
    return (
        '''<div class="fa-wrap">'''
        + '''<div class="fa-banner"><span class="fa-banner-title">Current Focus Areas &amp; Challenges</span></div>'''
        + f'''<div class="fa-grid">{cards}</div>'''
        + '''<button class="fa-add-btn" onclick="faAddCard(this)">+ Add Focus Area</button>'''
        + '''</div>'''
    )


@app.route("/new-focus-areas/<sec>", methods=["POST"])
def new_focus_areas(sec):
    if "email" not in session: return redirect("/")
    g = _require_write()
    if g: return g
    note = Note(
        title="Current Focus Areas & Challenges",
        content=_build_focus_areas_html(),
        section=sec,
        note_order=_next_order(sec),
    )
    db.session.add(note); db.session.commit()
    return redirect(f"/note/{note.id}")



@app.route("/clone/<int:note_id>", methods=["POST"])
def clone_note(note_id):
    if "email" not in session: return redirect("/")
    g = _require_write()
    if g: return g
    original = Note.query.get_or_404(note_id)
    clone = Note(
        title=original.title + " (Copy)",
        content=original.content,
        section=original.section,
        note_order=_next_order(original.section),
    )
    db.session.add(clone); db.session.commit()
    return redirect(f"/note/{clone.id}")


@app.route("/clone-section/<int:section_id>", methods=["POST"])
def clone_section(section_id):
    if "email" not in session: return redirect("/")
    g = _require_write()
    if g: return g
    # Find the original section
    from sqlalchemy import func
    orig_sec = Section.query.get_or_404(section_id)
    # Generate a unique new name
    base_name = orig_sec.name + " (Copy)"
    new_name  = base_name
    counter   = 1
    existing  = {s.name for s in Section.query.all()}
    while new_name in existing:
        counter  += 1
        new_name  = base_name + " " + str(counter)
    # Create new section
    max_order = db.session.query(func.max(Section.order)).scalar() or 0
    new_sec = Section(name=new_name, order=max_order + 1)
    db.session.add(new_sec)
    db.session.flush()  # get new_sec.id
    # Copy all notes from original section
    orig_notes = Note.query.filter_by(section=orig_sec.name).order_by(Note.note_order).all()
    for i, note in enumerate(orig_notes):
        new_note = Note(
            title=note.title,
            content=note.content,
            section=new_name,
            note_order=i,
        )
        db.session.add(new_note)
    db.session.commit()
    return redirect(f"/section/{new_name}")


@app.route("/section/add-sub", methods=["POST"])
def add_sub_section():
    if "email" not in session: return redirect("/")
    g = _require_write()
    if g: return g
    data = request.get_json()
    name = (data.get("name") or "").strip()
    parent_id = data.get("parent_id")
    if not name:
        return jsonify(status="error", message="Name required")
    parent = db.session.get(Section, parent_id) if parent_id else None
    # Store as "ParentName > MonthName" to ensure global uniqueness
    # while allowing same month name under different clients
    full_name = f"{parent.name} > {name}" if parent else name
    # If somehow still exists, append counter
    base = full_name; counter = 1
    while Section.query.filter_by(name=base).first():
        counter += 1; base = f"{full_name} ({counter})"
    full_name = base
    max_order = db.session.query(db.func.max(Section.order)).scalar() or 0
    sec = Section(name=full_name, order=max_order+1, parent_id=parent.id if parent else None)
    db.session.add(sec); db.session.commit()
    return jsonify(status="ok", id=sec.id, name=full_name, display=name)


@app.route("/clone-month/<int:section_id>", methods=["POST"])
def clone_month_section(section_id):
    if "email" not in session: return redirect("/")
    g = _require_write()
    if g: return g
    orig = db.session.get(Section, section_id)
    if not orig: return jsonify(status="error", message="Section not found")
    # Generate unique name for the clone
    base = orig.name + " (Copy)"
    new_name = base; counter = 1
    while Section.query.filter_by(name=new_name).first():
        counter += 1; new_name = base + f" {counter}"
    max_order = db.session.query(db.func.max(Section.order)).scalar() or 0
    clone_sec = Section(name=new_name, order=max_order+1, parent_id=orig.parent_id)
    db.session.add(clone_sec); db.session.flush()
    # Copy all notes
    for note in Note.query.filter_by(section=orig.name).order_by(Note.note_order).all():
        db.session.add(Note(title=note.title, content=note.content,
                            section=new_name, note_order=note.note_order))
    db.session.commit()
    return jsonify(status="ok", name=new_name)

if __name__ == "__main__":

    app.run(debug=True, host="0.0.0.0", port=5000)
