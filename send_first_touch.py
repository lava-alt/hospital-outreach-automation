#!/usr/bin/env python3
"""
Boston Hospital Outreach — First Touch Sender
==============================================
Adapted from Alumni Leader Outreach Automation (send_first_touch.py).
Handles: CSV queue management, template filling, 12-check validation, CRM updates.
The personalization hooks (p1, p2) and Chrome MCP Outlook flow are handled by Claude
reading outreach/SKILL.md.

Usage:
  python3 send_first_touch.py list                        # Show all contacts in send queue
  python3 send_first_touch.py next                        # Print next eligible contact as JSON
  python3 send_first_touch.py preview <row_idx>           # Preview filled email (Claude provides p1/p2)
  python3 send_first_touch.py validate                    # Validate draft (reads JSON from stdin)
  python3 send_first_touch.py compute-send-time <row_idx> <days> <time_range> [--recipient-tz TZ]
                                                          # Convert recipient-local time → ET for Outlook
  python3 send_first_touch.py mark-sent <row_idx>         # Status → Contacted no reply
  python3 send_first_touch.py mark-failed <row_idx>       # Status → Not sent (validation failure)

validate stdin JSON format:
  {"row_idx": N, "p1": "...", "p2": "...", "geo": "..."}   # geo optional; auto-derived from Location if omitted

compute-send-time examples:
  python3 send_first_touch.py compute-send-time 15 monday,tuesday 12:30-14:00
  python3 send_first_touch.py compute-send-time 15 2026-06-02 12:30-14:00 --recipient-tz America/Chicago
"""

import csv
import fcntl
import json
import re
import sys
import random
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

# ── Paths ──────────────────────────────────────────────────────────────────────
# Override CRM with env var OUTREACH_CRM_PATH (used for test runs against a fake CRM).
import os
BASE = Path("/Users/lavapanta/Desktop/Hospital prospecting, outreach reachout automation")
_default_crm = BASE / "prospecting" / "Boston Hospital Outreach - CRM.csv"
CRM_PATH = Path(os.environ.get("OUTREACH_CRM_PATH", _default_crm))
LOCK_PATH = str(CRM_PATH) + ".lock"   # Separate lock file (borrowed from alumni script)

# Session gate + last-session memory. Tied to CRM path so test runs (env override)
# get isolated session files, same as the lock pattern.
SESSION_PATH = str(CRM_PATH) + ".session.json"        # active session (expires)
LAST_SESSION_PATH = str(CRM_PATH) + ".last_session.json"  # persists across runs
SESSION_TTL_HOURS = 12
LAST_SESSION_TTL_DAYS = 30   # forget the repeat-offer record after this many days

# ── Sender identity ────────────────────────────────────────────────────────────
SENDER_NAME = "Lava Panta"

# ── Email template ─────────────────────────────────────────────────────────────
# {name}  → recipient first name only
# {p1}    → opener hook (continues "I saw ___")
# {p2}    → P.S. hook (different angle)
# {geo}   → geography phrase, e.g. "Boston medical system", "New York medical system"
EMAIL_TEMPLATE = """\
Hey {name},

I'm Lava Panta. I saw {p1}. I'm a Northeastern undergrad working with Barock Tesfaye, who shadowed doctors in Ethiopia and made a documentary about what they go through. The biggest issue was a shortage of supplies. We've raised money through the documentary, purchased ultrasound machines for those hospitals, and gotten in touch with Project C.U.R.E.

Now we're trying to learn from people in the {geo} who understand this work, to better understand what more could realistically be done for hospitals back home.

We'd love your advice. A 10-minute call this week would be invaluable. Thanks for taking the time to read this!

Sincerely,
Lava Panta
Barock C. Tesfaye

P.S. {p2}"""

SUBJECT = "Trying to help people back home, would love your advice!"

# ── Doctoral credentials that warrant "Dr. LastName" greeting ─────────────────
DOCTORAL_CREDENTIALS = {"MD", "DO", "PHD", "DNP", "DDS", "DMD", "EDD", "DVM", "MBBS"}


# ═══════════════════════════════════════════════════════════════════════════════
# CSV helpers (lock pattern adapted from alumni script)
# ═══════════════════════════════════════════════════════════════════════════════

def read_csv():
    """Read CRM and return (rows, fieldnames). Rows are list of dicts."""
    with open(CRM_PATH, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fieldnames = list(reader.fieldnames)
    return rows, fieldnames


def write_csv(rows, fieldnames):
    """Write rows back with exclusive lock on separate .lock file."""
    lock_fd = open(LOCK_PATH, "w")
    try:
        fcntl.flock(lock_fd, fcntl.LOCK_EX)
        with open(CRM_PATH, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    finally:
        fcntl.flock(lock_fd, fcntl.LOCK_UN)
        lock_fd.close()


# ═══════════════════════════════════════════════════════════════════════════════
# Session gate + last-session memory
# ═══════════════════════════════════════════════════════════════════════════════
# `next` refuses to dispense contacts until an active session is set by `session-start`.
# This forces the Interactive Start Flow (mode / timezone / scope) to run every session,
# even in headless/scheduled-task mode where prose instructions might be skipped.
# Active session expires after SESSION_TTL_HOURS so a stale config can't silently
# auto-run a future batch. Each start also writes a durable last-session record used
# to offer "rerun the same session as last time".

VALID_MODES = {"immediate", "schedule", "draft"}
VALID_TZ = {"same", "different"}
VALID_SCOPE = {"queue", "count", "names"}


def _read_json(path: str):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def load_active_session():
    """Return (is_active, config_or_None). Expired sessions are treated as inactive."""
    cfg = _read_json(SESSION_PATH)
    if not cfg:
        return False, None
    expires = cfg.get("expires_at")
    if expires:
        try:
            if datetime.now() > datetime.fromisoformat(expires):
                return False, cfg  # return cfg so caller can report expiry
        except ValueError:
            return False, cfg
    return True, cfg


def cmd_session_start():
    """
    Read session config JSON from stdin, validate, write active session + last-session record.
    stdin JSON:
      {
        "mode": "immediate|schedule|draft",
        "timezone": "same|different",     (optional, default "same")
        "scope": "queue|count|names",
        "days": "monday,tuesday",          (required when mode == schedule)
        "time_range": "12:30-14:00",       (required when mode == schedule)
        "count": 5,                          (used when scope == count)
        "names": ["Gene Bukhman", ...]      (used when scope == names)
      }
    """
    try:
        payload = json.load(sys.stdin)
    except json.JSONDecodeError as e:
        print(json.dumps({"ok": False, "error": f"Invalid JSON on stdin: {e}"}))
        sys.exit(1)

    mode = str(payload.get("mode", "")).strip().lower()
    tz = str(payload.get("timezone", "same")).strip().lower()
    scope = str(payload.get("scope", "queue")).strip().lower()

    errs = []
    if mode not in VALID_MODES:
        errs.append(f"mode must be one of {sorted(VALID_MODES)} (got '{mode}')")
    if tz not in VALID_TZ:
        errs.append(f"timezone must be one of {sorted(VALID_TZ)} (got '{tz}')")
    if scope not in VALID_SCOPE:
        errs.append(f"scope must be one of {sorted(VALID_SCOPE)} (got '{scope}')")
    if mode == "schedule":
        if not str(payload.get("days", "")).strip():
            errs.append("days required when mode == schedule (e.g. 'monday,tuesday')")
        if not str(payload.get("time_range", "")).strip():
            errs.append("time_range required when mode == schedule (e.g. '12:30-14:00')")
    if scope == "count" and not str(payload.get("count", "")).strip().isdigit():
        errs.append("count (positive int) required when scope == count")
    if scope == "names" and not payload.get("names"):
        errs.append("names (non-empty list) required when scope == names")

    if errs:
        print(json.dumps({"ok": False, "error": "; ".join(errs)}))
        sys.exit(1)

    now = datetime.now()
    cfg = {
        "mode": mode,
        "timezone": tz,
        "scope": scope,
        "days": str(payload.get("days", "")).strip(),
        "time_range": str(payload.get("time_range", "")).strip(),
        "count": int(payload["count"]) if scope == "count" else None,
        "names": list(payload.get("names", [])) if scope == "names" else [],
        "created_at": now.isoformat(timespec="seconds"),
        "expires_at": (now + timedelta(hours=SESSION_TTL_HOURS)).isoformat(timespec="seconds"),
    }

    with open(SESSION_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)

    # Durable last-session record (drop expiry; add saved_at).
    last = {k: v for k, v in cfg.items() if k != "expires_at"}
    last["saved_at"] = now.isoformat(timespec="seconds")
    with open(LAST_SESSION_PATH, "w", encoding="utf-8") as f:
        json.dump(last, f, indent=2)

    print(json.dumps({"ok": True, "session": cfg}, indent=2))


def load_last_session():
    """Return last-session record, or None if absent or older than LAST_SESSION_TTL_DAYS.
    Expired records are deleted so they stop being offered."""
    last = _read_json(LAST_SESSION_PATH)
    if not last:
        return None
    saved = last.get("saved_at")
    if saved:
        try:
            if datetime.now() - datetime.fromisoformat(saved) > timedelta(days=LAST_SESSION_TTL_DAYS):
                os.remove(LAST_SESSION_PATH)
                return None
        except (ValueError, OSError):
            pass
    return last


def cmd_session_show():
    """Print the active session (if any) and the last-session record for the repeat offer."""
    active, cfg = load_active_session()
    last = load_last_session()
    print(json.dumps({
        "active": active,
        "active_session": cfg if active else None,
        "expired_session": cfg if (cfg and not active) else None,
        "last_session": last,
        "has_last": last is not None,
    }, indent=2))


def cmd_session_end():
    """Clear the active session (last-session record is kept)."""
    existed = os.path.exists(SESSION_PATH)
    if existed:
        os.remove(SESSION_PATH)
    print(json.dumps({"ok": True, "cleared": existed}))


# ═══════════════════════════════════════════════════════════════════════════════
# Personalization helpers
# ═══════════════════════════════════════════════════════════════════════════════

def extract_first_name(full_name: str) -> str:
    """
    'Michelle L. Niescierenko, MD, MPH'  → 'Michelle'
    'Gene Bukhman, MD, PhD'              → 'Gene'
    'Lewis Seton'                         → 'Lewis'
    """
    base = full_name.split(",")[0].strip()
    tokens = base.split()
    if not tokens:
        return full_name
    honorifics = {"Dr.", "Prof.", "Mr.", "Ms.", "Mrs.", "Sir", "Rev."}
    first = tokens[0]
    if first in honorifics and len(tokens) > 1:
        first = tokens[1]
    return first


def extract_last_name(full_name: str) -> str:
    """
    'Michelle L. Niescierenko, MD, MPH'  → 'Niescierenko'
    'Gene Bukhman, MD, PhD'              → 'Bukhman'
    'Thea L. James, MD, MBA'             → 'James'
    """
    base = full_name.split(",")[0].strip()
    tokens = base.split()
    honorifics = {"Dr.", "Prof.", "Mr.", "Ms.", "Mrs.", "Sir", "Rev."}
    tokens = [t for t in tokens if t not in honorifics]
    return tokens[-1] if tokens else full_name


def is_doctor(full_name: str) -> bool:
    """Return True if name contains any doctoral credential (MD, DO, PhD, DNP, etc.)."""
    parts = full_name.split(",")
    if len(parts) < 2:
        return False
    credentials_str = " ".join(parts[1:]).upper()
    for cred in DOCTORAL_CREDENTIALS:
        if re.search(r'\b' + cred + r'\b', credentials_str):
            return True
    return False


def extract_greeting_name(full_name: str) -> str:
    """
    Returns the correct greeting name:
    - Doctor (MD/DO/PhD/DNP/etc.) → 'Dr. LastName'
    - Everyone else               → 'FirstName'

    Examples:
      'Emily Wroe, MD, MPH'              → 'Dr. Wroe'
      'Gene Bukhman, MD, PhD'            → 'Dr. Bukhman'
      'Alice Tang, PhD, MSc'             → 'Dr. Tang'
      'Lewis Seton'                       → 'Lewis'
      'Christine Brown'                   → 'Christine'
    """
    if is_doctor(full_name):
        return f"Dr. {extract_last_name(full_name)}"
    return extract_first_name(full_name)


DEFAULT_GEO = "Boston medical system"

# Suburb / borough → metro hub. A contact in a hub's metro gets the hub's name
# (Cambridge → Boston, Brooklyn → New York). Keys are lowercased city names.
METRO_ALIASES = {
    # Greater Boston
    "cambridge": "Boston", "somerville": "Boston", "brookline": "Boston",
    "newton": "Boston", "quincy": "Boston", "medford": "Boston",
    "malden": "Boston", "watertown": "Boston", "waltham": "Boston",
    "chelsea": "Boston", "revere": "Boston", "everett": "Boston",
    "boston area": "Boston", "greater boston": "Boston",
    # NYC metro
    "brooklyn": "New York", "manhattan": "New York", "queens": "New York",
    "the bronx": "New York", "bronx": "New York", "staten island": "New York",
    "new york city": "New York", "nyc": "New York",
}


def derive_geo(personalization: str) -> str:
    """
    Build the geography phrase from the Location bullet in the personalization blob.

    'Location: Boston, MA (America/New_York)'     → 'Boston medical system'
    'Location: Cambridge, MA (America/New_York)'  → 'Boston medical system'   (metro rollup)
    'Location: New York, NY (America/New_York)'   → 'New York medical system'
    'Location: Brooklyn, NY (America/New_York)'   → 'New York medical system' (metro rollup)

    Falls back to DEFAULT_GEO when no Location line is present (preserves the
    original Boston wording for legacy rows).
    """
    m = re.search(r'Location:\s*(.+)', personalization or "")
    if not m:
        return DEFAULT_GEO
    loc = m.group(1)
    # Strip trailing timezone paren and any '• ...' source notes.
    loc = re.split(r'[(•]', loc)[0]
    # City portion is everything before the first comma (drops state/country).
    city = loc.split(",")[0].strip()
    if not city:
        return DEFAULT_GEO
    # Roll suburbs/boroughs up to their metro hub.
    city = METRO_ALIASES.get(city.lower(), city)
    if re.search(r'\bsystems?\b', city, re.IGNORECASE):
        return city  # already phrased as a system
    return f"{city} medical system"


def fill_template(name: str, p1: str, p2: str, geo: str = DEFAULT_GEO) -> str:
    return EMAIL_TEMPLATE.format(name=name, p1=p1, p2=p2, geo=geo)


def count_words(text: str) -> int:
    return len(text.split())


def is_eligible(row: dict) -> bool:
    """Contact is in the send queue if: status blank, email present, personalization present."""
    return (
        row.get("Status", "").strip() == ""
        and bool(row.get("Email", "").strip())
        and bool(row.get("Personalization", "").strip())
    )


# ═══════════════════════════════════════════════════════════════════════════════
# Validation — 12 deterministic checks
# ═══════════════════════════════════════════════════════════════════════════════

def validate_draft(body: str, first_name: str, p1: str, p2: str) -> list[dict]:
    """
    Returns list of issue dicts. Empty list = all checks passed.
    Each issue: {"check": N, "severity": "high"|"medium", "message": "...", "fix": "..."}

    Checks:
     1  {name} placeholder filled
     2  {p1} placeholder filled
     3  {p2} placeholder filled
     4  No other stray {…} tokens
     5  Word count <= 150 (greeting through P.S. inclusive)
     6a No em dash (—)
     6b No double-hyphen (--)
     7  First name clean (no credentials, <= 2 tokens)
     8  p1 substantive (>= 15 chars)
     9  p2 substantive (>= 25 chars)
    10  p1 and p2 are different hooks
    11  No ask-level violations (donation / device / partnership-request language)
    12  p1 not a generic title-only opener (heuristic)
    """
    issues = []

    def fail(n, msg, fix, severity="high"):
        issues.append({"check": n, "severity": severity, "message": msg, "fix": fix})

    # 1
    if "{name}" in body or "{Name}" in body:
        fail(1, "Template {name} placeholder still present",
             "Extract first name and pass to fill_template()")

    # 2
    if "{p1}" in body or "{personalisation}" in body:
        fail(2, "Template {p1} placeholder still present",
             "Generate p1 from personalization blob")

    # 3
    if "{p2}" in body or "{personalisation2}" in body:
        fail(3, "Template {p2} placeholder still present",
             "Generate p2 from personalization blob")

    # 4
    stray = re.findall(r"\{[^}]{1,40}\}", body)
    if stray:
        fail(4, f"Stray placeholder(s) in draft: {stray}",
             "Remove or fill all {…} tokens")

    # 5
    wc = count_words(body)
    if wc > 150:
        fail(5, f"Word count {wc} exceeds 150-word hard limit",
             "Shorten p1 or p2 — aim for 8-12 words each; every word counts")

    # 6a
    if "—" in body:
        fail(6, "Em dash (—) found in draft",
             "Replace with a single hyphen (-) or rephrase")

    # 6b — em-dash substitute, also rejected
    if "--" in body:
        fail(6, "Double-hyphen (--) found — em-dash substitute not allowed",
             "Replace with a single hyphen (-) or rephrase")

    # 7. Greeting name clean — allow "Dr. LastName" or plain "FirstName"
    if len(first_name.split()) > 3:
        fail(7, f"Greeting name '{first_name}' too long — may contain credentials",
             "Should be 'Dr. LastName' for doctors or 'FirstName' for everyone else")
    if any(ch in first_name for ch in [",", ";", "("]):
        fail(7, f"Greeting name '{first_name}' contains unexpected punctuation",
             "Use extract_greeting_name() to get 'Dr. LastName' or 'FirstName'")
    if "." in first_name and not first_name.startswith("Dr."):
        fail(7, f"Greeting name '{first_name}' has unexpected period — may contain credentials",
             "Use extract_greeting_name() to get clean greeting name")

    # 8
    if len(p1.strip()) < 15:
        fail(8, f"p1 too short ({len(p1.strip())} chars) — must be specific",
             "Write 1 concrete clause about a real project, paper, or program from their bio")

    # 9
    if len(p2.strip()) < 25:
        fail(9, f"p2 too short ({len(p2.strip())} chars) — P.S. hook must be specific",
             "Write 1-2 sentences: a different credential, named program, or pointed question",
             severity="medium")

    # 10
    if p1.strip().lower() == p2.strip().lower():
        fail(10, "p1 and p2 are identical — must be two distinct personalization hooks",
             "Use a different angle for p2 (e.g. p1 = project they led, p2 = credential/origin detail)")

    # 11 — Ask-level check (scans p1 + p2 only; template body is fixed and clean)
    combined = (p1 + " " + p2).lower()
    ask_patterns = [
        (r"\bdonat\w*", "donation/donate language"),
        (r"\bsponsor\w*", "sponsor language"),
        (r"\bfund\w* (us|this|the)", "funding-us language"),
        (r"\bpartner with (us|you|lava)", "partnership-with-us language"),
        (r"\bseeking (a partner|partner\b)", "seeking-partner language"),
        (r"\bprovide (us|equipment|devices|supplies)", "provide-us language"),
        (r"\bsupply (us|equipment|devices)", "supply-us language"),
        (r"\bdevice donation\b", "device-donation language"),
    ]
    violations = [label for pat, label in ask_patterns if re.search(pat, combined)]
    if violations:
        fail(11,
             f"Ask-level violation in p1/p2: [{', '.join(violations)}]",
             "First email must be advice-only. The ask is a 10-min call to learn — nothing more.")

    # 12 — Generic title-only opener heuristic (medium — LLM judgment call flagged)
    generic_pats = [
        r"^you (are|work|serve) as (the |a )?director",
        r"^you lead (the |a )?global health",
        r"^you (are|work) (at|in) (the |a )?",
        r"^your (role|position|title|work) (at|in|as)",
        r"^you have (been|worked)",
    ]
    p1_lower = p1.strip().lower()
    for pat in generic_pats:
        if re.match(pat, p1_lower):
            fail(12,
                 f"p1 looks like a generic title-only opener: '{p1[:70]}'",
                 "Lead with something specific: a named project, paper, partnership, place, or recent post",
                 severity="medium")
            break

    return issues


# ═══════════════════════════════════════════════════════════════════════════════
# Commands
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_list():
    """List all contacts in the send queue (blank status, email present, personalization present)."""
    rows, _ = read_csv()
    queue = [(i, r) for i, r in enumerate(rows) if is_eligible(r)]

    print(f"Send queue: {len(queue)} contact(s) ready\n")
    for csv_idx, row in queue:
        name = row.get("Name", "").strip()
        system = row.get("System", "").strip()
        title = row.get("Title", "").strip()
        email = row.get("Email", "").strip()
        print(f"  [row {csv_idx}] {name}")
        print(f"           {title} | {system}")
        print(f"           {email}")
        print()


def cmd_next():
    """Find the next eligible contact and print as JSON with a pre-computed send slot."""
    # Session gate — refuse to dispense contacts without an active session.
    active, cfg = load_active_session()
    if not active:
        expired = cfg is not None
        print(json.dumps({
            "found": False,
            "gated": True,
            "message": (
                "Active session expired — rerun the Interactive Start Flow and `session-start`."
                if expired else
                "No active session. Run the Interactive Start Flow, then `session-start` "
                "before requesting contacts."
            ),
        }, indent=2))
        return

    rows, _ = read_csv()
    for csv_idx, row in enumerate(rows):
        if is_eligible(row):
            name = row.get("Name", "").strip()
            first_name = extract_greeting_name(name)
            result = {
                "found": True,
                "csv_row_idx": csv_idx,
                "crm_row_number": csv_idx + 1,   # 1-based (header = row 0)
                "name": name,
                "first_name": first_name,
                "system": row.get("System", "").strip(),
                "title": row.get("Title", "").strip(),
                "email": row.get("Email", "").strip(),
                "personalization": row.get("Personalization", "").strip(),
                "geo": derive_geo(row.get("Personalization", "")),
                "subject": SUBJECT,
                "session": cfg,
            }
            print(json.dumps(result, indent=2))
            return

    print(json.dumps({
        "found": False,
        "message": (
            "No eligible contacts. All contacts with emails are either already "
            "contacted, missing email, or missing personalization."
        ),
    }))


def cmd_preview(row_idx: int):
    """
    Preview a contact's data for a given row index.
    Prints the contact JSON so Claude can write p1/p2 and then call validate.
    Does NOT fill the template (that requires Claude-generated hooks).
    """
    rows, _ = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range (CSV has {len(rows)} rows)"}))
        sys.exit(1)

    row = rows[row_idx]
    name = row.get("Name", "").strip()
    print(json.dumps({
        "csv_row_idx": row_idx,
        "name": name,
        "first_name": extract_first_name(name),
        "system": row.get("System", "").strip(),
        "title": row.get("Title", "").strip(),
        "email": row.get("Email", "").strip(),
        "status": row.get("Status", "").strip(),
        "personalization": row.get("Personalization", "").strip(),
        "geo": derive_geo(row.get("Personalization", "")),
        "subject": SUBJECT,
        "template_hint": "Fill p1 and p2 from the personalization blob, then run: validate (geo auto-derived from Location; pass \"geo\" to override)",
    }, indent=2))


def cmd_validate():
    """
    Read JSON from stdin: {"row_idx": N, "p1": "...", "p2": "...", "geo": "..."}
    geo is optional — auto-derived from the Location bullet when omitted.
    Fill template, run 12 validation checks, print result JSON.
    Passes if no HIGH severity issues (medium issues are warnings).
    """
    try:
        payload = json.load(sys.stdin)
    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"Invalid JSON on stdin: {e}"}))
        sys.exit(1)

    row_idx = int(payload["row_idx"])
    p1 = payload.get("p1", "")
    p2 = payload.get("p2", "")

    rows, _ = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range (CSV has {len(rows)} rows)"}))
        sys.exit(1)

    row = rows[row_idx]
    name = row.get("Name", "").strip()
    first_name = extract_greeting_name(name)
    # geo: explicit override from stdin, else derived from the Location bullet.
    geo = str(payload.get("geo", "")).strip() or derive_geo(row.get("Personalization", ""))
    body = fill_template(first_name, p1, p2, geo)
    issues = validate_draft(body, first_name, p1, p2)

    high = [i for i in issues if i["severity"] == "high"]

    print(json.dumps({
        "valid": len(high) == 0,
        "csv_row_idx": row_idx,
        "first_name": first_name,
        "geo": geo,
        "email": row.get("Email", "").strip(),
        "subject": SUBJECT,
        "body": body,
        "word_count": count_words(body),
        "issues": issues,
        "high_issue_count": len(high),
        "total_issue_count": len(issues),
    }, indent=2))


# ═══════════════════════════════════════════════════════════════════════════════
# Timezone / scheduling helpers
# ═══════════════════════════════════════════════════════════════════════════════

def extract_timezone_from_personalization(personalization: str) -> str:
    """
    Extract IANA timezone from Location bullet in personalization blob.
    Looks for pattern: Location: Boston, MA (America/New_York)
    Returns 'America/New_York' by default if not found.
    """
    m = re.search(r'\(([A-Za-z_]+/[A-Za-z_]+)\)', personalization)
    if m:
        tz_str = m.group(1)
        try:
            ZoneInfo(tz_str)  # validate
            return tz_str
        except (ZoneInfoNotFoundError, KeyError):
            pass
    return "America/New_York"


def parse_time_range(time_range: str) -> tuple[int, int, int, int]:
    """
    Parse a time range string into (start_h, start_m, end_h, end_m) in 24hr.
    Accepts: '12:30-14:00' or '12:30 PM-2:00 PM' or '12:30PM-2:00PM'
    """
    def parse_single(t: str) -> tuple[int, int]:
        t = t.strip()
        pm = "PM" in t.upper()
        am = "AM" in t.upper()
        t_clean = t.upper().replace("PM", "").replace("AM", "").strip()
        if ":" in t_clean:
            h, m = t_clean.split(":")
        else:
            h, m = t_clean, "0"
        h, m = int(h), int(m)
        if pm and h != 12:
            h += 12
        if am and h == 12:
            h = 0
        return h, m

    parts = time_range.replace(" - ", "-").split("-")
    if len(parts) != 2:
        raise ValueError(f"Cannot parse time range: '{time_range}'. Expected 'HH:MM-HH:MM'.")
    # Handle cases like "12:30 PM-2:00 PM" where the split yields "12:30 PM" and "2:00 PM"
    start_str, end_str = parts[0], parts[1]
    sh, sm = parse_single(start_str)
    eh, em = parse_single(end_str)
    return sh, sm, eh, em


def random_minute_in_range(start_h: int, start_m: int, end_h: int, end_m: int) -> tuple[int, int]:
    """Pick a random minute within [start, end) range."""
    start_total = start_h * 60 + start_m
    end_total = end_h * 60 + end_m
    if end_total <= start_total:
        raise ValueError("End time must be after start time.")
    chosen = random.randint(start_total, end_total - 1)
    return chosen // 60, chosen % 60


def parse_days(days_str: str) -> list:
    """
    Parse comma-separated day names or dates into a list of date objects.
    Accepts: 'monday,tuesday' or '2026-06-02,2026-06-03' or mix.
    For named days: returns next occurrence of that weekday (today counts if today is that day).
    """
    DAY_MAP = {
        "monday": 0, "mon": 0,
        "tuesday": 1, "tue": 1, "tues": 1,
        "wednesday": 2, "wed": 2,
        "thursday": 3, "thu": 3, "thur": 3, "thurs": 3,
        "friday": 4, "fri": 4,
        "saturday": 5, "sat": 5,
        "sunday": 6, "sun": 6,
    }
    today = date.today()
    result = []
    for part in days_str.split(","):
        part = part.strip()
        if not part:
            continue
        lower = part.lower()
        if lower in DAY_MAP:
            target_weekday = DAY_MAP[lower]
            days_ahead = (target_weekday - today.weekday()) % 7
            result.append(today + timedelta(days=days_ahead))
        else:
            # Try parsing as YYYY-MM-DD
            try:
                result.append(date.fromisoformat(part))
            except ValueError:
                raise ValueError(f"Cannot parse day: '{part}'. Use a day name (monday) or date (2026-06-02).")
    return result


def cmd_compute_send_time(row_idx: int, days_str: str, time_range: str, recipient_tz_override=None):
    """
    For each day in days_str, pick a random slot in time_range (recipient local time),
    then convert to Eastern Time for Outlook scheduling.
    Prints JSON array of slot objects.
    """
    rows, _ = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range (CSV has {len(rows)} rows)"}))
        sys.exit(1)

    row = rows[row_idx]
    personalization = row.get("Personalization", "").strip()

    if recipient_tz_override:
        try:
            ZoneInfo(recipient_tz_override)
            recipient_tz_str = recipient_tz_override
        except (ZoneInfoNotFoundError, KeyError):
            print(json.dumps({"error": f"Unknown timezone: '{recipient_tz_override}'"}))
            sys.exit(1)
    else:
        recipient_tz_str = extract_timezone_from_personalization(personalization)

    recipient_tz = ZoneInfo(recipient_tz_str)
    et_tz = ZoneInfo("America/New_York")

    try:
        sh, sm, eh, em = parse_time_range(time_range)
    except ValueError as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)

    try:
        days = parse_days(days_str)
    except ValueError as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)

    slots = []
    for target_date in days:
        try:
            rh, rm = random_minute_in_range(sh, sm, eh, em)
        except ValueError as e:
            print(json.dumps({"error": str(e)}))
            sys.exit(1)

        # Build naive datetime in recipient's local time, then localize
        local_dt = datetime(target_date.year, target_date.month, target_date.day, rh, rm,
                            tzinfo=recipient_tz)
        # Convert to ET
        et_dt = local_dt.astimezone(et_tz)

        # Format for Outlook UI
        hour12 = et_dt.hour % 12 or 12
        ampm = "AM" if et_dt.hour < 12 else "PM"
        outlook_time = f"{hour12}:{et_dt.minute:02d} {ampm}"
        outlook_date = et_dt.strftime("%B %-d, %Y")

        # Recipient local display
        local_hour12 = local_dt.hour % 12 or local_dt.hour
        local_ampm = "AM" if local_dt.hour < 12 else "PM"
        tz_abbr = local_dt.strftime("%Z")
        local_display = f"{local_dt.strftime('%Y-%m-%d')} {local_hour12}:{local_dt.minute:02d} {local_ampm} {tz_abbr}"

        slots.append({
            "row_idx": row_idx,
            "recipient_tz": recipient_tz_str,
            "target_local": local_display,
            "outlook_et": f"{et_dt.strftime('%Y-%m-%d')} {outlook_time} ET",
            "outlook_date": outlook_date,
            "outlook_time": outlook_time,
        })

    print(json.dumps(slots if len(slots) > 1 else slots[0], indent=2))


def cmd_pace_plan(count: int, end_time: str):
    """
    Immediate-send pacing. Spread `count` direct sends across [now, end_time] today,
    roughly evenly with jitter. First send is immediate (now). Prints JSON list of
    targets the agent waits for before each direct Send click.

    end_time: 'HH:MM' 24hr or '2:00 PM'. Must be later than now today.
    """
    et = ZoneInfo("America/New_York")
    now = datetime.now(et)

    try:
        eh, em = None, None
        t = end_time.strip().upper()
        pm, am = "PM" in t, "AM" in t
        t_clean = t.replace("PM", "").replace("AM", "").strip()
        eh, em = (t_clean.split(":") + ["0"])[:2]
        eh, em = int(eh), int(em)
        if pm and eh != 12:
            eh += 12
        if am and eh == 12:
            eh = 0
    except (ValueError, AttributeError):
        print(json.dumps({"error": f"Cannot parse end_time '{end_time}'. Use 'HH:MM' or '2:00 PM'."}))
        sys.exit(1)

    end = now.replace(hour=eh, minute=em, second=0, microsecond=0)
    if end <= now:
        print(json.dumps({"error": f"end_time {end_time} is not after now ({now.strftime('%H:%M')} ET). Pick a later cutoff or use schedule mode."}))
        sys.exit(1)
    if count < 1:
        print(json.dumps({"error": "count must be >= 1"}))
        sys.exit(1)

    window = (end - now).total_seconds()
    seg = window / count                      # one segment per send
    targets = []
    prev = now
    for i in range(count):
        if i == 0:
            t_i = now                          # first send fires immediately
        else:
            base = now + timedelta(seconds=seg * i)
            jitter = random.uniform(0, seg * 0.5)
            t_i = min(base + timedelta(seconds=jitter), end)
        wait_s = max(0, int((t_i - prev).total_seconds()))
        hour12 = t_i.hour % 12 or 12
        ampm = "AM" if t_i.hour < 12 else "PM"
        targets.append({
            "n": i + 1,
            "send_at": t_i.strftime("%Y-%m-%d %H:%M:%S"),
            "send_at_display": f"{hour12}:{t_i.minute:02d} {ampm} ET",
            "wait_seconds_from_prev": wait_s,
        })
        prev = t_i

    print(json.dumps({
        "count": count,
        "window_start": now.strftime("%H:%M ET"),
        "window_end": end.strftime("%H:%M ET"),
        "avg_interval_seconds": int(seg),
        "targets": targets,
    }, indent=2))


def cmd_mark_sent(row_idx: int):
    """Mark contact as 'Contacted no reply'. Updates date and increments touch points."""
    rows, fieldnames = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range"}))
        sys.exit(1)

    today = date.today().strftime("%Y-%m-%d")
    current_tp = rows[row_idx].get("Number of Touch Points", "").strip()
    tp = (int(current_tp) + 1) if current_tp.isdigit() else 1

    rows[row_idx]["Status"] = "Contacted no reply"
    rows[row_idx]["Date of Last Contact"] = today
    rows[row_idx]["Number of Touch Points"] = str(tp)

    write_csv(rows, fieldnames)
    print(json.dumps({
        "ok": True,
        "csv_row_idx": row_idx,
        "name": rows[row_idx].get("Name", ""),
        "status": "Contacted no reply",
        "date": today,
        "touch_points": tp,
    }))


def cmd_mark_failed(row_idx: int):
    """Mark contact as 'Not sent (validation failure)'."""
    rows, fieldnames = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range"}))
        sys.exit(1)

    rows[row_idx]["Status"] = "Not sent (validation failure)"
    write_csv(rows, fieldnames)
    print(json.dumps({
        "ok": True,
        "csv_row_idx": row_idx,
        "name": rows[row_idx].get("Name", ""),
        "status": "Not sent (validation failure)",
    }))


# ═══════════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════════

USAGE = """\
Usage:
  python3 send_first_touch.py session-start                        (reads session JSON from stdin)
  python3 send_first_touch.py session-show
  python3 send_first_touch.py session-end
  python3 send_first_touch.py list
  python3 send_first_touch.py next                                 (gated: needs active session)
  python3 send_first_touch.py preview <row_idx>
  python3 send_first_touch.py validate                             (reads JSON from stdin)
  python3 send_first_touch.py pace-plan <count> <end_time>         (immediate-mode pacing schedule)
  python3 send_first_touch.py compute-send-time <row_idx> <days> <time_range> [--recipient-tz TZ]
  python3 send_first_touch.py mark-sent <row_idx>
  python3 send_first_touch.py mark-failed <row_idx>
"""

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(USAGE)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "session-start":
        cmd_session_start()
    elif cmd == "session-show":
        cmd_session_show()
    elif cmd == "session-end":
        cmd_session_end()
    elif cmd == "list":
        cmd_list()
    elif cmd == "next":
        cmd_next()
    elif cmd == "preview":
        if len(sys.argv) < 3:
            print("Usage: preview <row_idx>")
            sys.exit(1)
        cmd_preview(int(sys.argv[2]))
    elif cmd == "validate":
        cmd_validate()
    elif cmd == "pace-plan":
        if len(sys.argv) < 4:
            print("Usage: pace-plan <count> <end_time>   e.g. pace-plan 30 14:00")
            sys.exit(1)
        cmd_pace_plan(int(sys.argv[2]), sys.argv[3])
    elif cmd == "compute-send-time":
        if len(sys.argv) < 5:
            print("Usage: compute-send-time <row_idx> <days> <time_range> [--recipient-tz TZ]")
            sys.exit(1)
        _row_idx = int(sys.argv[2])
        _days = sys.argv[3]
        _time_range = sys.argv[4]
        _tz_override = None
        # Parse optional --recipient-tz
        for i, arg in enumerate(sys.argv[5:], start=5):
            if arg == "--recipient-tz" and i + 1 < len(sys.argv):
                _tz_override = sys.argv[i + 1]
                break
        cmd_compute_send_time(_row_idx, _days, _time_range, _tz_override)
    elif cmd == "mark-sent":
        if len(sys.argv) < 3:
            print("Usage: mark-sent <row_idx>")
            sys.exit(1)
        cmd_mark_sent(int(sys.argv[2]))
    elif cmd == "mark-failed":
        if len(sys.argv) < 3:
            print("Usage: mark-failed <row_idx>")
            sys.exit(1)
        cmd_mark_failed(int(sys.argv[2]))
    else:
        print(f"Unknown command: {cmd}\n{USAGE}")
        sys.exit(1)
