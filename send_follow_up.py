#!/usr/bin/env python3
"""
Boston Hospital Outreach — Follow-Up Sender
============================================
Sibling to send_first_touch.py. Runs the 3-touch follow-up sequence on contacts
who were already first-touched and have NOT replied. Lives on the same Outlook
inbox and reuses the same CSV / session / name / Outlook-compose structure as the
first-touch sender. Schedule-send and pacing logic are intentionally dropped:
follow-ups send immediately (direct Send) or to Drafts.

Cadence (1 week between touches, only while no reply):
  Touch points already sent → follow-up this run sends
    1 (initial only)        → Follow-up 1
    2 (initial + FU1)       → Follow-up 2
    3 (initial + FU2)       → Follow-up 3 (respectful close)
    >= 4                    → sequence complete, contact drops out

A contact is "due" when Status == "Contacted no reply" AND today >= due date,
where due date = Next Follow-Up Date if set, else Date of Last Contact + 7 days.

Reply handling: when Claude finds a reply in Outlook, call `mark-replied` — the
contact's Status becomes "Replied" and it never re-enters the follow-up queue.
This script does NOT read Outlook; Claude does that via Chrome MCP (see
outreach/FOLLOWUP_SKILL.md) and feeds decisions back through these commands.

Usage:
  python3 send_follow_up.py session-start          # Open session (stdin JSON) — REQUIRED before next
  python3 send_follow_up.py session-show           # Show active + last session (for repeat offer)
  python3 send_follow_up.py session-end            # Close session (re-gates next)
  python3 send_follow_up.py list                   # Show full follow-up queue (who's due, which FU#)
  python3 send_follow_up.py next                    # Next due contact as JSON (GATED: needs active session)
  python3 send_follow_up.py preview <row_idx>       # Inspect a specific row + its current FU#
  python3 send_follow_up.py validate                # Validate draft (stdin JSON)
  python3 send_follow_up.py mark-sent <row_idx>     # Increment touch points, stamp date, set +7 due date
  python3 send_follow_up.py mark-replied <row_idx>  # Status → Replied (drops out of sequence)
  python3 send_follow_up.py mark-failed <row_idx>   # Status → Not sent (validation failure)

validate stdin JSON:
  {"row_idx": N, "p1": "...", "p2": "..."}   # p1/p2 only needed for Follow-up 2
"""

import csv
import fcntl
import json
import os
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

# ── Paths (shared with first-touch sender) ──────────────────────────────────────
BASE = Path("/Users/lavapanta/Desktop/Hospital prospecting, outreach reachout automation")
_default_crm = BASE / "prospecting" / "Boston Hospital Outreach - CRM.csv"
CRM_PATH = Path(os.environ.get("OUTREACH_CRM_PATH", _default_crm))
LOCK_PATH = str(CRM_PATH) + ".lock"

# Follow-up gets its OWN session files so a first-touch session can't unlock it
# (and vice versa). Same TTL discipline as first-touch.
SESSION_PATH = str(CRM_PATH) + ".followup_session.json"
LAST_SESSION_PATH = str(CRM_PATH) + ".followup_last_session.json"
SESSION_TTL_HOURS = 12
LAST_SESSION_TTL_DAYS = 30

FOLLOWUP_INTERVAL_DAYS = 7
MAX_TOUCH_POINTS = 4  # initial + FU1 + FU2 + FU3 — once reached, sequence is done

SENDER_NAME = "Lava Panta"
CC_ADDRESS = os.environ.get("CC_ADDRESS", "")  # collaborator CC, set via env

# ── Follow-up templates (verbatim from Ethiopia_Device_Pipeline_Email_Sequence.md)
# Touch 2 = Follow-up 1, Touch 3 = Follow-up 2, Touch 4 = Follow-up 3.
FU1_SUBJECT = "Quick update on the Ethiopia device pipeline"
FU_THREAD_SUBJECT = "Re: Quick update on the Ethiopia device pipeline"

FU1_TEMPLATE = """\
Hi {name},

Hope you've had an amazing week. Wanted to share a quick update: we recently got in touch with BHI, a biomedical engineering organization that helps set up medical devices in low-resource areas, as well as Mass General. We're very excited about where this is going, and I'd love to hear your thoughts and get your advice as we build this device pipeline to the hospitals in Ethiopia that need it.

Sincerely,
Lava Panta"""

# Follow-up 2 — only the {p1}/{p2} span changes per recipient.
#   {p1} = the pipeline question their work speaks to (a problem on the ground)
#   {p2} = their specific work, named, tied to that question
FU2_TEMPLATE = """\
Hi {name},

I know your time is tight, so I'll keep this short. We're still building out the device pipeline to hospitals in Ethiopia, and the piece we keep coming back to is {p1}. That's exactly {p2}.

Even 10 minutes of your perspective would shape how we approach this. Happy to work entirely around your schedule.

Sincerely,
Lava Panta
Barock C. Tesfaye"""

FU3_TEMPLATE = """\
Hi {name},

I don't want to keep cluttering your inbox, so this will be my last note. I completely understand that the timing may not be right.

The offer stands open whenever it suits you, and if there's ever a moment down the line where a quick conversation makes sense, I'd be grateful for it. Thank you for the work you do in this space, it matters more than most people see.

Sincerely,
Lava Panta
Barock C. Tesfaye"""

DOCTORAL_CREDENTIALS = {"MD", "DO", "PHD", "DNP", "DDS", "DMD", "EDD", "DVM", "MBBS"}


# ═══════════════════════════════════════════════════════════════════════════════
# CSV helpers (identical lock pattern to first-touch sender)
# ═══════════════════════════════════════════════════════════════════════════════

def read_csv():
    with open(CRM_PATH, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fieldnames = list(reader.fieldnames)
    return rows, fieldnames


def write_csv(rows, fieldnames):
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
# Session gate + last-session memory (copied from first-touch; schedule fields dropped)
# ═══════════════════════════════════════════════════════════════════════════════

VALID_MODES = {"immediate", "draft"}   # no schedule mode for follow-ups
VALID_SCOPE = {"queue", "count", "names"}


def _read_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def load_active_session():
    cfg = _read_json(SESSION_PATH)
    if not cfg:
        return False, None
    expires = cfg.get("expires_at")
    if expires:
        try:
            if datetime.now() > datetime.fromisoformat(expires):
                return False, cfg
        except ValueError:
            return False, cfg
    return True, cfg


def cmd_session_start():
    try:
        payload = json.load(sys.stdin)
    except json.JSONDecodeError as e:
        print(json.dumps({"ok": False, "error": f"Invalid JSON on stdin: {e}"}))
        sys.exit(1)

    mode = str(payload.get("mode", "")).strip().lower()
    scope = str(payload.get("scope", "queue")).strip().lower()

    errs = []
    if mode not in VALID_MODES:
        errs.append(f"mode must be one of {sorted(VALID_MODES)} (got '{mode}')")
    if scope not in VALID_SCOPE:
        errs.append(f"scope must be one of {sorted(VALID_SCOPE)} (got '{scope}')")
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
        "scope": scope,
        "count": int(payload["count"]) if scope == "count" else None,
        "names": list(payload.get("names", [])) if scope == "names" else [],
        "created_at": now.isoformat(timespec="seconds"),
        "expires_at": (now + timedelta(hours=SESSION_TTL_HOURS)).isoformat(timespec="seconds"),
    }

    with open(SESSION_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)

    last = {k: v for k, v in cfg.items() if k != "expires_at"}
    last["saved_at"] = now.isoformat(timespec="seconds")
    with open(LAST_SESSION_PATH, "w", encoding="utf-8") as f:
        json.dump(last, f, indent=2)

    print(json.dumps({"ok": True, "session": cfg}, indent=2))


def load_last_session():
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
    existed = os.path.exists(SESSION_PATH)
    if existed:
        os.remove(SESSION_PATH)
    print(json.dumps({"ok": True, "cleared": existed}))


# ═══════════════════════════════════════════════════════════════════════════════
# Name helpers (copied verbatim from first-touch sender)
# ═══════════════════════════════════════════════════════════════════════════════

def extract_first_name(full_name):
    base = full_name.split(",")[0].strip()
    tokens = base.split()
    if not tokens:
        return full_name
    honorifics = {"Dr.", "Prof.", "Mr.", "Ms.", "Mrs.", "Sir", "Rev."}
    first = tokens[0]
    if first in honorifics and len(tokens) > 1:
        first = tokens[1]
    return first


def extract_last_name(full_name):
    base = full_name.split(",")[0].strip()
    tokens = base.split()
    honorifics = {"Dr.", "Prof.", "Mr.", "Ms.", "Mrs.", "Sir", "Rev."}
    tokens = [t for t in tokens if t not in honorifics]
    return tokens[-1] if tokens else full_name


def is_doctor(full_name):
    parts = full_name.split(",")
    if len(parts) < 2:
        return False
    credentials_str = " ".join(parts[1:]).upper()
    for cred in DOCTORAL_CREDENTIALS:
        if re.search(r'\b' + cred + r'\b', credentials_str):
            return True
    return False


def extract_greeting_name(full_name):
    if is_doctor(full_name):
        return f"Dr. {extract_last_name(full_name)}"
    return extract_first_name(full_name)


# ═══════════════════════════════════════════════════════════════════════════════
# Follow-up eligibility / cadence
# ═══════════════════════════════════════════════════════════════════════════════

def current_touch_count(row):
    """How many emails have already been sent to this contact.
    Blank touch-points on a 'Contacted no reply' row is treated as 1 (initial sent),
    matching legacy first-touch rows that never stamped the counter."""
    tp_str = row.get("Number of Touch Points", "").strip()
    if tp_str.isdigit():
        return int(tp_str)
    if row.get("Status", "").strip() == "Contacted no reply":
        return 1
    return 0


def followup_due_date(row):
    """Date this contact becomes due for its next follow-up.
    Prefers an explicit Next Follow-Up Date; else Date of Last Contact + 7 days."""
    nfu = row.get("Next Follow-Up Date", "").strip()
    if nfu:
        try:
            return date.fromisoformat(nfu)
        except ValueError:
            pass
    dlc = row.get("Date of Last Contact", "").strip()
    if dlc:
        try:
            return date.fromisoformat(dlc) + timedelta(days=FOLLOWUP_INTERVAL_DAYS)
        except ValueError:
            pass
    return None


def is_followup_due(row, today=None):
    """Due = sent-and-silent, still has touches left, and the 7-day window has elapsed."""
    today = today or date.today()
    if row.get("Status", "").strip() != "Contacted no reply":
        return False
    if not row.get("Email", "").strip():
        return False
    tp = current_touch_count(row)
    if tp < 1 or tp >= MAX_TOUCH_POINTS:
        return False
    due = followup_due_date(row)
    if due is None:
        return False
    return today >= due


def followup_number(row):
    """Which follow-up to send next (1, 2, or 3) = number of emails already sent."""
    return current_touch_count(row)


def subject_for(fu_no):
    return FU1_SUBJECT if fu_no == 1 else FU_THREAD_SUBJECT


# ═══════════════════════════════════════════════════════════════════════════════
# Template fill + validation
# ═══════════════════════════════════════════════════════════════════════════════

def fill_template(fu_no, name, p1="", p2=""):
    if fu_no == 1:
        return FU1_TEMPLATE.format(name=name)
    if fu_no == 2:
        return FU2_TEMPLATE.format(name=name, p1=p1, p2=p2)
    if fu_no == 3:
        return FU3_TEMPLATE.format(name=name)
    raise ValueError(f"No template for follow-up #{fu_no}")


def count_words(text):
    return len(text.split())


def validate_draft(fu_no, body, first_name, p1, p2):
    """
    Deterministic checks. Empty list = pass (no HIGH issues).
    FU1/FU3 are fixed bodies → structural checks only.
    FU2 adds the personalized {p1}/{p2} span checks.
    """
    issues = []

    def fail(n, msg, fix, severity="high"):
        issues.append({"check": n, "severity": severity, "message": msg, "fix": fix})

    # 1 — greeting filled
    if "{name}" in body or "{Name}" in body:
        fail(1, "Template {name} placeholder still present", "Pass greeting name to fill_template()")

    # 2 — no stray placeholders
    stray = re.findall(r"\{[^}]{1,40}\}", body)
    if stray:
        fail(2, f"Stray placeholder(s) in draft: {stray}", "Fill or remove all {…} tokens")

    # 3 — word count guard
    wc = count_words(body)
    if wc > 160:
        fail(3, f"Word count {wc} exceeds 160-word guard", "Shorten the personalized span")

    # 4a/4b — no em dash / double hyphen (sequence rule: no em dashes anywhere)
    if "—" in body:
        fail(4, "Em dash (—) found in draft", "Replace with a single hyphen (-) or rephrase")
    if "--" in body:
        fail(4, "Double-hyphen (--) found — em-dash substitute not allowed",
             "Replace with a single hyphen (-) or rephrase")

    # 5 — greeting name clean
    if len(first_name.split()) > 3:
        fail(5, f"Greeting name '{first_name}' too long — may contain credentials",
             "Use 'Dr. LastName' or 'FirstName'")
    if any(ch in first_name for ch in [",", ";", "("]):
        fail(5, f"Greeting name '{first_name}' has unexpected punctuation",
             "Use extract_greeting_name()")
    if "." in first_name and not first_name.startswith("Dr."):
        fail(5, f"Greeting name '{first_name}' has unexpected period",
             "Use extract_greeting_name()")

    # FU2-only personalization checks
    if fu_no == 2:
        if len(p1.strip()) < 15:
            fail(6, f"p1 too short ({len(p1.strip())} chars) — the pipeline question must be concrete",
                 "Phrase a real on-the-ground problem (catching failures early, what equipment survives, training gaps)")
        if len(p2.strip()) < 20:
            fail(7, f"p2 too short ({len(p2.strip())} chars) — must name their specific work",
                 "Name the actual program/research and tie it to p1 (e.g. 'the problem your GeoSentinel work has solved across 70 sites')")
        if p1.strip().lower() == p2.strip().lower():
            fail(8, "p1 and p2 are identical — they are two halves of one sentence, not duplicates",
                 "p1 = the question; p2 = their named work that answers it")
        # ask-level guard (advice only — same spirit as first touch)
        combined = (p1 + " " + p2).lower()
        ask_patterns = [
            (r"\bdonat\w*", "donation/donate language"),
            (r"\bsponsor\w*", "sponsor language"),
            (r"\bfund\w* (us|this|the)", "funding-us language"),
            (r"\bpartner with (us|you|lava)", "partnership-with-us language"),
            (r"\bprovide (us|equipment|devices|supplies)", "provide-us language"),
            (r"\bsupply (us|equipment|devices)", "supply-us language"),
        ]
        violations = [label for pat, label in ask_patterns if re.search(pat, combined)]
        if violations:
            fail(9, f"Ask-level violation in p1/p2: [{', '.join(violations)}]",
                 "Follow-up is advice-only — a 10-min call to learn, nothing more.")

    return issues


# ═══════════════════════════════════════════════════════════════════════════════
# Commands
# ═══════════════════════════════════════════════════════════════════════════════

def _scope_filter(queue, cfg):
    """Apply session scope to an ordered (csv_idx, row) queue."""
    scope = (cfg or {}).get("scope", "queue")
    if scope == "names":
        wanted = {n.strip().lower() for n in cfg.get("names", [])}
        return [(i, r) for i, r in queue if r.get("Name", "").strip().lower() in wanted
                or extract_first_name(r.get("Name", "")).lower() in wanted]
    # 'count' is enforced by the agent across sends; 'queue' = everything.
    return queue


def cmd_list():
    rows, _ = read_csv()
    today = date.today()
    queue = [(i, r) for i, r in enumerate(rows) if is_followup_due(r, today)]

    print(f"Follow-up queue: {len(queue)} contact(s) due (as of {today})\n")
    for csv_idx, row in queue:
        fu = followup_number(row)
        due = followup_due_date(row)
        print(f"  [row {csv_idx}] {row.get('Name','').strip()}  → Follow-up {fu}")
        print(f"           {row.get('Title','').strip()} | {row.get('System','').strip()}")
        print(f"           {row.get('Email','').strip()}")
        print(f"           last contact {row.get('Date of Last Contact','').strip() or '?'} | due {due}")
        print()

    # Surface not-yet-due so Claude can report "X waiting until <date>".
    upcoming = [(i, r) for i, r in enumerate(rows)
                if r.get("Status", "").strip() == "Contacted no reply"
                and r.get("Email", "").strip()
                and 1 <= current_touch_count(r) < MAX_TOUCH_POINTS
                and not is_followup_due(r, today)]
    if upcoming:
        print(f"Not yet due: {len(upcoming)} contact(s) waiting on the 7-day window")
        for csv_idx, row in upcoming:
            due = followup_due_date(row)
            print(f"  [row {csv_idx}] {row.get('Name','').strip()} → due {due}")


def cmd_next():
    active, cfg = load_active_session()
    if not active:
        expired = cfg is not None
        print(json.dumps({
            "found": False,
            "gated": True,
            "message": (
                "Active follow-up session expired — rerun the Interactive Start Flow and `session-start`."
                if expired else
                "No active session. Run the Interactive Start Flow, then `session-start` before requesting contacts."
            ),
        }, indent=2))
        return

    rows, _ = read_csv()
    today = date.today()
    queue = [(i, r) for i, r in enumerate(rows) if is_followup_due(r, today)]
    queue = _scope_filter(queue, cfg)

    if not queue:
        print(json.dumps({"found": False, "message": "No follow-ups due right now."}, indent=2))
        return

    csv_idx, row = queue[0]
    name = row.get("Name", "").strip()
    fu = followup_number(row)
    result = {
        "found": True,
        "csv_row_idx": csv_idx,
        "crm_row_number": csv_idx + 1,
        "name": name,
        "first_name": extract_greeting_name(name),
        "system": row.get("System", "").strip(),
        "title": row.get("Title", "").strip(),
        "email": row.get("Email", "").strip(),
        "follow_up_number": fu,
        "subject": subject_for(fu),
        "is_thread_reply": fu >= 2,   # FU2/FU3 are "Re:" — reply in the existing thread
        "needs_personalization": fu == 2,
        "personalization": row.get("Personalization", "").strip(),
        "date_of_last_contact": row.get("Date of Last Contact", "").strip(),
        "touch_points": current_touch_count(row),
        "session": cfg,
    }
    if fu == 2:
        result["template_hint"] = (
            "Follow-up 2: write p1 (the pipeline question their work speaks to) and "
            "p2 (their specific work, named, that answers it). Pull the angle from the "
            "Personalization blob / original Touch-1 hook. Then run validate."
        )
    print(json.dumps(result, indent=2))


def cmd_preview(row_idx):
    rows, _ = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range (CSV has {len(rows)} rows)"}))
        sys.exit(1)
    row = rows[row_idx]
    name = row.get("Name", "").strip()
    fu = followup_number(row)
    print(json.dumps({
        "csv_row_idx": row_idx,
        "name": name,
        "first_name": extract_greeting_name(name),
        "system": row.get("System", "").strip(),
        "title": row.get("Title", "").strip(),
        "email": row.get("Email", "").strip(),
        "status": row.get("Status", "").strip(),
        "touch_points": current_touch_count(row),
        "follow_up_number": fu if 1 <= fu < MAX_TOUCH_POINTS else None,
        "subject": subject_for(fu) if 1 <= fu < MAX_TOUCH_POINTS else None,
        "due_date": str(followup_due_date(row)) if followup_due_date(row) else None,
        "due_now": is_followup_due(row),
        "personalization": row.get("Personalization", "").strip(),
    }, indent=2))


def cmd_validate():
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
    fu = followup_number(row)
    if not (1 <= fu < MAX_TOUCH_POINTS):
        print(json.dumps({"error": f"Row {row_idx} has no follow-up due (touch count {fu})."}))
        sys.exit(1)

    name = row.get("Name", "").strip()
    first_name = extract_greeting_name(name)
    body = fill_template(fu, first_name, p1, p2)
    issues = validate_draft(fu, body, first_name, p1, p2)
    high = [i for i in issues if i["severity"] == "high"]

    print(json.dumps({
        "valid": len(high) == 0,
        "csv_row_idx": row_idx,
        "follow_up_number": fu,
        "first_name": first_name,
        "email": row.get("Email", "").strip(),
        "subject": subject_for(fu),
        "is_thread_reply": fu >= 2,
        "body": body,
        "word_count": count_words(body),
        "issues": issues,
        "high_issue_count": len(high),
        "total_issue_count": len(issues),
    }, indent=2))


def cmd_mark_sent(row_idx):
    rows, fieldnames = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range"}))
        sys.exit(1)

    today = date.today()
    new_tp = current_touch_count(rows[row_idx]) + 1
    rows[row_idx]["Status"] = "Contacted no reply"
    rows[row_idx]["Date of Last Contact"] = today.strftime("%Y-%m-%d")
    rows[row_idx]["Number of Touch Points"] = str(new_tp)
    # Schedule next window, or clear it once the sequence is complete.
    if new_tp < MAX_TOUCH_POINTS:
        next_due = today + timedelta(days=FOLLOWUP_INTERVAL_DAYS)
        rows[row_idx]["Next Follow-Up Date"] = next_due.strftime("%Y-%m-%d")
    else:
        rows[row_idx]["Next Follow-Up Date"] = ""

    write_csv(rows, fieldnames)
    print(json.dumps({
        "ok": True,
        "csv_row_idx": row_idx,
        "name": rows[row_idx].get("Name", ""),
        "follow_up_sent": new_tp - 1,
        "touch_points": new_tp,
        "date": rows[row_idx]["Date of Last Contact"],
        "next_follow_up_date": rows[row_idx]["Next Follow-Up Date"] or None,
        "sequence_complete": new_tp >= MAX_TOUCH_POINTS,
    }, indent=2))


def cmd_mark_replied(row_idx):
    rows, fieldnames = read_csv()
    if row_idx >= len(rows):
        print(json.dumps({"error": f"row_idx {row_idx} out of range"}))
        sys.exit(1)
    rows[row_idx]["Status"] = "Replied"
    rows[row_idx]["Next Follow-Up Date"] = ""
    write_csv(rows, fieldnames)
    print(json.dumps({
        "ok": True,
        "csv_row_idx": row_idx,
        "name": rows[row_idx].get("Name", ""),
        "status": "Replied",
        "note": "Dropped from the follow-up sequence.",
    }, indent=2))


def cmd_mark_failed(row_idx):
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
    }, indent=2))


# ═══════════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════════

USAGE = """\
Usage:
  python3 send_follow_up.py session-start          (reads session JSON from stdin)
  python3 send_follow_up.py session-show
  python3 send_follow_up.py session-end
  python3 send_follow_up.py list
  python3 send_follow_up.py next                    (gated: needs active session)
  python3 send_follow_up.py preview <row_idx>
  python3 send_follow_up.py validate               (reads JSON from stdin)
  python3 send_follow_up.py mark-sent <row_idx>
  python3 send_follow_up.py mark-replied <row_idx>
  python3 send_follow_up.py mark-failed <row_idx>
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
    elif cmd == "mark-sent":
        if len(sys.argv) < 3:
            print("Usage: mark-sent <row_idx>")
            sys.exit(1)
        cmd_mark_sent(int(sys.argv[2]))
    elif cmd == "mark-replied":
        if len(sys.argv) < 3:
            print("Usage: mark-replied <row_idx>")
            sys.exit(1)
        cmd_mark_replied(int(sys.argv[2]))
    elif cmd == "mark-failed":
        if len(sys.argv) < 3:
            print("Usage: mark-failed <row_idx>")
            sys.exit(1)
        cmd_mark_failed(int(sys.argv[2]))
    else:
        print(f"Unknown command: {cmd}\n{USAGE}")
        sys.exit(1)
