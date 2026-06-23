---
name: hospital-outreach
description: |
  Boston Hospital Outreach sending automation. Use this skill whenever the user asks to send emails, run the outreach, "send the next one", "process the queue", "send to [name]", or any reference to scheduling emails in Outlook. Also trigger on "check validation failures", "retry failed sends", "how many are queued", or "draft mode". Covers the full send pipeline: contact selection → personalization hook generation → email validation (12 checks, 3-retry loop) → Outlook compose via Chrome → CRM update.
---

# Boston Hospital Outreach — Sending Skill

You are sending cold outreach emails on behalf of Lava Panta and Barock C. Tesfaye to global health stakeholders at Boston-area hospitals. The ask is **advice only** — a 10-minute call from a Northeastern student working on Ethiopian hospital supply shortages.

## The Script Helper

All mechanical CRM operations go through:
```
outreach/scripts/send_first_touch.py
```

Always `cd` to the base path first:
```bash
cd "/Users/lavapanta/Desktop/Hospital prospecting, outreach reachout automation"
```

Available commands:
```bash
python3 outreach/scripts/send_first_touch.py session-start         # Open session (stdin JSON) — REQUIRED before next
python3 outreach/scripts/send_first_touch.py session-show          # Show active + last session (for repeat offer)
python3 outreach/scripts/send_first_touch.py session-end           # Close session (re-gates next)
python3 outreach/scripts/send_first_touch.py list                  # Show full send queue
python3 outreach/scripts/send_first_touch.py next                  # Get next contact as JSON (GATED: needs active session)
python3 outreach/scripts/send_first_touch.py preview <row_idx>     # Inspect a specific row
python3 outreach/scripts/send_first_touch.py validate              # Validate draft (stdin JSON)
python3 outreach/scripts/send_first_touch.py mark-sent <row_idx>   # Status → Contacted no reply
python3 outreach/scripts/send_first_touch.py mark-failed <row_idx> # Status → Not sent (validation failure)
```

---

## Interactive Start Flow

**HARD GATE.** `next` refuses to return contacts until you call `session-start`. This forces
the start questions to run every session — including headless/scheduled-task runs. Do not
attempt to bypass by editing the CSV directly; just run the flow.

Run these steps at the beginning of every send session, in order.

---

### Step 0 — Read last session (repeat offer)

```bash
python3 outreach/scripts/send_first_touch.py session-show
```

- `has_last: true` → there is a remembered prior session (`last_session`: mode, timezone, scope, days, time_range, count/names). Offer it as the first widget option so Lava can rerun the exact same setup in one tap.
- `has_last: false` → no prior session; skip the repeat option.

This last-session record is the only persistent memory between runs — treat it as a convenience default, never an auto-run. Confirm via the widget every time.

---

### Step A — Ask the start questions (single widget)

Use **one AskUserQuestion** call with these questions (do not ask in prose — the widget keeps it cheap and consistent):

1. **Send mode** — `Schedule send` / `Immediate send` / `Draft mode`
   - If Step 0 found a last session, prepend a **`Rerun last session`** option that restates it (e.g. "Schedule · Mon,Tue · 12:30–14:00 · count 3").
2. **Timezone** — `Same (ET)` / `Different` (only needed for Schedule/Immediate; skip for Draft if reviewing only)
3. **Batch scope** — `Whole queue` / `Specific count` / `Specific names`

If scope is `count` or `names`, or if the chosen mode is `Schedule` and days/time_range weren't carried from a repeat, collect those values too (free-text follow-up is fine for days/time/count/names — the widget can't capture them cleanly).

**Draft mode** note: compose in Outlook, save to Drafts, do NOT call mark-sent. Still requires a session (`session-start` with `mode: draft`).

---

### Step B — Open the session

Translate the answers to JSON and open the session:

```bash
python3 outreach/scripts/send_first_touch.py session-start << 'ENDJSON'
{
  "mode": "schedule",          // immediate | schedule | draft
  "timezone": "same",          // same | different
  "scope": "count",            // queue | count | names
  "days": "monday,tuesday",    // required when mode == schedule
  "time_range": "12:30-14:00", // required when mode == schedule
  "count": 3,                  // when scope == count
  "names": []                  // when scope == names
}
ENDJSON
```

For a **`Rerun last session`** choice, copy the fields straight from `last_session` (Step 0) into this payload.

Session lives 12h, then `next` re-gates. `next` echoes the session config back in its output (`session` key) so you don't re-ask mode/scope/days mid-batch. Enforce scope yourself: stop after `count` sends, or only process `names`.

When the batch is done (or Lava stops early), close it:
```bash
python3 outreach/scripts/send_first_touch.py session-end
```

---

### Routing by mode

- **Immediate** → [Immediate pacing](#immediate-pacing). Clicks **Send** directly (no Schedule dialog), paced now→cutoff until quota. Use this when Lava wants real sends spread over a window.
- **Schedule + Same (ET)** → Step C-same below. Uses the Schedule-send dialog.
- **Schedule + Different** → Step C-diff below. Schedule-send dialog + TZ conversion.
- **Draft** → Full Send Pipeline, but Draft-mode compose (save, no mark-sent).

**Immediate vs Schedule — do not confuse:**
- Immediate = press the Send button at paced live intervals. No future timestamp set on the message.
- Schedule = open the Schedule-send dialog, set a future date/time, let Outlook hold it.
Picking the wrong one wastes Lava's intent (and the Schedule path costs more). For Immediate, NEVER open the Schedule-send dialog.

---

### Immediate pacing

For **Immediate** mode, the session `time_range` end is the cutoff. Quota = session `count` (scope `count`) or the whole queue.

1. Build the pacing schedule once at session start:
   ```bash
   python3 outreach/scripts/send_first_touch.py pace-plan <count> <end_time>
   # e.g. pace-plan 30 14:00   → 30 targets spread now→2:00 PM ET, first fires immediately
   ```
   Returns `targets[]`: each has `send_at`, `send_at_display`, `wait_seconds_from_prev`.

2. For each target in order:
   - Wait `wait_seconds_from_prev` before composing (target 1 = 0, send now).
   - Short gaps (≤ ~270s): `sleep` in one Bash call.
   - Long gaps: `ScheduleWakeup` so context doesn't burn idle, then resume.
   - Run Full Send Pipeline Steps 1–4 (find contact → hooks → validate → compose-fill).
   - **Direct-send branch** in Step 4 (click Send, NOT Schedule dialog).
   - `mark-sent` after the send toast confirms.

3. Stop when quota hit, cutoff passed, or queue empty — whichever first. Then `session-end`.

If `now` is already past the cutoff, `pace-plan` errors — tell Lava to pick a later cutoff or switch to Schedule mode.

---

### Step C-same — Schedule in ET

Use the `days` and `time_range` from the session. For each contact, call:
```bash
python3 outreach/scripts/send_first_touch.py compute-send-time <row_idx> <days> <time_range> --recipient-tz America/New_York
```

Use the returned `outlook_date` and `outlook_time` in the Outlook schedule-send picker. No conversion needed — recipient is in ET.

---

### Step C-diff — Schedule with timezone conversion

When timezone is `Different`, the session's `time_range` is interpreted as the recipient's **local** arrival time (make this explicit in the Step A follow-up). Use the session `days` and `time_range`. For each contact, call:
```bash
python3 outreach/scripts/send_first_touch.py compute-send-time <row_idx> <days> <time_range>
```

The script reads the timezone from the personalization blob's `Location:` line. It returns `outlook_date` and `outlook_time` **already converted to ET** — use those values directly in Outlook. No manual conversion needed.

If the personalization blob has no timezone, the script defaults to `America/New_York`.

---

## Full Send Pipeline

### Step 1: Find the next contact

```bash
python3 outreach/scripts/send_first_touch.py next
```

Returns JSON. Key fields:
- `found` — true/false. If false, stop and report queue is empty.
- `csv_row_idx` — use this in all subsequent commands
- `first_name` — pre-extracted
- `email` — send to this address
- `personalization` — full research blob (your raw material for writing hooks)
- `geo` — derived geography phrase for `{geo}` (e.g. "Boston medical system")
- `subject` — fixed subject line

Send timing is NOT set here — it comes from the interactive start flow: `pace-plan` (immediate) or `compute-send-time` (schedule).

To check the full queue first: `python3 outreach/scripts/send_first_touch.py list`

### Step 2: Generate personalization hooks

Read the `personalization` blob. Write two distinct hooks — p1 and p2. These are judgment calls; the validator catches structure problems but you must enforce quality yourself.

---

**p1 — continues "I saw ___"**

Default rule: pick the single most relevant angle from all available info (bio, lab/personal site, About section, role, focus, projects) that best connects to the Ethiopia hospital supply work.

LinkedIn sub-rule (only applies when blob notes source was LinkedIn):
- Post within 14 days → lead with that post
- Post older than 14 days → fall back to mission/passion angle from bio

Hard rules:
- **Never a generic title-only opener.** "I saw you lead the global health program" is not acceptable — it must be specific enough that only this person could have done it.
- 1 short clause or phrase, not a full paragraph
- No em dashes (—), no double-hyphens (--)

Good examples:
- `the Ethiopia Surgical Training Partnership you co-led at MGH`
- `your 2023 Lancet paper on surgical capacity gaps in sub-Saharan Africa`
- `the PEN-Plus program scaling NCD care across East Africa`
- `your post last week on ultrasound access in rural Ethiopia` (LinkedIn recency rule)

Bad examples (all rejected):
- `your work in global health` — generic
- `your experience with international medicine` — generic
- `your role as director of the program` — title-only

---

**p2 — P.S. hook**

A second, different specific hook. Must not repeat p1's angle.

Good types:
- A different credential: named award, fellowship, or paper not used in p1
- A named program or partnership: something concrete from their portfolio
- An origin or background detail that shows you read their full bio
- A pointed question relevant to the Ethiopia/supply work

Good examples:
- `Your Rwanda Ministry of Health NCD work is directly relevant to what we're seeing — curious whether supply chain was the main bottleneck there too.`
- `The CAMTech model of co-designing diagnostics with local teams is exactly what we're trying to understand better.`
- `If you have 10 minutes, would love to know whether Project C.U.R.E. was a useful partner for the surgical training program.`

---

**Ask level — hard rule for both hooks:**
First email is advice-only. A 10-minute call to learn. Never: donation ask, device request, partnership pitch, supply ask, sponsorship. The validator catches obvious violations; use judgment for edge cases.

---

### Step 3: Validate the draft (up to 3 attempts)

```bash
python3 outreach/scripts/send_first_touch.py validate << 'ENDJSON'
{
  "row_idx": <csv_row_idx>,
  "p1": "<your p1>",
  "p2": "<your p2>"
}
ENDJSON
```

`{geo}` (the "...people in the ___ who understand this work" phrase) is auto-derived from the contact's `Location:` bullet — Boston contacts get "Boston medical system", New York contacts get "New York medical system", etc. Add `"geo": "<phrase>"` to the JSON above only when you want to override the derived value.

Returns JSON with `valid` (true/false), `body` (full filled email), `geo` (the phrase used), `word_count`, and `issues`.

**Retry logic:**
- `valid: false` → read issues, fix p1/p2, run validate again
- Max 3 attempts total
- After 3 failures → `mark-failed` and report to user

**The 12 checks:**
1. `{name}` placeholder filled
2. `{p1}` placeholder filled
3. `{p2}` placeholder filled
4. No stray `{…}` tokens
5. Word count ≤ 150 (greeting through P.S. inclusive — hard limit)
6a. No em dash (—)
6b. No double-hyphen (--)
7. First name clean (no credentials)
8. p1 substantive (≥ 15 chars)
9. p2 substantive (≥ 25 chars) — medium severity
10. p1 and p2 are different hooks
11. No ask-level violations (donation/sponsor/device/partner-with-us language)
12. p1 not a generic title-only opener — medium severity (LLM judgment)

HIGH severity issues block sending. MEDIUM severity issues are warnings — fix if easy, proceed if the draft is genuinely strong.

---

### Step 4: Compose in Outlook via Chrome

Use the Chrome MCP tools. JS-first approach — minimize round-trips, two screenshots max per email.

#### Screenshot discipline (HARD RULE)

Exactly **2 screenshots per email**:
1. **Pre-send** — after compose fully filled, before pressing Send / opening Schedule dialog. Catches typos, mangled recipients, body issues.
2. **Post-send** — after the Send (immediate) or Send-in-dialog (schedule) click. Confirms sent/scheduled toast.

Do NOT screenshot for: opening menus, finding buttons, picking dropdown options, verifying clicks. Use `find` + accessibility tree instead.

#### Compose flow (per email)

1. **First email only**: navigate to `https://outlook.office.com/mail/`.
2. Click "New mail" (use `find` to get button ref).
3. **JS template** — fill To, Cc, Subject, body in one `javascript_tool` call:

   ```javascript
   (() => {
     const TO = 'recipient@example.com';
     const CC = 'collaborator@example.edu';
     const SUBJECT = 'Trying to help people back home, would love your advice!';
     const BODY = `Hey {name},\n\nI'm Lava Panta. I saw {p1}. ...full body...`;

     // Body — MUST target the editable compose textbox, not the read-only
     // original message body. In reply mode both elements share
     // aria-label="Message body"; the compose one has role="textbox" and
     // isContentEditable=true. A naked [aria-label="Message body"] selector
     // silently picks the original-message DIV and the reply sends BLANK.
     const bodyEl =
       document.querySelector('[aria-label="Message body"][role="textbox"]')
       || Array.from(document.querySelectorAll('[aria-label="Message body"]'))
            .find(e => e.isContentEditable)
       || document.querySelector('[contenteditable="true"][role="textbox"]');
     if (bodyEl) {
       bodyEl.focus();
       bodyEl.innerText = BODY;
       bodyEl.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: BODY }));
     }

     // Subject
     const subjEl = document.querySelector('input[aria-label="Subject"]');
     if (subjEl) {
       const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
       setter.call(subjEl, SUBJECT);
       subjEl.dispatchEvent(new Event('input', { bubbles: true }));
     }

     return JSON.stringify({ body: bodyEl?.innerText.length, subject: subjEl?.value });
   })()
   ```

   **HARD CHECK after this call**: the returned `body` length must be > 0 and
   roughly match `BODY.length`. If it's 0 or missing, the wrong element was
   targeted — DO NOT click Send. Re-inspect:
   ```javascript
   Array.from(document.querySelectorAll('[aria-label="Message body"]'))
     .map(e => ({ role: e.getAttribute('role'), ce: e.isContentEditable, len: e.innerText.length }))
   ```
   and adjust the selector before proceeding.

   To and Cc still need typing + click-suggestion (Outlook's pill component rejects JS sets). Pattern with explicit 6s waits (validated for Lava's 8080-unread inbox; 2s/4s race-failed):

   ```
   click field ref → type email → wait 6s → find pill → click pill ref
   ```

   Then immediately:
   ```
   find next field (Cc or Subject) → click ref → repeat pattern
   ```

   Why 6s: Outlook's heavy DOM + autocomplete + React renders are slow. Without enough delay, autocomplete races with the next action and corrupts the address mid-string (e.g. "collaborator@example.edu" → "derbilt.edu"). The first session's ~50 screenshots provided accidental 1-2s buffers; removing them surfaced the race. Explicit waits replace screenshot timing without vision-token cost.

   The `find` after wait acts as a serialization barrier (forces accessibility tree settle) AND returns the ref needed for the next click. Two roles, one tool call.

   Never use Tab or Return after typing an address — both trigger truncation regardless of waits. Always commit via click on the autocomplete suggestion pill.

4. **Pre-send screenshot** — verify To pill, Cc pill, subject, body are correct.

5. **Branch by mode:**

   **Immediate (direct send):**
   - find "Send button" (the main compose Send, aria-label `Send`) → click ref
   - Do NOT open the More-send-options chevron / Schedule dialog.

   **Schedule (Schedule-send dialog):** find→ref chain:
   - find "More send options chevron" → click ref
   - find "Schedule send menu item" → click ref
   - find "Custom time button" → click ref
   - find "Select a time combobox" → triple-click ref → type target slot (must be a 30-min preset, e.g. "2:00 PM" or "3:30 PM")
   - find "<target time> option in listbox" → click ref to commit
   - find "Send button in custom date dialog" → click ref

6. **Post-send screenshot** — confirm "Sending..."/sent toast (immediate) or "Send scheduled for ..." toast (schedule).

**Time picker constraints (HARD):**
- Outlook only accepts 30-min preset slots (12:00, 12:30, 1:00, ...).
- Round `outlook_time` to nearest preset before typing.
- After typing, MUST click the matching dropdown option — typing alone doesn't commit; Outlook reverts to default 8:00 AM if you skip the click.

**Address-field pitfalls:**
- Tab after typing email = address gets mangled mid-string (e.g. "vanderbilt.edu" becomes "derbilt.edu").
- ALWAYS use Return + click suggestion pill.
- If pill shows "Use this address: <email>" instead of contact card, click that — autocomplete had no match but address is valid.

**Recovery if compose corrupted:**
- Click "Remove recipient" button (top of compose) to clear bad pill.
- Refill the field with Return + suggestion-click pattern.
- Body usually survives — don't re-paste unless visibly broken.

**Fallback if Chrome MCP fails:**
- Try `https://outlook.office.com/mail/deeplink/compose`
- If persistent failure, show user the full draft body and subject for manual send

#### Draft mode (save to Drafts, no send)

Steps 1-4 same as above. Then:
5. Close compose window via the **X** (auto-saves as draft) OR click trash icon to discard if draft was placeholder.
6. Confirm draft appears in Drafts folder.
7. Report to user: draft saved, ready for review. Do NOT call `mark-sent`.

---

### Step 5: Update CRM

**Draft mode:** do NOT mark as sent. The draft hasn't been sent yet. Just report to user.

**Immediate or Schedule-send mode:** after Outlook confirms the send/schedule toast:
```bash
python3 outreach/scripts/send_first_touch.py mark-sent <csv_row_idx>
```
Sets Status → "Contacted no reply", stamps today's date, increments touch points.

---

### Step 6: Report to user

**Draft mode:**
- Draft saved for: [Name] ([Title] at [System])
- Email address: [address]
- Word count: [N] words
- Check Drafts folder in Outlook to review

**Immediate mode:**
- Sent to: [Name] ([Title] at [System])
- Email: [address]
- Sent at: [target send_at_display] (pacing target [n]/[count])
- Word count: [N] words
- CRM row [crm_row_number] → "Contacted no reply"
- Queue remaining: run `list` to check

**Schedule-send mode:**
- Sent to: [Name] ([Title] at [System])
- Email: [address]
- Scheduled: [compute-send-time `outlook_et`]
- Word count: [N] words
- CRM row [crm_row_number] → "Contacted no reply"
- Queue remaining: run `list` to check

---

## Batch sends

Run pipeline sequentially per contact (Outlook compose handles one at a time).

After each: mark-sent → call `next` → repeat.

Send slots are randomized per `next` call — collisions are rare but possible. If two contacts land on the same slot, note it in your report.

---

## Edge cases

**Contact has two emails** (e.g. `email1@hospital.org , email2@gmail.com`):
- Use institutional email first (hospital/university domain)
- Fall back to second if institutional bounces

**Email blank:** `next` skips it. Tell user: "[Name] has no email — fill manually or use Apollo to enter the send queue."

**Personalization blank:** `next` skips it. Tell user: "[Name] needs a research pass — run the prospecting skill first."

**Retry a validation failure:**
1. Read the personalization blob (use `preview <row_idx>`)
2. Generate fresh p1/p2 with a different angle
3. Have user reset status to blank manually (or do it in the CSV directly)
4. Run from Step 3

---

## Email template (reference — do not modify)

```
Hey {name},

I'm Lava Panta. I saw {p1}. I'm a Northeastern undergrad working with Barock Tesfaye, who shadowed doctors in Ethiopia and made a documentary about what they go through. The biggest issue was a shortage of supplies. We've raised money through the documentary, purchased ultrasound machines for those hospitals, and gotten in touch with Project C.U.R.E.

Now we're trying to learn from people in the {geo} who understand this work, to better understand what more could realistically be done for hospitals back home.

We'd love your advice. A 10-minute call this week would be invaluable. Thanks for taking the time to read this!

Sincerely,
Lava Panta
Barock C. Tesfaye

P.S. {p2}
```

`{geo}` is the geography phrase (e.g. "Boston medical system", "New York medical system"). `next`/`preview`/`validate` auto-derive it from the `Location:` bullet in the personalization blob (city → "<City> medical system"), defaulting to "Boston medical system" when no Location is present. Suburbs/boroughs roll up to their metro hub via `METRO_ALIASES` (Cambridge → Boston, Brooklyn → New York). To override or add a region, pass `"geo"` in the validate stdin JSON alongside `p1`/`p2`. Do NOT hardcode "Boston" — use the derived/overridden value.

Hard rules for the filled draft:
- Under 150 words (greeting through P.S. inclusive)
- No em dashes (—) anywhere
- No double-hyphens (--)
- No leftover `{placeholder}` text
- p1 and p2 must be specific enough that you could only write them for this person
