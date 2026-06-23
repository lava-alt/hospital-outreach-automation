---
name: hospital-followup
description: |
  Boston Hospital Outreach follow-up automation. Use this skill whenever the user asks to run follow-ups, "send the next follow-up", "follow up with people who haven't replied", "check for replies", "run the follow-up sequence", or any reference to the 3-touch follow-up cadence. Reads the same Boston Hospital Outreach CRM and sends on the same Outlook inbox as the first-touch sender. Covers: reply detection in Outlook → CRM reconcile → due-contact selection → follow-up template fill → validation → Outlook compose (reply-in-thread) → CRM update.
---

# Boston Hospital Outreach — Follow-Up Skill

You run the **3-touch follow-up sequence** on contacts who were already first-touched
and have **not replied**. This lives on the SAME Outlook inbox as the first-touch sender
(`outreach/SKILL.md`) and deliberately reuses its Outlook-compose structure. The only
things dropped: **schedule-send dialog and pacing** — follow-ups send immediately
(direct Send) or to Drafts.

The ask never changes: advice only, a 10-minute call. Never a donation/device/partnership ask.

## The Sequence

| Touch points already sent | Follow-up this run sends | Subject | Signed |
|---|---|---|---|
| 1 (initial only) | **Follow-up 1** — progress update | `Quick update on the Ethiopia device pipeline` | Lava only |
| 2 (initial + FU1) | **Follow-up 2** — personalized nudge | `Re: Quick update on the Ethiopia device pipeline` | Lava + Barock |
| 3 (initial + FU2) | **Follow-up 3** — respectful close | `Re: Quick update on the Ethiopia device pipeline` | Lava + Barock |
| 4 (FU3 sent) | nothing — sequence complete | — | — |

Cadence: each follow-up fires **1 week** after the last contact, only while the contact
has **not replied**. FU2 and FU3 are `Re:` replies and should go **in the existing thread**.

## The Script Helper

All mechanical CRM operations go through:
```
outreach/scripts/send_follow_up.py
```

Always `cd` to the base path first:
```bash
cd "/Users/lavapanta/Desktop/Hospital prospecting, outreach reachout automation"
```

Commands:
```bash
python3 outreach/scripts/send_follow_up.py session-start        # Open session (stdin JSON) — REQUIRED before next
python3 outreach/scripts/send_follow_up.py session-show         # Show active + last session (repeat offer)
python3 outreach/scripts/send_follow_up.py session-end          # Close session
python3 outreach/scripts/send_follow_up.py list                 # Due queue + not-yet-due waiting list
python3 outreach/scripts/send_follow_up.py next                 # Next due contact as JSON (GATED)
python3 outreach/scripts/send_follow_up.py preview <row_idx>    # Inspect a row + its current FU#
python3 outreach/scripts/send_follow_up.py validate            # Validate draft (stdin JSON)
python3 outreach/scripts/send_follow_up.py mark-sent <row_idx>  # +1 touch point, stamp date, set +7 due date
python3 outreach/scripts/send_follow_up.py mark-replied <row_idx> # Status → Replied (drops out forever)
python3 outreach/scripts/send_follow_up.py mark-failed <row_idx>  # Status → Not sent (validation failure)
```

A contact is **due** when: `Status == "Contacted no reply"` AND email present AND
`1 <= touch points < 4` AND `today >= due date` (due date = `Next Follow-Up Date` if set,
else `Date of Last Contact + 7 days`). The script enforces all of this — trust `list`/`next`.

---

## Reply Detection (do this FIRST, every run)

The CRM is the source of truth for *when* an email was sent. **Outlook is the source of
truth for whether someone replied.** Before sending any follow-up, reconcile the two so
you never follow up with someone who already wrote back.

For each contact the script reports as **due** (`list` or `next`):

1. In Outlook (Chrome MCP), search the inbox for the contact's email address or the thread
   subject (`Quick update on the Ethiopia device pipeline`).
2. Open the thread. Determine if the contact sent a reply (a message **from their address**,
   not your own sent copies / auto-replies / bounces).
   - **Out-of-office / auto-reply / delivery-failure** → NOT a real reply. Proceed with the
     follow-up. (A bounce means bad address — flag to the user instead of resending.)
3. **If a real reply exists** → call `mark-replied <row_idx>`. Status becomes `Replied`,
   the contact drops out of the sequence permanently. Do nothing else for them. Report it.
4. **If no reply** → the contact stays due. Continue to the send flow below.

This catches contacts the CRM marked wrong. Do the reply check even though `Status` is
still `Contacted no reply` — the CRM may simply be stale.

---

## Interactive Start Flow

**HARD GATE.** `next` refuses to return contacts until `session-start` runs. This forces
the start questions every session, including headless/scheduled runs. Do not bypass by
editing the CSV. (Follow-up sessions use their own session files, separate from first-touch.)

### Step 0 — Read last session (repeat offer)
```bash
python3 outreach/scripts/send_follow_up.py session-show
```
`has_last: true` → offer "Rerun last session" as the first widget option.

### Step A — Ask the start questions (single AskUserQuestion)
1. **Send mode** — `Immediate send` / `Draft mode` (plus `Rerun last session` if Step 0 had one).
   No schedule mode here — follow-ups send now or save as drafts.
2. **Batch scope** — `Whole queue` / `Specific count` / `Specific names`.

If scope is `count` or `names`, collect the value (free-text follow-up is fine).

### Step B — Open the session
```bash
python3 outreach/scripts/send_follow_up.py session-start << 'ENDJSON'
{
  "mode": "immediate",     // immediate | draft
  "scope": "count",        // queue | count | names
  "count": 5,              // when scope == count
  "names": []              // when scope == names
}
ENDJSON
```
Session lives 12h. `next` echoes the session back in its `session` key. **Enforce scope
yourself**: stop after `count` sends; only process the listed `names`. Close when done:
```bash
python3 outreach/scripts/send_follow_up.py session-end
```

**Immediate** = send right now, sequentially, one contact at a time (no pacing, no future
timestamp). **Draft** = compose + save to Drafts, do NOT call `mark-sent`.

---

## Send Pipeline (per contact)

### Step 1 — Get the next due contact
```bash
python3 outreach/scripts/send_follow_up.py next
```
Key fields: `csv_row_idx`, `first_name` (already "Dr. LastName" or "FirstName"),
`email`, `follow_up_number` (1/2/3), `subject`, `is_thread_reply` (true for FU2/FU3),
`needs_personalization` (true only for FU2), `personalization`.

### Step 2 — Hooks (Follow-up 2 ONLY)
FU1 and FU3 are fixed bodies — skip to Step 3. For **FU2**, write the one personalized
sentence-span from the `personalization` blob (and the contact's original Touch-1 angle):
- **p1** = the pipeline question their work speaks to, phrased as a problem on the ground
  (catching failures early, what equipment survives the conditions, training gaps, etc.).
- **p2** = their specific work, **named**, that answers p1.

The two are halves of one sentence: *"...the piece we keep coming back to is **{p1}**.
That's exactly **{p2}**."* Concrete over abstract; name the actual program/research.
No em dashes. Example:
- p1 = `how you catch what's failing in a hospital before it turns into a crisis`
- p2 = `the problem your GeoSentinel surveillance work has spent years solving across 70 sites`

### Step 3 — Validate (up to 3 attempts)
```bash
python3 outreach/scripts/send_follow_up.py validate << 'ENDJSON'
{ "row_idx": <csv_row_idx>, "p1": "<FU2 only>", "p2": "<FU2 only>" }
ENDJSON
```
(For FU1/FU3 omit p1/p2.) Returns `valid`, `body` (full filled email), `subject`,
`is_thread_reply`, `word_count`, `issues`. `valid:false` → fix and retry. After 3 fails →
`mark-failed` and report.

### Step 4 — Compose in Outlook via Chrome
Reuse the first-touch compose structure exactly (`outreach/SKILL.md` Step 4): JS-first,
**2 screenshots max** (pre-send + post-send), address fields via Return + click suggestion
pill (never Tab), 6s waits between address actions.

**Threading — the one difference from first-touch:**
- **FU1** (`is_thread_reply: false`) → "New mail". To = contact email, Cc = `collaborator@example.edu`,
  Subject = `Quick update on the Ethiopia device pipeline`. Fill body via the JS template.
- **FU2 / FU3** (`is_thread_reply: true`) → open the contact's existing thread (you already
  found it during reply detection) and click **Reply All** to preserve the threading and the
  Cc to Barock. To/Cc/Subject auto-fill — only set the **body** via the JS template
  (`bodyEl.innerText = BODY`). Verify To still shows the contact and Cc still has Barock in
  the pre-send screenshot.
  - If you cannot locate the original thread, fall back to a New mail with subject
    `Re: Quick update on the Ethiopia device pipeline`, To = contact, Cc = Barock.

**CRITICAL — reply-compose has two body elements.** In any Reply / Reply All
view, `[aria-label="Message body"]` matches BOTH the original message DIV
(read-only, `role="document"`) AND the compose textbox (`role="textbox"`,
`isContentEditable=true`). Naked `document.querySelector('[aria-label="Message body"]')`
returns the original — your body lands in a non-editable element and the reply
sends BLANK. Always scope: `[aria-label="Message body"][role="textbox"]` or
`Array.from(...).find(e => e.isContentEditable)`. Use the JS template in
`outreach/SKILL.md` Step 4 (already fixed) and verify the returned `body` length
is > 0 BEFORE clicking Send.

**Send branch (HARD — never the schedule dialog):**
- find "Send button" (aria-label `Send`) → click ref. Do NOT open the More-send-options
  chevron / Schedule dialog. Follow-ups are immediate only.

**Draft mode:** fill compose, close via **X** to auto-save, confirm it lands in Drafts, do
NOT `mark-sent`.

### Step 5 — Update CRM
**Immediate** (after the sent toast confirms):
```bash
python3 outreach/scripts/send_follow_up.py mark-sent <csv_row_idx>
```
Increments touch points, stamps today's `Date of Last Contact`, sets `Next Follow-Up Date`
= today + 7 (cleared once FU3 sent → tp 4 → sequence complete).

**Draft:** do NOT mark-sent. Report the draft is ready for review.

### Step 6 — Report
- Followed up: [Name] ([Title] at [System]) — **Follow-up [N]**
- Email: [address] | threaded reply: [yes/no]
- Word count: [N]
- CRM row [crm_row_number]: touch points → [N], next due [date or "sequence complete"]
- Replies found this run: [list of names marked Replied]
- Queue remaining: run `list`

---

## Batch runs
Process sequentially (Outlook composes one at a time): reply-check → send or mark-replied →
`next` → repeat. Stop at scope `count`, when the queue empties, or when the user stops.
Then `session-end`.

## Edge cases
- **Reply found** → `mark-replied`, never send. Positive vs. neutral does not matter here;
  any genuine reply ends the sequence (the user handles the conversation from there).
- **Bounce / bad address** → do not resend. Flag the contact to the user to fix the email.
- **Two emails on a row** (`a@x.org , b@gmail.com`) → reply-check and reply-in-thread on the
  institutional address first.
- **`due None` in `list`** → contacted but missing both `Date of Last Contact` and
  `Next Follow-Up Date`. Never auto-fires. Flag to the user to set a date.
- **Sequence complete (tp 4)** → contact is silently excluded; nothing more to do.

## Template reference (verbatim — do not modify; no em dashes anywhere)

**Follow-up 1** — Subject: `Quick update on the Ethiopia device pipeline`
```
Hi {name},

Hope you've had an amazing week. Wanted to share a quick update: we recently got in touch with BHI, a biomedical engineering organization that helps set up medical devices in low-resource areas, as well as Mass General. We're very excited about where this is going, and I'd love to hear your thoughts and get your advice as we build this device pipeline to the hospitals in Ethiopia that need it.

Sincerely,
Lava Panta
```

**Follow-up 2** — Subject: `Re: Quick update on the Ethiopia device pipeline` (reply in thread)
```
Hi {name},

I know your time is tight, so I'll keep this short. We're still building out the device pipeline to hospitals in Ethiopia, and the piece we keep coming back to is {p1}. That's exactly {p2}.

Even 10 minutes of your perspective would shape how we approach this. Happy to work entirely around your schedule.

Sincerely,
Lava Panta
Barock C. Tesfaye
```

**Follow-up 3** — Subject: `Re: Quick update on the Ethiopia device pipeline` (reply in thread)
```
Hi {name},

I don't want to keep cluttering your inbox, so this will be my last note. I completely understand that the timing may not be right.

The offer stands open whenever it suits you, and if there's ever a moment down the line where a quick conversation makes sense, I'd be grateful for it. Thank you for the work you do in this space, it matters more than most people see.

Sincerely,
Lava Panta
Barock C. Tesfaye
```
