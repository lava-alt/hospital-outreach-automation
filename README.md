# Hospital Outreach Automation

A 4-touch cold outreach system for global-health stakeholders at Boston-area hospitals, run by an
AI agent driving a real Outlook inbox. The ask was advice, not money: a Northeastern undergrad
working on Ethiopian hospital supply shortages, looking for ten minutes from people who understand
the space.

**It worked.** This sequence is how the project reached Project C.U.R.E., the largest medical device
donation nonprofit in the world.

## How it works

A CSV acts as the CRM. Python owns state and validation; the agent owns judgment and the browser.

```bash
python3 send_first_touch.py next          # next contact, as JSON
python3 send_first_touch.py validate      # validate a drafted email (stdin JSON)
python3 send_first_touch.py mark-sent 12  # stamp the touch, schedule the next
python3 send_follow_up.py list            # who is due, and for which follow-up
```

- **Sessions are gated.** `next` refuses to hand out a contact without an open session, so a crashed
  run cannot silently double-send.
- **Drafts are validated before sending** — personalization tokens must be filled, banned phrases
  (`partnership-with-us` language, anything that reads as a pitch rather than a request for advice)
  are rejected.
- **Cadence is enforced in data, not memory.** One week between touches, four touches maximum, and a
  reply drops the contact out of the sequence permanently.

## What the skill files are for

`SKILL.md` and `FOLLOWUP_SKILL.md` are the agent's operating instructions for driving Outlook through
a browser. Most of their length is hard-won failure handling:

> Never use Tab or Return after typing an address — both truncate it regardless of waits. Commit via
> a click on the autocomplete suggestion pill.

Outlook's DOM races with its own autocomplete, and an address silently mangled mid-string is worse
than a crash, because the send still succeeds. The waits and the `find`-as-serialization-barrier
pattern exist because that failure happened and had to be designed out.

## Configuration

```bash
export CC_ADDRESS="collaborator@example.edu"   # collaborator CC'd on every send
```

Contact data lives in a local CSV that is not published.
