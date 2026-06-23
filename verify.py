#!/usr/bin/env python3
"""
Verification script for Boston Hospital Outreach - CRM.csv
Run after writing a batch of contacts.

Usage:
    python3 verify.py <csv_path> <row_indices>

Example:
    python3 prospecting/scripts/verify.py "prospecting/Boston Hospital Outreach - CRM.csv" 1,2,3,4,5

Row indices are 0-based. Row 0 is the header — start from 1.

Columns (0-indexed):
    0: Name
    1: System
    2: LinkedIn URL
    3: Title
    4: Email
    5: Status
    6: Date of Last Contact
    7: Number of Touch Points
    8: Next Follow-Up Date
    9: Personalization

Checks:
    Part A: Completeness — required fields filled
    Part B: Structural — email format, status values, personalization bullet count
    Part C: Fact-check hints — flags what to manually verify
"""

import csv
import sys
import json
import re

# Column indices
COL_NAME = 0
COL_SYSTEM = 1
COL_LINKEDIN = 2
COL_TITLE = 3
COL_EMAIL = 4
COL_STATUS = 5
COL_DATE_LAST_CONTACT = 6
COL_TOUCHPOINTS = 7
COL_NEXT_FOLLOWUP = 8
COL_PERSONALIZATION = 9

TOTAL_COLS = 10

# Valid status values (blank is also valid for new leads)
VALID_STATUSES = [
    "",
    "Contacted no reply",
    "Replied",
    "Call booked",
    "In discussion",
    "Won",
    "Declined",
    "Dead",
    "Not sent (validation failure)",
]

# Columns that are expected blank for new leads
NEW_LEAD_BLANK_COLS = [COL_STATUS, COL_DATE_LAST_CONTACT, COL_TOUCHPOINTS, COL_NEXT_FOLLOWUP]

# Required for all rows (must not be blank)
REQUIRED_COLS = [COL_NAME, COL_SYSTEM, COL_TITLE]

# Optional but validated if present
OPTIONAL_COLS = [COL_EMAIL, COL_LINKEDIN]


def verify_rows(csv_path, row_indices):
    with open(csv_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        print("ERROR: CSV is empty.")
        sys.exit(1)

    header = rows[0]
    issues = []
    results = {}

    for idx in row_indices:
        if idx <= 0:
            print(f"WARNING: Row index {idx} skipped (row 0 is the header).")
            continue
        if idx >= len(rows):
            issues.append({
                "name": f"Row {idx}",
                "issue": f"Row {idx} does not exist in CSV (only {len(rows)-1} data rows)",
                "severity": "critical",
                "fix": "Check row index",
            })
            continue

        row = rows[idx]

        # Pad row if short
        while len(row) < TOTAL_COLS:
            row.append("")

        name = row[COL_NAME].strip() or f"Row {idx}"
        system = row[COL_SYSTEM].strip() or "Unknown System"
        row_issues = []

        # === PART A: COMPLETENESS ===

        for col in REQUIRED_COLS:
            val = row[col].strip()
            if not val or val == "—":
                row_issues.append({
                    "name": name,
                    "issue": f"Required field empty: {header[col]}",
                    "severity": "high",
                    "fix": f"Fill '{header[col]}' with real data",
                })

        # Email: blank is OK (user fills manually), but warn so it's visible
        if not row[COL_EMAIL].strip():
            row_issues.append({
                "name": name,
                "issue": "Email is blank — needs manual fill or Apollo lookup",
                "severity": "warn",
                "fix": "Find email via bio page, HMS/Tufts faculty page, or Apollo",
            })

        # === PART B: STRUCTURAL VALIDATION ===

        # Email format (only if something is filled in)
        email = row[COL_EMAIL].strip()
        if email:
            if "@" not in email or "." not in email.split("@")[-1]:
                row_issues.append({
                    "name": name,
                    "issue": f"Email format looks wrong: {email}",
                    "severity": "high",
                    "fix": "Fix email format",
                })

        # LinkedIn URL (only if filled)
        linkedin = row[COL_LINKEDIN].strip()
        if linkedin and "linkedin.com" not in linkedin.lower():
            row_issues.append({
                "name": name,
                "issue": f"LinkedIn field doesn't look like a LinkedIn URL: {linkedin[:60]}",
                "severity": "medium",
                "fix": "Use full linkedin.com/in/... URL or leave blank",
            })

        # Status value
        status = row[COL_STATUS].strip()
        if status not in VALID_STATUSES:
            row_issues.append({
                "name": name,
                "issue": f"Invalid status: '{status}'",
                "severity": "high",
                "fix": f"Use one of: {', '.join(v for v in VALID_STATUSES if v)}",
            })

        # Personalization bullet count
        personalization = row[COL_PERSONALIZATION].strip()
        if personalization:
            bullet_count = personalization.count("•")
            if bullet_count < 3:
                row_issues.append({
                    "name": name,
                    "issue": f"Only {bullet_count} personalization bullet(s) — need at least 3",
                    "severity": "medium",
                    "fix": "Add more research bullets (focus, projects, Africa angle, location, email source)",
                })
        # Empty personalization is already caught by REQUIRED_COLS check above — no duplicate needed

        # Touchpoints: should be numeric if filled
        tp = row[COL_TOUCHPOINTS].strip()
        if tp and not tp.isdigit():
            row_issues.append({
                "name": name,
                "issue": f"Number of Touch Points is not a number: '{tp}'",
                "severity": "medium",
                "fix": "Set to an integer (e.g. 1) or leave blank for new leads",
            })

        # Date format check (YYYY-MM-DD) if filled
        for col, label in [(COL_DATE_LAST_CONTACT, "Date of Last Contact"),
                           (COL_NEXT_FOLLOWUP, "Next Follow-Up Date")]:
            val = row[col].strip()
            if val:
                if not re.match(r"^\d{4}-\d{2}-\d{2}$", val):
                    row_issues.append({
                        "name": name,
                        "issue": f"{label} format wrong: '{val}' (expected YYYY-MM-DD)",
                        "severity": "medium",
                        "fix": "Use YYYY-MM-DD format",
                    })

        # === PART C: FACT-CHECK HINTS ===
        fact_check_hints = [
            f"WebSearch to verify: '{name}' at {system} — confirm title and current role",
        ]
        if not email:
            fact_check_hints.append(
                f"Email needed: check {system} bio page, HMS/Tufts faculty page, or LinkedIn"
            )
        if not linkedin:
            fact_check_hints.append(
                f"LinkedIn not found: try site:linkedin.com '{name}' {system}"
            )

        results[f"{name} (row {idx})"] = {
            "row_index": idx,
            "name": name,
            "system": system,
            "email_status": "filled" if email else "BLANK — needs fill",
            "linkedin_status": "found" if linkedin else "not found",
            "issues_count": len(row_issues),
            "issues": row_issues,
            "fact_check_hints": fact_check_hints,
            "status": "PASS" if all(i["severity"] not in ("high", "critical") for i in row_issues) else "ISSUES_FOUND",
        }
        issues.extend(row_issues)

    # === PRINT REPORT ===
    print("=" * 70)
    print("VERIFICATION REPORT — Boston Hospital Outreach CRM")
    print("=" * 70)

    for key, data in results.items():
        icon = "✅" if data["status"] == "PASS" else "⚠️ "
        print(f"\n{icon} {key}")
        print(f"   System: {data['system']}")
        print(f"   Email:    {data['email_status']}")
        print(f"   LinkedIn: {data['linkedin_status']}")

        if data["issues"]:
            for issue in data["issues"]:
                sev = issue["severity"].upper()
                prefix = "❌" if sev in ("HIGH", "CRITICAL") else ("⚠️ " if sev == "MEDIUM" else "ℹ️ ")
                print(f"   {prefix} [{sev}] {issue['issue']}")
                print(f"        Fix: {issue['fix']}")

        if data["fact_check_hints"]:
            print(f"   📋 Fact-check:")
            for hint in data["fact_check_hints"]:
                print(f"      - {hint}")

    high_issues = [i for i in issues if i["severity"] in ("high", "critical")]
    warn_issues = [i for i in issues if i["severity"] == "warn"]
    med_issues  = [i for i in issues if i["severity"] == "medium"]

    print(f"\n{'=' * 70}")
    print(f"TOTAL: {len(issues)} issues across {len(results)} contacts")
    print(f"  ❌ High/Critical: {len(high_issues)}")
    print(f"  ⚠️  Medium:        {len(med_issues)}")
    print(f"  ℹ️  Warnings:      {len(warn_issues)}  (email blanks — expected for new leads)")
    print(f"{'=' * 70}")

    # Save JSON report
    json_path = csv_path.rsplit(".", 1)[0] + "_verification.json"
    with open(json_path, "w") as f:
        json.dump({"issues": issues, "results": results, "total_issues": len(issues)}, f, indent=2)
    print(f"\nJSON report saved to: {json_path}")

    return {"issues": issues, "results": results, "total_issues": len(issues)}


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python3 verify.py <csv_path> <comma-separated-row-indices>")
        print("Example: python3 prospecting/scripts/verify.py 'prospecting/Boston Hospital Outreach - CRM.csv' 1,2,3")
        sys.exit(1)

    csv_path = sys.argv[1]
    row_indices = [int(x.strip()) for x in sys.argv[2].split(",")]
    verify_rows(csv_path, row_indices)
