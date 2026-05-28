#!/usr/bin/env python3
"""
extract_registers.py
=====================
Extract structured data from TVET CDACC "Assessment Registration" register PDFs
(the kind produced by dompdf, one unit per section, candidates in a table).

Strategy
--------
Rather than relying on fragile column whitespace, we:
  1. Pull every word from each page together with its (x, y) position.
  2. Re-group words into lines by their vertical position (`top`), sorting
     each line left-to-right by `x0`. This faithfully rebuilds lines such as
     "UNIT NAME : ...   UNIT CODE :..." that are visually side-by-side.
  3. Walk the lines as a small state machine. The registration number
     (e.g. 0320134P/CSC/6/2024/004) is a very distinctive anchor, so a
     candidate row is matched as:  <S/N>  <NAME...>  <REG NO>
  4. Units are keyed by their UNIT CODE, so headers repeated on continuation
     pages — and "orphan" rows on pages with no header at all — merge into
     the same unit. Candidates are de-duplicated by registration number.
  5. Centre name and course name overflow (text that wraps to the next line
     because the value is too long for the header row) is handled by
     expect_centre_wrap / expect_course_wrap flags, analogous to the
     existing unit-name continuation logic.

Usage
-----
    python extract_registers.py input.pdf --json out.json --csv out.csv
    python extract_registers.py input.pdf --roster roster.csv
    python extract_registers.py input.pdf            # prints a summary

Requires: pdfplumber  (pip install pdfplumber)
"""

import argparse
import csv
import json
import re
import sys
from collections import defaultdict, OrderedDict

import pdfplumber

# --------------------------------------------------------------------------- #
# Patterns
# --------------------------------------------------------------------------- #

# Registration number, e.g. 0320134P/CSC/6/2024/004
REG_RE = re.compile(r"[A-Z0-9]{6,}/[A-Z]+/\d+/\d{4}/\d+")

# A candidate table row:  serial   name (greedy-but-minimal)   reg-number
ROW_RE = re.compile(r"^\s*(\d{1,4})\s+(.+?)\s+(" + REG_RE.pattern + r")\s*$")

# Header fields. These appear on a single rebuilt line, often with the
# left field and right field side-by-side.
# Exam series header, e.g. "MARCH/APRIL 2026" or "JULY/AUGUST 2026"
SERIES_RE  = re.compile(r"^[A-Z]+/[A-Z]+\s+\d{4}$")

REPORT_RE = re.compile(r"Report Type:\s*(.+)", re.I)
CENTRE_RE = re.compile(r"CENTRE NAME\s*:\s*(.+?)\s*(?:CENTRE CODE\s*:\s*(\S+))?\s*$", re.I)
CENTRE_CODE_RE = re.compile(r"CENTRE CODE\s*:\s*(\S+)", re.I)
COURSE_RE = re.compile(r"COURSE NAME\s*:\s*(.+?)\s+COURSE LEVEL\s*:\s*(.+?)\s*$", re.I)
UNIT_RE = re.compile(r"UNIT NAME\s*:\s*(.+?)\s+UNIT CODE\s*:\s*(\S+)\s*$", re.I)
TABLE_HDR_RE = re.compile(r"^\s*S/N\b", re.I)

# Lines that are structural noise and must never be treated as name wraps.
NOISE_RE = re.compile(
    r"(MARCH/APRIL|Report Type|CENTRE|COURSE|UNIT|National ID|Signature|"
    r"Date\s*:|Name of Cent|^\s*\d+\s*/\s*\d+\s*$)",
    re.I,
)


# --------------------------------------------------------------------------- #
# Line reconstruction
# --------------------------------------------------------------------------- #

def page_lines(page, y_tol=3):
    """Return the page's text as a list of lines, rebuilt from word positions.

    Words whose vertical positions are within `y_tol` points are treated as
    belonging to the same line and concatenated left-to-right.
    """
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    buckets = defaultdict(list)
    for w in words:
        buckets[round(w["top"] / y_tol)].append(w)
    lines = []
    for key in sorted(buckets):
        ws = sorted(buckets[key], key=lambda w: w["x0"])
        lines.append(" ".join(w["text"] for w in ws))
    return lines


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #

def extract(pdf_path, on_log=None, on_progress=None):
    """Parse the register PDF and return a dict with centre info + units.

    Optional callbacks (called from the calling thread — use a queue when
    driving from a GUI worker thread):
        on_log(level: str, message: str)   level in: info|step|found|success|warn
        on_progress(current: int, total: int)
    """
    def _log(level, msg):
        if on_log:
            on_log(level, msg)

    state = {
        "report_type": None,
        "centre_name": None,
        "centre_code": None,
        "course_name": None,
        "course_level": None,
        "series":      None,   # e.g. "MARCH/APRIL 2026"
    }
    units = OrderedDict()           # unit_code -> unit dict
    current_unit = None             # the unit currently receiving rows
    expect_wrap = False             # between UNIT NAME and S/N → unit-name continuation
    expect_centre_wrap = False      # after CENTRE NAME line → centre-name continuation
    expect_course_wrap = False      # after COURSE NAME line → course-name continuation

    def get_unit(code, name):
        if code not in units:
            units[code] = {
                "unit_code": code,
                "unit_name": name,
                "report_type": state["report_type"],
                "course_name": state["course_name"],
                "course_level": state["course_level"],
                "centre_name": state["centre_name"],
                "centre_code": state["centre_code"],
                "candidates": [],
                "_reg_seen": set(),
            }
        return units[code]

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        _log("step", f"Opened — {total_pages} page{'s' if total_pages != 1 else ''} detected")
        for pg_idx, page in enumerate(pdf.pages, 1):
            _log("info", f"Reading page {pg_idx} / {total_pages}")
            if on_progress:
                on_progress(pg_idx, total_pages)
            for line in page_lines(page):
                stripped = line.strip()
                if not stripped:
                    continue

                # --- candidate row? (check first: most common line) ---------
                m = ROW_RE.match(line)
                if m and current_unit is not None:
                    expect_wrap = False
                    expect_centre_wrap = False
                    expect_course_wrap = False
                    serial, name, reg = m.group(1), m.group(2).strip(), m.group(3)
                    name = re.sub(r"\s+", " ", name)
                    if reg not in current_unit["_reg_seen"]:
                        current_unit["_reg_seen"].add(reg)
                        current_unit["candidates"].append(
                            {"sn": int(serial), "name": name, "reg_no": reg}
                        )
                    continue

                # --- header fields ------------------------------------------
                # Capture exam series (e.g. "MARCH/APRIL 2026") on first sight
                if state["series"] is None and SERIES_RE.match(stripped):
                    state["series"] = stripped
                    continue

                if (mm := REPORT_RE.search(stripped)):
                    state["report_type"] = mm.group(1).strip()
                    continue

                if (mm := CENTRE_RE.match(stripped)):
                    state["centre_name"] = mm.group(1).strip()
                    if mm.group(2):
                        state["centre_code"] = mm.group(2).strip()
                    elif (cc := CENTRE_CODE_RE.search(stripped)):
                        state["centre_code"] = cc.group(1).strip()
                    # Centre name may overflow to the next line —
                    # flag it so the continuation can be appended.
                    expect_centre_wrap = True
                    expect_course_wrap = False
                    continue

                if (mm := COURSE_RE.match(stripped)):
                    state["course_name"] = mm.group(1).strip()
                    state["course_level"] = mm.group(2).strip()
                    # Course name may overflow to the next line.
                    expect_centre_wrap = False
                    expect_course_wrap = True
                    continue

                if (mm := UNIT_RE.match(stripped)):
                    code = mm.group(2).strip()
                    name = re.sub(r"\s+", " ", mm.group(1).strip())
                    is_new = code not in units
                    current_unit = get_unit(code, name)
                    expect_centre_wrap = False
                    expect_course_wrap = False
                    # Only a freshly-seen unit may still need its wrapped name
                    # continuation; repeated headers on continuation pages must not.
                    expect_wrap = is_new
                    if is_new:
                        _log("found", f"Unit found: {code}  —  {name[:55]}")
                    continue

                if TABLE_HDR_RE.match(stripped):
                    expect_wrap = False
                    expect_centre_wrap = False
                    expect_course_wrap = False
                    continue

                # --- centre-name continuation --------------------------------
                # e.g. "SCIENCE AND TECHNOLOGY(RVIST)" after the truncated
                # "CENTRE NAME: RIFT VALLEY INSTITUTE OF  CENTRE CODE :XXXXX"
                if expect_centre_wrap and not NOISE_RE.search(stripped):
                    state["centre_name"] = (
                        state["centre_name"] + " " + stripped
                    ).strip()
                    expect_centre_wrap = False
                    continue

                # --- course-name continuation --------------------------------
                # e.g. "Technology" after the truncated
                # "COURSE NAME : Information and Communication  COURSE LEVEL :Level 5"
                if expect_course_wrap and not NOISE_RE.search(stripped):
                    state["course_name"] = (
                        state["course_name"] + " " + stripped
                    ).strip()
                    expect_course_wrap = False
                    continue

                # --- unit-name continuation (e.g. the lone word "Systems") --
                if expect_wrap and current_unit is not None and not NOISE_RE.search(stripped):
                    current_unit["unit_name"] = (
                        current_unit["unit_name"] + " " + stripped
                    ).strip()
                    continue

    # tidy up internal bookkeeping
    result_units = []
    for u in units.values():
        u.pop("_reg_seen", None)
        u["candidate_count"] = len(u["candidates"])
        result_units.append(u)

    return {
        "centre_name": state["centre_name"],
        "centre_code": state["centre_code"],
        "course_name": state["course_name"],
        "course_level": state["course_level"],
        "series":      state["series"],
        "unit_count":  len(result_units),
        "units":       result_units,
    }


# --------------------------------------------------------------------------- #
# Outputs
# --------------------------------------------------------------------------- #

def write_json(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def write_csv(data, path):
    """Flat CSV: one row per (unit, candidate)."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
            "centre_code", "course_name", "course_level", "report_type",
            "unit_code", "unit_name", "sn", "candidate_name", "reg_no",
        ])
        for u in data["units"]:
            for c in u["candidates"]:
                w.writerow([
                    data["centre_code"], u["course_name"], u["course_level"],
                    u["report_type"], u["unit_code"], u["unit_name"],
                    c["sn"], c["name"], c["reg_no"],
                ])


def build_roster(data):
    """Collapse to unique candidates, with the list of units each is sitting."""
    roster = OrderedDict()  # reg_no -> {name, units, unit_names, course_name, course_level}
    for u in data["units"]:
        for c in u["candidates"]:
            entry = roster.setdefault(
                c["reg_no"],
                {
                    "name":         c["name"],
                    "units":        [],
                    "unit_names":   [],
                    # course info captured on first encounter for this reg no
                    "course_name":  (u.get("course_name")  or "").strip(),
                    "course_level": (u.get("course_level") or "").strip(),
                },
            )
            entry["units"].append(u["unit_code"])
            entry["unit_names"].append(u["unit_name"])
    return roster


def write_roster(data, path):
    roster = build_roster(data)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["reg_no", "candidate_name", "num_units", "units"])
        for reg, info in roster.items():
            w.writerow([reg, info["name"], len(info["units"]), "; ".join(info["units"])])


def print_summary(data):
    print(f"Centre : {data['centre_name']} ({data['centre_code']})")
    print(f"Course : {data['course_name']} — {data['course_level']}")
    print(f"Units  : {data['unit_count']}")
    roster = build_roster(data)
    print(f"Unique candidates across all units : {len(roster)}\n")
    print(f"{'UNIT CODE':<26} {'#':>4}  UNIT NAME")
    print("-" * 78)
    total = 0
    for u in data["units"]:
        total += u["candidate_count"]
        rt = "" if (u["report_type"] or "").lower().startswith("assessment") else " [RE-ASSESS]"
        print(f"{u['unit_code']:<26} {u['candidate_count']:>4}  {u['unit_name']}{rt}")
    print("-" * 78)
    print(f"{'TOTAL registration rows':<26} {total:>4}")


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def main(argv=None):
    p = argparse.ArgumentParser(description="Extract TVET CDACC register data from a PDF.")
    p.add_argument("pdf", help="path to the register PDF")
    p.add_argument("--json", metavar="FILE", help="write full structured data to JSON")
    p.add_argument("--csv", metavar="FILE", help="write flat per-candidate CSV")
    p.add_argument("--roster", metavar="FILE", help="write unique-candidate roster CSV")
    p.add_argument("--quiet", action="store_true", help="suppress the summary table")
    args = p.parse_args(argv)

    data = extract(args.pdf)

    if args.json:
        write_json(data, args.json)
        print(f"Wrote JSON   -> {args.json}", file=sys.stderr)
    if args.csv:
        write_csv(data, args.csv)
        print(f"Wrote CSV    -> {args.csv}", file=sys.stderr)
    if args.roster:
        write_roster(data, args.roster)
        print(f"Wrote roster -> {args.roster}", file=sys.stderr)

    if not args.quiet:
        print_summary(data)

    return data


if __name__ == "__main__":
    main()
