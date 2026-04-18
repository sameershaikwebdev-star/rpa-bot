"""
RPA Bot - Main Runner
=====================
Orchestrates: Load CSV/Excel → Launch Chrome → Fill forms → Generate reports

Supports two CSV formats, auto-detected by column names:
  • Type A — Student data  : name, id, marks, percentage
  • Type B — Form data     : custname, custtel, custemail, size, toppings, delivery, comments

Grading system (student CSV):
  marks >= 90 → S (large)    marks >= 80 → A (large)
  marks >= 70 → B (medium)   marks >= 60 → C (medium)
  marks >= 50 → D (small)    marks >= 40 → E (small)
  marks <  40 → F (small)    → result = FAIL

Field validation:
  Any row with a missing/invalid required field is marked FAILED and skipped.
  A validation_failures CSV is ALWAYS written when any row fails (not just all-fail).

Usage:
  python main.py --file sample_data/sample_input.csv
  python main.py --file sample_data/student.csv
  python main.py --file sample_data/sample_input.csv --headless
  python main.py --file sample_data/sample_input.csv --dry-run
  python main.py --file sample_data/sample_input.csv --config config.json
"""

import re
import sys
import json
import logging
import argparse
from pathlib import Path
from datetime import datetime

import pandas as pd

from core.data_loader import DataLoader
from core.bot_engine import RPABotEngine
from core.report_generator import ReportGenerator

Path("reports").mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(
            f"reports/bot_{datetime.now():%Y%m%d}.log", mode="a", encoding="utf-8"
        ),
    ],
)
logger = logging.getLogger(__name__)

# ─── Config ───────────────────────────────────────────────────────────────────

DEFAULT_CONFIG = {
    "form_url":               "https://httpbin.org/forms/post",
    "timeout":                15,
    "page_load_delay":        1.5,
    "submit_delay":           2.0,
    "delay_between_records":  1.0,
    "submit_selector": {"by": "xpath", "value": "//button[contains(text(),'Submit')]"},
}

# ─── Field Maps ───────────────────────────────────────────────────────────────

FIELD_MAP_FORM = [
    {"column": "custname",  "selector": {"by": "name", "value": "custname"},  "type": "text"},
    {"column": "custtel",   "selector": {"by": "name", "value": "custtel"},   "type": "text"},
    {"column": "custemail", "selector": {"by": "name", "value": "custemail"}, "type": "text"},
    {"column": "size",      "radio_name": "size",                             "type": "radio"},
    {"column": "toppings",  "checkbox_name": "topping",                       "type": "checkbox"},
    {"column": "delivery",  "selector": {"by": "name", "value": "delivery"},  "type": "time"},
    {"column": "comments",  "selector": {"by": "name", "value": "comments"},  "type": "text"},
]

# Student CSV → httpbin form mapping:
#   name       → custname
#   id         → custtel
#   custemail  → custemail  (generated as name@example.com)
#   size       → size radio (derived from grade)
#   comments   → comments   ("Grade: B | Result: pass | Marks: 78")
FIELD_MAP_STUDENT = [
    {"column": "name",      "selector": {"by": "name", "value": "custname"},  "type": "text"},
    {"column": "id",        "selector": {"by": "name", "value": "custtel"},   "type": "text"},
    {"column": "custemail", "selector": {"by": "name", "value": "custemail"}, "type": "text"},
    {"column": "size",      "radio_name": "size",                             "type": "radio"},
    {"column": "comments",  "selector": {"by": "name", "value": "comments"},  "type": "text"},
]

# ─── Internal keys stripped from reports ─────────────────────────────────────

# These are added by detect_and_prepare() for tracking but must NOT appear
# as columns in the HTML/Excel report.
_INTERNAL_KEYS = {"_validation_errors", "_row_index"}

# ─── Valid values ─────────────────────────────────────────────────────────────

VALID_SIZES = {"small", "medium", "large"}

SIZE_NORMALISE = {
    "extra_large": "large", "extralarge": "large",
    "xl":          "large", "x-large":   "large", "lg": "large",
    "med":         "medium", "md":        "medium",
    "sm":          "small",
}

# ─── Grading system ───────────────────────────────────────────────────────────

def get_grade(marks: int) -> str:
    if marks >= 90: return "S"
    if marks >= 80: return "A"
    if marks >= 70: return "B"
    if marks >= 60: return "C"
    if marks >= 50: return "D"
    if marks >= 40: return "E"
    return "F"

def grade_to_size(grade: str) -> str:
    if grade in ("S", "A"): return "large"
    if grade in ("B", "C"): return "medium"
    return "small"   # D, E, F

# ─── Field validators ─────────────────────────────────────────────────────────

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _validate_form_row(row: dict) -> list:
    errors = []
    if not str(row.get("custname", "")).strip():
        errors.append("custname is empty")
    if not str(row.get("custtel", "")).strip():
        errors.append("custtel is empty")
    email = str(row.get("custemail", "")).strip()
    if not email:
        errors.append("custemail is empty")
    elif not _EMAIL_RE.match(email):
        errors.append(f"custemail '{email}' is not a valid email address")
    size = str(row.get("size", "")).strip().lower()
    if not size:
        errors.append("size is empty")
    elif size not in VALID_SIZES:
        errors.append(f"size '{size}' is invalid — must be one of {sorted(VALID_SIZES)}")
    if not str(row.get("delivery", "")).strip():
        errors.append("delivery time is empty")
    return errors

def _validate_student_row(row: dict) -> list:
    errors = []
    if not str(row.get("name", "")).strip():
        errors.append("name is empty")
    if not str(row.get("id", "")).strip():
        errors.append("id is empty")
    try:
        m = int(str(row.get("marks", "")).strip())
        if not (0 <= m <= 100):
            errors.append(f"marks '{m}' out of valid range 0–100")
    except (ValueError, TypeError):
        errors.append(f"marks '{row.get('marks')}' is not a valid integer")
    return errors

# ─── Size normaliser ──────────────────────────────────────────────────────────

def _normalise_size(val: str) -> str:
    v = str(val).strip().lower()
    return SIZE_NORMALISE.get(v, v)

# ─── CSV Format Detection + Preparation ──────────────────────────────────────

def detect_and_prepare(df: pd.DataFrame):
    """
    Auto-detect CSV type, apply transforms and field validation.
    Returns: (data_type, valid_records, invalid_records, field_map)
    """
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    cols = set(df.columns)

    # ── Type A: Student CSV ───────────────────────────────────────────────────
    if {"name", "marks"}.issubset(cols):
        logger.info("📊 Detected CSV type: STUDENT")

        def _enrich(row):
            try:
                m = int(str(row["marks"]).strip())
            except (ValueError, TypeError):
                m = -1
            grade  = get_grade(m)
            result = "pass" if m >= 40 else "fail"
            size   = grade_to_size(grade)
            email  = (
                re.sub(r"[^a-z0-9.]", "",
                       str(row["name"]).strip().lower().replace(" ", "."))
                + "@example.com"
            )
            comments = f"Grade: {grade} | Result: {result} | Marks: {m}"
            return pd.Series({
                "grade": grade, "result": result,
                "size": size, "custemail": email, "comments": comments,
            })

        enriched = df.apply(_enrich, axis=1)
        df = pd.concat([df, enriched], axis=1)

        if "percentage" in df.columns:
            df["percentage"] = df["percentage"].astype(str).str.strip()

        valid_rows, invalid_records = [], []
        for idx, row in df.iterrows():
            errs = _validate_student_row(row.to_dict())
            if errs:
                rec = row.to_dict()
                rec["_validation_errors"] = " | ".join(errs)
                rec["_row_index"] = idx + 2
                invalid_records.append(rec)
                logger.warning(f"⚠️  Row {idx+2} INVALID: {errs}")
            else:
                valid_rows.append(row.to_dict())

        logger.info(
            f"   Valid: {len(valid_rows)}  Invalid: {len(invalid_records)}  "
            f"pass={(sum(1 for r in valid_rows if r.get('result')=='pass'))}  "
            f"fail={(sum(1 for r in valid_rows if r.get('result')=='fail'))}"
        )
        return "student", valid_rows, invalid_records, FIELD_MAP_STUDENT

    # ── Type B: Form CSV ──────────────────────────────────────────────────────
    elif {"custname", "custemail"}.issubset(cols):
        logger.info("📋 Detected CSV type: FORM")

        if "size" in df.columns:
            original = df["size"].tolist()
            df["size"] = df["size"].apply(_normalise_size)
            changed = [
                f"'{o}' → '{n}'"
                for o, n in zip(original, df["size"])
                if str(o).strip().lower() != n
            ]
            if changed:
                logger.info(f"   Size normalised: {changed}")

        valid_rows, invalid_records = [], []
        for idx, row in df.iterrows():
            errs = _validate_form_row(row.to_dict())
            if errs:
                rec = row.to_dict()
                rec["_validation_errors"] = " | ".join(errs)
                rec["_row_index"] = idx + 2
                invalid_records.append(rec)
                logger.warning(f"⚠️  Row {idx+2} INVALID: {errs}")
            else:
                valid_rows.append(row.to_dict())

        logger.info(f"   Valid: {len(valid_rows)}  Invalid: {len(invalid_records)}")
        return "form", valid_rows, invalid_records, FIELD_MAP_FORM

    # ── Unknown ───────────────────────────────────────────────────────────────
    else:
        raise ValueError(
            f"Unrecognised CSV format. Got columns: {sorted(cols)}.\n"
            "  Expected {{name, marks, ...}} for student data\n"
            "  or {{custname, custemail, ...}} for form data."
        )

# ─── Helpers ─────────────────────────────────────────────────────────────────

def _stats(results: list) -> dict:
    total   = len(results)
    success = sum(1 for r in results if r["status"] == "SUCCESS")
    return {"total": total, "success": success, "failed": total - success}

def _strip_internal_keys(records: list) -> list:
    """Remove _validation_errors and _row_index from row_data before reporting."""
    cleaned = []
    for rec in records:
        r = dict(rec)
        r["row_data"] = {k: v for k, v in r.get("row_data", {}).items()
                         if k not in _INTERNAL_KEYS}
        cleaned.append(r)
    return cleaned

def _write_validation_report(invalid_records: list):
    """
    FIX: Always write validation_failures CSV when ANY rows fail validation
    (not just when ALL rows fail).
    """
    if not invalid_records:
        return
    path = Path("reports") / f"validation_failures_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    rows = [
        {k: v for k, v in r.items() if k not in _INTERNAL_KEYS - {"_validation_errors"}}
        for r in invalid_records
    ]
    pd.DataFrame(rows).to_csv(path, index=False)
    logger.info(f"📄 Validation failure report → {path}")

# ─── Pipeline ─────────────────────────────────────────────────────────────────

def run_pipeline(input_file, config=None, field_map=None, headless=False, dry_run=False):
    merged_config = dict(DEFAULT_CONFIG)
    if config:
        merged_config.update(config)

    p = Path(input_file)
    if not p.exists():
        raise FileNotFoundError(f"Input file not found: {p.resolve()}")
    if p.suffix.lower() not in {".xlsx", ".xls", ".csv"}:
        raise ValueError(f"Unsupported file type '{p.suffix}'")

    logger.info("=" * 60)
    logger.info("🤖  RPA BOT  —  Pipeline Starting")
    logger.info("=" * 60)
    logger.info(f"📁 Input      : {p.resolve()}")
    logger.info(f"🌐 Target URL : {merged_config['form_url']}")
    logger.info(f"🖱️  Submit sel : {merged_config['submit_selector']}")

    loader = DataLoader(input_file)
    loader.load()
    loader.clean()

    if not hasattr(loader, "df") or loader.df is None or loader.df.empty:
        logger.warning("⚠️  No records found after loading. Exiting.")
        return {"results": [], "stats": {"total": 0, "success": 0, "failed": 0}}

    logger.info(f"📋 Data summary: {loader.summary()}")

    invalid_records = []

    if field_map is None:
        data_type, records, invalid_records, field_map = detect_and_prepare(loader.df)
        logger.info(
            f"✅ Format: {data_type.upper()} | "
            f"{len(records)} valid | {len(invalid_records)} invalid"
        )
    else:
        records = loader.to_records()
        logger.info("ℹ️  Using caller-supplied field_map — skipping auto-detection")

    # FIX: always write validation CSV when there are any invalid rows
    if invalid_records:
        _write_validation_report(invalid_records)
        logger.warning(f"🚫 {len(invalid_records)} row(s) failed validation and will be skipped:")
        for r in invalid_records:
            logger.warning(
                f"   Row {r.get('_row_index', '?')}: {r.get('_validation_errors', '')}"
            )

    # Build pre-failed results for report (strip internal keys first)
    pre_failed = [
        {
            "row_data":  {k: v for k, v in r.items() if k not in _INTERNAL_KEYS},
            "status":    "FAILED",
            "error":     f"Validation: {r.get('_validation_errors', 'unknown')}",
            "timestamp": datetime.now().isoformat(),
        }
        for r in invalid_records
    ]

    if not records:
        logger.warning("⚠️  No valid records to process.")
        return {"results": pre_failed, "stats": _stats(pre_failed)}

    if dry_run:
        logger.info("🔍 DRY RUN — data valid, skipping browser")
        for i, r in enumerate(records, 1):
            logger.info(f"  Record {i}: {r}")
        if invalid_records:
            logger.warning(f"  {len(invalid_records)} record(s) would be rejected:")
            for r in invalid_records:
                logger.warning(f"    → Row {r.get('_row_index')}: {r.get('_validation_errors')}")
        return {"records": records, "invalid": invalid_records, "dry_run": True}

    bot = RPABotEngine(merged_config, headless=headless)
    try:
        bot.start_driver()
        bot_results = bot.process_batch(records, field_map)
    finally:
        bot.stop_driver()

    # Merge bot results + pre-failed; strip internal keys before reporting
    all_results = _strip_internal_keys(bot_results) + pre_failed

    reporter = ReportGenerator(all_results)
    reports  = reporter.generate_all()

    stats = _stats(all_results)
    logger.info("=" * 60)
    logger.info(f"✅ Done! {stats}")
    logger.info(f"📄 HTML  → {reports['html']}")
    logger.info(f"📊 Excel → {reports['excel']}")
    logger.info("=" * 60)

    return {
        "results":      all_results,
        "report_html":  reports["html"],
        "report_excel": reports["excel"],
        "stats":        stats,
    }

# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="RPA Bot — CSV/Excel → Web Form Automation")
    ap.add_argument("--file",     required=True,       help="Path to input CSV or Excel file")
    ap.add_argument("--config",   default=None,        help="Path to config.json (optional)")
    ap.add_argument("--headless", action="store_true", help="Run Chrome in headless mode")
    ap.add_argument("--dry-run",  action="store_true", help="Validate data without opening browser")
    args = ap.parse_args()

    config = None
    if args.config:
        cp = Path(args.config)
        if not cp.exists():
            logger.error(f"❌ Config file not found: {cp.resolve()}")
            sys.exit(1)
        try:
            config = json.loads(cp.read_text(encoding="utf-8"))
            logger.info(f"⚙️  Loaded config: {cp.resolve()}")
        except json.JSONDecodeError as e:
            logger.error(f"❌ Invalid JSON in config file: {e}")
            sys.exit(1)

    try:
        run_pipeline(
            args.file,
            config=config,
            headless=args.headless,
            dry_run=args.dry_run,
        )
    except (FileNotFoundError, ValueError) as e:
        logger.error(f"❌ {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()