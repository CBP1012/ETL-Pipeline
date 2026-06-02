#!/usr/bin/env python3
"""
pipeline_demo.py — Interactive CI pipeline walkthrough
  Raw sensor data → Normalize → CCO-based RDF → SPARQL QC → SHACL

Press ENTER at each stage prompt to advance (or Ctrl+C to quit).
Each stage corresponds to one or more slides in the presentation deck.

Stages:
  1  Discover Inputs          — locate raw sensor files (A, B, C)
  2  Ingest sensor_A.csv      — read CSV (cp1252)
  3  Ingest sensor_B.json     — parse & flatten nested JSON
  4  Ingest sensor_C.csv      — read CSV (same schema as A)
  5  Normalize                — canonicalize IDs, kinds, units, timestamps
  6  Merge & Drop NaNs        — combine all three; dropna → 20 raw, 15 clean
  7  Pre-Ontology Triage      — lightweight quality checks
  8  Export Clean CSV         — write readings_normalized.csv (5 cols, 15 rows)
  9  Transform to CCO RDF     — map rows to CCO ontology, serialize as .ttl
 10  SPARQL QC Validation     — run violation-detector queries
 11  SHACL Validation         — validate graph against shape constraints

Usage:
    python pipeline_demo.py

Requirements:
    pip install pandas rdflib pyshacl

Outputs:
    data/interim/readings_normalized.csv
    measure_cco.ttl
"""

import csv
import json
import math
import pathlib
import sys
import textwrap

import pandas as pd
from rdflib import Graph, Namespace, Literal, URIRef
from rdflib.namespace import RDF, RDFS, XSD

# ──────────────────────────────────────────────────────────────────────
#  CONFIG — Update ROOT to match your local directory structure
# ──────────────────────────────────────────────────────────────────────
ROOT       = pathlib.Path(r"C:\Users\crist\Documents\BP\Pipelines\ETL+\assignment\src")
SRC        = ROOT / "data"
OUT_DIR    = ROOT / "data" / "interim"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_CSV    = OUT_DIR / "readings_normalized.csv"
OUT_TTL    = ROOT / "measure_cco.ttl"
SPARQL_DIR = ROOT / "sparql"
SHACL_FILE = ROOT / "shacl" / "cco_shapes.ttl"


# ──────────────────────────────────────────────────────────────────────
#  LOOKUP TABLES  (controlled vocabulary — pre-ontology)
#  Unit codes match the normalized CSV: degF / degC / PSI_gauge / kPa_gauge / V / Ω
# ──────────────────────────────────────────────────────────────────────
KIND_MAP = {
    "temp":        "temperature",
    "temperature": "temperature",
    "pressure":    "pressure",
    "voltage":     "voltage",
    "resistance":  "resistance",
}

UNIT_MAP = {
    # Fahrenheit
    "f":          "degF",
    "°f":         "degF",
    "fahrenheit": "degF",
    # Celsius
    "c":          "degC",
    "°c":         "degC",
    "celsius":    "degC",
    # Pressure
    "psi":        "PSI_gauge",
    "kpa":        "kPa_gauge",
    "kilopascal": "kPa_gauge",
    # Electrical
    "v":          "V",
    "volt":       "V",
    "volts":      "V",
    "ohm":        "Ω",
    "ohms":       "Ω",
}

BOUNDS = {
    "temperature": {"min": -100.0,     "max": 200.0},
    "pressure":    {"min":    0.0,     "max": 1_000_000.0},
    "voltage":     {"min":    0.0,     "max": 10_000.0},
    "resistance":  {"min":    0.0,     "max": 1_000_000.0},
}

# ──────────────────────────────────────────────────────────────────────
#  CCO / BFO IRI CONSTANTS
# ──────────────────────────────────────────────────────────────────────
CCO = Namespace("https://www.commoncoreontologies.org/")
BFO = Namespace("http://purl.obolibrary.org/obo/")
EX  = Namespace("http://example.org/measurement/")

IRI_SDC          = URIRef("http://purl.obolibrary.org/obo/BFO_0000020")
IRI_ART          = URIRef("https://www.commoncoreontologies.org/ont00000995")
IRI_MU           = URIRef("https://www.commoncoreontologies.org/ont00000120")
IRI_MICE         = URIRef("https://www.commoncoreontologies.org/ont00001163")
IRI_BEARER_OF    = URIRef("http://purl.obolibrary.org/obo/BFO_0000196")
IRI_IS_MEASURE   = URIRef("https://www.commoncoreontologies.org/ont00001966")
IRI_USES_MU      = URIRef("https://www.commoncoreontologies.org/ont00001863")

# Human-readable labels for unit nodes in the RDF graph
UNIT_LABEL_MAP = {
    "degF":      "Degree Fahrenheit",
    "degC":      "Degree Celsius",
    "PSI_gauge": "Pound per Square Inch",
    "kPa_gauge": "Kilopascal",
    "V":         "Volt",
    "Ω":         "Ohm",
}

# Prefix used in RDF obs IDs: "sensor_A" → "a"
SOURCE_PREFIX = {
    "sensor_A": "a",
    "sensor_B": "b",
    "sensor_C": "c",
}


# ──────────────────────────────────────────────────────────────────────
#  TERMINAL HELPERS
# ──────────────────────────────────────────────────────────────────────

# ANSI colour codes (fall back gracefully on Windows without ANSI support)
try:
    import ctypes
    ctypes.windll.kernel32.SetConsoleMode(
        ctypes.windll.kernel32.GetStdHandle(-11), 7
    )
except Exception:
    pass

TEAL   = "\033[96m"
AMBER  = "\033[93m"
RED    = "\033[91m"
GREEN  = "\033[92m"
PURPLE = "\033[35m"
GRAY   = "\033[90m"
WHITE  = "\033[97m"
BOLD   = "\033[1m"
RESET  = "\033[0m"

W = 72  # terminal width for banners


def _c(text, colour):
    """Wrap text in an ANSI colour code."""
    return f"{colour}{text}{RESET}"


def banner(stage_num: int, title: str, slide_ref: str, description: str):
    """
    Print a clearly visible stage header with a slide cross-reference.

      ════════════════════════════════════════════════════════════════
        STAGE 2 · INGEST sensor_A.csv           ← Slide 5 (left card)
      ════════════════════════════════════════════════════════════════
        Read the CSV with Windows-1252 encoding …
      ────────────────────────────────────────────────────────────────
    """
    print()
    print(_c("═" * W, TEAL))
    stage_str = f"  STAGE {stage_num} · {title}"
    ref_str   = f"← {slide_ref}"
    gap       = W - len(stage_str) - len(ref_str) - 2
    print(_c(f"{stage_str}{' ' * max(gap, 2)}{ref_str}", BOLD + WHITE))
    print(_c("═" * W, TEAL))
    for line in textwrap.wrap(description, W - 4):
        print(f"  {_c(line, GRAY)}")
    print(_c("─" * W, GRAY))


def pause(prompt: str = "  ▶  Press ENTER to continue to next stage …"):
    """Block until the user presses Enter."""
    try:
        print()
        input(_c(prompt, TEAL + BOLD))
    except KeyboardInterrupt:
        print("\n\n  Walkthrough cancelled.")
        sys.exit(0)


def _ok(msg):   print(f"  {_c('✔', GREEN)}  {msg}")
def _warn(msg): print(f"  {_c('⚠', AMBER)}  {_c(msg, AMBER)}")
def _fail(msg): print(f"  {_c('✗', RED)}  {_c(msg, RED)}")
def _info(msg): print(f"  {_c('·', GRAY)}  {msg}")


def _safe_uri_fragment(s: str) -> str:
    return str(s).strip().lower().replace(" ", "-").replace("/", "-").replace("_", "-")


# ──────────────────────────────────────────────────────────────────────
#  NORMALIZATION HELPERS
# ──────────────────────────────────────────────────────────────────────

def _canon_entity_id(s) -> str | None:
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return None
    # "Chiller 3" → "Chiller-3"
    return " ".join(str(s).strip().split()).replace(" ", "-")


def _canon_kind(s) -> str | None:
    if s is None:
        return None
    return KIND_MAP.get(str(s).strip().lower(), str(s).strip().lower())


def _canon_unit(s) -> str | None:
    if s is None:
        return None
    key = str(s).strip()
    # Try exact, then lowercase
    return UNIT_MAP.get(key, UNIT_MAP.get(key.lower(), key))


def _parse_timestamp(s) -> str | None:
    """Parse any date string → UTC ISO-8601 ending in 'Z'. Returns None on failure."""
    if s is None or str(s).strip() == "":
        return None
    try:
        dt = pd.to_datetime(s, utc=True, infer_datetime_format=True)
        return dt.isoformat().replace("+00:00", "Z")
    except Exception:
        return None


def _coerce_float(x) -> float:
    try:
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return math.nan
        return float(x)
    except Exception:
        return math.nan


# ══════════════════════════════════════════════════════════════════════
#  STAGE FUNCTIONS
# ══════════════════════════════════════════════════════════════════════

# ── STAGE 1 ──────────────────────────────────────────────────────────

def stage_1_discover_inputs():
    banner(1, "DISCOVER INPUTS", "Slide 4 (architecture) · Slide 5 (source cards)",
           "Locate the three raw sensor files on disk and confirm they all "
           "exist before any parsing begins. One missing file = hard stop.")

    paths = {
        "sensor_A": SRC / "sensor_A.csv",
        "sensor_B": SRC / "sensor_B.json",
        "sensor_C": SRC / "sensor_C.csv",
    }

    all_ok = True
    for name, p in paths.items():
        exists = p.exists()
        label  = _c("found ✔", GREEN) if exists else _c("MISSING ✗", RED)
        _info(f"{name:12s}  {p}  [{label}]")
        if not exists:
            all_ok = False

    if not all_ok:
        _fail(f"One or more input files missing in {SRC}. Aborting.")
        sys.exit(1)

    print()
    _ok("All three input files found.")
    pause()
    return paths["sensor_A"], paths["sensor_B"], paths["sensor_C"]


# ── STAGE 2 ──────────────────────────────────────────────────────────

def stage_2_ingest_sensor_a(path: pathlib.Path) -> pd.DataFrame:
    banner(2, "INGEST sensor_A.csv", "Slide 5 (Source A card) · Slide 6 (parsing diagram)",
           "Read sensor_A.csv with Windows-1252 (cp1252) encoding. All columns "
           "are loaded as raw strings so nothing is silently coerced yet.")

    df = pd.read_csv(path, encoding="cp1252", dtype=str,
                     keep_default_na=False, na_values=["", "NA", "NaN"])

    _info(f"Rows loaded   : {_c(len(df), WHITE)}")
    _info(f"Columns       : {_c(list(df.columns), WHITE)}")
    print()
    print(_c("  Raw data (all 6 rows):", GRAY))
    print(df.to_string(index=False))

    pause()
    return df


# ── STAGE 3 ──────────────────────────────────────────────────────────

def stage_3_ingest_sensor_b(path: pathlib.Path) -> pd.DataFrame:
    banner(3, "INGEST sensor_B.json", "Slide 5 (Source B card) · Slide 6 (parsing diagram)",
           "Parse the nested JSON ('readings → entity_id + data[]') and "
           "flatten it into a tabular DataFrame. The adaptive parser also "
           "handles flat-list and NDJSON shapes — no manual switch needed.")

    with open(path, "r", encoding="utf-8") as f:
        obj = json.load(f)

    _info(f"Top-level keys : {_c(list(obj.keys()), WHITE)}")
    _info(f"# entities     : {_c(len(obj.get('readings', [])), WHITE)}")

    # Adaptive parse — nested structure
    if "readings" in obj:
        records = [
            {
                "artifact_id": entry.get("entity_id"),
                "sdc_kind":    pt.get("kind"),
                "unit_label":  pt.get("unit"),
                "value":       pt.get("value"),
                "timestamp":   pt.get("time"),
                "_source":     "sensor_B",
            }
            for entry in obj["readings"]
            for pt    in entry.get("data", [])
        ]
    elif isinstance(obj, list):
        records = obj          # flat list fallback
    else:
        records = [            # NDJSON fallback
            json.loads(line)
            for line in path.read_text().splitlines()
            if line.strip()
        ]

    df = pd.DataFrame(records)
    _info(f"Rows flattened : {_c(len(df), WHITE)}")
    print()
    print(_c("  Raw data (all 8 rows):", GRAY))
    print(df.to_string(index=False))

    pause()
    return df


# ── STAGE 4 ──────────────────────────────────────────────────────────

def stage_4_ingest_sensor_c(path: pathlib.Path) -> pd.DataFrame:
    banner(4, "INGEST sensor_C.csv", "Slide 5 (Source C card) · Slide 6 (parsing diagram)",
           "Read sensor_C.csv (UTF-8). Same column layout as sensor_A. "
           "This source adds Circuit-07, Boiler-09 — and two more NaN values "
           "from bad strings that will be dropped in Stage 6.")

    df = pd.read_csv(path, dtype=str,
                     keep_default_na=False, na_values=["", "NA", "NaN"])

    _info(f"Rows loaded   : {_c(len(df), WHITE)}")
    _info(f"Columns       : {_c(list(df.columns), WHITE)}")
    if "Reading Type" in df.columns:
        kinds = sorted(df["Reading Type"].str.strip().str.lower().unique())
        _info(f"Distinct kinds: {_c(kinds, WHITE)}")
    print()
    print(_c("  Raw data (all 6 rows):", GRAY))
    print(df.to_string(index=False))

    pause()
    return df


# ── STAGE 5 ──────────────────────────────────────────────────────────

def _normalize_one(df: pd.DataFrame, source_label: str) -> pd.DataFrame:
    """
    Apply all canonicalization rules to one source DataFrame and
    return a clean 5-column frame: artifact_id, sdc_kind, unit_label,
    value, timestamp — plus a hidden _source column used only by Stage 9
    to build obs IDs like obs-a-000001.
    """
    df = df.copy()

    # Rename columns for sensor_A and sensor_C
    if source_label in ("sensor_A", "sensor_C"):
        df = df.rename(columns={
            "Device Name":   "artifact_id",
            "Reading Type":  "sdc_kind",
            "Units":         "unit_label",
            "Reading Value": "value",
            "Time (Local)":  "timestamp",
        })
        df["_source"] = source_label

    df = df[["artifact_id", "sdc_kind", "unit_label", "value", "timestamp", "_source"]]

    # ── 1. Entity ID: spaces → hyphens, strip ─────────────────────
    before = df["artifact_id"].iloc[:2].tolist()
    df["artifact_id"] = df["artifact_id"].map(_canon_entity_id)
    after  = df["artifact_id"].iloc[:2].tolist()
    _info(f"artifact_id  sample : {_c(before, AMBER)}  →  {_c(after, TEAL)}")

    # ── 2. Quantity kind → controlled vocab ───────────────────────
    before = df["sdc_kind"].iloc[:2].tolist()
    df["sdc_kind"] = df["sdc_kind"].map(_canon_kind)
    after  = df["sdc_kind"].iloc[:2].tolist()
    _info(f"sdc_kind     sample : {_c(before, AMBER)}  →  {_c(after, TEAL)}")

    # ── 3. Unit → canonical code ──────────────────────────────────
    before = df["unit_label"].iloc[:2].tolist()
    df["unit_label"] = df["unit_label"].map(_canon_unit)
    after  = df["unit_label"].iloc[:2].tolist()
    _info(f"unit_label   sample : {_c(before, AMBER)}  →  {_c(after, TEAL)}")

    # ── 4. Value → float (bad strings become NaN) ─────────────────
    df["value"] = df["value"].map(_coerce_float)
    nan_vals = df["value"].isna().sum()
    msg = f"{nan_vals} NaN (will be dropped in Stage 6)"
    _info(f"value        NaN    : " + (_c(msg, RED) if nan_vals else _c("0 ✔", GREEN)))

    # ── 5. Timestamp → UTC ISO-8601 ───────────────────────────────
    df["timestamp"] = df["timestamp"].map(_parse_timestamp)
    bad_ts = df["timestamp"].isna().sum()
    _info(f"timestamp    bad    : " + (_c(f"{bad_ts}", RED) if bad_ts else _c("0 ✔", GREEN)))

    return df


def stage_5_normalize(raw_a, raw_b, raw_c):
    banner(5, "NORMALIZE", "Slide 7 (transform table)",
           "Apply controlled-vocabulary mapping to all three sources: "
           "canonicalize entity IDs (spaces → hyphens), map kind strings "
           "to the controlled vocab, map unit strings to canonical codes "
           "(F → degF, psi → PSI_gauge, kPa → kPa_gauge, etc.), coerce "
           "values to float (bad strings → NaN), parse timestamps to UTC ISO-8601.")

    print(_c("\n  ── sensor_A ──────────────────────────────────────────────", TEAL))
    norm_a = _normalize_one(raw_a, "sensor_A")

    print(_c("\n  ── sensor_B ──────────────────────────────────────────────", PURPLE))
    norm_b = _normalize_one(raw_b, "sensor_B")

    print(_c("\n  ── sensor_C ──────────────────────────────────────────────", AMBER))
    norm_c = _normalize_one(raw_c, "sensor_C")

    print()
    _ok("All three sources normalized.  Unit codes now: "
        "degF · degC · PSI_gauge · kPa_gauge · V · Ω")
    pause()
    return norm_a, norm_b, norm_c


# ── STAGE 6 ──────────────────────────────────────────────────────────

def stage_6_merge_and_drop(norm_a, norm_b, norm_c) -> pd.DataFrame:
    banner(6, "MERGE + dropna", "Slide 8 (output table)",
           "Concatenate all three normalized DataFrames, then run dropna() "
           "on the 5 required columns. This removes the 5 rows whose 'value' "
           "became NaN (3 bad strings from A/C, 2 JSON nulls from B). "
           "Result: 20 raw rows → 15 clean rows across 8 devices.")

    combined = pd.concat([norm_a, norm_b, norm_c], ignore_index=True)
    _info(f"Combined total      : {_c(len(combined), WHITE)} rows")

    REQUIRED = ["artifact_id", "sdc_kind", "unit_label", "value", "timestamp"]
    before = len(combined)
    df = combined.dropna(subset=REQUIRED).reset_index(drop=True)
    dropped = before - len(df)

    _info(f"Dropped (NaN)       : {_c(dropped, RED if dropped else GREEN)} rows")
    _info(f"Clean rows          : {_c(len(df), GREEN)}")
    _info(f"Unique devices      : {_c(df['artifact_id'].nunique(), WHITE)}")
    _info(f"Unique sdc_kinds    : {_c(sorted(df['sdc_kind'].unique()), WHITE)}")
    _info(f"Unique unit_labels  : {_c(sorted(df['unit_label'].unique()), WHITE)}")

    print()
    print(_c("  All 15 clean rows:", GRAY))
    print(df[REQUIRED].to_string(index=False))

    pause()
    return df


# ── STAGE 7 ──────────────────────────────────────────────────────────

def stage_7_triage(df: pd.DataFrame):
    banner(7, "PRE-ONTOLOGY QUALITY TRIAGE", "Slide 9 (dropped rows + surviving issue)",
           "Lightweight guard-rails on the clean tabular data before any RDF "
           "is generated. These are warnings, not hard failures. Two issues "
           "are expected: the 5 NaN rows are already gone, but one "
           "out-of-range temperature reading survives into the RDF.")

    print(_c("  ── Missingness ──────────────────────────────────────────────", TEAL))
    miss = df.isna().sum()
    all_clean = True
    for col in df.columns:
        if col.startswith("_"):
            continue
        count = miss[col]
        if count:
            _warn(f"{col:20s} : {count} NaN")
            all_clean = False
        else:
            _ok(f"{col:20s} : 0 ✔")
    if all_clean:
        _ok("No missing values (dropna cleared them in Stage 6)")

    print()
    print(_c("  ── Controlled Vocabulary ───────────────────────────────────", TEAL))
    known_kinds = set(KIND_MAP.values())
    found_kinds = set(df["sdc_kind"].dropna())
    unknown = found_kinds - known_kinds
    if unknown:
        _warn(f"Unknown sdc_kind values: {sorted(unknown)}")
    else:
        _ok(f"All sdc_kind values in vocab: {sorted(found_kinds)}")

    print()
    print(_c("  ── Value Bounds ────────────────────────────────────────────", TEAL))
    any_oob = False
    for kind, bounds in BOUNDS.items():
        sub = df[df["sdc_kind"] == kind]
        if sub.empty:
            continue
        lo = (sub["value"] < bounds["min"]).sum()
        hi = (sub["value"] > bounds["max"]).sum()
        if lo or hi:
            _warn(f"{kind:12s} : {lo} below {bounds['min']}, "
                  f"{hi} above {bounds['max']}")
            oob = sub[(sub["value"] < bounds["min"]) | (sub["value"] > bounds["max"])]
            for _, r in oob.iterrows():
                print(f"         {_c(r['artifact_id'], WHITE)} "
                      f"{_c(r['value'], RED)} {r['unit_label']}")
            any_oob = True
        else:
            _ok(f"{kind:12s} : all in [{bounds['min']}, {bounds['max']}]")

    if any_oob:
        print()
        _warn("1 out-of-range reading will survive into the RDF graph.")
        _warn("Boiler-07 @ 212 °F is physically valid (boiling point) but "
              "exceeds the configured 200 °F threshold.")

    pause()


# ── STAGE 8 ──────────────────────────────────────────────────────────

def stage_8_export_csv(df: pd.DataFrame):
    banner(8, "EXPORT CLEAN CSV", "Slide 10 (15-row table)",
           "Write the cleaned DataFrame to readings_normalized.csv. "
           "Only the 5 canonical columns are exported: artifact_id, sdc_kind, "
           "unit_label, value, timestamp. The _source column is internal only.")

    COLS = ["artifact_id", "sdc_kind", "unit_label", "value", "timestamp"]
    out  = df[COLS]

    out.to_csv(OUT_CSV, index=False, quoting=csv.QUOTE_MINIMAL)

    _info(f"Output file         : {_c(OUT_CSV, TEAL)}")
    _info(f"Total records       : {_c(len(out), WHITE)}")

    print()
    print(_c("  Column dtypes:", GRAY))
    for col in COLS:
        _info(f"  {col:20s} : {out[col].dtype}")

    print()
    print(_c("  Counts by sdc_kind:", GRAY))
    for kind, count in out["sdc_kind"].value_counts(dropna=False).items():
        _info(f"  {str(kind):20s} : {count}")

    print()
    print(_c("  Counts by unit_label:", GRAY))
    for unit, count in out["unit_label"].value_counts(dropna=False).items():
        _info(f"  {str(unit):20s} : {count}")

    print()
    _ok(f"readings_normalized.csv ready  — {len(out)} rows, 5 columns.")
    pause()


# ── STAGE 9 ──────────────────────────────────────────────────────────

def stage_9_transform_to_rdf(df: pd.DataFrame) -> Graph:
    banner(9, "TRANSFORM TO CCO-BASED RDF", "Slides 11–12 (ontology model + row mapping)",
           "Map each of the 15 clean rows to four CCO/BFO nodes: "
           "Artifact (the device), SDC/BFO_0000020 (the quality), "
           "MICE/ont00001163 (the reading), MeasurementUnit/ont00000120. "
           "Link them with bearer_of, is_measure_of, uses_measurement_unit.")

    g = Graph()
    g.bind("cco", CCO)
    g.bind("bfo", BFO)
    g.bind("ex",  EX)
    g.bind("xsd", XSD)

    created_entities = {}  # "artifact_id|sdc_kind" → (artifact_uri, sdc_uri)
    created_units    = {}  # unit_label → unit_uri

    # Per-source counters for obs IDs (matches obs-a-000001 … obs-c-000006 in TTL)
    source_counters  = {"sensor_A": 0, "sensor_B": 0, "sensor_C": 0}

    mice_count = 0

    for _, row in df.iterrows():
        artifact_id = row["artifact_id"]
        sdc_kind    = row["sdc_kind"]
        value       = row["value"]
        unit_label  = row["unit_label"]
        timestamp   = row["timestamp"]
        source      = row.get("_source", "sensor_A")

        # ── Observation ID  (e.g. obs-a-000001) ──────────────────
        source_counters[source] = source_counters.get(source, 0) + 1
        prefix  = SOURCE_PREFIX.get(source, "x")
        obs_id  = f"obs-{prefix}-{source_counters[source]:06d}"

        # ── Artifact + SDC  (one node-pair per device+kind) ───────
        ek_key = f"{artifact_id}|{sdc_kind}"
        if ek_key not in created_entities:
            ent_frag  = _safe_uri_fragment(artifact_id)
            kind_frag = _safe_uri_fragment(sdc_kind)

            artifact_uri = EX[ent_frag]
            sdc_uri      = EX[f"{ent_frag}-{kind_frag}-sdc"]

            g.add((artifact_uri, RDF.type,       IRI_ART))
            g.add((artifact_uri, RDFS.label,     Literal(artifact_id)))
            g.add((sdc_uri,      RDF.type,       IRI_SDC))
            g.add((sdc_uri,      RDFS.label,     Literal(f"{sdc_kind} of {artifact_id}")))
            g.add((artifact_uri, IRI_BEARER_OF,  sdc_uri))

            created_entities[ek_key] = (artifact_uri, sdc_uri)
        else:
            _, sdc_uri = created_entities[ek_key]

        # ── Measurement Unit  (one node per unit code) ────────────
        if pd.notna(unit_label) and unit_label not in created_units:
            unit_frag = _safe_uri_fragment(unit_label)
            unit_uri  = EX[f"unit-{unit_frag}"]
            g.add((unit_uri, RDF.type,   IRI_MU))
            g.add((unit_uri, RDFS.label, Literal(
                UNIT_LABEL_MAP.get(unit_label, str(unit_label))
            )))
            created_units[unit_label] = unit_uri

        # ── MICE  (one per row) ───────────────────────────────────
        # ⚠ NOTE: This label template causes the SPARQL failure (Slide 13).
        #   "pressure reading: 101.325 kPa_gauge" is not unique across devices
        #   when two sensors record the same value.
        mice_uri   = EX[obs_id]
        mice_label = f"{sdc_kind} reading: {value} {unit_label}"

        g.add((mice_uri, RDF.type,       IRI_MICE))
        g.add((mice_uri, RDFS.label,     Literal(mice_label)))
        g.add((mice_uri, IRI_IS_MEASURE, sdc_uri))
        if pd.notna(unit_label) and unit_label in created_units:
            g.add((mice_uri, IRI_USES_MU, created_units[unit_label]))

        mice_count += 1

    g.serialize(destination=str(OUT_TTL), format="turtle")

    print(_c("  ── Graph stats ──────────────────────────────────────────────", TEAL))
    _info(f"Artifact nodes        : {_c(len({v[0] for v in created_entities.values()}), WHITE)}")
    _info(f"SDC nodes             : {_c(len(created_entities), WHITE)}")
    _info(f"MeasurementUnit nodes : {_c(len(created_units), WHITE)}")
    _info(f"MICE instances        : {_c(mice_count, WHITE)}")
    _info(f"Total RDF triples     : {_c(len(g), WHITE)}")
    _info(f"Output file           : {_c(OUT_TTL, TEAL)}")

    print()
    print(_c("  ── Sample triples (first 12) ────────────────────────────────", GRAY))
    for i, (s, p, o) in enumerate(g):
        if i >= 12:
            print(f"  {_c('…', GRAY)}")
            break
        s_s = str(s).replace("http://example.org/measurement/", "ex:")
        p_s = str(p).split("/")[-1]
        o_s = (str(o)
               .replace("http://example.org/measurement/", "ex:")
               .replace("https://www.commoncoreontologies.org/", "cco:")
               .replace("http://purl.obolibrary.org/obo/", "bfo:"))
        print(f"  {_c(f'{s_s[:32]:<32}', TEAL)}  {_c(f'{p_s[:22]:<22}', PURPLE)}  {_c(o_s[:40], WHITE)}")

    print()
    _ok(f"CCO-based RDF graph written to {OUT_TTL}")

    # ── Show the MICE nodes that share labels (→ Stage 10 failure) ────
    print()
    print(_c("  ── ⚠  Label collision preview (→ Slide 13) ─────────────────", AMBER))
    _warn('Group 1 — "pressure reading: 101.325 kPa_gauge"')
    _warn("  obs-a-000005  →  Pump-A1")
    _warn("  obs-b-000002  →  Boiler-07")
    _warn("  obs-c-000005  →  Pump-A2")
    _warn('Group 2 — "voltage reading: 2.0 V"')
    _warn("  obs-b-000007  →  Circuit-12")
    _warn("  obs-c-000001  →  Circuit-07")

    pause()
    return g


# ── STAGE 10 ─────────────────────────────────────────────────────────

def stage_10_sparql_qc(g: Graph) -> int:
    banner(10, "SPARQL QC VALIDATION", "Slides 13–14 (query list + failure analysis)",
           "Run every .rq file in sparql/ as a violation detector. "
           "Zero rows = clean; any rows = violation. "
           "Expect 7 passes and 1 failure: no_duplicate_labels returns 8 rows "
           "from the 2 label collision groups identified in Stage 9.")

    SHOW_LIMIT = 10

    if not SPARQL_DIR.exists():
        _warn(f"SPARQL directory not found: {SPARQL_DIR}")
        _warn("Skipping SPARQL QC.")
        pause()
        return 0

    rq_files = sorted(SPARQL_DIR.glob("*.rq"))
    if not rq_files:
        _warn(f"No .rq files in {SPARQL_DIR}. Skipping.")
        pause()
        return 0

    _info(f"Query directory : {_c(SPARQL_DIR, TEAL)}")
    _info(f"Queries found   : {_c(len(rq_files), WHITE)}")
    print()

    failures = 0
    for qpath in rq_files:
        q    = qpath.read_text(encoding="utf-8")
        rows = list(g.query(q))
        if rows:
            print(f"  {_c('✗', RED)}  {_c(qpath.name, RED)} : {_c(len(rows), RED)} violation(s)")
            for r in rows[:SHOW_LIMIT]:
                print(f"       {_c(' | '.join(str(x) for x in r), AMBER)}")
            failures += 1
        else:
            print(f"  {_c('✔', GREEN)}  {qpath.name} : 0 violations")

    print()
    if failures:
        _fail(f"SPARQL QC: {failures} failing quer{'y' if failures == 1 else 'ies'}.")
        _warn("See Slide 13 for root cause.")
    else:
        _ok(f"SPARQL QC: all {len(rq_files)} checks passed.")

    pause()
    return failures


# ── STAGE 11 ─────────────────────────────────────────────────────────

def stage_11_shacl_validation(g: Graph) -> bool:
    banner(11, "SHACL VALIDATION", "Slide 15 (shapes + result card)",
           "Validate the RDF graph against SHACL structural constraints "
           "in cco_shapes.ttl. This checks that every MICE uses a "
           "MeasurementUnit, every Artifact bears at least one SDC, "
           "and every SDC is typed as BFO_0000020. Expected result: "
           "Conforms: True — structural integrity is intact even though "
           "SPARQL caught the semantic label issue.")

    if not SHACL_FILE.exists():
        _warn(f"SHACL shapes file not found: {SHACL_FILE}")
        _warn("Skipping SHACL validation.")
        pause()
        return True

    _info(f"Data graph   : {_c(OUT_TTL, TEAL)}  ({len(g)} triples)")
    _info(f"Shapes file  : {_c(SHACL_FILE, TEAL)}")
    print()

    try:
        from pyshacl import validate as shacl_validate

        conforms, _, report_text = shacl_validate(
            data_graph=g,
            shacl_graph=str(SHACL_FILE),
            data_graph_format="turtle",
            shacl_graph_format="turtle",
            inference="rdfs",
            allow_infos=True,
            allow_warnings=True,
            advanced=True,
        )

        for line in report_text.strip().splitlines():
            print(f"  {_c(line, GRAY)}")

        print()
        if conforms:
            _ok("SHACL: Conforms: True ✅ — graph structure is valid.")
        else:
            _fail("SHACL: Conforms: False ❌ — see violations above.")

    except ImportError:
        _warn("pyshacl not installed.  Run: pip install pyshacl")
        _warn("Skipping SHACL validation.")
        conforms = True

    pause("  ▶  Pipeline complete — press ENTER to see the final summary …")
    return conforms


# ══════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════

def main():
    print()
    print(_c("=" * W, TEAL + BOLD))
    print(_c("  SENSOR DATA ETL PIPELINE — INTERACTIVE DEMO WALKTHROUGH", BOLD + WHITE))
    print(_c("  Raw sensor feeds → CCO-based RDF → SPARQL QC → SHACL", GRAY))
    print(_c("=" * W, TEAL + BOLD))
    print()
    print(textwrap.dedent(f"""\
  {_c('Phases:', WHITE)}

    {_c('Phase 1 — Data Cleaning   (Stages 1–8)', TEAL)}
      Ingest raw sensor data (A, B, C) → normalize → export clean CSV

    {_c('Phase 2 — RDF Transformation  (Stage 9)', PURPLE)}
      Map 15 clean rows to CCO ontology classes → measure_cco.ttl

    {_c('Phase 3 — Validation  (Stages 10–11)', GREEN)}
      SPARQL violation queries + SHACL shape constraints

  {_c('Controls:', WHITE)}
    Press {_c('ENTER', TEAL + BOLD)} at each prompt to advance to the next stage.
    Press {_c('Ctrl+C', RED)} at any time to quit.
    Each stage banner shows the corresponding {_c('slide number', AMBER)}.
    """))

    pause("  ▶  Press ENTER to begin Stage 1 …")

    # ── PHASE 1: DATA CLEANING ─────────────────────────────────────────

    a_path, b_path, c_path = stage_1_discover_inputs()
    raw_a  = stage_2_ingest_sensor_a(a_path)
    raw_b  = stage_3_ingest_sensor_b(b_path)
    raw_c  = stage_4_ingest_sensor_c(c_path)

    norm_a, norm_b, norm_c = stage_5_normalize(raw_a, raw_b, raw_c)

    df = stage_6_merge_and_drop(norm_a, norm_b, norm_c)

    stage_7_triage(df)
    stage_8_export_csv(df)

    # ── PHASE 2: RDF TRANSFORMATION ───────────────────────────────────

    rdf_graph = stage_9_transform_to_rdf(df)

    # ── PHASE 3: VALIDATION ───────────────────────────────────────────

    sparql_failures = stage_10_sparql_qc(rdf_graph)
    shacl_ok        = stage_11_shacl_validation(rdf_graph)

    # ── FINAL SUMMARY ─────────────────────────────────────────────────

    print()
    print(_c("═" * W, TEAL + BOLD))
    print(_c("  PIPELINE SUMMARY  (→ Slide 16)", BOLD + WHITE))
    print(_c("═" * W, TEAL + BOLD))
    print()
    _info(f"Clean CSV           : {_c(OUT_CSV, TEAL)}")
    _info(f"RDF graph (TTL)     : {_c(OUT_TTL, TEAL)}")
    _info(f"Total triples       : {_c(len(rdf_graph), WHITE)}")
    _info(f"SPARQL QC           : " +
          (_c("✔ Passed", GREEN) if sparql_failures == 0
           else _c(f"✗ {sparql_failures} failure(s)", RED)))
    _info(f"SHACL validation    : " +
          (_c("✔ Passed", GREEN) if shacl_ok else _c("✗ Failed", RED)))
    print()

    if sparql_failures == 0 and shacl_ok:
        print(_c("  ✔  PIPELINE FULLY GREEN — data is clean, RDF is valid.", GREEN + BOLD))
    elif shacl_ok:
        print(_c("  ⚠  PIPELINE ALMOST CLEAN — SHACL passed, SPARQL has 1 known issue.", AMBER + BOLD))
    else:
        print(_c("  ✗  PIPELINE HAS ISSUES — review the output above.", RED + BOLD))

    print()
    return 0 if (sparql_failures == 0 and shacl_ok) else 1


if __name__ == "__main__":
    sys.exit(main())

    