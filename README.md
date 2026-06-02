# CCO Semantic ETL Pipeline Demo

A CI-driven pipeline that ingests raw sensor data, normalizes it, transforms it into a [Common Core Ontologies (CCO)](https://github.com/CommonCoreOntology/CommonCoreOntologies)-aligned RDF knowledge graph, and validates the output using SPARQL and SHACL.

## Pipeline Overview

```
sensor_A.csv ─┐                                  ┌─ SPARQL QC (8 queries)
               ├→ normalize_readings.py → readings_normalized.csv → measure_rdflib.py → measure_cco.ttl ─┤
sensor_B.json ─┘                                  └─ SHACL validation
```

1. **Extract & Transform** — Raw sensor files with inconsistent schemas, units, and encodings are cleaned and merged into a single normalized CSV using Pandas.
2. **Map to Ontology** — Normalized readings are mapped to CCO's Measurement Design Pattern via RDFlib, producing an RDF/Turtle graph.
3. **Validate** — The graph is checked for structural and datatype conformance through 8 SPARQL quality-control queries and a set of SHACL shapes.
4. **Automate** — A GitHub Actions workflow triggers the full pipeline whenever a new dataset is added to `src/data/`.

## Repository Structure

```
├── src/
│   ├── data/
│   │   ├── sensor_A.csv            # Raw source A
│   │   ├── sensor_B.json           # Raw source B
│   │   ├── sensor_C.csv            # Additional dataset (triggers CI)
│   │   └── readings_normalized.csv # Pipeline output — cleaned data
│   ├── scripts/
│   │   ├── normalize_readings.py   # ETL script (Pandas)
│   │   ├── measure_rdflib.py       # RDF construction script (RDFlib)
│   │   ├── run_sparql_qc.py        # Runs SPARQL QC queries
│   │   └── shacl_validate.py       # Runs SHACL validation
│   ├── sparql/
│   │   └── *.rq                    # 8 SPARQL quality-control queries
│   ├── shacl/
│   │   └── cco_shapes.ttl          # SHACL shapes for graph validation
│   ├── measure_cco.ttl             # Pipeline output — RDF graph
│   ├── cco_merged.ttl              # Merged CCO modules (reference)
│   ├── ETL_Guide.md                # Guidance for the ETL script
│   └── Workflow_Guide.md           # Guidance for GitHub Actions setup
└── requirements.txt
```

## Quick Start

### Prerequisites

- Python 3.10+
- Dependencies listed in `requirements.txt`

### Install

```bash
pip install -r requirements.txt
```

### Run Locally

```bash
# 1. Normalize raw sensor data
python src/scripts/normalize_readings.py

# 2. Build the CCO-aligned RDF graph
python src/scripts/measure_rdflib.py

# 3. Run SPARQL quality-control queries (expect 0 results per query)
python src/scripts/run_sparql_qc.py

# 4. Run SHACL validation (expect conformance)
python src/scripts/shacl_validate.py
```

## Design Pattern

The RDF graph follows the **CCO Measurement Design Pattern**, which models measurements using BFO and CCO classes such as measurement information content entities, specifically dependent continuants, and information bearing entities. See `cco_measurement_design_pattern.pptx` for the full pattern diagram.

## Validation Strategy

| Method | Purpose | Pass Criteria |
|--------|---------|---------------|
| **SPARQL QC** | Structural checks (e.g., unique `rdfs:label` per entity, required properties) | Each query returns **0 results** |
| **SHACL Shapes** | Datatype and cardinality conformance | Graph **conforms** with no violations |
