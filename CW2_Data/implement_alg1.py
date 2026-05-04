"""
implement_alg1.py  –  Naive column type annotation (Alg 1)

Algorithm overview
------------------
Step 1  Cell_Sampling   : randomly pick k unique, non-empty cells from the column
Step 2  Entity_Lookup   : map each cell to a DBpedia entity via the Lookup API (top-1)
Step 3  Class_Querying  : retrieve every dbo: rdf:type declared for that entity
Step 4  Class_Selection : pick the most frequent class across all sampled entities
"""

import csv
import random
import time
from collections import Counter

import requests

# ── Configuration ──────────────────────────────────────────────────────────
LOOKUP_URL   = "https://lookup.dbpedia.org/api/search"
SPARQL_URL   = "https://dbpedia.org/sparql"
DBO_PREFIX   = "http://dbpedia.org/ontology/"
TABLES_DIR   = "tables"
TARGETS_FILE = "Targets.csv"
RESULTS_FILE = "results_alg1.csv"
K            = 5          # number of cells sampled per column
SLEEP_SEC    = 0.4        # polite delay between API calls


# ── Step 1 ─────────────────────────────────────────────────────────────────

def Cell_Sampling(column_values: list, k: int) -> list:
    """Randomly select k distinct, non-empty cell values from the column."""
    unique_non_empty = list(dict.fromkeys(v for v in column_values if v.strip()))
    return random.sample(unique_non_empty, min(k, len(unique_non_empty)))


# ── Step 2 ─────────────────────────────────────────────────────────────────

def Entity_Lookup(cell: str) -> str | None:
    """Return the top-1 DBpedia resource URI for a cell string, or None on failure."""
    try:
        resp = requests.get(
            LOOKUP_URL,
            params={"query": cell, "format": "json", "maxResults": 1},
            timeout=10,
        )
        resp.raise_for_status()
        docs = resp.json().get("docs", [])
        if docs:
            resource = docs[0].get("resource", [])
            if resource:
                return resource[0]
    except Exception:
        pass
    return None


# ── Step 3 ─────────────────────────────────────────────────────────────────

def Class_Querying(entity_uri: str) -> list:
    """Return all directly declared DBpedia ontology (dbo:) classes for an entity."""
    query = f"""
    SELECT DISTINCT ?type WHERE {{
      <{entity_uri}> a ?type .
      FILTER (STRSTARTS(STR(?type), "{DBO_PREFIX}"))
    }}
    """
    try:
        resp = requests.get(
            SPARQL_URL,
            params={"query": query, "format": "application/sparql-results+json"},
            timeout=15,
        )
        resp.raise_for_status()
        bindings = resp.json()["results"]["bindings"]
        return [b["type"]["value"] for b in bindings]
    except Exception:
        return []


# ── Step 4 ─────────────────────────────────────────────────────────────────

def Class_Selection(classes: list) -> str | None:
    """Return the most frequent class in the list, or None if the list is empty."""
    if not classes:
        return None
    return Counter(classes).most_common(1)[0][0]


# ── Pipeline ───────────────────────────────────────────────────────────────

def annotate_column(table_name: str, col_id: int) -> str:
    """Run the full four-step algorithm for one target column."""
    path = f"{TABLES_DIR}/{table_name}.csv"
    with open(path, encoding="utf-8") as f:
        rows = list(csv.reader(f))

    # Extract raw column values (skip header row)
    column_values = [row[col_id] for row in rows[1:] if len(row) > col_id]

    # Step 1 – sample
    cells = Cell_Sampling(column_values, K)

    # Steps 2 & 3 – lookup then query for each sampled cell
    all_classes = []
    for cell in cells:
        entity_uri = Entity_Lookup(cell)
        if entity_uri:
            all_classes.extend(Class_Querying(entity_uri))
        time.sleep(SLEEP_SEC)

    # Step 4 – select
    prediction = Class_Selection(all_classes)
    return prediction if prediction else "NONE"


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    # Read target columns
    targets = []
    with open(TARGETS_FILE) as f:
        for line in f:
            parts = line.strip().split(",")
            tbl = parts[0].strip('"')
            col = int(parts[1].strip('"'))
            targets.append((tbl, col))

    results = []
    for idx, (tbl, col) in enumerate(targets, 1):
        print(f"[{idx:>2}/{len(targets)}] {tbl}  col={col} ... ", end="", flush=True)
        pred = annotate_column(tbl, col)
        print(pred)
        results.append((tbl, col, pred))

    # Write results CSV
    with open(RESULTS_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        for tbl, col, cls in results:
            writer.writerow([tbl, str(col), cls])

    print(f"\nDone. Results saved to {RESULTS_FILE}")


if __name__ == "__main__":
    main()
