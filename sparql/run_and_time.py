"""
Run all four Task-1 SPARQL queries and print execution time + result count.
Requires:  pip install SPARQLWrapper
"""
import time
from SPARQLWrapper import SPARQLWrapper, JSON

DBPEDIA   = "https://dbpedia.org/sparql"
WIKIDATA  = "https://query.wikidata.org/sparql"

QUERIES = [
    {
        "label": "Task 1.1 – DBpedia country list",
        "endpoint": DBPEDIA,
        "sparql": open("task1_1_dbpedia_countries.sparql").read(),
    },
    {
        "label": "Task 1.2 – DBpedia country attributes",
        "endpoint": DBPEDIA,
        "sparql": open("task1_2_dbpedia_country_attrs.sparql").read(),
    },
    {
        "label": "Task 1.3 (3rd) – Wikidata country list",
        "endpoint": WIKIDATA,
        "sparql": open("task1_3_wikidata_countries.sparql").read(),
    },
    {
        "label": "Task 1.3 (4th) – Wikidata country attributes + provinces",
        "endpoint": WIKIDATA,
        "sparql": open("task1_3_wikidata_country_attrs.sparql").read(),
    },
]

for q in QUERIES:
    sparql = SPARQLWrapper(q["endpoint"])
    sparql.setQuery(q["sparql"])
    sparql.setReturnFormat(JSON)

    print(f"\n{'='*60}")
    print(f"Running: {q['label']}")
    print(f"Endpoint: {q['endpoint']}")

    t0 = time.time()
    try:
        results = sparql.query().convert()
        elapsed = time.time() - t0
        rows = results["results"]["bindings"]
        print(f"  Results : {len(rows)} rows")
        print(f"  Time    : {elapsed:.2f} seconds")
        if rows:
            print(f"  Sample  : {list(rows[0].values())[0]['value']}")
    except Exception as e:
        elapsed = time.time() - t0
        print(f"  ERROR after {elapsed:.2f}s: {e}")
