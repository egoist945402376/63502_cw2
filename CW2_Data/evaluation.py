import os
import pandas as pd

DBO_PREFIX = "http://dbpedia.org/ontology/"
_DIR = os.path.dirname(os.path.abspath(__file__))

def normalize_label(x: str) -> str:
    """Return a clean DBpedia ontology URI or 'NONE'."""
    s = str(x).strip().strip('"').strip("'")
    return s if s.startswith(DBO_PREFIX) else "NONE"

def checker():
    gt_df = pd.read_csv(os.path.join(_DIR, "GT.csv"), header=None)
    y_true = gt_df[2].astype(str).str.strip().str.strip('"').str.strip("'").tolist()
    y_true = [normalize_label(x) for x in y_true]

    pred_df = pd.read_csv(os.path.join(_DIR, "results_alg1.csv"), header=None)
    y_pred = pred_df[2].astype(str).str.strip().str.strip('"').str.strip("'").tolist()
    y_pred = [normalize_label(x) for x in y_pred]

    #with open("results_alg1.txt", "r", encoding="utf-8") as f:
    #    y_pred = [line.strip() for line in f]
    #y_pred = [normalize_label(x) for x in y_pred]

    print("Length of Ground Truth: ", len(y_true))
    print("Length of Predictions:  ", len(y_pred))

    if len(y_true) != len(y_pred):
        m = min(len(y_true), len(y_pred))
        print(f"WARNING: length mismatch, truncating both to {m}")
        y_true = y_true[:m]
        y_pred = y_pred[:m]

    tp = 0
    for i, c in enumerate(y_true):
        if y_pred[i] == c:
            tp += 1
    acc = tp / len(y_true)
    print(f"Accuracy: {acc}")
   
checker()
