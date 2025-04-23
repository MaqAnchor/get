#!/usr/bin/env python3
"""
Full pipeline for auto-filling 'Application' from 'Short description',
with duplicate auditing, enriched features, rule overrides, embeddings+knn,
and low-confidence flagging.

Prerequisites:
    pip install pandas scikit-learn openpyxl sentence-transformers



This version prepends your user‐site install paths so you can import packages
from your roaming profile without touching system PATH.
"""

import sys
import os
import time

# ── allow user‐site installs without modifying PATH ──
user_site = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\site-packages")
if os.path.isdir(user_site):
    sys.path.insert(0, user_site)
scripts_dir = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\Scripts")
if os.path.isdir(scripts_dir):
    sys.path.insert(0, scripts_dir)

# now safe to import everything else
import pandas as pd
import numpy as np

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.pipeline import Pipeline, FeatureUnion
from sklearn.linear_model import LogisticRegression
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

# —— Helpers —— #

def normalize_columns(df, mapping=None):
    """
    Strip whitespace and unify casing in column names.
    Optionally remap any variants via `mapping` dict.
    """
    df = df.rename(columns=lambda c: c.strip())
    if mapping:
        df = df.rename(columns=mapping)
    return df

def audit_conflicting_duplicates(df, sim_threshold=0.9):
    texts = df["Short description"].astype(str).tolist()
    apps  = df["Application"].astype(str).tolist()
    tfidf = TfidfVectorizer(stop_words="english").fit_transform(texts)
    sims  = cosine_similarity(tfidf)
    conflicts = []
    n = sims.shape[0]
    for i in range(n):
        for j in range(i+1, n):
            if sims[i,j] > sim_threshold and apps[i] != apps[j]:
                conflicts.append((i, j, sims[i,j], texts[i], apps[i], apps[j]))
    print(f"\n▶ Audit: {len(conflicts)} conflicts at sim>{sim_threshold}")
    for i,j,s,txt,a,b in conflicts[:5]:
        print(f"  • [{i}]≃[{j}] sim={s:.2f} → '{txt}'  labels: '{a}' vs '{b}'")

def build_ml_pipeline():
    word_ngram = TfidfVectorizer(
        analyzer="word", stop_words="english",
        ngram_range=(1,3), min_df=2
    )
    char_ngram = TfidfVectorizer(
        analyzer="char_wb", ngram_range=(3,5), min_df=2
    )
    features = FeatureUnion([("word", word_ngram), ("char", char_ngram)])
    return Pipeline([("features", features),
                     ("clf", LogisticRegression(max_iter=1000))])

def rule_based_override(text):
    t = text.lower()
    if "database" in t and "backup" in t:
        return "DBBackupApp"
    if t.startswith("crm:"):
        return "CRMSystem"
    return None

def build_embeddings_model():
    return SentenceTransformer("all-MiniLM-L6-v2")

def predict_with_embeddings(model, train_texts, train_labels, query, emb_threshold=0.75):
    q_emb = model.encode([query], normalize_embeddings=True)
    t_emb = model.encode(train_texts, normalize_embeddings=True)
    sims  = cosine_similarity(q_emb, t_emb)[0]
    idx   = int(np.argmax(sims))
    if sims[idx] >= emb_threshold:
        return train_labels[idx]
    return None

# —— Main —— #

def main():
    start = time.time()

    # 1) Load & normalize training data
    print("1) Loading App.xlsx → Sheet1…")
    df_train = pd.read_excel("App.xlsx", sheet_name="Sheet1")
    col_map = {
        # if your headers were e.g. 'short description' or ' Short description '
        "short description":  "Short description",
        "application":        "Application"
    }
    df_train = normalize_columns(df_train, mapping=col_map)
    print(f"   • Columns = {df_train.columns.tolist()}")

    # 1a) Audit for conflicting near-duplicates
    audit_conflicting_duplicates(df_train)

    texts  = df_train["Short description"].astype(str).tolist()
    labels = df_train["Application"].astype(str).tolist()

    # 2) Train enriched TF-IDF + LR
    print("\n2) Training ML pipeline…")
    ml_pipeline = build_ml_pipeline()
    ml_pipeline.fit(texts, labels)

    # 3) Prepare embeddings model
    print("3) Loading embeddings model…")
    emb_model = build_embeddings_model()

    # 4) Prompt & load new data
    fname = input("\n4) Enter Excel file to process (e.g. NewData.xlsx): ").strip()
    print(f"   • Reading '{fname}' sheet 'Page 1'…")
    df_new = pd.read_excel(fname, sheet_name="Page 1")
    df_new = normalize_columns(df_new, mapping=col_map)
    print(f"   • Columns = {df_new.columns.tolist()}")
    print(f"   • {len(df_new)} rows to predict.\n")

    # 5) Hybrid prediction
    print("5) Predicting…")
    proba = ml_pipeline.predict_proba(df_new["Short description"].astype(str))
    preds = []
    for i, desc in enumerate(df_new["Short description"].astype(str)):
        # (a) rule
        rule = rule_based_override(desc)
        if rule:
            preds.append(rule)
            continue
        # (b) embeddings
        emb_lbl = predict_with_embeddings(emb_model, texts, labels, desc)
        if emb_lbl:
            preds.append(emb_lbl)
            continue
        # (c) ML + threshold
        rowp = proba[i]
        best = ml_pipeline.classes_[np.argmax(rowp)]
        preds.append(best if rowp.max() >= 0.60 else "REVIEW MANUALLY")

    # 6) Insert & save
    print(f"   • Flagged = {preds.count('REVIEW MANUALLY')}")
    idx = df_new.columns.get_loc("Short description")
    df_new.insert(idx+1, "Application", preds)
    out = fname.replace(".xlsx", "_with_Apps.xlsx")
    df_new.to_excel(out, sheet_name="Page 1", index=False)
    print(f"\n6) Saved → {out}")

    # 7) Export to-review
    review = df_new[df_new["Application"]=="REVIEW MANUALLY"]
    if not review.empty:
        review.to_excel("to_review.xlsx", index=False)
        print(f"   • to_review.xlsx ({len(review)} rows)")

    print(f"\n✔ Done in {time.time()-start:.2f}s")

if __name__=="__main__":
    main()
