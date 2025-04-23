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

# ── allow user-site installs without modifying PATH ──
user_site = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\site-packages")
if os.path.isdir(user_site):
    sys.path.insert(0, user_site)
scripts_dir = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\Scripts")
if os.path.isdir(scripts_dir):
    sys.path.insert(0, scripts_dir)

# now safe to import everything else
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.pipeline import Pipeline, FeatureUnion
from sklearn.linear_model import LogisticRegression
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

# —— Helpers —— #

def normalize_columns(df):
    """
    Strip whitespace, lowercase all column names, then remap known keys to canonical names.
    """
    df = df.rename(columns=lambda c: c.strip().lower())
    mapping = {
        'short description': 'Short Description',
        'application name': 'Application Name'
    }
    return df.rename(columns=mapping)


def audit_conflicting_duplicates(df, sim_threshold=0.9):
    print(f"\n▶ Auditing for near-duplicates (sim > {sim_threshold}) with conflicting labels…")
    texts = df['Short Description'].astype(str).tolist()
    apps = df['Application Name'].astype(str).tolist()
    tfidf = TfidfVectorizer(stop_words='english').fit_transform(texts)
    sims = cosine_similarity(tfidf)
    conflicts = []
    n = sims.shape[0]
    for i in range(n):
        for j in range(i+1, n):
            if sims[i, j] > sim_threshold and apps[i] != apps[j]:
                conflicts.append((i, j, sims[i,j], texts[i], apps[i], apps[j]))
    if not conflicts:
        print("   ✔ No conflicting near-duplicates found.")
    else:
        print(f"   ⚠ Found {len(conflicts)} conflicting pairs; showing up to 5:")
        for i, j, s, txt, a, b in conflicts[:5]:
            print(f"     • [{i}]≃[{j}] sim={s:.2f} → '{txt}'  labels: '{a}' vs '{b}'")


def build_ml_pipeline():
    word_ngram = TfidfVectorizer(analyzer='word', stop_words='english', ngram_range=(1,3), min_df=2)
    char_ngram = TfidfVectorizer(analyzer='char_wb', ngram_range=(3,5), min_df=2)
    features = FeatureUnion([('word', word_ngram), ('char', char_ngram)])
    clf = LogisticRegression(max_iter=1000)
    return Pipeline([('features', features), ('clf', clf)])


def rule_based_override(text):
    t = text.lower()
    if 'database' in t and 'backup' in t:
        return 'DBBackupApp'
    if t.startswith('crm:'):
        return 'CRMSystem'
    return None


def build_embeddings_model():
    return SentenceTransformer('all-MiniLM-L6-v2')


def predict_with_embeddings(model, train_texts, train_labels, query, emb_threshold=0.75):
    q_emb = model.encode([query], normalize_embeddings=True)
    t_emb = model.encode(train_texts, normalize_embeddings=True)
    sims = cosine_similarity(q_emb, t_emb)[0]
    best_idx = int(np.argmax(sims))
    if sims[best_idx] >= emb_threshold:
        return train_labels[best_idx]
    return None


def main():
    start = time.time()

    # 1) Load & normalize training data
    print('1) Loading App.xlsx → Sheet1...')
    df_train = pd.read_excel('App.xlsx', sheet_name='Sheet1')
    df_train = normalize_columns(df_train)
    print(f"   • Columns = {df_train.columns.tolist()}")
    print(f"   • {len(df_train)} labeled rows loaded.")

    # 1a) Audit for conflicts
    audit_conflicting_duplicates(df_train)

    texts = df_train['Short Description'].astype(str).tolist()
    labels = df_train['Application Name'].astype(str).tolist()

    # 2) Train ML pipeline
    print('\n2) Training TF-IDF + LR pipeline...')
    ml_pipeline = build_ml_pipeline()
    ml_pipeline.fit(texts, labels)
    print('   ✔ Model trained.')

    # 3) Load embeddings model
    print('3) Loading embeddings model...')
    emb_model = build_embeddings_model()

    # 4) GUI file picker for new data
    print('\n4) Please select the target Excel file...')
    root = tk.Tk()
    root.withdraw()
    fname = filedialog.askopenfilename(title='Select Excel file', filetypes=[('Excel files', '*.xlsx')])
    root.destroy()
    if not fname:
        print('No file selected, exiting.')
        return
    print(f"   • Reading '{fname}' sheet 'Page1'...")
    df_new = pd.read_excel(fname, sheet_name='Page1')
    df_new = normalize_columns(df_new)
    print(f"   • Columns = {df_new.columns.tolist()}")
    print(f"   • {len(df_new)} rows to predict.\n")

    # 5) Hybrid prediction
    print('5) Predicting...')
    proba = ml_pipeline.predict_proba(df_new['Short Description'].astype(str))
    predictions = []
    for i, desc in enumerate(df_new['Short Description'].astype(str)):
        rule = rule_based_override(desc)
        if rule:
            predictions.append(rule)
            continue
        emb_label = predict_with_embeddings(emb_model, texts, labels, desc)
        if emb_label:
            predictions.append(emb_label)
            continue
        rowp = proba[i]
        top = ml_pipeline.classes_[np.argmax(rowp)]
        predictions.append(top if rowp.max() >= 0.60 else 'REVIEW MANUALLY')

    # 6) Insert & save
    flagged = predictions.count('REVIEW MANUALLY')
    print(f"   • Flagged for review = {flagged}")
    idx = df_new.columns.get_loc('Short Description')
    df_new.insert(idx+1, 'Application Name', predictions)
    out = os.path.splitext(fname)[0] + '_with_ApplicationName.xlsx'
    df_new.to_excel(out, sheet_name='Page1', index=False)
    print(f"\n6) Saved → {out}")

    # 7) Export low-confidence for manual labeling
    review_df = df_new[df_new['Application Name']=='REVIEW MANUALLY']
    if not review_df.empty:
        review_df.to_excel('to_review.xlsx', index=False)
        print(f"   • Exported to_review.xlsx ({len(review_df)} rows)")

    print(f"\n✔ Done in {time.time() - start:.2f}s!")

if __name__ == '__main__':
    main()
