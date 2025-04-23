#!/usr/bin/env python3
"""
Full pipeline for auto-filling 'Application Name' from 'Short Description',
optimized for speed: column normalization, sampled duplicate audit,
enriched TF-IDF, parallelized pipeline, rule overrides,
precomputed embeddings + NearestNeighbors, and vectorized ML predictions.

Uses a GUI file picker and supports user-site installs without PATH changes.

Prerequisites:
    pip install pandas scikit-learn openpyxl sentence-transformers
"""
import sys
import os
import time

# Allow --user installs from roaming profile
user_site = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\site-packages")
if os.path.isdir(user_site): sys.path.insert(0, user_site)
scripts_dir = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\Scripts")
if os.path.isdir(scripts_dir): sys.path.insert(0, scripts_dir)

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.pipeline import Pipeline, FeatureUnion
from sklearn.linear_model import LogisticRegression
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.neighbors import NearestNeighbors
from sentence_transformers import SentenceTransformer

# —— Helpers ——

def normalize_columns(df):
    df = df.rename(columns=lambda c: c.strip().lower())
    mapping = {'short description': 'Short Description',
               'application name':  'Application Name'}
    return df.rename(columns=mapping)

# Sampled audit to avoid O(n^2) on large data
def audit_conflicting_duplicates(df, sim_threshold=0.9, sample_size=1000):
    n = len(df)
    sample = df.sample(n=min(n, sample_size), random_state=42)
    texts = sample['Short Description'].astype(str).tolist()
    apps  = sample['Application Name'].astype(str).tolist()
    tfidf = TfidfVectorizer(stop_words='english').fit_transform(texts)
    sims  = cosine_similarity(tfidf)
    conflicts = []
    m = len(texts)
    for i in range(m):
        for j in range(i+1, m):
            if sims[i,j] > sim_threshold and apps[i] != apps[j]:
                conflicts.append((i, j, sims[i,j], texts[i], apps[i], apps[j]))
    print(f"▶ Audit: {len(conflicts)} conflicts (sample size {m}) at sim>{sim_threshold}")
    for i,j,s,txt,a,b in conflicts[:5]:
        print(f"  • [{i}]≃[{j}] sim={s:.2f} → '{txt}' labels: '{a}' vs '{b}'")

# Parallelized TF-IDF + LogisticRegression
def build_ml_pipeline():
    word = TfidfVectorizer(analyzer='word', stop_words='english', ngram_range=(1,3), min_df=2)
    char = TfidfVectorizer(analyzer='char_wb', ngram_range=(3,5), min_df=2)
    feats = FeatureUnion([('w', word), ('c', char)], n_jobs=-1)
    clf   = LogisticRegression(solver='saga', max_iter=1000, n_jobs=-1)
    return Pipeline([('feats', feats), ('clf', clf)])

# Simple rule overrides
def rule_based_override(text):
    t = text.lower()
    if 'database' in t and 'backup' in t:
        return 'DBBackupApp'
    if t.startswith('crm:'):
        return 'CRMSystem'
    return None

# Preload SentenceTransformer and NearestNeighbors
def build_embeddings_model():
    return SentenceTransformer('all-MiniLM-L6-v2')

# —— Main ——

def main():
    t0 = time.time()

    # 1) Load & normalize training data
    print('1) Loading App.xlsx…')
    df_train = pd.read_excel('App.xlsx', sheet_name='Sheet1')
    df_train = normalize_columns(df_train)
    print(f"   • {len(df_train)} rows, cols: {df_train.columns.tolist()}")

    # 1a) Sampled audit
    audit_conflicting_duplicates(df_train)

    texts = df_train['Short Description'].astype(str).tolist()
    labels= df_train['Application Name'].astype(str).tolist()

    # 2) Train ML pipeline
    print('\n2) Training ML pipeline…')
    ml_pipe = build_ml_pipeline()
    ml_pipe.fit(texts, labels)
    print('   ✔ Pipeline trained.')

    # 3) Precompute train embeddings & NN index
    print('3) Loading embeddings & building index…')
    emb_model = build_embeddings_model()
    train_emb = emb_model.encode(texts, normalize_embeddings=True)
    nn = NearestNeighbors(metric='cosine', algorithm='brute', n_jobs=-1)
    nn.fit(train_emb)

    # 4) GUI for target file
    print('\n4) Select target file…')
    root = tk.Tk(); root.withdraw()
    fname = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')])
    root.destroy()
    if not fname:
        print('No file selected, exiting.'); return
    print(f"   • Reading {fname}…")
    df_new = pd.read_excel(fname, sheet_name='Page1')
    df_new = normalize_columns(df_new)
    new_texts = df_new['Short Description'].astype(str).tolist()
    print(f"   • {len(df_new)} rows to process.")

    # 5) Bulk encode new texts & query NN
    print('5) Encoding new descriptions & querying NN…')
    new_emb = emb_model.encode(new_texts, normalize_embeddings=True)
    dist, idx = nn.kneighbors(new_emb, n_neighbors=1)
    sims = 1 - dist[:,0]

    # 6) Bulk ML predict
    print('6) Predicting ML probabilities…')
    proba = ml_pipe.predict_proba(new_texts)
    ml_preds = ml_pipe.classes_[np.argmax(proba, axis=1)]
    ml_conf  = proba.max(axis=1)

    # 7) Assemble final predictions
    print('7) Assembling final predictions…')
    thr_ml  = 0.60
    thr_emb = 0.75
    results = []
    for i, desc in enumerate(new_texts):
        # rule
        r = rule_based_override(desc)
        if r:
            results.append(r); continue
        # embedding
        if sims[i] >= thr_emb:
            results.append(labels[idx[i,0]]); continue
        # ML
        if ml_conf[i] >= thr_ml:
            results.append(ml_preds[i])
        else:
            results.append('REVIEW MANUALLY')

    # 8) Insert and save
    df_new.insert(df_new.columns.get_loc('Short Description')+1,
                  'Application Name', results)
    out = os.path.splitext(fname)[0] + '_with_ApplicationName.xlsx'
    df_new.to_excel(out, sheet_name='Page1', index=False)
    print(f"\n✔ Saved predictions to {out}")

    # 9) Export low-confidence
    rev = df_new[df_new['Application Name']=='REVIEW MANUALLY']
    if not rev.empty:
        rev.to_excel('to_review.xlsx', index=False)
        print(f"   • {len(rev)} rows exported to to_review.xlsx")

    print(f"\nTotal time: {time.time()-t0:.2f}s")

if __name__ == '__main__':
    main()
