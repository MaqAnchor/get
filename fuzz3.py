#!/usr/bin/env python3
"""
High-performance pipeline for auto-filling 'Application Name' from 'Short Description',
optimized for large datasets (~60k+ rows). Supports on-the-fly model download for embeddings:
- Column normalization
- Fast HashingVectorizer features with reduced dimensionality
- MultinomialNB for lightning-fast training
- Batched sentence-transformers encoding with automatic download
- Approximate k-NN via NearestNeighbors
- Vectorized predictions with predict_proba
- GUI file picker

Prerequisites:
    pip install pandas scikit-learn openpyxl sentence-transformers
"""
import sys
import os
import time
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog

# allow --user installs in roaming profile
user_site = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\site-packages")
if os.path.isdir(user_site): sys.path.insert(0, user_site)
scripts_dir = os.path.expanduser(r"~\AppData\Roaming\Python\Python312\Scripts")
if os.path.isdir(scripts_dir): sys.path.insert(0, scripts_dir)

from sklearn.feature_extraction.text import HashingVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import Pipeline, FeatureUnion
from sklearn.neighbors import NearestNeighbors
from sentence_transformers import SentenceTransformer

# —— Config ——
SKIP_AUDIT = True          # disable duplicate audit for speed
EMB_BATCH_SIZE = 256       # batch size for embedding encoding
THR_ML = 0.60
THR_EMB = 0.75
HASH_BITS = 16             # features = 2**HASH_BITS (65k dims)
# Model name and local cache directory for embeddings
MODEL_NAME = 'sentence-transformers/all-MiniLM-L6-v2'
CACHE_DIR = os.path.join(os.getcwd(), 'model_cache')

# —— Helpers ——

def normalize_columns(df):
    df = df.rename(columns=lambda c: c.strip().lower())
    mapping = {'short description': 'Short Description',
               'application name':  'Application Name'}
    return df.rename(columns=mapping)

# fast feature pipeline using HashingVectorizer + MultinomialNB
def build_ml_pipeline():
    dims = 2**HASH_BITS
    word_hash = HashingVectorizer(analyzer='word', ngram_range=(1,2), n_features=dims, alternate_sign=False)
    char_hash = HashingVectorizer(analyzer='char_wb', ngram_range=(3,5), n_features=dims, alternate_sign=False)
    feats = FeatureUnion([('w', word_hash), ('c', char_hash)], n_jobs=-1)
    clf = MultinomialNB()
    return Pipeline([('feats', feats), ('clf', clf)])

# rule-based overrides for known patterns
def rule_based_override(text):
    t = text.lower()
    if 'database' in t and 'backup' in t:
        return 'DBBackupApp'
    if t.startswith('crm:'):
        return 'CRMSystem'
    return None

# batched embedding encoder
def encode_batches(model, texts, batch_size=EMB_BATCH_SIZE):
    embs = []
    for i in range(0, len(texts), batch_size):
        batch = texts[i:i+batch_size]
        embs.append(model.encode(batch, normalize_embeddings=True))
    return np.vstack(embs)

# load or download SentenceTransformer model
def load_embedding_model():
    os.makedirs(CACHE_DIR, exist_ok=True)
    try:
        # attempt offline load
        model = SentenceTransformer(MODEL_NAME, cache_folder=CACHE_DIR, local_files_only=True)
    except Exception:
        # fallback to download
        print(f"Model not found locally, downloading {MODEL_NAME} to {CACHE_DIR}...")
        model = SentenceTransformer(MODEL_NAME, cache_folder=CACHE_DIR)
    return model

# —— Main ——

def main():
    t0 = time.time()
    print('1) Loading and normalizing App.xlsx → Sheet1...')
    df_train = pd.read_excel('App.xlsx', sheet_name='Sheet1')
    df_train = normalize_columns(df_train)
    print(f'   • {len(df_train)} rows, cols: {df_train.columns.tolist()}')

    # Prepare data
    texts = df_train['Short Description'].astype(str).tolist()
    labels= df_train['Application Name'].astype(str).tolist()

    print('\n2) Building and training ML pipeline (MultinomialNB)...')
    ml_pipe = build_ml_pipeline()
    ml_pipe.fit(texts, labels)

    print('3) Loading embeddings model (offline or download)...')
    emb_model = load_embedding_model()
    train_emb = encode_batches(emb_model, texts)
    nn = NearestNeighbors(metric='cosine', algorithm='brute', n_jobs=-1)
    nn.fit(train_emb)

    print('\n4) Select target file via GUI...')
    root = tk.Tk(); root.withdraw()
    fname = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')])
    root.destroy()
    if not fname:
        print('No file selected, exiting.'); return

    print(f'   • Reading {fname} → Page1')
    df_new = pd.read_excel(fname, sheet_name='Page1')
    df_new = normalize_columns(df_new)
    new_texts = df_new['Short Description'].astype(str).tolist()
    print(f'   • {len(new_texts)} rows to process')

    print('5) Encoding new texts and querying k-NN...')
    new_emb = encode_batches(emb_model, new_texts)
    dist, idx = nn.kneighbors(new_emb, n_neighbors=1)
    sims_emb = 1 - dist[:,0]

    print('6) ML predict_proba for NB...')
    ml_probs = ml_pipe.predict_proba(new_texts)
    ml_preds = ml_pipe.classes_[np.argmax(ml_probs, axis=1)]
    ml_conf  = ml_probs.max(axis=1)

    print('7) Merging rule, embedding, ML')
    results = []
    for i, desc in enumerate(new_texts):
        r = rule_based_override(desc)
        if r:
            results.append(r)
            continue
        if sims_emb[i] >= THR_EMB:
            results.append(labels[idx[i,0]])
            continue
        if ml_conf[i] < THR_ML:
            results.append('REVIEW MANUALLY')
        else:
            results.append(ml_preds[i])

    print('8) Inserting and saving results')
    df_new.insert(df_new.columns.get_loc('Short Description')+1,
                  'Application Name', results)
    out = os.path.splitext(fname)[0] + '_with_ApplicationName.xlsx'
    df_new.to_excel(out, sheet_name='Page1', index=False)
    print(f'   ✔ Saved → {out}')

    review = df_new[df_new['Application Name']=='REVIEW MANUALLY']
    if not review.empty:
        review.to_excel('to_review.xlsx', index=False)
        print(f'   • {len(review)} for manual review -> to_review.xlsx')

    print(f'Total time: {time.time()-t0:.2f}s')

if __name__ == '__main__':
    main()
