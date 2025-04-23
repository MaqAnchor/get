#!/usr/bin/env python3
"""
Full pipeline for auto-filling 'Application' from 'Short description',
with duplicate auditing, enriched features, rule overrides, embeddings+knn,
and low-confidence flagging.

Prerequisites:
    pip install pandas scikit-learn openpyxl sentence-transformers
"""

import time
import pandas as pd
import numpy as np

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.pipeline import Pipeline, FeatureUnion
from sklearn.linear_model import LogisticRegression
from sklearn.metrics.pairwise import cosine_similarity

from sentence_transformers import SentenceTransformer

# —— 1. AUDIT & CLEAN ——
def audit_conflicting_duplicates(df, sim_threshold=0.9):
    """
    Prints out any pairs of descriptions whose TF-IDF cosine similarity
    exceeds sim_threshold but whose Application labels differ.
    """
    print(f"\n▶ Auditing for near-duplicates (sim > {sim_threshold}) with conflicting labels…")
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
    if not conflicts:
        print("   ✔ No conflicting near-duplicates found.")
    else:
        print(f"   ⚠ Found {len(conflicts)} conflicting pairs; showing up to 5:")
        for i, j, s, txt, a, b in conflicts[:5]:
            print(f"     • [{i}]≃[{j}] sim={s:.2f} → '{txt}'  labels: '{a}' vs '{b}'")
    return

# —— 2. FEATURE-ENRICHED PIPELINE ——  
def build_ml_pipeline():
    """
    Combines word-level and character-level TF-IDF features, then LogisticRegression.
    """
    word_ngram = TfidfVectorizer(
        analyzer="word", stop_words="english",
        ngram_range=(1,3), min_df=2
    )
    char_ngram = TfidfVectorizer(
        analyzer="char_wb",
        ngram_range=(3,5), min_df=2
    )
    features = FeatureUnion([
        ("word", word_ngram),
        ("char", char_ngram)
    ])
    clf = LogisticRegression(max_iter=1000)
    pipeline = Pipeline([
        ("features", features),
        ("clf", clf)
    ])
    return pipeline

# —— 3. RULE-BASED OVERRIDES ——
def rule_based_override(text: str):
    """
    Return an Application if `text` matches a hard rule; else None.
    Customize these rules for your domain.
    """
    t = text.lower()
    if "database" in t and "backup" in t:
        return "DBBackupApp"
    if t.startswith("crm:"):
        return "CRMSystem"
    # TODO: add your own patterns here
    return None

# —— 4. EMBEDDINGS + k-NN ——  
def build_embeddings_model():
    """
    Load a lightweight SentenceTransformer for encoding.
    """
    return SentenceTransformer("all-MiniLM-L6-v2")

def predict_with_embeddings(model, train_texts, train_labels, query, emb_threshold=0.75):
    """
    If the query's embedding has cosine similarity > emb_threshold
    with any training example, return that example's label; else None.
    """
    q_emb = model.encode([query], normalize_embeddings=True)
    t_emb = model.encode(train_texts, normalize_embeddings=True)
    sims  = cosine_similarity(q_emb, t_emb)[0]
    best_idx = int(np.argmax(sims))
    best_sim = sims[best_idx]
    if best_sim >= emb_threshold:
        return train_labels[best_idx]
    return None

def main():
    start = time.time()
    # -- Load & audit training data --
    print("1) Loading training data from App.xlsx → Sheet1…")
    df_train = pd.read_excel("App.xlsx", sheet_name="Sheet1")
    print(f"   • {len(df_train)} labeled rows loaded.")
    audit_conflicting_duplicates(df_train)

    texts = df_train["Short description"].astype(str).tolist()
    labels = df_train["Application"].astype(str).tolist()

    # -- Train ML pipeline --
    print("\n2) Training enriched TF-IDF + LogisticRegression model…")
    ml_pipeline = build_ml_pipeline()
    ml_pipeline.fit(texts, labels)
    print("   ✔ Model trained.")

    # -- Prepare embeddings model once --
    print("\n3) Loading sentence-transformers model for k-NN fallback…")
    emb_model = build_embeddings_model()

    # -- Prompt & load new data --
    fname = input("\n4) Enter the Excel filename to process (e.g. NewData.xlsx): ").strip()
    print(f"   • Reading '{fname}' sheet 'Page 1'…")
    df_new = pd.read_excel(fname, sheet_name="Page 1")
    print(f"   • {len(df_new)} rows to predict.\n")

    # -- Predict each row using hybrid strategy --
    print("5) Running hybrid prediction…")
    ml_probs = None
    predictions = []
    for desc in df_new["Short description"].astype(str):
        # 5a) Rule-based?
        rule = rule_based_override(desc)
        if rule:
            predictions.append(rule)
            continue

        # 5b) Embeddings + k-NN?
        emb_label = predict_with_embeddings(emb_model, texts, labels, desc)
        if emb_label:
            predictions.append(emb_label)
            continue

        # 5c) ML model + confidence threshold
        if ml_probs is None:
            ml_probs = ml_pipeline.predict_proba(df_new["Short description"].astype(str))
        # find this row's index
        idx = len(predictions)
        row_proba = ml_probs[idx]
        best_class = ml_pipeline.classes_[np.argmax(row_proba)]
        if row_proba.max() >= 0.60:
            predictions.append(best_class)
        else:
            predictions.append("REVIEW MANUALLY")

    # -- Insert, save, and export low-confidence cases --
    print(f"   • Total flagged for review: {predictions.count('REVIEW MANUALLY')}")
    sd_i = df_new.columns.get_loc("Short description")
    df_new.insert(sd_i+1, "Application", predictions)

    out_name = fname.replace(".xlsx", "_with_Apps.xlsx")
    print(f"\n6) Saving full predictions to '{out_name}'…")
    df_new.to_excel(out_name, sheet_name="Page 1", index=False)

    # Export only the rows needing review
    review_df = df_new[df_new["Application"] == "REVIEW MANUALLY"]
    if not review_df.empty:
        review_name = "to_review.xlsx"
        print(f"   • Exporting {len(review_df)} rows to '{review_name}' for manual labeling.")
        review_df.to_excel(review_name, index=False)

    print(f"\n✔ Done in {time.time() - start:.2f}s!")

if __name__ == "__main__":
    main()
