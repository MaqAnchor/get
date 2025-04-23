import pandas as pd
import numpy as np
import re
import os
import time
import tkinter as tk
from tkinter import filedialog
from rapidfuzz import process, fuzz
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from datetime import datetime
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

# Start tracking execution time
start_time = time.time()

print("Starting Advanced Application Mapping Process...")
print("-----------------------------------------------")

# Initialize NLTK resources - download if needed
try:
    nltk.data.find('tokenizers/punkt')
    nltk.data.find('corpora/stopwords')
    nltk.data.find('corpora/wordnet')
except LookupError:
    print("Downloading required NLTK resources...")
    nltk.download('punkt')
    nltk.download('stopwords')
    nltk.download('wordnet')

# Create a file picker dialog
root = tk.Tk()
root.withdraw()  # Hide the main window

print("Please select your 'Page 1' Excel file...")
page1_file_path = filedialog.askopenfilename(
    title="Select your 'Page 1' Excel file",
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
)

if not page1_file_path:
    print("No file selected. Exiting program.")
    exit()

print(f"Selected file: {page1_file_path}")

# Step 2: Load the three data sources
print("\nLoading master application and TCS group data...")
try:
    # Load application master list
    app_groups_df = pd.read_excel("Application_Groups.xlsx", sheet_name="ApplicationName")
    print(f"Loaded {len(app_groups_df)} master application records")
    
    # Load TCS group lookup
    tcs_groups_df = pd.read_excel("Application_Groups.xlsx", sheet_name="TCSGroups")
    print(f"Loaded {len(tcs_groups_df)} TCS group mappings")
    
    # Load historic mappings
    historic_apps_df = pd.read_excel("App.xlsx", sheet_name="Sheet1")
    print(f"Loaded {len(historic_apps_df)} historic application mappings")
except Exception as e:
    print(f"Error loading reference data: {e}")
    exit()

# Step 3: Read the "Page 1" sheet of the selected file
print("\nLoading your selected Page 1 data...")
try:
    page1_df = pd.read_excel(page1_file_path, sheet_name="Page 1")
    original_row_count = len(page1_df)
    print(f"Loaded {original_row_count} rows from Page 1")
except Exception as e:
    print(f"Error loading Page 1 data: {e}")
    exit()

# Step 4: Add the necessary columns in the right positions
print("\nPreparing target dataframe with additional columns...")
try:
    # Find the index for column G and the "Assignment group" column
    column_list = page1_df.columns.tolist()
    col_g_index = column_list.index(column_list[6])  # Index of column G (zero-based)
    
    # Find assignment group column
    assignment_group_col = None
    for col in column_list:
        if "Assignment group" in col or "assignment group" in str(col).lower():
            assignment_group_col = col
            break
    
    if assignment_group_col:
        assignment_group_index = column_list.index(assignment_group_col)
        print(f"Found 'Assignment group' column at position {assignment_group_index+1}")
    else:
        print("Warning: 'Assignment group' column not found. Using default position")
        assignment_group_index = col_g_index + 2  # Just a fallback position
    
    # Insert the new columns
    # Add Application Name after column G
    new_cols = column_list[:col_g_index+1] + ["Application Name"] + column_list[col_g_index+1:]
    
    # Find the new position of Assignment group column after inserting Application Name
    new_assignment_index = new_cols.index(assignment_group_col)
    
    # Add TCS Group after the Assignment group column
    final_cols = new_cols[:new_assignment_index+1] + ["TCS Group"] + new_cols[new_assignment_index+1:]
    
    # Create a new DataFrame with the new column structure
    new_df = pd.DataFrame(columns=final_cols)
    
    # Copy data from old DataFrame to new DataFrame, accounting for new columns
    for col in column_list:
        new_df[col] = page1_df[col]
    
    # Initialize the new columns with empty values
    new_df["Application Name"] = ""
    new_df["TCS Group"] = ""
    
    page1_df = new_df
    print("Successfully added 'Application Name' and 'TCS Group' columns")
    
except Exception as e:
    print(f"Error adding columns: {e}")
    # Fallback: Just add the columns at the end if positioning fails
    if "Application Name" not in page1_df.columns:
        page1_df["Application Name"] = ""
    if "TCS Group" not in page1_df.columns:
        page1_df["TCS Group"] = ""
    print("Used fallback method to add columns")

# Step 5: Build a dictionary of historic exact matches
print("\nBuilding lookup dictionary for exact matches...")
try:
    # Filter out rows with null values in either column
    valid_mappings = historic_apps_df.dropna(subset=['Short description', 'Application Name'])
    
    # Keep first occurrence of each 'Short description' to handle duplicates
    unique_mappings = valid_mappings.drop_duplicates(subset=['Short description'])
    
    # Create dictionary for exact matching
    exact_match_dict = dict(zip(unique_mappings['Short description'], unique_mappings['Application Name']))
    print(f"Created lookup dictionary with {len(exact_match_dict)} unique short descriptions")
    
except Exception as e:
    print(f"Error building lookup dictionary: {e}")
    exact_match_dict = {}

# Text preprocessing function for better matching
def preprocess_text(text):
    if not isinstance(text, str) or pd.isna(text):
        return ""
    
    # Tokenize and lowercase
    tokens = word_tokenize(text.lower())
    
    # Remove stopwords and punctuation
    stop_words = set(stopwords.words('english'))
    tokens = [token for token in tokens if token.isalnum() and token not in stop_words]
    
    # Lemmatize
    lemmatizer = WordNetLemmatizer()
    tokens = [lemmatizer.lemmatize(token) for token in tokens]
    
    return " ".join(tokens)

# Prepare a preprocessed version of our lookup keys for advanced matching
print("\nPreprocessing text data for semantic matching...")
preprocessed_keys = {k: preprocess_text(k) for k in exact_match_dict.keys() if isinstance(k, str)}
original_keys = list(preprocessed_keys.keys())
preprocessed_values = list(preprocessed_keys.values())

# Step 6a: First fill by exact lookup
print("\nPerforming exact matches...")
exact_match_count = 0
try:
    for i, row in page1_df.iterrows():
        short_desc = row.get('Short description')
        if pd.notna(short_desc) and short_desc in exact_match_dict:
            page1_df.at[i, 'Application Name'] = exact_match_dict[short_desc]
            exact_match_count += 1
    
    print(f"Applied {exact_match_count} exact matches")
except Exception as e:
    print(f"Error during exact matching: {e}")

# Step 6b: For remaining blanks, use regex to find substrings
print("\nPerforming substring matching for remaining blanks...")
substring_match_count = 0
try:
    # Create a list of patterns to look for
    patterns = list(exact_match_dict.keys())
    # Escape special regex characters and filter out non-string or null values
    escaped_patterns = [re.escape(p) for p in patterns if isinstance(p, str) and pd.notna(p)]
    
    # Create one big regex pattern with all possible matches
    combined_pattern = '|'.join(f'({p})' for p in escaped_patterns)
    
    # Apply regex matching to rows where Application Name is still blank
    blank_app_mask = page1_df['Application Name'] == ""
    blank_rows = page1_df[blank_app_mask]
    
    for i, row in blank_rows.iterrows():
        short_desc = row.get('Short description')
        if pd.notna(short_desc) and isinstance(short_desc, str):
            # Try to find the first match
            match = re.search(combined_pattern, short_desc)
            if match:
                # Get the matching text
                matched_text = next((group for group in match.groups() if group is not None), None)
                if matched_text:
                    # Map to application name
                    page1_df.at[i, 'Application Name'] = exact_match_dict[matched_text]
                    substring_match_count += 1
    
    print(f"Applied {substring_match_count} substring matches")
except Exception as e:
    print(f"Error during substring matching: {e}")

# Step 6c: For remaining blanks, use advanced semantic matching
print("\nPerforming advanced semantic matching for remaining blanks...")
semantic_match_count = 0
try:
    # Get remaining blank rows
    blank_app_mask = page1_df['Application Name'] == ""
    blank_rows = page1_df[blank_app_mask]
    
    # Only proceed if we have blank rows and reference data
    if not blank_rows.empty and preprocessed_values:
        # Create TF-IDF vectorizer
        vectorizer = TfidfVectorizer(min_df=1)
        
        # Fit the vectorizer on our preprocessed reference texts
        tfidf_matrix = vectorizer.fit_transform(preprocessed_values)
        
        # Define threshold for TF-IDF cosine similarity
        TFIDF_THRESHOLD = 0.65  # Adjust as needed based on testing
        
        for i, row in blank_rows.iterrows():
            short_desc = row.get('Short description')
            if pd.notna(short_desc) and isinstance(short_desc, str):
                # Preprocess the query text
                preprocessed_query = preprocess_text(short_desc)
                
                # Vectorize the query
                query_vector = vectorizer.transform([preprocessed_query])
                
                # Calculate cosine similarity with all reference texts
                cosine_similarities = cosine_similarity(query_vector, tfidf_matrix).flatten()
                
                # Get the best match
                best_match_index = cosine_similarities.argmax()
                best_match_score = cosine_similarities[best_match_index]
                
                if best_match_score >= TFIDF_THRESHOLD:
                    # Map to application name using the original key
                    original_key = original_keys[best_match_index]
                    page1_df.at[i, 'Application Name'] = exact_match_dict[original_key]
                    semantic_match_count += 1
                else:
                    # Try RapidFuzz as a fallback
                    # This combines token set ratio with partial ratio for better matching
                    match, score, _ = process.extractOne(
                        short_desc,
                        original_keys,
                        scorer=lambda x, y: (fuzz.token_set_ratio(x, y) + fuzz.partial_ratio(x, y)) / 2
                    )
                    
                    FUZZY_THRESHOLD = 80  # Threshold for the combined score
                    if score >= FUZZY_THRESHOLD:
                        page1_df.at[i, 'Application Name'] = exact_match_dict[match]
                        semantic_match_count += 1
                    else:
                        # Still no good match
                        page1_df.at[i, 'Application Name'] = "Not Available"
        
        print(f"Applied {semantic_match_count} semantic matches")
    else:
        print("No blank rows to process or no reference data available")
except Exception as e:
    print(f"Error during semantic matching: {e}")
    print("Falling back to basic fuzzy matching...")
    
    # Fallback to basic fuzzy matching if semantic matching fails
    try:
        fuzzy_match_count = 0
        # Get remaining blank rows again in case some were filled
        blank_app_mask = page1_df['Application Name'] == ""
        blank_rows = page1_df[blank_app_mask]
        
        for i, row in blank_rows.iterrows():
            short_desc = row.get('Short description')
            if pd.notna(short_desc) and isinstance(short_desc, str):
                # Get the best match
                match_result = process.extractOne(
                    short_desc,
                    [k for k in exact_match_dict.keys() if isinstance(k, str) and pd.notna(k)],
                    scorer=fuzz.token_set_ratio
                )
                
                if match_result:
                    match, score, _ = match_result
                    FALLBACK_THRESHOLD = 75
                    if score >= FALLBACK_THRESHOLD:
                        page1_df.at[i, 'Application Name'] = exact_match_dict[match]
                        fuzzy_match_count += 1
        
        print(f"Applied {fuzzy_match_count} fallback fuzzy matches")
        semantic_match_count = fuzzy_match_count  # Update count for the summary
    except Exception as e:
        print(f"Error during fallback fuzzy matching: {e}")

# Step 7: Fill any remaining gaps with "Not Available"
print("\nFilling remaining gaps...")
remaining_count = 0
try:
    blank_app_mask = page1_df['Application Name'] == ""
    remaining_count = blank_app_mask.sum()
    page1_df.loc[blank_app_mask, 'Application Name'] = "Not Available"
    print(f"Filled {remaining_count} remaining blank Application Name entries with 'Not Available'")
except Exception as e:
    print(f"Error filling remaining gaps: {e}")

# Step 8: Populate TCS Group using Assignment group mapping
print("\nMapping TCS Groups from Assignment group values...")
tcs_mapping_count = 0
try:
    # Create TCS mapping dictionary
    if 'Assignment group' in tcs_groups_df.columns and 'TCS Group' in tcs_groups_df.columns:
        tcs_dict = dict(zip(tcs_groups_df['Assignment group'], tcs_groups_df['TCS Group']))
        
        for i, row in page1_df.iterrows():
            assignment_group = row.get('Assignment group')
            if pd.notna(assignment_group) and assignment_group in tcs_dict:
                page1_df.at[i, 'TCS Group'] = tcs_dict[assignment_group]
                tcs_mapping_count += 1
            else:
                page1_df.at[i, 'TCS Group'] = "Not Found"
                
        print(f"Mapped {tcs_mapping_count} TCS Groups from Assignment groups")
        print(f"Set {len(page1_df) - tcs_mapping_count} TCS Groups to 'Not Found'")
    else:
        print("Required columns for TCS mapping not found in reference data")
        page1_df['TCS Group'] = "Not Found"
except Exception as e:
    print(f"Error mapping TCS Groups: {e}")
    page1_df['TCS Group'] = "Not Found"

# Step 10: Save the updated sheet to a new file
print("\nSaving updated data to file...")
try:
    # Generate output filename based on input
    file_dir = os.path.dirname(page1_file_path)
    file_name = os.path.basename(page1_file_path)
    output_file_name = os.path.splitext(file_name)[0] + "_with_advanced_mapping.xlsx"
    output_path = os.path.join(file_dir, output_file_name)
    
    # Save to Excel
    page1_df.to_excel(output_path, sheet_name="Page 1", index=False)
    print(f"Successfully saved mapped data to: {output_path}")
except Exception as e:
    print(f"Error saving output file: {e}")
    # Try fallback save location
    try:
        fallback_path = f"Page1_mapped_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        page1_df.to_excel(fallback_path, sheet_name="Page 1", index=False)
        print(f"Used fallback save location: {fallback_path}")
    except:
        print("Could not save file. Please check permissions and disk space.")

# Final statistics
print("\n-----------------------------------------------")
print("Summary of Application Name Mapping:")
print(f"  - Exact matches: {exact_match_count}")
print(f"  - Substring matches: {substring_match_count}")
print(f"  - Semantic matches: {semantic_match_count}")
print(f"  - Not Available entries: {remaining_count}")
print(f"  - Total rows processed: {original_row_count}")

# Calculate and print total runtime
end_time = time.time()
total_time = end_time - start_time
print(f"\nTotal execution time: {total_time:.2f} seconds")
print("Advanced mapping process complete!")
