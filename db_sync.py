import pandas as pd
import re
 
df_master = pd.read_excel(r"C:\Users\AdityaGarg\Dump\RK Diagnostic db.xlsx")
df_lab = pd.read_excel(r"C:\Users\AdityaGarg\Dump\R.K. Diagnostic centre ratelist.xlsx")

df_master.columns = df_master.columns.str.strip()

matched_df = pd.DataFrame()
unmatched_df = pd.DataFrame()

master_unique = df_master.drop_duplicates(subset=["Test"]).copy()

# --- Step 1: Direct match on test_name_lab with test ---
merged = pd.merge(
    df_lab,
    master_unique[['Test', 'Master_Test_Name']],
    left_on='Test Name',
    right_on='Test',
    how='left'
)

# Create helper column for matched test (direct)
merged["Matched_Test"] = merged["Test Name"]

# Fill Master Test Name where matched directly
merged["Mapped_Master_Test_Name"] = merged["Master_Test_Name"]

Final_Matched = merged


def clean_text_series(series: pd.Series) -> pd.Series:

    series = series.str.lower()

    # Remove "outlab", "inlab", "center"
    series = series.str.replace(r'\b(outlab|inlab|\(outlab\)|\(inlab\)|center)\b\s*$', '', regex=True)

    # Replace time variations with "hours"
    series = series.str.replace(r'\b(hr|hrs|hour|hr\.|hrs\.)\b', 'hours', regex=True)

    # Remove unwanted patterns
    series = series.str.replace(r'\(sr\)|\(l\)|\(m\)|\(*m\)', '', regex=True)
    series = series.str.replace(r'\bantibody\b', 'antibodies', regex=True)
    series = series.str.replace(r'\b&\b', 'and', regex=True)
    series = series.str.replace(r'\bserum\b', '', regex=True)
    series = series.str.replace(r'\bfor\b', '', regex=True)
    series = series.str.replace(r'\bby\b', '', regex=True)
    series = series.str.replace(r'\bwith\b', '', regex=True)

    # Replace special characters with space
    series = series.str.replace(r'[-–—_;/,()\r\n]', ' ', regex=True)

    # Collapse multiple spaces
    series = series.str.replace(r'\s+', ' ', regex=True).str.strip()

    # Remove non-alphanumeric
    series = series.str.replace(r'[^a-zA-Z0-9]', '', regex=True)

    return series

df_master["Cleaned_Test"] = clean_text_series(df_master["Test"])

df_master["Cleaned_Test_name"] = clean_text_series(df_master["Master_Test_Name"])

df_master["Cleaned_lab_Test"] = clean_text_series(df_lab["Test Name"])

df_master.columns = df_master.columns.str.strip()

if merged['Mapped_Master_Test_Name'].isnull().sum() != 0:

    if "Test" in df_master.columns:
        df_master["Cleaned_Test"] = clean_text_series(df_master["Test"])
    else:
        raise KeyError("df_master is missing the required column 'Test'")

    if "Master_Test_Name" in df_master.columns:
        df_master["Cleaned_Test_name"] = clean_text_series(df_master["Master_Test_Name"])
    else:
        raise KeyError("df_master is missing the required column 'Master_Test_Name'")


    # Create dict from master_cleaned_unique for fast lookup
    master_cleaned_unique = df_master.drop_duplicates(subset=["Cleaned_Test"])[
    ["Test", "Master_Test_Name", "Cleaned_Test", "Cleaned_Test_name"]]

    # Step 2: Map using Cleaned_Test
    mask = merged["Mapped_Master_Test_Name"].isna()
    
    # Step 3: Clean test names for those rows
    merged.loc[mask, "Cleaned_lab_Test"] = clean_text_series(merged.loc[mask, "Test Name"])
    
    # Step 4a: Build dictionary for Cleaned_Test_name -> Master_Test_Name
    map_test = dict(zip(master_cleaned_unique["Cleaned_Test"],
                        master_cleaned_unique["Test"]))
    map_master = dict(zip(master_cleaned_unique["Cleaned_Test"],
                        master_cleaned_unique["Master_Test_Name"]))
    
    # Step 4b: Map Cleaned_Test_name2 to Mapped_Master_Test_Name
    merged.loc[mask, "Matched_Test"] = merged.loc[mask, "Cleaned_lab_Test"].map(map_test)
    merged.loc[mask, "Mapped_Master_Test_Name"] = merged.loc[mask, "Cleaned_lab_Test"].map(map_master)  
    
    # Final mapped rows
    mapped_only_df1 = merged[merged["Mapped_Master_Test_Name"].notna()].copy()
    
    # Step 3: Fallback mapping using Cleaned_Test_name
    mask2 = merged["Mapped_Master_Test_Name"].isna()
    if mask2.any():
        # Create dicts for Cleaned_Test lookups
        map_test = dict(zip(master_cleaned_unique["Cleaned_Test_name"],
                        master_cleaned_unique["Test"]))
        map_master = dict(zip(master_cleaned_unique["Cleaned_Test_name"],
                        master_cleaned_unique["Master_Test_Name"]))

        merged.loc[mask2, "Matched_Test"] = merged.loc[mask2, "Cleaned_lab_Test"].map(map_test)
        merged.loc[mask2, "Mapped_Master_Test_Name"] = merged.loc[mask2, "Cleaned_lab_Test"].map(map_master) 
    
        # Final mapped rows
        mapped_only_df2 = merged[merged["Mapped_Master_Test_Name"].notna()].copy()
    
        # Remaining unmatched for fallback/fuzzy
        need_fallback3 = merged[merged["Mapped_Master_Test_Name"].isna()].copy()


            
        if merged['Mapped_Master_Test_Name'].isnull().sum() != 0:
            
            from rapidfuzz import process, fuzz
            import numpy as np
                       
            # Lab file subset
            Lab_File_Fuzzy = need_fallback3.copy()
            
            master_clean_names = master_cleaned_unique["Cleaned_Test"].tolist()
            master_clean_names_name = master_cleaned_unique["Cleaned_Test_name"].tolist()
            
            cutoff = 70

            # Step 1: Match against Cleaned_Test
            scores = process.cdist(
                Lab_File_Fuzzy["Cleaned_lab_Test"],
                master_clean_names,
                scorer=fuzz.token_sort_ratio,
                workers=-1
            )
            
            best_idx = scores.argmax(axis=1)
            best_scores = scores.max(axis=1)
            
            best_matches = [
                master_clean_names[i] if score >= cutoff else None
                for i, score in zip(best_idx, best_scores)
            ]
            
            Lab_File_Fuzzy["Best Match Cleaned"] = best_matches
            Lab_File_Fuzzy["Matched Percentage"] = best_scores
            
            # Deduplicated lookup Series (avoids InvalidIndexError)
            map_test = (
                master_cleaned_unique.drop_duplicates("Cleaned_Test")
                .set_index("Cleaned_Test")["Test"]
            )
            map_master = (
                master_cleaned_unique.drop_duplicates("Cleaned_Test")
                .set_index("Cleaned_Test")["Master_Test_Name"]
            )
            Lab_File_Fuzzy["Matched_Test"] = Lab_File_Fuzzy["Best Match Cleaned"].map(map_test)
            Lab_File_Fuzzy["Mapped_Master_Test_Name"] = Lab_File_Fuzzy["Best Match Cleaned"].map(map_master)

            
            # Step 2: For unmatched rows, match against Cleaned_Test_name
            mask_unmatched = Lab_File_Fuzzy["Best Match Cleaned"].isna()
            if mask_unmatched.any():
                scores2 = process.cdist(
                    Lab_File_Fuzzy.loc[mask_unmatched, "Cleaned_lab_Test"],
                    master_clean_names_name,
                    scorer=fuzz.token_sort_ratio,
                    workers=-1
                )
                best_idx2 = scores2.argmax(axis=1)
                best_scores2 = scores2.max(axis=1)
            
                best_matches2 = [
                    master_clean_names_name[i] if score >= cutoff else None
                    for i, score in zip(best_idx2, best_scores2)
                ]
            
                Lab_File_Fuzzy.loc[mask_unmatched, "Best Match Cleaned"] = best_matches2
                Lab_File_Fuzzy.loc[mask_unmatched, "Matched Percentage"] = best_scores2
            
                map_test2 = (
                master_cleaned_unique.drop_duplicates("Cleaned_Test_name")
                    .set_index("Cleaned_Test_name")["Test"]
                )
                map_master2 = (
                    master_cleaned_unique.drop_duplicates("Cleaned_Test_name")
                    .set_index("Cleaned_Test_name")["Master_Test_Name"]
                )
                
                Lab_File_Fuzzy.loc[mask_unmatched, "Matched_Test"] = (
                    Lab_File_Fuzzy.loc[mask_unmatched, "Best Match Cleaned"].map(map_test2)
                )
                Lab_File_Fuzzy.loc[mask_unmatched, "Mapped_Master_Test_Name"] = (
                    Lab_File_Fuzzy.loc[mask_unmatched, "Best Match Cleaned"].map(map_master2)
                )
                            
            # Separate matched vs unmatched
            matched_mask = Lab_File_Fuzzy["Mapped_Master_Test_Name"].notna()
            matched_df = Lab_File_Fuzzy.loc[matched_mask].reset_index(drop=True)
            unmatched_df = Lab_File_Fuzzy.loc[~matched_mask].reset_index(drop=True)
            Final_Matched = pd.concat([mapped_only_df1, mapped_only_df2], ignore_index=True)
            
        else:
            Final_Matched = mapped_only_df2

    else:
        Final_Matched = mapped_only_df1

else:
    Final_Matched = merged

# Save to Excel with matched and unmatched

Final_Match = Final_Matched.drop_duplicates(subset=['Test Name'], keep='first')
with pd.ExcelWriter(r"C:\Users\AdityaGarg\Dump\Output\output.xlsx", engine="openpyxl") as writer:
    Final_Match.to_excel(writer, sheet_name="Matched", index=False)
    matched_df.to_excel(writer, sheet_name="Fuzzy Match", index=False)
    unmatched_df.to_excel(writer, sheet_name="Unmatched", index=False)