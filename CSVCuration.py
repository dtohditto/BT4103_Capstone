import pandas as pd
import numpy as np
from sentence_transformers import SentenceTransformer
from umap import UMAP
import hdbscan
from bertopic import BERTopic
import seaborn as sns
from collections import Counter
import re
import os


from typing import Optional, Tuple, Union

def curate_programme_and_cost_data(programme_file, cost_file, output_path: Optional[str] = None, return_csv_bytes: bool = False) -> Union[pd.DataFrame, bytes, Tuple[pd.DataFrame, bytes]]:
    """
    Curate programme and cost data.

    Parameters
    - programme_file, cost_file: file-like or path accepted by pandas.read_excel
    - output_path: optional path to write the resulting CSV to disk
    - return_csv_bytes: if True, return CSV data as bytes (utf-8-sig) in addition to or instead of the DataFrame

    Return
    - By default returns the curated DataFrame.
    - If return_csv_bytes=True and output_path is None, returns bytes.
    - If return_csv_bytes=True and output_path is provided, returns (DataFrame, bytes).
    """

    if programme_file is None or cost_file is None:
        return None

    os.environ["TOKENIZERS_PARALLELISM"] = "false" 

    file_path_programme = 'Capstone_Project_AnonyVBA_V4A.xlsm'
    file_path_cost = 'Capstone Project 2025 - Programme_V4C.xlsx'

    # Use the 'header' parameter to specify the row number (0-based index) to use as the header.`
    df_programme = pd.read_excel(programme_file, header=12)
    df_cost = pd.read_excel(cost_file, header=0)

    columns_to_drop_programme = df_programme.columns[:3] # Drop the first 3 columns
    rows_to_drop_programme = df_programme.index[31981:] # Drop rows showing total

    df_programme = df_programme.drop(columns=columns_to_drop_programme).drop(index=rows_to_drop_programme)

    # Convert data types for df_programme (assuming similar types as df_summary where applicable)
    df_programme['Application ID'] = df_programme['Application ID'].astype(str)
    df_programme['Application Status'] = df_programme['Application Status'].astype('category')
    df_programme['Applicant Type'] = df_programme['Applicant Type'].astype('category')
    df_programme['Contact ID'] = df_programme['Contact ID'].astype(str)
    df_programme['Organisation Name: Organisation Name'] = df_programme['Organisation Name: Organisation Name'].astype(str).str.upper()
    df_programme['Job Title'] = df_programme['Job Title'].astype(str).str.upper()
    df_programme['Truncated Programme Name'] = df_programme['Truncated Programme Name'].astype(str)
    df_programme['Truncated Programme Run'] = df_programme['Truncated Programme Run'].astype(str)
    df_programme['Primary Category'] = df_programme['Primary Category'].astype('category')
    df_programme['Secondary Category'] = df_programme['Secondary Category'].astype('category')
    df_programme['Programme Start Date'] = pd.to_datetime(df_programme['Programme Start Date'], errors='coerce', format='%d/%m/%Y')
    df_programme['Programme End Date'] = pd.to_datetime(df_programme['Programme End Date'], errors='coerce', format='%d/%m/%Y')
    df_programme['How You Learnt About This Programme'] = df_programme['How You Learnt About This Programme'].astype('category').str.upper()
    # Fill missing values with 0 before converting to integer (Age might have non-finite values)
    # df_programme['Age'] = df_programme['Age'].fillna(0).astype(int)
    df_programme['Gender'] = df_programme['Gender'].astype('category')
    df_programme['Country Of Residence'] = df_programme['Country Of Residence'].astype('category')
    df_programme['Educational_Qualification'] = df_programme['Educational_Qualification'].astype('category')

    # ---- Age Cleaning ----
    # Rule: Age <= 15 or Age > 100 → set to NaN
    df_programme['Age'] = df_programme['Age'].apply(
        lambda x: np.nan if pd.notnull(x) and (x <= 15 or x > 100) else x
    )

    # Define mapping to reduce number of categories in Application Status
    pending = [
        "Pending", "Provisionally Approved", "HR Verification", "Waiting List"
    ]

    attended = [
        "Approved", "Accepted", "Admitted", "Pass", "Attended", "Makeup", "AD - Admit (same as Offer Made)"
    ]

    withdrawn = [
        "Cancelled", "Rejected", "DE - Deny (same as Rejected)", "DC - Decline (Offer Refused) (same as Offer Declined)",
        "WI - Withdraw (same as Withdrawn)", "Postponed"
    ]

    failed = [
        "Failed", "Fail", "Failed - Not awarded e-certificate"
    ]

    # Build mapping dictionary
    status_mapping = {}
    status_mapping.update(dict.fromkeys(pending, "Pending"))
    status_mapping.update(dict.fromkeys(attended, "Attended"))
    status_mapping.update(dict.fromkeys(withdrawn, "Withdrawn"))
    status_mapping.update(dict.fromkeys(failed, "Failed"))

    # Apply mapping (Edit to modify csv instead of changing df)
    df_programme["Application Status"] = (
        df_programme["Application Status"]
        .map(status_mapping)
        .fillna("Unknown")
    )

    df_programme["Secondary Category"] = df_programme["Secondary Category"].astype(str)
    df_programme["Secondary Category"] = df_programme["Secondary Category"].replace("nO", "Non-AI")
    df_programme["Secondary Category"] = df_programme["Secondary Category"].replace("0", "Non-AI")
    df_programme["Secondary Category"] = df_programme["Secondary Category"].astype("category")


    # Job Title cleanup + Seniority (cleaner "Top X Titles", more insight)
    raw = df_programme["Job Title"].astype("string")

    # 1) Trim + normalize obvious placeholders to NA (case-insensitive)
    # covers: "", "-", "N/A", "N. A", "NA", "NIL", "NONE", "NULL", "NaN", em-dash
    placeholder_pat = re.compile(r"^\s*(?:N/?A|N\.?\s*A\.?|NIL|NONE|NULL|NAN|—|-)?\s*$", re.IGNORECASE)
    raw = raw.where(~raw.fillna("").str.fullmatch(placeholder_pat), other=pd.NA)

    # 2) Uppercase and strip non-meaningful chars
    clean = raw.str.strip().str.upper()
    clean = clean.str.replace(r"[^A-Z0-9\s/&\-]", "", regex=True)

    # 3) If cleaning wiped the string, treat as NA
    clean = clean.replace("", pd.NA)

    # 4) Final coalesce: NA -> "UNKNOWN"
    df_programme["Job Title Clean"] = clean.fillna("UNKNOWN")

    def seniority_from_title(t):
        if t == "UNKNOWN": return "UNKNOWN"
        if any(k in t for k in ["CEO","CHIEF","CFO","COO","PRESIDENT","VICE CHAIR","CTO","CIO","CDO","CMO"]):
            return "C-LEVEL"
        if any(k in t for k in ["EVP","SVP","VICE PRESIDENT","VP","DIRECTOR","HEAD","GENERAL MANAGER","GM"]):
            return "DIRECTOR/VP"
        if any(k in t for k in ["ASSISTANT MANAGER","ASST MANAGER","DEPUTY MANAGER"]):
            return "MANAGER"
        if "MANAGER" in t:
            return "MANAGER"
        if any(k in t for k in ["LEAD","SENIOR","SR ","SR."]):
            return "SENIOR/LEAD"
        return "INDIVIDUAL CONTRIBUTOR"

    df_programme["Seniority"] = df_programme["Job Title Clean"].apply(seniority_from_title)


    #Run_Month
    def parse_run_month(s):
        if pd.isna(s): return pd.NaT
        s = str(s).strip()
        for fmt in ("%b-%Y", "%b %Y", "%b-%y"):
            try:
                return pd.to_datetime(s, format=fmt).to_period("M").to_timestamp()
            except Exception:
                pass
        return pd.NaT

    run_month_from_str = df_programme["Truncated Programme Run"].map(parse_run_month) \
        if "Truncated Programme Run" in df_programme.columns else pd.Series(pd.NaT, index=df_programme.index)

    df_programme["Run_Month"] = run_month_from_str.fillna(
        pd.to_datetime(df_programme["Programme Start Date"], errors="coerce")
    ).dt.to_period("M").dt.to_timestamp()

    # Optional pretty label for charts / csv
    df_programme["Run_Month_Label"] = df_programme["Run_Month"].dt.strftime("%b-%Y")


    # Country strings: light cleanup
    df_programme["Country Of Residence"] = (
        df_programme["Country Of Residence"].astype("string").str.strip()
    )
    country_map = {"Hong Kong SAR": "Hong Kong"}
    df_programme["Country Of Residence"] = df_programme["Country Of Residence"].replace(country_map)


    # Normalise organisation names
    df_programme["Organisation Name: Organisation Name"] = (
        df_programme["Organisation Name: Organisation Name"]
        .fillna("Unknown")
        .astype("string")
        .str.strip()
        .str.upper()
        .replace({
            "N.A.": "UNKNOWN",
            "N. A": "UNKNOWN",
            "NIL": "UNKNOWN",
            "NA": "UNKNOWN"
        })
    )

    # Gender
    df_programme["Gender"] = (
        df_programme["Gender"].astype("string")
        .fillna("Unknown")
        .str.strip()
        .str.capitalize()
    )
    df_programme["Gender"] = df_programme["Gender"].astype("category")


    # Make final types dashboard-friendly (categories for slicers)
    for col in ["Application Status","Applicant Type","Primary Category","Secondary Category","Seniority"]:
        if col in df_programme.columns:
            df_programme[col] = df_programme[col].astype("category")

    cols_for_dashboard = [
        "Application ID","Contact ID","Application Status","Applicant Type",
        "Organisation Name: Organisation Name","Job Title Clean","Seniority",
        "Truncated Programme Name","Truncated Programme Run","Primary Category","Secondary Category",
        "Programme Start Date","Programme End Date","Run_Month","Run_Month_Label","Programme_Duration",
        "Gender","Age","Country Of Residence"
    ]

    curated = df_programme[[c for c in cols_for_dashboard if c in df_programme.columns]].copy()
    
    # Has to go after seniority
    def preprocess_job_titles(df: pd.DataFrame, title_col='Job Title Clean', seniority_col='Seniority') -> Tuple[pd.DataFrame, pd.DataFrame]:
      """
      Preprocess job titles for BERTopic clustering.
      Returns a tuple: (cleaned_df_for_clustering, dropped_df)
      All rows that are generic or turned 'UNKNOWN' are returned as dropped_df.
      """
      df = df.copy()
      
      # Track rows that will eventually be dropped
      df['_drop_flag'] = False

      # 1. Replace non-letter titles with "UNKNOWN"
      mask_non_letters = ~df[title_col].str.contains(r'[A-Za-z]', na=False)
      df.loc[mask_non_letters, [title_col, seniority_col]] = "UNKNOWN"
      df.loc[mask_non_letters, '_drop_flag'] = True

      # 2. Replace code-like titles (digits but no spaces) with "UNKNOWN"
      code_mask = df[title_col].str.contains(r'\d', na=False) & ~df[title_col].str.contains(r'\s', na=False)
      df.loc[code_mask, [title_col, seniority_col]] = "UNKNOWN"
      df.loc[code_mask, '_drop_flag'] = True

      # 3. Replace single-letter titles with "UNKNOWN"
      single_letter_mask = df[title_col].str.strip().str.len() == 1
      df.loc[single_letter_mask, [title_col, seniority_col]] = "UNKNOWN"
      df.loc[single_letter_mask, '_drop_flag'] = True

      # 4. Replace short titles (<3 letters) not in meaningful list
      meaningful_titles = ["CEO", "VP", "AVP", "CFO", "COO", "GM", "MD", "SVP", "CTO"]
      short_mask = (df[title_col].str.len() < 3) & (~df[title_col].isin(meaningful_titles))
      df.loc[short_mask, [title_col, seniority_col]] = "UNKNOWN"
      df.loc[short_mask, '_drop_flag'] = True

      # 5. Track and drop fully generic titles
      generic_terms = [
          "UNKNOWN","CEO","CHIEF","CFO","COO","PRESIDENT","VICE CHAIR","CTO","CIO","CDO","CMO",
          "EVP","SVP","VICE PRESIDENT","VP","DIRECTOR","HEAD","GENERAL MANAGER","GM","AVP","MD",
          "ASSISTANT MANAGER","ASST MANAGER","DEPUTY MANAGER",
          "LEAD","SENIOR","SR ","SR.","MANAGER","COMMANDER","EXEC","MGR",
          "STAFF","PRINCIPAL","ASSOCIATE",
          "DEPUTY","VICE","ASSISTANT","ASST","EXECUTIVE","SECRETARY","CHAIRMAN","PERSONNEL",
          "MANAGING","OFFICER","OWNER","CHEIF","PARTNER","FOUNDER","COFOUNDER",
          "REGIONAL","APAC","SINGAPORE","REGION","ASEAN","ASIA","PACIFIC","GLOBAL","CHINA","GROUP","BRANCH","LOCAL",
          "COUNTRY","DIVISION","SECTION","AREA",
          "GENERAL","SERVICES","PROFESSIONAL","STRATEGIC"
      ]
      generic_mask = df[title_col].isin(generic_terms)
      df.loc[generic_mask, '_drop_flag'] = True

      # Split into cleaned and dropped rows 
      dropped_rows = df[df['_drop_flag']].copy()
      cleaned_df = df[~df['_drop_flag']].copy()

      # Drop the temporary flag
      cleaned_df.drop(columns=['_drop_flag'], inplace=True)
      dropped_rows.drop(columns=['_drop_flag'], inplace=True)

      return cleaned_df.reset_index(drop=True), dropped_rows.reset_index(drop=True)

    curated_cleaned, dropped_titles = preprocess_job_titles(curated)

    def preprocess_title(title):
        if pd.isna(title):
            return ""
        title = str(title).lower()
        title = re.sub(r'[^\w\s]', ' ', title)
        title = re.sub(r'\b(?:of|and|for|the)\b', '', title)
        title = re.sub(r'\s+', ' ', title).strip()
        return title

    curated_cleaned['Job Title Clean'] = curated_cleaned['Job Title Clean'].apply(preprocess_title)

    # Domain tagging
    domain_keywords = {
        'Sales': ['sales', 'business development', 'account'],
        'Marketing': ['marketing', 'brand', 'communications', 'content'],
        'Finance': ['finance', 'accounting', 'audit', 'analyst','actuarial','actuary','risk'],
        'HR': ['hr', 'talent', 'people', 'customer service'],
        'Analytics': ['data', 'intelligence', 'analytics'],
        'Engineering': ['engineer', 'engineering', 'technical', 'developer', 'it', 'technology'],
        'Accounting/Audit': ['tax', 'audit', 'accounting', 'accountant'],
        'Operations': ['operations', 'project', 'logistics', 'admin', 'administrative'],
        'Legal': ['legal', 'counsel', 'compliance'],
        'Strategy': ['product', 'strategy'],
        'Education': ['lecturer', 'professor', 'tutor', 'teacher'],
        'Healthcare': ['nurse', 'health', 'healthcare', 'medical', 'allied'],
        'Advisory': ['client', 'advisory', 'advisor', 'relations', 'relationships']
    }

    def add_domain_keywords(title):
        title_lower = title.lower()
        added_domains = []
        for domain, keywords in domain_keywords.items():
            if any(k in title_lower for k in keywords):
                added_domains.append(domain)
        return title + " " + " ".join(added_domains) if added_domains else title

    curated_cleaned['Job Title Tagged'] = curated_cleaned['Job Title Clean'].apply(add_domain_keywords)
 
    model = SentenceTransformer('all-MiniLM-L6-v2')
    embeddings = model.encode(curated_cleaned['Job Title Tagged'].tolist(), show_progress_bar=False)
 
    umap_model = UMAP(n_neighbors=50, n_components=5, metric='cosine', random_state=42)
    hdbscan_model = hdbscan.HDBSCAN(min_cluster_size=50, min_samples=1, metric='euclidean', cluster_selection_method='eom')

    topic_model = BERTopic(umap_model=umap_model, hdbscan_model=hdbscan_model, embedding_model=model)
    topics, probs = topic_model.fit_transform(curated_cleaned['Job Title Tagged'].tolist(), embeddings=embeddings)
    topic_model.reduce_topics(curated_cleaned['Job Title Tagged'].tolist(), nr_topics=20)
 
    curated_cleaned['Topic_Label'] = topic_model.topics_
  
    dropped_titles['Domain'] = "Others"
    dropped_titles['Topic_Label'] = -1  # for consistency
    dropped_titles['Job Title Tagged'] = pd.NA

    curated = pd.concat([curated_cleaned, dropped_titles], ignore_index=True)

    # Some topics e.g. topic 16 will not be useful, so classify them under "Others"
    topic_mapping = {
        -1: "Others",
        0: "Engineering",
        1: "Finance",
        2: "Logistics/Supply Chain",
        3: "Sales",
        4: "Marketing",
        6: "HR",
        7: "Operations/Admin",
        8: "Education",
        9: "Strategy/Product",
        10: "Audit/Accounting",
        11: "Advisory/Client Relations",
        12: "Legal",
        13: "Analytics"
    }

    # Map topic numbers to names, fallback to "Others" for anything else
    curated['Domain'] = curated['Topic_Label'].apply(lambda x: topic_mapping.get(x, "Others"))

    curated.drop(columns=['Topic_Label', 'Job Title Tagged', 'Job Title Clean'], inplace=True)
    curated.rename(columns={'Organisation Name: Organisation Name': 'Organisation Name'}, inplace=True)

    df_cost.rename(columns={'Programme Name': 'Truncated Programme Name'}, inplace=True)

    df_cost.to_csv("cost.csv", index=False)
    curated.to_csv("programme_curated.csv", index=False)

    curated = curated.merge(df_cost, how='left', on='Truncated Programme Name')

    cost_zero = curated[curated['Programme Cost']==0]
    cost_zero = (
        cost_zero
        .groupby('Truncated Programme Name')['Run_Month_Label']
        .nunique()
        .reset_index(name='Unique_Run_Month_Count')
    )
    df_dashboard_filtered = curated[~curated['Truncated Programme Name'].isin(cost_zero['Truncated Programme Name'])]

    # Optionally write to disk
    if output_path:
        try:
            df_dashboard_filtered.to_csv(output_path, index=False)
        except Exception as e:
            print(f"Failed to write CSV to {output_path}: {e}")

    # Optionally return CSV bytes (useful for web UIs like Streamlit)
    if return_csv_bytes:
        csv_bytes = df_dashboard_filtered.to_csv(index=False).encode("utf-8-sig")
        if output_path:
            return df_dashboard_filtered, csv_bytes
        return csv_bytes

    return df_dashboard_filtered 
