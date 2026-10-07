import pandas as pd
import numpy as np

df = pd.read_csv("DataScience Jobs.csv")

df = df.reset_index(drop=True)

df["job_id"] = range(1, len(df) + 1)

df["company"] = (
    df["company"]
    .astype(str)
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)

df["title"] = (
    df["title"]
    .astype(str)
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)

df["experience"] = (
    df["experience"]
    .astype(str)
    .str.strip()
    .str.lower()
    .str.replace(r"\s+", " ", regex=True)
)

df["avg_salary"] = (
    df["avg_salary"]
    .astype(str)
    .str.replace("L", "", regex=False)
    .str.strip()
    .astype(float)
)

df["min_salary"] = (
    df["min_salary"]
    .astype(str)
    .str.replace("L", "", regex=False)
    .str.strip()
    .astype(float)
)

df["max_salary"] = (
    df["max_salary"]
    .astype(str)
    .str.replace("L", "", regex=False)
    .str.strip()
    .astype(float)
)

df["num_of_jobs"] = pd.to_numeric(
    df["num_of_jobs"],
    errors="coerce"
)

df["experience_min"] = pd.to_numeric(
    df["experience"].str.extract(r"(\d+)")[0],
    errors="coerce"
)

df["experience_group"] = np.where(
    df["experience_min"] >= 5,
    "5+",
    df["experience_min"].astype("Int64").astype(str)
)

df["title_lower"] = df["title"].str.lower()

df["senior_title"] = df["title_lower"].str.contains(
    r"\bsenior\b|\bsr\.?\b|\blead\b|\bprincipal\b|\barchitect\b|\bmanager\b|\bdirector\b",
    regex=True,
    na=False
)

df["senior_low_experience"] = (
    df["senior_title"] &
    (df["experience_min"] <= 1)
)

df["non_senior_high_experience"] = (
    ~df["senior_title"] &
    (df["experience_min"] >= 8)
)

df["salary_band_ratio"] = (
    df["max_salary"] / df["min_salary"]
)

df["wide_salary_band"] = (
    df["salary_band_ratio"] > 5
)

df["num_of_jobs_log"] = np.log1p(
    df["num_of_jobs"]
)

df["avg_salary_log"] = np.log1p(
    df["avg_salary"]
)

df["company_normalized"] = (
    df["company"]
    .str.lower()
    .str.replace(r"[^\w\s]", "", regex=True)
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
)

df = df.drop(columns=["title_lower"])

df.to_csv(
    "DataScience_Jobs_Cleaned.csv",
    index=False
)

print("Cleaned dataset shape:", df.shape)

print("\nMissing values:")
print(df.isna().sum())

print("\nDuplicate rows:")
print(df.duplicated().sum())

print("\nSalary columns:")
print(df[[
    "min_salary",
    "avg_salary",
    "max_salary"
]].dtypes)

print("\nSenior roles with 0-1 years:")
print(df["senior_low_experience"].sum())

print("\nNon-senior roles with 8+ years:")
print(df["non_senior_high_experience"].sum())

print("\nWide salary bands:")
print(df["wide_salary_band"].sum())

print("\nCleaned file saved as DataScience_Jobs_Cleaned.csv")

