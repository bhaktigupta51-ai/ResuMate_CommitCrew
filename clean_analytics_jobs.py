import pandas as pd
import re

df = pd.read_csv("Analytics Jobs.csv")

df = df.drop(columns=["s_no"])
df = df.drop_duplicates()

df["job_description"] = df["job_description"].fillna("")
df["job_desig"] = df["job_desig"].fillna("")
df["job_type"] = df["job_type"].fillna("")
df["key_skills"] = df["key_skills"].fillna("")
df["location"] = df["location"].fillna("")
df["salary"] = df["salary"].fillna("")
df["experience"] = df["experience"].fillna("")

text_columns = [
    "job_description",
    "job_desig",
    "job_type",
    "key_skills",
    "location",
    "salary",
    "experience"
]

for column in text_columns:
    df[column] = (
        df[column]
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )

df["job_description"] = (
    df["job_description"]
    .str.replace("&amp;", "&", regex=False)
    .str.replace("Ã©", "é", regex=False)
    .str.replace("â€™", "'", regex=False)
    .str.replace("â€“", "-", regex=False)
    .str.replace("â€”", "-", regex=False)
    .str.replace("â€œ", '"', regex=False)
    .str.replace("â€", '"', regex=False)
)

df["job_type"] = (
    df["job_type"]
    .str.lower()
    .replace({
        "analytics": "Analytics",
        "analytic": "Analytics",
        "": "Not Specified"
    })
)

df = df.drop(columns=["job_type"])

df["salary"] = (
    df["salary"]
    .str.lower()
    .str.replace(" ", "", regex=False)
    .str.replace("to", "-", regex=False)
)

salary_values = df["salary"].str.extract(
    r"(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)"
)

df["salary_min"] = pd.to_numeric(
    salary_values[0],
    errors="coerce"
)

df["salary_max"] = pd.to_numeric(
    salary_values[1],
    errors="coerce"
)

invalid_salary = df["salary_min"] > df["salary_max"]

df.loc[
    invalid_salary,
    ["salary_min", "salary_max"]
] = pd.NA

df = df.drop(columns=["salary"])

df["experience"] = (
    df["experience"]
    .str.lower()
    .str.replace("years", "", regex=False)
    .str.replace("year", "", regex=False)
    .str.replace("yrs", "", regex=False)
    .str.replace("yr", "", regex=False)
    .str.strip()
)

experience_values = df["experience"].str.extract(
    r"(\d+(?:\.\d+)?)\s*[-to]+\s*(\d+(?:\.\d+)?)"
)

df["exp_min"] = pd.to_numeric(
    experience_values[0],
    errors="coerce"
)

df["exp_max"] = pd.to_numeric(
    experience_values[1],
    errors="coerce"
)

invalid_experience = df["exp_min"] > df["exp_max"]

df.loc[
    invalid_experience,
    ["exp_min", "exp_max"]
] = pd.NA

df = df.drop(columns=["experience"])

df["location"] = (
    df["location"]
    .str.replace(r"\s*,\s*", ", ", regex=True)
    .str.strip()
)

location_replacements = {
    r"\bBangalore\b": "Bengaluru",
    r"\bGurgaon\b": "Gurugram",
    r"\bTrivandrum\b": "Thiruvananthapuram",
    r"\bMysore\b": "Mysuru",
    r"\bDelhi NCR\b": "Delhi",
    r"\bNCR\b": "Delhi"
}

for pattern, replacement in location_replacements.items():
    df["location"] = df["location"].str.replace(
        pattern,
        replacement,
        case=False,
        regex=True
    )

df["is_multi_location"] = df["location"].str.contains(
    ",",
    regex=False
)

df["location_count"] = (
    df["location"]
    .str.split(",")
    .str.len()
)

def clean_skill(skill):
    skill = skill.strip().lower()

    if "sql" in skill:
        return "SQL"
    elif "python" in skill:
        return "Python"
    elif "excel" in skill:
        return "Excel"
    elif "power bi" in skill or "powerbi" in skill:
        return "Power BI"
    elif "tableau" in skill:
        return "Tableau"
    elif "machine learning" in skill or skill == "ml":
        return "Machine Learning"
    elif "tensorflow" in skill:
        return "TensorFlow"
    elif "pytorch" in skill:
        return "PyTorch"

    return skill

df["key_skills"] = (
    df["key_skills"]
    .str.split(",")
    .apply(
        lambda skills: ", ".join(
            dict.fromkeys(
                clean_skill(skill)
                for skill in skills
                if skill.strip()
            )
        )
    )
)

df["skills_truncated"] = df["key_skills"].str.contains(
    r"\.\.\.$",
    regex=True
)

df["job_desig"] = (
    df["job_desig"]
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
)

df["job_role"] = "Other"

title = df["job_desig"].str.lower()

df.loc[
    title.str.contains(
        r"\bdata scientist\b|\bdata science\b",
        regex=True
    ),
    "job_role"
] = "Data Scientist"

df.loc[
    title.str.contains(
        r"\bdata analyst\b|\banalytics analyst\b|\banalyst\b",
        regex=True
    ),
    "job_role"
] = "Data Analyst"

df.loc[
    title.str.contains(
        r"\bbusiness analyst\b|\bbusiness intelligence\b|\bbi analyst\b",
        regex=True
    ),
    "job_role"
] = "Business Analyst"

df.loc[
    title.str.contains(
        r"\bmachine learning\b|\bml engineer\b",
        regex=True
    ),
    "job_role"
] = "Machine Learning"

df.loc[
    title.str.contains(
        r"\bdata engineer\b",
        regex=True
    ),
    "job_role"
] = "Data Engineer"

df.loc[
    title.str.contains(
        r"\bsoftware engineer\b|\bsoftware developer\b|\bapplication developer\b",
        regex=True
    ),
    "job_role"
] = "Software/Developer"

spam_patterns = (
    r"home base job|data entry|online work|"
    r"part time work|freelancer work"
)

df["is_spam_or_irrelevant"] = title.str.contains(
    spam_patterns,
    regex=True,
    na=False
)

df = df[~df["is_spam_or_irrelevant"]]

df = df.drop(columns=["is_spam_or_irrelevant"])

df = df.reset_index(drop=True)

print("Cleaned dataset shape:", df.shape)
print("\nMissing values:")
print(df.isna().sum())

print("\nDuplicate rows:")
print(df.duplicated().sum())

print("\nJob roles:")
print(df["job_role"].value_counts())

print("\nMulti-location jobs:")
print(df["is_multi_location"].value_counts())

df.to_csv(
    "Analytics_Jobs_Cleaned.csv",
    index=False
)

print("\nCleaned file saved as Analytics_Jobs_Cleaned.csv")

