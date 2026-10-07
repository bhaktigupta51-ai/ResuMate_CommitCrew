import pandas as pd

df = pd.read_excel("JDS Skill Traits.xlsx")

df["data_analytics_score"] = (
    df["big_data_skills"] + df["maths-stats_skills"]
) / 2

df["technical_ai_score"] = (
    df["coding_skills"] + df["ai_and_ml_skills"]
) / 2

df["visualization_score"] = df["dashboard_and_storytelling_skills"]

skill_columns = [
    "big_data_skills",
    "maths-stats_skills",
    "coding_skills",
    "ai_and_ml_skills",
    "dashboard_and_storytelling_skills"
]

df["overall_skill_score"] = df[skill_columns].mean(axis=1)

def skill_band(score):
    if score < 3:
        return "Beginner"
    elif score < 4:
        return "Developing"
    elif score < 4.5:
        return "Proficient"
    else:
        return "Advanced"

df["overall_skill_band"] = df["overall_skill_score"].apply(skill_band)
df["data_analytics_band"] = df["data_analytics_score"].apply(skill_band)
df["technical_ai_band"] = df["technical_ai_score"].apply(skill_band)
df["visualization_band"] = df["visualization_score"].apply(skill_band)

df["high_coding_flag"] = (df["coding_skills"] >= 4).astype(int)
df["high_ai_ml_flag"] = (df["ai_and_ml_skills"] >= 4).astype(int)
df["high_data_analytics_flag"] = (df["data_analytics_score"] >= 4).astype(int)
df["high_visualization_flag"] = (df["visualization_score"] >= 4).astype(int)
df["overall_strong_skill_flag"] = (df["overall_skill_score"] >= 4).astype(int)

df["weakest_skill"] = df[skill_columns].idxmin(axis=1)

skill_name_mapping = {
    "big_data_skills": "Big Data",
    "maths-stats_skills": "Maths & Statistics",
    "coding_skills": "Coding",
    "ai_and_ml_skills": "AI & ML",
    "dashboard_and_storytelling_skills": "Dashboard & Storytelling"
}

df["weakest_skill"] = df["weakest_skill"].map(skill_name_mapping)

df["skill_gap_flag"] = (df[skill_columns].min(axis=1) < 3).astype(int)

df.to_excel("JDS_Skill_Traits_Analysed.xlsx", index=False)

print("Analysis completed successfully!")
print(df.head())