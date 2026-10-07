import pandas as pd

df = pd.read_excel("SDS Personality Traits.xlsx")

df.columns = df.columns.str.strip()

personality_columns = [
    "neuroticism",
    "extraversion",
    "openness_to_experience",
    "agreeableness",
    "conscientiousness"
]

def personality_band(score):
    if score < 20:
        return "Low"
    elif score < 40:
        return "Moderate"
    else:
        return "High"

for col in personality_columns:
    df[col + "_band"] = df[col].apply(personality_band)

df["high_neuroticism_flag"] = (df["neuroticism"] >= 40).astype(int)
df["high_extraversion_flag"] = (df["extraversion"] >= 40).astype(int)
df["high_openness_flag"] = (df["openness_to_experience"] >= 40).astype(int)
df["high_agreeableness_flag"] = (df["agreeableness"] >= 40).astype(int)
df["high_conscientiousness_flag"] = (df["conscientiousness"] >= 40).astype(int)

df["overall_personality_score"] = df[personality_columns].mean(axis=1)

df["overall_personality_band"] = df["overall_personality_score"].apply(
    personality_band
)

df["dominant_personality"] = df[personality_columns].idxmax(axis=1)

personality_name_mapping = {
    "neuroticism": "Neuroticism",
    "extraversion": "Extraversion",
    "openness_to_experience": "Openness",
    "agreeableness": "Agreeableness",
    "conscientiousness": "Conscientiousness"
}

df["dominant_personality"] = df["dominant_personality"].map(
    personality_name_mapping
)

df["low_conscientiousness_flag"] = (
    df["conscientiousness"] < 20
).astype(int)

df["high_stability_flag"] = (
    df["neuroticism"] < 20
).astype(int)

df.to_excel("SDS_Personality_Traits_Analysed.xlsx", index=False)

print("Analysis completed successfully!")
print(df.head())