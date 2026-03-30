#!/usr/bin/env python
# coding: utf-8

import pandas as pd

# === Step 1: Load Excel File ===
file_path = r"C:\Users\PIA - Laptop 080\Downloads\AES cumulative raw data - Jan-Feb 2026.xlsx"
df = pd.read_excel(file_path)

# === Step 2: Clean column names ===
df.columns = df.columns.str.strip()

# === Step 3: Ensure Region column exists ===
possible_region_cols = [
    "Location of the activity/Lokasyon ng aktibidad",
    "Region",
    "region"
]

region_col = None
for col in possible_region_cols:
    if col in df.columns:
        region_col = col
        break

if region_col is None:
    raise Exception(f"No Region column found. Available columns: {df.columns.tolist()}")

df = df.rename(columns={region_col: "Region"})

# === Step 4: Define survey items ===
survey_items = [
    "The activity was well-organized. (Maayos ang pagkaka-organisa ng aktibidad.)",
    "The activity was relevant to me. (Mahalaga sa akin ang aktibidad.)",
    "The activity was relevant to my community. (Mahalaga sa aking komunidad ang aktibidad.)",
    "The objectives of the activity were met. (Natupad ang mga layunin ng aktibidad.)",
    "I gained knowledge and understanding. (Nadagdagan ang aking kaalaman at pag-unawa.)",
    "I want to learn more about the topic. (Gusto ko pang matutunan ang tungkol sa paksa.)",
    "I became more aware of PIA. (Mas nakilala ko ang PIA.)"
]

# === Step 5: Keep only needed columns ===
needed_cols = [
    "Activity/Event Name (Pangalan ng Aktibidad)",
    "Date of the Activity (Petsa ng Aktibidad)",
    "Type of Activity (Uri ng aktibidad)",
    "Region",
    "Province",
    "How could we further improve the activity? (Paano pa namin mapapabuti ang aktibidad?)"
] + survey_items

survey_df = df[[col for col in needed_cols if col in df.columns]]

# === Step 6: Summaries ===
def summarize_region(region_df):
    summary = []
    for item in survey_items:
        if item not in region_df.columns:
            continue

        counts = region_df[item].value_counts().reindex([1,2,3,4,5], fill_value=0)
        total_responses = counts.sum()
        agree_percent = ((counts[4] + counts[5]) / total_responses * 100) if total_responses > 0 else 0

        summary.append([
            item,
            counts[1],
            counts[2],
            counts[3],
            counts[4],
            counts[5],
            round(agree_percent, 2)
        ])

    return pd.DataFrame(summary, columns=[
        "Statement",
        "Strongly Disagree (1)",
        "Disagree (2)",
        "Neither (3)",
        "Agree (4)",
        "Strongly Agree (5)",
        "% Agree + Strongly Agree"
    ])

# === OVERALL SUMMARY ===
def overall_summary(full_df):
    summary = []
    for item in survey_items:
        if item not in full_df.columns:
            continue

        counts = full_df[item].value_counts().reindex([1,2,3,4,5], fill_value=0)
        total_responses = counts.sum()
        agree_percent = ((counts[4] + counts[5]) / total_responses * 100) if total_responses > 0 else 0

        summary.append([
            item,
            counts[1],
            counts[2],
            counts[3],
            counts[4],
            counts[5],
            round(agree_percent, 2)
        ])

    return pd.DataFrame(summary, columns=[
        "Statement",
        "Strongly Disagree (1)",
        "Disagree (2)",
        "Neither (3)",
        "Agree (4)",
        "Strongly Agree (5)",
        "% Agree + Strongly Agree"
    ])

# === Activity tally ===
def tally_activities(region_df):
    respondent_counts = region_df.groupby(
        [
            "Activity/Event Name (Pangalan ng Aktibidad)",
            "Type of Activity (Uri ng aktibidad)",
            "Province",
            "Date of the Activity (Petsa ng Aktibidad)"
        ]
    )["The activity was well-organized. (Maayos ang pagkaka-organisa ng aktibidad.)"].count().reset_index()

    respondent_counts = respondent_counts.rename(
        columns={"The activity was well-organized. (Maayos ang pagkaka-organisa ng aktibidad.)": "Number of Respondents"}
    )

    total = pd.DataFrame([{
        "Activity/Event Name (Pangalan ng Aktibidad)": "TOTAL",
        "Type of Activity (Uri ng aktibidad)": "",
        "Province": "",
        "Date of the Activity (Petsa ng Aktibidad)": "",
        "Number of Respondents": respondent_counts["Number of Respondents"].sum()
    }])

    return pd.concat([respondent_counts, total], ignore_index=True)

# === Suggestions ===
def improvement_suggestions(region_df):
    col_name = "How could we further improve the activity? (Paano pa namin mapapabuti ang aktibidad?)"

    if col_name not in region_df.columns:
        return pd.DataFrame()

    comments = region_df[region_df[col_name].notna()].copy()

    comments = comments.groupby(
        ["Activity/Event Name (Pangalan ng Aktibidad)",
         "Date of the Activity (Petsa ng Aktibidad)",
         "Province"]
    )[col_name].apply(
        lambda x: "• " + "\n• ".join(x.astype(str).str.strip())
    ).reset_index()

    comments = comments.rename(columns={
        col_name: "Consolidated Comments/Suggestions"
    })

    return comments

# === Step 9: Write Excel output ===
output_path = "AES_2026_Output.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

    if "Region" not in survey_df.columns:
        raise Exception("Region column missing after cleaning.")

    for region in survey_df["Region"].dropna().unique():
        region_df = survey_df[survey_df["Region"] == region]

        if region_df.empty:
            continue

        sheet_name = str(region)[:31]

        survey_summary = summarize_region(region_df)
        activity_tally = tally_activities(region_df)
        suggestions_table = improvement_suggestions(region_df)

        survey_summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
        activity_tally.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(survey_summary) + 3)
        suggestions_table.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(survey_summary) + len(activity_tally) + 6)

    overall = overall_summary(survey_df)
    overall.to_excel(writer, sheet_name="Overall Summary", index=False)

print("Excel report successfully generated:", output_path)


