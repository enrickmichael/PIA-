#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import xlsxwriter

# === Load Data ===
df = pd.read_excel("PIACSMJanuary.xlsx")

# === Step 1: Clean column names ===
df.columns = df.columns.str.strip()

# === Step 2: Rename columns ===
rename_map = {
    "PIA Office Visited/Transacted With": "PIA Office",
    "Client Type ": "Client Type",
    "External Services ": "External Services",
    "Internal Service": "Internal Services",
    # CC
    "CC1. Which of the following best describes your awareness of a Citizen's Charter (CC)?": "CC1",
    "CC2. If aware of CC (answered 1-3 in CC1), would you say that the CC of this office was...?": "CC2",
    "CC3. If aware of CC (answered codes 1-3 in CC1), how much did the CC help you in your transaction?": "CC3",
    # SQDs
    "SQD0. I am satisfied with the service that I availed.": "SQD0",
    "SQD1. I spent a reasonable amount of time for my transaction.": "SQD1",
    "SQD2. The office followed the transaction's requirements and steps based on the information provided.": "SQD2",
    "SQD3. The steps (including payment) I needed to do for my transaction were easy and simple.": "SQD3",
    "SQD4. I easily found information about my transaction from the office or its website.": "SQD4",
    "SQD5. I paid a reasonable amount of fees for my transaction. (If service was free, mark â€˜N/A')": "SQD5",
    "SQD6. I feel the office was fair and my transaction was secure.": "SQD6",
    "SQD7. The office's support staff was courteous, helpful, and quick to respond.": "SQD7",
    "SQD8. I got what I needed from the government office, or (if denied) denial of request was sufficiently explained to me.": "SQD8",
    # PIA follow-ups
    "PIA1. Are you going to engage the service of PIA again?": "PIA1",
    "PIA2. Would you recommend PIA to another colleague or agency/organization?": "PIA2",
    "PIA3. How could we further improve our service?": "PIA3",
    "PIA3. Paano pa namin mapapabuti ang aming serbisyo?": "PIA3_alt",
}
df = df.rename(columns=rename_map)

# === Step 3: Replace responses (CC2 + CC3 only) ===
cc2_map = {
    "1. Madaling makita": "1. Easy to see",
    "2. Medyo madaling makita": "2. Somewhat easy to see",
    "3. Mahirap makita": "3. Difficult to see",
    "4. Hindi makita": "4. Not visible at all"
}
cc3_map = {
    "1. Sobrang nakatulong": "1. Helped very much",
    "2. Nakatulong naman": "2. Somewhat helped",
    "3. Hindi nakatulong": "3. Did not help"
}
df["CC2"] = df["CC2"].astype(str).str.strip().replace(cc2_map)
df["CC3"] = df["CC3"].astype(str).str.strip().replace(cc3_map)

# === Step 4: Replace SQD responses ===
sqd_map = {
    "Lubos na hindi sang-ayon": "Strongly Disagree",
    "Hindi sang-ayon": "Disagree",
    "Walang kinikilingan": "Neither Agree nor Disagree",
    "Sang-ayon": "Agree",
    "Labis na sang-ayon Sang-ayon": "Strongly Agree",
    "Hindi angkop (N/A)": "Not applicable (N/A)"
}
for col in [f"SQD{i}" for i in range(9)]:
    df[col] = df[col].astype(str).str.strip().replace(sqd_map)

# === Step 5: Define response orders ===
cc1_order = [
    "1. I know what a CC is, and I saw this office's CC.",
    "2. I know what a CC is, but I did NOT see this office's CC.",
    "3. I learned of the CC only when I saw this office's CC.",
    "4. I do not know what a CC is, and I did not see one in this office. (Answer 'N/A' in CC2 and CC3)"
]
cc2_order = ["1. Easy to see","2. Somewhat easy to see","3. Difficult to see","4. Not visible at all","5. N/A"]
cc3_order = ["1. Helped very much","2. Somewhat helped","3. Did not help","4. N/A"]
resp_order = ["Strongly Agree","Agree","Neither Agree nor Disagree","Disagree","Strongly Disagree","Not applicable (N/A)"]

sqd_labels = {
    "SQD1": "SQD1 - Responsiveness","SQD2": "SQD2 - Reliability","SQD3": "SQD3 - Access",
    "SQD4": "SQD4 - Communication","SQD5": "SQD5 - Cost","SQD6": "SQD6 - Integrity",
    "SQD7": "SQD7 - Assurance","SQD8": "SQD8 - Outcome","SQD0": "Overall"
}
cc_labels = {
    "CC1": "CC1. Which of the following best describes your awareness of a Citizen’s Charter (CC)?",
    "CC2": "CC2. If aware of CC, would you say that the CC of this office was...?",
    "CC3": "CC3. If aware of CC, how much did the CC help you in your transaction?"
}

# === Step 6: Normalize CC1 and filter invalid ===
df["CC1"] = df["CC1"].astype(str).str.strip()
df["CC1"] = df["CC1"].where(df["CC1"].isin(cc1_order))

# === Step 7: Helpers ===
def normalize_text(s):
    if pd.isna(s):
        return ''
    return str(s).strip().replace("’","'")

def autofit_columns(ws, headers, start_col=0):
    for i, col in enumerate(headers, start=start_col):
        max_len = max(len(str(col)), 1)
        ws.set_column(i, i, max_len + 2)

# === Step 8: Create Excel ===
out_file = "PIA_January 2026 updated_withComments.xlsx"
writer = pd.ExcelWriter(out_file, engine="xlsxwriter")
workbook = writer.book

bold_fmt = workbook.add_format({'bold': True,'font_name':'Arial','font_size':9})
center_fmt = workbook.add_format({'align':'center','font_name':'Arial','font_size':9})
percent_fmt = workbook.add_format({'num_format':'0.00','align':'center','font_name':'Arial','font_size':9})
border_fmt = workbook.add_format({'border':1,'font_name':'Arial','font_size':9})
wrap_fmt = workbook.add_format({'text_wrap': True,'valign': 'top','border': 1,'font_name':'Arial','font_size':9})

offices = df["PIA Office"].dropna().unique()

# === MAIN PER-OFFICE TABLES (with comments section) ===
for office in offices:
    if office=="PIA Main - Central Office":
        df_off = df[(df["PIA Office"]==office) & ((df["External Services"].notna()) | (df["Internal Services"].notna()))]
    else:
        df_off = df[(df["PIA Office"]==office) & (df["External Services"].notna())]

    for q in ["CC1","CC2","CC3"]:
        df_off.loc[:, q] = df_off[q].apply(normalize_text)

    ws = workbook.add_worksheet(office[:31])
    writer.sheets[office[:31]] = ws
    row = 0

    # === CC Tables ===
    ws.write(row,0,"Citizen Charter Awareness",bold_fmt)
    row+=1
    for q, order in zip(["CC1","CC2","CC3"],[cc1_order,cc2_order,cc3_order]):
        ws.write(row,0,cc_labels[q],bold_fmt)
        ws.write(row,1,"Response",border_fmt)
        ws.write(row,2,"Total responses",border_fmt)
        ws.write(row,3,"%",border_fmt)
        row+=1
        counts = df_off[q].value_counts()
        total = counts.sum()
        for resp in order:
            cnt = counts.get(resp,0)
            ws.write(row,1,resp,border_fmt)
            ws.write(row,2,cnt,center_fmt)
            ws.write(row,3,round(cnt/total*100,2) if total>0 else 0,percent_fmt)
            row+=1
        for col_idx in range(4):
            ws.write(row, col_idx, "", border_fmt)
        row += 1

    # === SQD Table ===
    row += 1
    ws.write(row, 0, "Service Quality Dimensions", bold_fmt)
    row += 1
    ws.write_row(row, 1, resp_order + ["Total Responses", "Overall %"], border_fmt)
    row += 1
    for sqd in ["SQD1","SQD2","SQD3","SQD4","SQD5","SQD6","SQD7","SQD8","SQD0"]:
        counts = df_off[sqd].value_counts()
        total = sum(counts.get(x,0) for x in resp_order[:5])
        agree = counts.get("Strongly Agree",0) + counts.get("Agree",0)
        ws.write(row, 0, sqd_labels[sqd], border_fmt)
        for i, resp in enumerate(resp_order):
            ws.write(row, i+1, counts.get(resp,0), center_fmt)
        ws.write(row, len(resp_order)+1, total, center_fmt)
        ws.write(row, len(resp_order)+2, round(agree/total*100,2) if total>0 else 0, percent_fmt)
        row += 1

    # === Services Table ===
    row += 1
    ws.write(row, 0, "Services", bold_fmt)
    ws.write(row, 1, "Total Responses", border_fmt)
    row += 1
    if office == "PIA Main - Central Office":
        services = pd.concat([df_off["External Services"], df_off["Internal Services"]])
    else:
        services = df_off["External Services"]
    counts = services.value_counts()
    for serv, cnt in counts.items():
        ws.write(row, 0, serv, border_fmt)
        ws.write(row, 1, cnt, center_fmt)
        row += 1
    ws.write(row, 0, "N", border_fmt)
    ws.write(row, 1, counts.sum(), center_fmt)

    # === NEW: Client Suggestions for Service Improvement (PIA3) ===
    row += 3
    ws.write(row, 0, "Client Suggestions for Service Improvement", bold_fmt)
    row += 1
    ws.write(row, 0, "Service Availed", border_fmt)
    ws.write(row, 1, "Consolidated Client Comments", border_fmt)
    row += 1
    df_comments = df_off[df_off["PIA3"].notna() & (df_off["PIA3"].astype(str).str.strip() != "")]
    if not df_comments.empty:
        df_comments["Service Availed"] = df_comments["External Services"].fillna('') + " " + df_comments["Internal Services"].fillna('')
        df_comments["Service Availed"] = df_comments["Service Availed"].str.strip()
        grouped = df_comments.groupby("Service Availed")["PIA3"].apply(lambda x: "\n• " + "\n• ".join(x.astype(str).str.strip())).reset_index()
        for _, r in grouped.iterrows():
            ws.write(row, 0, str(r["Service Availed"]).strip(), border_fmt)
            ws.write(row, 1, str(r["PIA3"]).strip(), wrap_fmt)
            row += 1

    # === Autofit Columns ===
    autofit_columns(ws, ["Question","Response","Total responses","%"], start_col=0)
    autofit_columns(ws, ["Dimension"] + resp_order + ["Total Responses","Overall %"], start_col=0)
    autofit_columns(ws, ["Service Availed","Consolidated Client Comments"], start_col=0)

# === CENTRAL OFFICE BY SERVICE (same as before) ===
central_office = "PIA Main - Central Office"
df_central = df[(df["PIA Office"] == central_office) & ((df["External Services"].notna()) | (df["Internal Services"].notna()))]
df_central["Service Availed"] = df_central["External Services"].fillna('') + " " + df_central["Internal Services"].fillna('')
df_central["Service Availed"] = df_central["Service Availed"].str.strip()
services = df_central["Service Availed"].dropna().unique()

for service in services:
    df_service = df_central[df_central["Service Availed"] == service]
    sheet_name = f"Central - {service}"[:31]
    ws = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws
    row = 0

    ws.write(row,0,"Citizen Charter Awareness",bold_fmt)
    row+=1
    for q, order in zip(["CC1","CC2","CC3"],[cc1_order,cc2_order,cc3_order]):
        ws.write(row,0,cc_labels[q],bold_fmt)
        ws.write(row,1,"Response",border_fmt)
        ws.write(row,2,"Total responses",border_fmt)
        ws.write(row,3,"%",border_fmt)
        row+=1
        counts = df_service[q].value_counts()
        total = counts.sum()
        for resp in order:
            cnt = counts.get(resp,0)
            ws.write(row,1,resp,border_fmt)
            ws.write(row,2,cnt,center_fmt)
            ws.write(row,3,round(cnt/total*100,2) if total>0 else 0,percent_fmt)
            row+=1
        for col_idx in range(4):
            ws.write(row, col_idx, "", border_fmt)
        row += 1

    # === SQD Table ===
    row += 1
    ws.write(row, 0, "Service Quality Dimensions", bold_fmt)
    row += 1
    ws.write_row(row, 1, resp_order + ["Total Responses", "Overall %"], border_fmt)
    row += 1
    for sqd in ["SQD1","SQD2","SQD3","SQD4","SQD5","SQD6","SQD7","SQD8","SQD0"]:
        counts = df_service[sqd].value_counts()
        total = sum(counts.get(x,0) for x in resp_order[:5])
        agree = counts.get("Strongly Agree",0) + counts.get("Agree",0)
        ws.write(row, 0, sqd_labels[sqd], border_fmt)
        for i, resp in enumerate(resp_order):
            ws.write(row, i+1, counts.get(resp,0), center_fmt)
        ws.write(row, len(resp_order)+1, total, center_fmt)
        ws.write(row, len(resp_order)+2, round(agree/total*100,2) if total>0 else 0, percent_fmt)
        row += 1

    # === PIA3 Comments Table ===
    row += 3
    ws.write(row, 0, "Client Suggestions for Service Improvement", bold_fmt)
    row += 1
    ws.write(row, 0, "Service Availed", border_fmt)
    ws.write(row, 1, "Consolidated Client Comments", border_fmt)
    row += 1
    df_comments = df_service[df_service["PIA3"].notna() & (df_service["PIA3"].astype(str).str.strip() != "")]
    if not df_comments.empty:
        grouped = df_comments.groupby("Service Availed")["PIA3"].apply(lambda x: "\n• " + "\n• ".join(x.astype(str).str.strip())).reset_index()
        for _, r in grouped.iterrows():
            ws.write(row, 0, str(r["Service Availed"]).strip(), border_fmt)
            ws.write(row, 1, str(r["PIA3"]).strip(), wrap_fmt)
            row += 1

    autofit_columns(ws, ["Question","Response","Total responses","%"], start_col=0)
    autofit_columns(ws, ["Dimension"] + resp_order + ["Total Responses","Overall %"], start_col=0)
    autofit_columns(ws, ["Service Availed","Consolidated Client Comments"], start_col=0)

# === Save workbook ===
writer.close()
print(f"✅ Excel report generated with comments in all sheets: {out_file}")


# In[ ]:




