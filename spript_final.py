import pandas as pd
import win32com.client as win32

# ============================================================
# FILE PATHS
# ============================================================

DOCUMENT_FILE = r"Document_Catalogue_May_2026_V1.0.xlsx"

MEMBERS_FILE = r"GSPO Basis Members-Grid view.xlsx"

# ============================================================
# READ DOCUMENT CATALOGUE
# ============================================================

doc_df = pd.read_excel(
    DOCUMENT_FILE,
    sheet_name="Master Document",
    header=7,
    engine="openpyxl"
)

# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

doc_df.columns = (
    doc_df.columns
    .astype(str)
    .str.strip()
)

# ============================================================
# RENAME COLUMNS
# ============================================================

doc_df.columns = [
    "SL No",
    "Document Name",
    "Location of the Document with Link",
    "Responsible",
    "Review Cycle (days)",
    "Remarks",
    "Status",
    "Review Date",
    "Next Planned Review Date",
    "Extra"
]

# ============================================================
# REMOVE EXTRA COLUMN
# ============================================================

doc_df = doc_df.drop(columns=["Extra"])

# ============================================================
# REMOVE INVALID ROWS
# ============================================================

doc_df = doc_df[
    doc_df["Document Name"].notna()
].copy()

doc_df = doc_df[
    doc_df["Document Name"]
    .astype(str)
    .str.strip()
    != ""
].copy()

doc_df = doc_df[
    doc_df["Document Name"]
    .astype(str)
    .str.upper()
    != "PROCESS DOCUMENT LIST"
].copy()

# ============================================================
# DATE CONVERSION
# ============================================================

doc_df["Review Date"] = pd.to_datetime(
    doc_df["Review Date"],
    errors="coerce"
)

doc_df["Next Planned Review Date"] = pd.to_datetime(
    doc_df["Next Planned Review Date"],
    errors="coerce"
)

# ============================================================
# READ MEMBERS FILE
# ============================================================

members_df = pd.read_excel(
    MEMBERS_FILE,
    engine="openpyxl"
)

members_df.columns = (
    members_df.columns
    .astype(str)
    .str.strip()
)

# ============================================================
# FILTER LOGIC
# ============================================================

today = pd.Timestamp.today().normalize()

# Upcoming breach window
upcoming_limit = today + pd.Timedelta(days=4)

# Blank status means pending
pending_mask = (
    doc_df["Status"].isna()
    |
    (
        doc_df["Status"]
        .astype(str)
        .str.strip()
        == ""
    )
)

# Already breached
breached_mask = (
    doc_df["Review Date"].notna()
    &
    (
        doc_df["Review Date"] < today
    )
)

# Will breach within 4 days
upcoming_mask = (
    doc_df["Review Date"].notna()
    &
    (
        doc_df["Review Date"] >= today
    )
    &
    (
        doc_df["Review Date"] <= upcoming_limit
    )
)

# Final mask
final_mask = (
    pending_mask
    &
    (
        breached_mask
        |
        upcoming_mask
    )
)

# Final dataframe
review_df = doc_df[
    final_mask
].copy()

# ============================================================
# ADD ALERT TYPE
# ============================================================

review_df["Review Alert"] = review_df[
    "Review Date"
].apply(
    lambda x:
    "BREACHED"
    if x < today
    else "BREACHING IN 4 DAYS"
)

# ============================================================
# DEBUG
# ============================================================

print("\n==============================")
print("DOCUMENTS FOUND")
print("==============================")

print(
    review_df[[
        "SL No",
        "Document Name",
        "Responsible",
        "Review Date",
        "Review Alert"
    ]]
)

print("\nTotal Documents:", len(review_df))

# ============================================================
# STOP IF EMPTY
# ============================================================

if review_df.empty:

    print("\nNo pending reviews found.")

    exit()

# ============================================================
# GET MANAGER EMAILS
# ============================================================

manager_emails = members_df.loc[
    members_df["Kanban"]
    .astype(str)
    .str.upper()
    == "MANAGERS",
    "Linde_E-Mail"
].dropna().unique().tolist()

manager_cc = ";".join(manager_emails)

print("\nManager CC:")
print(manager_cc)

# ============================================================
# CREATE RECIPIENT LIST
# ============================================================

all_emails = (
    members_df["Linde_E-Mail"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

to_emails = ";".join(all_emails)

# ============================================================
# INITIALIZE OUTLOOK
# ============================================================

print("\nInitializing Outlook...")

try:

    outlook = win32.Dispatch(
        "Outlook.Application"
    )

    print("Outlook initialized successfully.")

except Exception as e:

    print("Outlook initialization failed.")
    print(e)

    exit()

# ============================================================
# SORT DATA
# ============================================================

review_df = review_df.sort_values(
    by=[
        "Responsible",
        "Review Date"
    ]
)

# ============================================================
# CREATE HTML TABLE
# ============================================================

html_table = """
<table border='1'
       cellpadding='5'
       cellspacing='0'
       style='border-collapse: collapse;
              font-family: Arial;
              font-size: 10pt;'>

    <tr style='background-color:#D9EAF7;'>

        <th>Document Name</th>
        <th>Location</th>
        <th>Responsible</th>
        <th>Review Cycle</th>
        <th>Remarks</th>
        <th>Status</th>
        <th>Review Alert</th>
        <th>Review Date</th>
        <th>Next Planned Review Date</th>

    </tr>
"""

# ============================================================
# LOOP THROUGH ALL DOCUMENTS
# ============================================================

for _, row in review_df.iterrows():

    review_date = ""
    next_review_date = ""

    if pd.notna(row["Review Date"]):

        review_date = (
            row["Review Date"]
            .strftime("%d-%m-%Y")
        )

    if pd.notna(
        row["Next Planned Review Date"]
    ):

        next_review_date = (
            row["Next Planned Review Date"]
            .strftime("%d-%m-%Y")
        )

    location = str(
        row[
            "Location of the Document with Link"
        ]
    )

    alert = row["Review Alert"]

    # ========================================================
    # ALERT COLOR
    # ========================================================

    if alert == "BREACHED":

        alert_color = "#FFB3B3"

    else:

        alert_color = "#FFE699"

    # ========================================================
    # ADD ROW
    # ========================================================

    html_table += f"""
    <tr>

        <td>{row['Document Name']}</td>

        <td>
            <a href="{location}">
                Open Document
            </a>
        </td>

        <td>{row['Responsible']}</td>

        <td>{row['Review Cycle (days)']}</td>

        <td>{row['Remarks']}</td>

        <td>{row['Status']}</td>

        <td style="background-color:{alert_color};
                   font-weight:bold;">
            {alert}
        </td>

        <td>{review_date}</td>

        <td>{next_review_date}</td>

    </tr>
    """

html_table += "</table>"

# ============================================================
# EMAIL BODY
# ============================================================

email_body = f"""
<html>

<body style='font-family: Arial;
             font-size: 10pt;'>

    <p>Dear Team,</p>

    <p>
    The following document reviews are either overdue
    or approaching breach within the next 4 days
    and require your attention.
    </p>

    <br>

    {html_table}

    <br>

    <p>
    Kindly complete the review activity
    at the earliest.
    </p>

    <br>

    <p>
    Regards,<br>
    SAP Basis
    </p>

</body>

</html>
"""

# ============================================================
# CREATE SINGLE OUTLOOK MAIL
# ============================================================

try:

    mail = outlook.CreateItem(0)

    mail.To = to_emails

    mail.CC = manager_cc

    mail.Subject = (
        "Pending Document Review - Action Required"
    )

    mail.HTMLBody = email_body

    # ========================================================
    # DISPLAY + SAVE DRAFT
    # ========================================================

    mail.Display()

    mail.Save()

    print("\nSingle draft mail created successfully.")

except Exception as e:

    print("\nFailed to create mail.")

    print(e)

# ============================================================
# FINISHED
# ============================================================

print("\nProcess completed successfully.")