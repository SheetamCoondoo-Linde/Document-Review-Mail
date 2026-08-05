import pandas as pd
import win32com.client as win32

# ============================================================
# FILE PATHS
# ============================================================

DOCUMENT_FILE = r"C:\Users\a8ti61\Linde Group\SAP Server and Technology Platform - Documents\8 - Repository and Best Practices\Inventory Basis Documents\\Document_Catalogue_2026_V1.0.xlsx"

MEMBERS_FILE = r"TEAM_AVAIL.xlsx"

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
]

# ============================================================
# REMOVE EXTRA COLUMN
# ============================================================

#doc_df = doc_df.drop(columns=["Extra"])

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

# Remove section title row
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
# READ TEAM AVAIL FILE
# ============================================================

members_df = pd.read_excel(
    MEMBERS_FILE,
    engine="openpyxl"
)

# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

members_df.columns = (
    members_df.columns
    .astype(str)
    .str.strip()
)

# ============================================================
# DEBUG TEAM FILE
# ============================================================

print("\n==============================")
print("TEAM AVAIL DATA")
print("==============================")

print(
    members_df.head(10)
)

# ============================================================
# CREATE FIRST NAME COLUMN
# ============================================================

members_df["FirstName_clean"] = (
    members_df["Name"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.split()
    .str[0]
    .str.lower()
)

# ============================================================
# CLEAN RESPONSIBLE COLUMN
# ============================================================

doc_df["Responsible_clean"] = (
    doc_df["Responsible"]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ============================================================
# TODAY
# ============================================================

today = pd.Timestamp.today().normalize()

# ============================================================
# REVIEW ALERT FUNCTION
# ============================================================

def get_review_alert(review_date):

    if pd.isna(review_date):

        return "", "#FFFFFF"

    days_left = (
        review_date.normalize()
        - today
    ).days

    # ========================================================
    # BREACHED
    # ========================================================

    if days_left < 0:

        return (
            f"BREACHED BY {abs(days_left)} DAY(S)",
            "#FF4D4D"
        )

    # ========================================================
    # TODAY
    # ========================================================

    elif days_left == 0:

        return (
            "BREACHING TODAY",
            "#FF9933"
        )

    # ========================================================
    # 1 DAY
    # ========================================================

    elif days_left == 1:

        return (
            "BREACHING IN 1 DAY",
            "#FFD966"
        )

    # ========================================================
    # WITHIN 4 DAYS
    # ========================================================

    elif days_left <= 4:

        return (
            f"BREACHING IN {days_left} DAYS",
            "#FFF2CC"
        )

    # ========================================================
    # IGNORE OTHERS
    # ========================================================

    else:

        return (
            "",
            "#FFFFFF"
        )

# ============================================================
# CREATE ALERT COLUMNS
# ============================================================

alerts = doc_df["Review Date"].apply(
    get_review_alert
)

doc_df["Review Alert"] = alerts.apply(
    lambda x: x[0]
)

doc_df["Alert Color"] = alerts.apply(
    lambda x: x[1]
)

# ============================================================
# FILTER:
# 1. BLANK STATUS
# 2. BREACHED OR WITHIN 4 DAYS
# ============================================================

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

days_diff = (
    doc_df["Review Date"]
    - today
).dt.days

review_window_mask = (
    days_diff <= 4
)

breached_df = doc_df[
    pending_mask
    &
    review_window_mask
].copy()

# ============================================================
# DEBUG OUTPUT
# ============================================================

print("\n==============================")
print("FINAL REVIEW ALERT DATA")
print("==============================")

print(
    breached_df[[
        "Document Name",
        "Responsible",
        "Review Alert",
        "Review Date"
    ]]
)

print("\nTotal Rows:", len(breached_df))

# ============================================================
# STOP IF EMPTY
# ============================================================

if breached_df.empty:

    print("\nNo breached reviews found.")

    upcoming_days_diff = (
        doc_df["Next Planned Review Date"]
        - today
    ).dt.days

    upcoming_df = doc_df[
        upcoming_days_diff > 0
    ].copy()

    upcoming_df = (
        upcoming_df
        .sort_values(
            by="Next Planned Review Date",
            ascending=True
        )
        .head(10)
        .copy()
    )

    upcoming_df["Review Alert"] = upcoming_df[
        "Next Planned Review Date"
    ].apply(
        lambda next_planned_review_date: (
            f"BREACHING IN {(next_planned_review_date.normalize() - today).days} DAYS"
            if pd.notna(next_planned_review_date)
            else ""
        )
    )

    upcoming_df["Alert Color"] = "#D9EAD3"

    upcoming_html_table = """
<table border='1'
       cellpadding='6'
       cellspacing='0'
       style='border-collapse: collapse;
              font-family: Arial;
              font-size: 10pt;
              width: 100%;'>

    <tr style='background-color:#1F4E78;
               color:white;
               font-weight:bold;
               text-align:center;'>

        <th>Document Name</th>
        <th>Location</th>
        <th>Responsible</th>
        <th>Review Cycle</th>
        <th>Remarks</th>
        <th>Review Alert</th>
        <th>Review Date</th>
        <th>Next Planned Review Date</th>

    </tr>
"""

    for _, row in upcoming_df.iterrows():

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

        upcoming_html_table += f"""
    <tr>

        <td>{row['Document Name']}</td>

        <td align='center'>
            <a href="{location}">
                Open Document
            </a>
        </td>

        <td>{row['Responsible']}</td>

        <td align='center'>
            {row['Review Cycle (days)']}
        </td>

        <td>{row['Remarks']}</td>

        <td style='background-color:{row['Alert Color']};
                   font-weight:bold;
                   text-align:center;'>

            {row['Review Alert']}

        </td>

        <td align='center'>
            {review_date}
        </td>

        <td align='center'>
            {next_review_date}
        </td>

    </tr>
    """

    upcoming_html_table += "</table>"

    try:

        print("\nInitializing Outlook...")

        outlook = win32.Dispatch("Outlook.Application")

        mail = outlook.CreateItem(0)

        mail.To = "si_basis@linde.com"
        mail.CC = (
            "sandeep.kumar.jha@linde.com;"
            "thomas.gerulat@linde.com;"
            "Steffen.Schnell-Kretschmer@linde.com;"
            "sumit.das@linde.com"
        )
        mail.Subject = "Document Review Alert - No Pending Reviews"

        mail.HTMLBody = f"""
        <html>
        <body style='font-family: Arial; font-size:10pt;'>

        <p>Dear Team,</p>

        <p>
        Good news! No breached reviews were found during today's document review check.
        </p>

        <p>
        All monitored documents are currently within their defined review cycles.
        </p>

        <p>
        Below are the next document(s) scheduled for review.
        </p>

        <br>

        {upcoming_html_table}

        <br>

        <p>
        Regards,<br>
        SAP Basis Automation
        </p>

        </body>
        </html>
        """

        # Display Draft
        mail.Display()

        # Save Draft
        mail.Save()

        print("No breached review notification draft created.")

    except Exception as e:

        print("Failed to create notification email.")
        print(e)

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

#To add managers to the mail CC
manager_cc = ";".join(manager_emails)

# ============================================================
# EXTRACT UNIQUE RESPONSIBLE PEOPLE
# ============================================================

responsible_people_raw = (
    breached_df["Responsible"]
    .dropna()
    .astype(str)
    .unique()
)

# ============================================================
# SPLIT RESPONSIBLE NAMES
# ============================================================

responsible_first_names = set()

for item in responsible_people_raw:

    split_names = (
        item.replace(",", " ")
        .replace("/", " ")
        .split()
    )

    # ONLY FIRST TOKEN
    if len(split_names) > 0:

        first_name = (
            split_names[0]
            .strip()
            .lower()
        )

        if first_name:

            responsible_first_names.add(
                first_name
            )

print("\nResponsible First Names:")
print(responsible_first_names)

# ============================================================
# EXTRACT TO EMAILS
# ============================================================

to_emails = []

for first_name in responsible_first_names:

    member_match = members_df[
        members_df["FirstName_clean"]
        == first_name
    ]

    if not member_match.empty:

        for _, match_row in member_match.iterrows():

            email = match_row["Linde_E-Mail"]

            if pd.notna(email):

                to_emails.append(email)

                print(
                    f"Matched: {first_name} -> {email}"
                )

    else:

        print(
            f"No email found for: {first_name}"
        )

# ============================================================
# REMOVE DUPLICATES
# ============================================================

to_emails = sorted(
    list(set(to_emails))
)

print("\nFINAL TO EMAILS:")
print(to_emails)

# ============================================================
# CREATE HTML TABLE
# ============================================================

html_table = """
<table border='1'
       cellpadding='6'
       cellspacing='0'
       style='border-collapse: collapse;
              font-family: Arial;
              font-size: 10pt;
              width: 100%;'>

    <tr style='background-color:#1F4E78;
               color:white;
               font-weight:bold;
               text-align:center;'>

        <th>Document Name</th>
        <th>Location</th>
        <th>Responsible</th>
        <th>Review Cycle</th>
        <th>Remarks</th>
        <th>Review Alert</th>
        <th>Review Date</th>
        <th>Next Planned Review Date</th>

    </tr>
"""

# ============================================================
# ADD ROWS
# ============================================================

for _, row in breached_df.iterrows():

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

    html_table += f"""
    <tr>

        <td>{row['Document Name']}</td>

        <td align='center'>
            <a href="{location}">
                Open Document
            </a>
        </td>

        <td>{row['Responsible']}</td>

        <td align='center'>
            {row['Review Cycle (days)']}
        </td>

        <td>{row['Remarks']}</td>

        <td style='background-color:{row['Alert Color']};
                   font-weight:bold;
                   text-align:center;'>

            {row['Review Alert']}

        </td>

        <td align='center'>
            {review_date}
        </td>

        <td align='center'>
            {next_review_date}
        </td>

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
    Kindly take the required action.
    </p>

    <br>

    {html_table}

    <br>

    <p>
    Regards,<br>
    SAP Basis Automation
    </p>

</body>

</html>
"""

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
# CREATE SINGLE CONSOLIDATED MAIL
# ============================================================

try:

    mail = outlook.CreateItem(0)

    mail.To = ";".join(to_emails)

    #mail.CC = manager_cc
    mail.CC = (
            "sandeep.kumar.jha@linde.com;"
            "si_basis@linde.com;"
            "thomas.gerulat@linde.com;"
            "Steffen.Schnell-Kretschmer@linde.com;"
            "sumit.das@linde.com"
    )

    mail.Subject = (
        "Document Review Alert - Action Required"
    )

    mail.HTMLBody = email_body

    # Display Draft
    mail.Display()

    # Save Draft
    mail.Save()

    print("\nSingle consolidated draft created.")

except Exception as e:

    print("\nFailed to create draft.")

    print(e)

# ============================================================
# FINISHED
# ============================================================

print("\nProcess completed.")
