import pandas as pd
import win32com.client as win32
from datetime import datetime
import sys

# ============================================================
# CONFIGURATION
# ============================================================

# -------------------------------
# Input Files
# -------------------------------

DOCUMENT_FILE = r"Document_Catalogue_2026_V1.0.xlsx"
MEMBERS_FILE = r"TEAM_AVAIL.xlsx"

# -------------------------------
# Excel Settings
# -------------------------------

DOCUMENT_SHEET = "Master Document"
DOCUMENT_HEADER_ROW = 7        # Row 8 in Excel

# -------------------------------
# Outlook Configuration
# -------------------------------

NO_BREACH_TO = [
    "si_basis@linde.com"
]

NO_BREACH_CC = [
    "sandeep.kumar.jha@linde.com",
    "thomas.gerulat@linde.com",
    "Steffen.Schnell-Kretschmer@linde.com",
    "sumit.das@linde.com"
]

BREACH_CC = [
    "sandeep.kumar.jha@linde.com",
    "si_basis@linde.com",
    "thomas.gerulat@linde.com",
    "Steffen.Schnell-Kretschmer@linde.com",
    "sumit.das@linde.com"
]

# -------------------------------
# Email Subjects
# -------------------------------

BREACH_SUBJECT = (
    "Document Review Alert - Action Required"
)

NO_BREACH_SUBJECT = (
    "Document Review Alert - No Pending Reviews"
)

# -------------------------------
# Alert Window
# -------------------------------

BREACH_WINDOW_DAYS = 4

# ============================================================
# HTML COLOURS
# ============================================================

COLOR_RED = "#FF4D4D"
COLOR_ORANGE = "#FF9933"
COLOR_YELLOW = "#FFD966"
COLOR_LIGHT_YELLOW = "#FFF2CC"
COLOR_GREEN = "#D9EAD3"
COLOR_WHITE = "#FFFFFF"

HEADER_COLOR = "#1F4E78"

# ============================================================
# COLUMN NAMES
# ============================================================

DOCUMENT_COLUMNS = [
    "SL No",
    "Document Name",
    "Location of the Document with Link",
    "Responsible",
    "Review Cycle (days)",
    "Remarks",
    "Status",
    "Review Date",
    "Next Planned Review Date"
]

# ============================================================
# STATUS VALUES
# ============================================================

STATUS_REVIEW_DONE = "REVIEW DONE"
STATUS_PENDING = "PENDING"

# ============================================================
# TODAY
# ============================================================

today = pd.Timestamp.today().normalize()

# ============================================================
# PANDAS SETTINGS
# ============================================================

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)
pd.set_option("display.max_colwidth", 80)

# ============================================================
# CONSOLE HEADER
# ============================================================

print("=" * 70)
print("SAP BASIS DOCUMENT REVIEW AUTOMATION")
print("=" * 70)

print(f"\nExecution Date : {today.strftime('%d-%m-%Y')}")

# ============================================================
# VERIFY INPUT FILES
# ============================================================

print("\nChecking input files...")

try:

    open(DOCUMENT_FILE).close()
    print(f"[OK] {DOCUMENT_FILE}")

except Exception:

    print(f"[ERROR] Cannot find {DOCUMENT_FILE}")
    sys.exit()

try:

    open(MEMBERS_FILE).close()
    print(f"[OK] {MEMBERS_FILE}")

except Exception:

    print(f"[ERROR] Cannot find {MEMBERS_FILE}")
    sys.exit()

print("\nConfiguration Loaded Successfully.")

# ============================================================
# READ DOCUMENT CATALOGUE
# ============================================================

print("\nReading Document Catalogue...")

try:

    doc_df = pd.read_excel(
        DOCUMENT_FILE,
        sheet_name=DOCUMENT_SHEET,
        header=DOCUMENT_HEADER_ROW,
        engine="openpyxl"
    )

    print(f"Loaded {len(doc_df)} rows.")

except Exception as e:

    print("Failed to read Document Catalogue.")
    print(e)
    sys.exit()

# ============================================================
# CLEAN DOCUMENT COLUMN NAMES
# ============================================================

doc_df.columns = (
    doc_df.columns
    .astype(str)
    .str.strip()
)

# Rename columns to standard names
doc_df.columns = DOCUMENT_COLUMNS

# ============================================================
# REMOVE INVALID ROWS
# ============================================================

doc_df = doc_df[
    doc_df["Document Name"].notna()
].copy()

doc_df = doc_df[
    doc_df["Document Name"]
    .astype(str)
    .str.strip() != ""
].copy()

doc_df = doc_df[
    doc_df["Document Name"]
    .astype(str)
    .str.upper()
    != "PROCESS DOCUMENT LIST"
].copy()

doc_df.reset_index(
    drop=True,
    inplace=True
)

print(f"Valid documents : {len(doc_df)}")

# ============================================================
# CONVERT DATE COLUMNS
# ============================================================

date_columns = [
    "Review Date",
    "Next Planned Review Date"
]

for column in date_columns:

    doc_df[column] = pd.to_datetime(
        doc_df[column],
        errors="coerce"
    )

# ============================================================
# CLEAN STRING COLUMNS
# ============================================================

string_columns = [
    "Document Name",
    "Location of the Document with Link",
    "Responsible",
    "Remarks",
    "Status"
]

for column in string_columns:

    doc_df[column] = (
        doc_df[column]
        .fillna("")
        .astype(str)
        .str.strip()
    )

# ============================================================
# NORMALIZE STATUS
# ============================================================

doc_df["Status"] = (
    doc_df["Status"]
    .str.upper()
)

# ============================================================
# REVIEW CYCLE
# ============================================================

doc_df["Review Cycle (days)"] = pd.to_numeric(
    doc_df["Review Cycle (days)"],
    errors="coerce"
)

# ============================================================
# READ TEAM FILE
# ============================================================

print("\nReading TEAM_AVAIL...")

try:

    members_df = pd.read_excel(
        MEMBERS_FILE,
        engine="openpyxl"
    )

    print(f"Loaded {len(members_df)} team members.")

except Exception as e:

    print("Failed to read TEAM_AVAIL.")
    print(e)
    sys.exit()

# ============================================================
# CLEAN TEAM COLUMN NAMES
# ============================================================

members_df.columns = (
    members_df.columns
    .astype(str)
    .str.strip()
)

# ============================================================
# CLEAN TEAM DATA
# ============================================================

members_df["Name"] = (
    members_df["Name"]
    .fillna("")
    .astype(str)
    .str.strip()
)

members_df["Linde_E-Mail"] = (
    members_df["Linde_E-Mail"]
    .fillna("")
    .astype(str)
    .str.strip()
)

members_df["Kanban"] = (
    members_df["Kanban"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
)

# ============================================================
# CREATE FIRST NAME COLUMN
# ============================================================

members_df["FirstName_clean"] = (

    members_df["Name"]

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
# DEBUG OUTPUT
# ============================================================

print("\n==============================")
print("DOCUMENT DATA")
print("==============================")

print(doc_df.head())

print("\nTotal Documents :", len(doc_df))

print("\n==============================")
print("TEAM DATA")
print("==============================")

print(

    members_df[
        [
            "Name",
            "Linde_E-Mail",
            "Kanban"
        ]
    ].head(10)

)

print("\nTotal Team Members :", len(members_df))

# ============================================================
# REVIEW ALERT CALCULATION
# ============================================================

def get_review_alert(review_date, mode="breach"):
    """
    Returns:
        alert_text,
        background_color
    """

    if pd.isna(review_date):

        return "", COLOR_WHITE

    days = (
        review_date.normalize() - today
    ).days

    # --------------------------------------------------------
    # BREACHED DOCUMENTS
    # --------------------------------------------------------

    if mode.lower() == "breach":

        if days < 0:

            return (
                f"BREACHED BY {abs(days)} DAY(S)",
                COLOR_RED
            )

        elif days == 0:

            return (
                "BREACHING TODAY",
                COLOR_ORANGE
            )

        elif days == 1:

            return (
                "BREACHING IN 1 DAY",
                COLOR_YELLOW
            )

        elif days <= BREACH_WINDOW_DAYS:

            return (
                f"BREACHING IN {days} DAYS",
                COLOR_LIGHT_YELLOW
            )

        else:

            return (
                "",
                COLOR_WHITE
            )

    # --------------------------------------------------------
    # UPCOMING DOCUMENTS
    # --------------------------------------------------------

    if days <= 14:

        return (
            f"REVIEW IN {days} DAY(S)",
            COLOR_RED
        )

    elif days < 31:

        return (
            f"REVIEW IN {days} DAY(S)",
            COLOR_YELLOW
        )

    else:

        return (
            f"REVIEW IN {days} DAY(S)",
            COLOR_GREEN
        )


# ============================================================
# FORMAT DATE
# ============================================================

def format_date(date_value):
    """
    Returns dd-mm-yyyy
    """

    if pd.isna(date_value):

        return ""

    return date_value.strftime("%d-%m-%Y")


# ============================================================
# EXTRACT RESPONSIBLE EMAILS
# ============================================================

def extract_to_emails(breached_df, members_df):

    responsible_people = (
        breached_df["Responsible"]
        .dropna()
        .astype(str)
        .unique()
    )

    first_names = set()

    for item in responsible_people:

        item = (
            item.replace(",", " ")
            .replace("/", " ")
            .replace("&", " ")
        )

        words = item.split()

        if words:

            first_names.add(
                words[0].lower().strip()
            )

    print("\nResponsible First Names")

    print(first_names)

    emails = []

    for first_name in first_names:

        match = members_df[
            members_df["FirstName_clean"]
            == first_name
        ]

        if not match.empty:

            for _, row in match.iterrows():

                email = row["Linde_E-Mail"]

                if email:

                    emails.append(email)

                    print(
                        f"Matched : {first_name} -> {email}"
                    )

        else:

            print(
                f"No email found for : {first_name}"
            )

    emails = sorted(
        list(set(emails))
    )

    print("\nFinal TO Emails")

    print(emails)

    return emails


# ============================================================
# HTML TABLE BUILDER
# ============================================================

def build_html_table(df):

    html = f"""
<table border="1"
       cellpadding="6"
       cellspacing="0"
       style="
            border-collapse:collapse;
            font-family:Arial;
            font-size:10pt;
            width:100%;
       ">

<tr style="
        background:{HEADER_COLOR};
        color:white;
        font-weight:bold;
        text-align:center;
">

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

    if df.empty:

        html += """
<tr>

<td colspan="9"
    align="center">

No records found.

</td>

</tr>
"""

    else:

        for _, row in df.iterrows():

            location = str(
                row["Location of the Document with Link"]
            )

            html += f"""

<tr>

<td>{row["Document Name"]}</td>

<td align="center">

<a href="{location}">
Open Document
</a>

</td>

<td>{row["Responsible"]}</td>

<td align="center">
{row["Review Cycle (days)"]}
</td>

<td>{row["Remarks"]}</td>

<td align="center">
{row["Status"]}
</td>

<td align="center"
style="
background:{row['Alert Color']};
font-weight:bold;
">

{row["Review Alert"]}

</td>

<td align="center">

{format_date(row["Review Date"])}

</td>

<td align="center">

{format_date(
row["Next Planned Review Date"]
)}

</td>

</tr>
"""

    html += "</table>"

    return html


# ============================================================
# OUTLOOK INITIALIZATION
# ============================================================

def initialize_outlook():

    try:

        print("\nInitializing Outlook...")

        outlook = win32.Dispatch(
            "Outlook.Application"
        )

        print("Outlook Initialized.")

        return outlook

    except Exception as e:

        print("\nFailed to initialize Outlook.")

        print(e)

        sys.exit()

# ============================================================
# BREACHED DOCUMENT FILTER
# ============================================================

print("\nChecking for breached document reviews...")

# ------------------------------------------------------------
# Status eligible for breach checking
# Blank
# Pending
# ------------------------------------------------------------

pending_mask = (

    doc_df["Status"].eq("")

    |

    doc_df["Status"].eq(STATUS_PENDING)

)

# ------------------------------------------------------------
# Calculate days remaining
# ------------------------------------------------------------

doc_df["Days Remaining"] = (

    doc_df["Review Date"]

    - today

).dt.days

# ------------------------------------------------------------
# Review Alert
# ------------------------------------------------------------

alerts = doc_df["Review Date"].apply(
    lambda x: get_review_alert(
        x,
        mode="breach"
    )
)

doc_df["Review Alert"] = alerts.apply(
    lambda x: x[0]
)

doc_df["Alert Color"] = alerts.apply(
    lambda x: x[1]
)

# ------------------------------------------------------------
# Breach Window
# ------------------------------------------------------------

breach_window_mask = (

    doc_df["Days Remaining"]

    <= BREACH_WINDOW_DAYS

)

# ------------------------------------------------------------
# Final breached dataframe
# ------------------------------------------------------------

breached_df = (

    doc_df[

        pending_mask

        &

        breach_window_mask

    ]

    .copy()

)

# ------------------------------------------------------------
# Sort
# ------------------------------------------------------------

breached_df.sort_values(

    by="Review Date",

    inplace=True

)

breached_df.reset_index(

    drop=True,

    inplace=True

)

# ============================================================
# DEBUG
# ============================================================

print("\n==============================")
print("FINAL REVIEW ALERT DATA")
print("==============================")

if breached_df.empty:

    print("No breached reviews found.")

else:

    print(

        breached_df[
            [

                "Document Name",

                "Responsible",

                "Status",

                "Review Date",

                "Review Alert"

            ]

        ]

    )

print(

    "\nTotal Breached Documents :",

    len(breached_df)

)

# ============================================================
# UPCOMING DOCUMENTS
# ============================================================

print("\nFinding next scheduled reviews...")

upcoming_df = (

    doc_df[

        doc_df["Next Planned Review Date"].notna()

        &

        (doc_df["Next Planned Review Date"] >= today)

        &

        (doc_df["Next Planned Review Date"] <= today + pd.Timedelta(days=31))

    ]

    .copy()

)

# ------------------------------------------------------------
# Sort by nearest Review Date
# ------------------------------------------------------------

upcoming_df.sort_values(

    by="Next Planned Review Date",

    ascending=True,

    inplace=True

)

# ------------------------------------------------------------
# Only next 10
# ------------------------------------------------------------

upcoming_df = upcoming_df.head(10)

# ------------------------------------------------------------
# Review Alerts
# ------------------------------------------------------------

alerts = upcoming_df["Next Planned Review Date"].apply(

    lambda x: get_review_alert(
        x,
        mode="upcoming"
    )

)

upcoming_df["Review Alert"] = alerts.apply(
    lambda x: x[0]
)

upcoming_df["Alert Color"] = alerts.apply(
    lambda x: x[1]
)

# ============================================================
# DEBUG
# ============================================================

print("\n==============================")
print("NEXT DOCUMENTS TO BE REVIEWED")
print("==============================")

if upcoming_df.empty:

    print(

        "No upcoming document reviews found."

    )

else:

    print(

        upcoming_df[
            [

                "Document Name",

                "Responsible",

                "Status",

                "Review Date",

                "Review Alert"

            ]

        ]

    )

print(

    "\nTotal Upcoming Documents :",

    len(upcoming_df)

)

# ============================================================
# CREATE OUTLOOK EMAIL
# ============================================================

def create_outlook_mail(
        outlook,
        subject,
        to_emails,
        cc_emails,
        body):

    mail = outlook.CreateItem(0)

    mail.To = ";".join(to_emails)

    mail.CC = ";".join(cc_emails)

    mail.Subject = subject

    mail.HTMLBody = body

    mail.Display()

    mail.Save()

    print("\nDraft created successfully.")

# ============================================================
# BREACHED EMAIL BODY
# ============================================================

def create_breached_email():

    print("\nPreparing Action Required email...")

    to_emails = extract_to_emails(
        breached_df,
        members_df
    )

    html_table = build_html_table(
        breached_df
    )

    body = f"""
<html>

<body style="font-family:Arial;font-size:10pt;">

<p>Dear Team,</p>

<p>

Kindly take the required action for the
following document reviews.

</p>

<br>

{html_table}

<br>

Regards,

<br>

SAP Basis Automation

</body>

</html>
"""

    create_outlook_mail(

        outlook=outlook,

        subject=BREACH_SUBJECT,

        to_emails=to_emails,

        cc_emails=BREACH_CC,

        body=body

    )

# ============================================================
# NO BREACHED EMAIL BODY
# ============================================================

def create_no_breach_email():

    print("\nPreparing No Pending Review email...")

    html_table = build_html_table(
        upcoming_df
    )

    body = f"""
<html>

<body style="font-family:Arial;font-size:10pt;">

<p>Dear Team,</p>

<p>

Good news!

No breached document reviews were found
during today's review cycle.

</p>

<p>

Below are the next scheduled document
reviews.

</p>

<br>

{html_table}

<br>

Regards,

<br>

SAP Basis Automation

</body>

</html>
"""

    create_outlook_mail(

        outlook=outlook,

        subject=NO_BREACH_SUBJECT,

        to_emails=NO_BREACH_TO,

        cc_emails=NO_BREACH_CC,

        body=body

    )


# ============================================================
# MAIN
# ============================================================

outlook = initialize_outlook()

# ------------------------------------------------------------
# Action Required
# ------------------------------------------------------------

if not breached_df.empty:

    create_breached_email()

# ------------------------------------------------------------
# No Pending Reviews
# ------------------------------------------------------------

else:

    create_no_breach_email()

print("\n===================================")
print("PROCESS COMPLETED")
print("===================================")