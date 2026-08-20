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

DOCUMENT_FILE = r"C:\Users\a8ti61\Linde Group\SAP Server and Technology Platform - Documents\8 - Repository and Best Practices\Inventory Basis Documents\Document_Catalogue_2026_V1.0.xlsx"
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
    "si_basis@linde.com",
    "sandeep.kumar.jha@linde.com",
    "praveen.verma@linde.com",
    "thomas.gerulat@linde.com",
    "Steffen.Schnell-Kretschmer@linde.com",
    "sumit.das@linde.com"
]

BREACH_CC = [
    "si_basis@linde.com",
    "sandeep.kumar.jha@linde.com",
    "praveen.verma@linde.com",
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

def extract_to_emails(df, members_df):

    print("\n==============================")
    print("EXTRACTING RESPONSIBLE EMAILS")
    print("==============================")

    emails = []

    # ------------------------------------------------------------
    # Create first-name lookup from TEAM_AVAIL
    # ------------------------------------------------------------

    members_df = members_df.copy()

    members_df["FirstName_clean"] = (
        members_df["Name"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.split()
        .str[0]
        .str.lower()
    )

    members_df["Email_Clean"] = (
        members_df["Linde_E-Mail"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ------------------------------------------------------------
    # Process Responsible names
    # ------------------------------------------------------------

    for responsible in (
        df["Responsible"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    ):

        if not responsible:
            continue

        responsible_clean = (
            responsible
            .lower()
            .strip()
            .split()[0]
        )

        print(
            f"\nResponsible : {responsible}"
        )

        # --------------------------------------------------------
        # Match Responsible first name
        # --------------------------------------------------------

        match = members_df[
            members_df["FirstName_clean"]
            == responsible_clean
        ]

        if match.empty:

            print(
                f"NOT FOUND : {responsible}"
            )

            continue

        # --------------------------------------------------------
        # Add matching email(s)
        # --------------------------------------------------------

        for _, row in match.iterrows():

            email = row["Email_Clean"]

            if email and email.lower() != "nan":

                emails.append(email)

                print(
                    f"MATCHED : "
                    f"{row['Name']} -> {email}"
                )

    # ------------------------------------------------------------
    # Remove duplicate emails
    # ------------------------------------------------------------

    emails = list(dict.fromkeys(emails))

    # ------------------------------------------------------------
    # Final output
    # ------------------------------------------------------------

    print("\n==============================")
    print("FINAL TO EMAILS")
    print("==============================")

    for email in emails:
        print(email)

    print(
        f"\nTotal recipients: {len(emails)}"
    )

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
# Calculate days remaining independently for both review dates
# ------------------------------------------------------------

doc_df["Review Date Days Remaining"] = (

    doc_df["Review Date"]

    - today

).dt.days

doc_df["Next Planned Review Days Remaining"] = (

    doc_df["Next Planned Review Date"]

    - today

).dt.days

# ------------------------------------------------------------
# Breach alert helpers
# ------------------------------------------------------------

COLOR_PRIORITY = {
    COLOR_RED: 1,
    COLOR_ORANGE: 2,
    COLOR_YELLOW: 3,
    COLOR_LIGHT_YELLOW: 4,
    COLOR_WHITE: 5
}


def get_breach_alert_for_row(row):
    """
    Returns a consolidated breach alert that identifies which date field
    triggered the alert. If both dates are within the breach window, both
    conditions are shown and the most urgent colour is selected.
    """

    alert_parts = []
    alert_colors = []

    date_checks = [
        (
            "REVIEW DATE",
            row["Review Date"]
        ),
        (
            "NEXT PLANNED REVIEW",
            row["Next Planned Review Date"]
        )
    ]

    for label, date_value in date_checks:

        alert_text, alert_color = get_review_alert(
            date_value,
            mode="breach"
        )

        if alert_text:

            alert_parts.append(
                f"{label} - {alert_text}"
            )

            alert_colors.append(alert_color)

    if not alert_parts:

        return "", COLOR_WHITE

    alert_color = min(
        alert_colors,
        key=lambda color: COLOR_PRIORITY.get(
            color,
            COLOR_PRIORITY[COLOR_WHITE]
        )
    )

    return "<br>".join(alert_parts), alert_color

# ------------------------------------------------------------
# Breach Window - Review Date OR Next Planned Review Date
# Missing/blank dates are ignored safely by requiring notna().
# ------------------------------------------------------------

review_date_breach_mask = (

    doc_df["Review Date"].notna()

    &

    (
        doc_df["Review Date Days Remaining"]
        <= BREACH_WINDOW_DAYS
    )

)

next_planned_review_breach_mask = (

    doc_df["Next Planned Review Date"].notna()

    &

    (
        doc_df["Next Planned Review Days Remaining"]
        <= BREACH_WINDOW_DAYS
    )

)

breach_window_mask = (

    review_date_breach_mask

    |

    next_planned_review_breach_mask

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
# Review Alert for breached documents
# ------------------------------------------------------------

if not breached_df.empty:

    alerts = breached_df.apply(
        get_breach_alert_for_row,
        axis=1
    )

    breached_df["Review Alert"] = alerts.apply(
        lambda x: x[0]
    )

    breached_df["Alert Color"] = alerts.apply(
        lambda x: x[1]
    )

else:

    breached_df["Review Alert"] = ""

    breached_df["Alert Color"] = COLOR_WHITE

# ------------------------------------------------------------
# Sort by the most urgent triggering date
# ------------------------------------------------------------

if not breached_df.empty:

    breached_df["Earliest Breach Days Remaining"] = breached_df[
        [
            "Review Date Days Remaining",
            "Next Planned Review Days Remaining"
        ]
    ].min(axis=1)

    breached_df.sort_values(

        by="Earliest Breach Days Remaining",

        inplace=True

    )

    breached_df.drop(
        columns=["Earliest Breach Days Remaining"],
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
print("BREACH CHECK")
print("==============================")

if breached_df.empty:

    print("No breached reviews found.")

else:

    print(

        breached_df[
            [

                "Document Name",

                "Status",

                "Review Date",

                "Next Planned Review Date",

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
print("UPCOMING DOCUMENTS")
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

                "Next Planned Review Date",

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
# EMAIL SECTION HELPERS
# ============================================================

def build_upcoming_section():

    if upcoming_df.empty:

        return """
<p>
No document reviews are scheduled within the next 31 days.
</p>
"""

    return build_html_table(
        upcoming_df
    )


# ============================================================
# CONSOLIDATED EMAIL BODY
# ============================================================

def create_consolidated_review_email():

    if not breached_df.empty:

        print("\nPreparing Action Required email...")

        subject = BREACH_SUBJECT

        to_emails = extract_to_emails(
            breached_df,
            members_df
        )

        cc_emails = BREACH_CC

        breach_html = build_html_table(
            breached_df
        )

        intro_html = f"""
<p>Dear Team,</p>

<p>
The following document reviews require attention.
</p>

<br>

{breach_html}

<br>

<p>
The following documents are also scheduled<br>
for review within the next 31 days.
</p>
"""

    else:

        print("\nPreparing No Pending Review email...")

        subject = NO_BREACH_SUBJECT

        to_emails = extract_to_emails(
            upcoming_df,
            members_df
        )

        cc_emails = NO_BREACH_CC

        intro_html = """
<p>Dear Team,</p>

<p>
Good news!
</p>

<p>
No breached document reviews were found<br>
during today's review cycle.
</p>

<p>
Below are the next scheduled document reviews.
</p>
"""

    upcoming_html = build_upcoming_section()

    body = f"""
<html>

<body style="font-family:Arial;font-size:10pt;">

{intro_html}

<br>

{upcoming_html}

<br>

Regards,

<br>

SAP Basis Automation

</body>

</html>
"""

    create_outlook_mail(
        outlook=outlook,
        subject=subject,
        to_emails=to_emails,
        cc_emails=cc_emails,
        body=body
    )

# ============================================================
# MAIN
# ============================================================

outlook = initialize_outlook()

create_consolidated_review_email()

print("\n===================================")
print("PROCESS COMPLETED")
print("===================================")