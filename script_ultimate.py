import pandas as pd
import win32com.client as win32
from datetime import datetime
import sys
import html

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
# PENDING SLA DASHBOARD - LIGHT THEME CONFIGURATION
# ============================================================

DASHBOARD_THEME = {
    "page": "#FFFFFF",
    "header": "#C8D9F5",
    "header_text": "#17345C",
    "section": "#EAF0F8",
    "section_text": "#253D5B",
    "text": "#26384D",
    "muted_text": "#52647B",
    "border": "#D9E0E8",
    "track": "#E8EBEF",
    "kpi_critical_bg": "#F4E8D9",
    "kpi_critical_text": "#B85A32",
    "kpi_attention_bg": "#F8EBD9",
    "kpi_attention_text": "#C17A25",
    "kpi_total_bg": "#E8EDF5",
    "kpi_total_text": "#234E85",
}

# Central colour convention: edit these values to change the bar colours.
SLA_COLOURS = {
    "critical": "#D94B45",   # 0-7 days
    "attention": "#E58B32",  # 8-14 days
    "upcoming": "#D5B52E",   # 15-31 days
    "on_track": "#36A36A",   # More than 31 days
    "due_today": "#D94B45",
}

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
STATUS_IN_PROGRESS = "IN PROGRESS"

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
# Status eligible for current review breach checking
# Blank
# Pending
# In Progress
# ------------------------------------------------------------

active_current_review_mask = (

    doc_df["Status"].eq("")

    |

    doc_df["Status"].eq(STATUS_PENDING)

    |

    doc_df["Status"].eq(STATUS_IN_PROGRESS)

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
    triggered the alert. Review Date only triggers when the current review
    status is active; Next Planned Review Date is monitored for all statuses.
    """

    alert_parts = []
    alert_colors = []

    if row["Current Review Action Required"]:

        alert_text, alert_color = get_review_alert(
            row["Review Date"],
            mode="breach"
        )

        if alert_text:

            alert_parts.append(
                f"REVIEW DATE - {alert_text}"
            )

            alert_colors.append(alert_color)

    if row["Next Planned Review Action Required"]:

        alert_text, alert_color = get_review_alert(
            row["Next Planned Review Date"],
            mode="breach"
        )

        if alert_text:

            alert_parts.append(
                f"NEXT PLANNED REVIEW - {alert_text}"
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
# Breach Window - Current Review OR Next Planned Review
# Missing/blank dates are ignored safely by requiring notna().
#
# Current Review uses Review Date only for active statuses.
# Next Planned Review Date is monitored for every status, including
# REVIEW DONE, because it represents the next scheduled review.
# ------------------------------------------------------------

current_review_action_required_mask = (

    active_current_review_mask

    &

    doc_df["Review Date"].notna()

    &

    (
        doc_df["Review Date Days Remaining"]
        <= BREACH_WINDOW_DAYS
    )

)

next_planned_review_action_required_mask = (

    doc_df["Next Planned Review Date"].notna()

    &

    (
        doc_df["Next Planned Review Days Remaining"]
        <= BREACH_WINDOW_DAYS
    )

)

doc_df["Current Review Action Required"] = (
    current_review_action_required_mask
)

doc_df["Next Planned Review Action Required"] = (
    next_planned_review_action_required_mask
)

breach_window_mask = (

    current_review_action_required_mask

    |

    next_planned_review_action_required_mask

)

# ------------------------------------------------------------
# Final breached dataframe
# ------------------------------------------------------------

breached_df = (

    doc_df[

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

    breached_df["Current Review Action Days Remaining"] = (
        breached_df["Review Date Days Remaining"]
        .where(breached_df["Current Review Action Required"])
    )

    breached_df["Next Planned Review Action Days Remaining"] = (
        breached_df["Next Planned Review Days Remaining"]
        .where(breached_df["Next Planned Review Action Required"])
    )

    breached_df["Earliest Breach Days Remaining"] = breached_df[
        [
            "Current Review Action Days Remaining",
            "Next Planned Review Action Days Remaining"
        ]
    ].min(axis=1)

    breached_df.sort_values(

        by="Earliest Breach Days Remaining",

        inplace=True

    )

    breached_df.drop(
        columns=[
            "Current Review Action Days Remaining",
            "Next Planned Review Action Days Remaining",
            "Earliest Breach Days Remaining"
        ],
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
# PENDING SLA DASHBOARD
# ============================================================

def get_sla_classification(days_remaining):
    """Classify a pending document by calendar days remaining."""
    if days_remaining <= 7:
        return "critical", "Critical"
    if days_remaining <= 14:
        return "attention", "Attention Required"
    if days_remaining <= 31:
        return "upcoming", "Upcoming"
    return "on_track", "On Track"


def get_pending_documents(source_df):
    """
    Dashboard-only filter: show exactly PENDING source statuses.
    Status is not recalculated here; the catalogue status is authoritative.
    """
    pending = source_df[
        source_df["Status"].fillna("").astype(str).str.strip().str.upper()
        .eq(STATUS_PENDING)
    ].copy()

    pending = pending[pending["Next Planned Review Date"].notna()].copy()

    pending["Days Remaining"] = (
        pending["Next Planned Review Date"].dt.normalize() - today
    ).dt.days

    pending["Pending Window (days)"] = (
        pd.to_numeric(pending["Review Cycle (days)"], errors="coerce") * 0.20
    )

    def remaining_percentage(row):
        window = row["Pending Window (days)"]
        days = row["Days Remaining"]
        if pd.isna(window) or window <= 0:
            return 0.0
        return min(100.0, max(0.0, (days / window) * 100.0))

    pending["SLA Remaining %"] = pending.apply(
        remaining_percentage, axis=1
    )

    classified = pending["Days Remaining"].apply(get_sla_classification)
    pending["SLA Key"] = classified.apply(lambda item: item[0])
    pending["SLA Classification"] = classified.apply(lambda item: item[1])

    pending.sort_values(
        by=["Days Remaining", "Next Planned Review Date"],
        ascending=[True, True],
        inplace=True
    )
    pending.reset_index(drop=True, inplace=True)
    return pending


def build_sla_bar(percentage, colour):
    """Build an Outlook-compatible horizontal bar using nested HTML tables."""
    percentage = max(0.0, min(100.0, float(percentage or 0.0)))
    return f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
           border="0" style="width:100%;border-collapse:collapse;">
      <tr>
        <td bgcolor="{DASHBOARD_THEME['track']}" height="10"
            style="background:{DASHBOARD_THEME['track']};height:10px;
                   padding:0;font-size:1px;line-height:1px;">
          <table role="presentation" width="{percentage:.1f}%"
                 cellpadding="0" cellspacing="0" border="0"
                 style="width:{percentage:.1f}%;border-collapse:collapse;">
            <tr>
              <td bgcolor="{colour}" height="10"
                  style="background:{colour};height:10px;
                         font-size:1px;line-height:1px;">&nbsp;</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    """


def build_pending_sla_dashboard(pending_df):
    """Render the light-theme dashboard containing pending documents only."""
    theme = DASHBOARD_THEME

    total = len(pending_df)
    critical = int((pending_df["Days Remaining"] <= 7).sum()) if total else 0
    attention = int(
        ((pending_df["Days Remaining"] >= 8) &
         (pending_df["Days Remaining"] <= 14)).sum()
    ) if total else 0

    def kpi_card(label, value, background, foreground, note):
        return f"""
        <td width="33.33%" valign="top"
            style="width:33.33%;padding:12px;background:{background};
                   border-right:2px solid #FFFFFF;">
          <div style="font-family:Arial,sans-serif;font-size:9pt;
                      font-weight:bold;color:{theme['text']};">{label}</div>
          <div style="font-family:Arial,sans-serif;font-size:25pt;
                      line-height:32pt;font-weight:bold;color:{foreground};">
            {value}
          </div>
          <div style="font-family:Arial,sans-serif;font-size:9pt;
                      color:{theme['muted_text']};">{note}</div>
        </td>
        """

    if pending_df.empty:
        rows_html = f"""
        <tr><td colspan="7" align="center"
            style="padding:20px;color:{theme['muted_text']};
                   border-bottom:1px solid {theme['border']};">
            There are currently no pending document reviews.
        </td></tr>
        """
    else:
        row_parts = []
        for index, row in pending_df.iterrows():
            days = int(row["Days Remaining"])
            if days == 0:
                days_label = "Due Today"
            elif days == 1:
                days_label = "1 day left"
            elif days < 0:
                days_label = f"{abs(days)} day(s) overdue"
            else:
                days_label = f"{days} days left"

            bar_colour = SLA_COLOURS.get(
                row["SLA Key"], SLA_COLOURS["on_track"]
            )
            document_name = html.escape(str(row["Document Name"]))
            reviewer = html.escape(str(row["Responsible"] or ""))
            cycle = row["Review Cycle (days)"]
            cycle_display = (
                str(int(cycle))
                if pd.notna(cycle) and float(cycle).is_integer()
                else (html.escape(str(cycle)) if pd.notna(cycle) else "")
            )

            location = str(
                row["Location of the Document with Link"] or ""
            ).strip()
            if location.lower().startswith(("https://", "http://")):
                safe_url = html.escape(location, quote=True)
                document_link = (
                    f'<a href="{safe_url}" style="color:#3269A5;'
                    f'text-decoration:underline;font-weight:bold;">'
                    f'{document_name}</a>'
                )
            else:
                document_link = document_name

            actual_date = html.escape(format_date(row["Review Date"]))
            next_date = html.escape(
                format_date(row["Next Planned Review Date"])
            )
            classification = html.escape(str(row["SLA Classification"]))
            bar = build_sla_bar(row["SLA Remaining %"], bar_colour)

            row_parts.append(f"""
            <tr>
              <td valign="top" style="padding:10px 8px;
                  border-bottom:1px solid {theme['border']};
                  color:{theme['muted_text']};font-size:10pt;">{index + 1}</td>
              <td valign="top" style="padding:10px 8px;
                  border-bottom:1px solid {theme['border']};
                  color:{theme['text']};font-size:10pt;">{document_link}</td>
              <td valign="top" style="padding:10px 8px;
                  border-bottom:1px solid {theme['border']};
                  color:{theme['text']};font-size:10pt;">{reviewer}</td>
              <td align="center" valign="top" style="padding:10px 8px;
                  border-bottom:1px solid {theme['border']};
                  color:{theme['text']};font-size:10pt;">{cycle_display}</td>
              <td align="center" valign="top" style="padding:10px 8px;
                  border-bottom:1px solid {theme['border']};
                  color:{theme['text']};font-size:9pt;">{actual_date}</td>
              <td align="center" valign="top" style="padding:10px 8px;
                  border-bottom:1px solid {theme['border']};
                  color:{theme['text']};font-size:9pt;">{next_date}</td>
              <td valign="top" style="padding:10px 8px;
                  border-bottom:1px solid {theme['border']};min-width:150px;">
                <div style="font-family:Arial,sans-serif;font-size:9pt;
                            color:{bar_colour};font-weight:bold;margin-bottom:5px;">
                  {html.escape(days_label)} | {classification}
                </div>
                {bar}
              </td>
            </tr>
            """)
        rows_html = "".join(row_parts)

    report_date = today.strftime("%d %b %Y")
    return f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
           border="0" style="width:100%;border-collapse:collapse;
           font-family:Arial,sans-serif;background:{theme['page']};">
      <tr>
        <td colspan="3" bgcolor="{theme['header']}"
            style="background:{theme['header']};padding:14px 16px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
                 border="0"><tr>
            <td style="font-size:16pt;font-weight:bold;color:{theme['header_text']};">
              DOCUMENT REVIEW – PENDING SLA STATUS
            </td>
            <td align="right" style="font-size:10pt;font-weight:bold;
                color:{theme['header_text']};">{report_date}</td>
          </tr></table>
        </td>
      </tr>
      <tr>
        {kpi_card("CRITICAL DOCUMENTS", critical, theme["kpi_critical_bg"],
                  theme["kpi_critical_text"], "0–7 days remaining")}
        {kpi_card("ATTENTION REQUIRED", attention, theme["kpi_attention_bg"],
                  theme["kpi_attention_text"], "8–14 days remaining")}
        {kpi_card("TOTAL PENDING DOCUMENTS", total, theme["kpi_total_bg"],
                  theme["kpi_total_text"], "Pending documents")}
      </tr>
      <tr>
        <td colspan="3" bgcolor="{theme['section']}"
            style="background:{theme['section']};padding:10px 12px;
                   color:{theme['section_text']};font-size:10pt;font-weight:bold;">
          PENDING DOCUMENTS – SORTED BY SHORTEST SLA FIRST
        </td>
      </tr>
      <tr><td colspan="3" style="padding:0;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0"
               style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;">
          <tr bgcolor="#F3F6FA">
            <th align="left" style="padding:9px 8px;color:{theme['muted_text']};
                font-size:9pt;border-bottom:1px solid {theme['border']};">ID</th>
            <th align="left" style="padding:9px 8px;color:{theme['muted_text']};
                font-size:9pt;border-bottom:1px solid {theme['border']};">DOCUMENT NAME</th>
            <th align="left" style="padding:9px 8px;color:{theme['muted_text']};
                font-size:9pt;border-bottom:1px solid {theme['border']};">REVIEWER</th>
            <th align="center" style="padding:9px 8px;color:{theme['muted_text']};
                font-size:9pt;border-bottom:1px solid {theme['border']};">CYCLE</th>
            <th align="center" style="padding:9px 8px;color:{theme['muted_text']};
                font-size:9pt;border-bottom:1px solid {theme['border']};">ACTUAL REVIEW</th>
            <th align="center" style="padding:9px 8px;color:{theme['muted_text']};
                font-size:9pt;border-bottom:1px solid {theme['border']};">NEXT REVIEW</th>
            <th align="left" style="padding:9px 8px;color:{theme['muted_text']};
                font-size:9pt;border-bottom:1px solid {theme['border']};">SLA COUNTDOWN</th>
          </tr>
          {rows_html}
        </table>
      </td></tr>
      <tr>
        <td colspan="3" bgcolor="{theme['section']}"
            style="background:{theme['section']};padding:9px 12px;
                   color:{theme['muted_text']};font-size:8pt;">
          Document Review Reporting System | Pending window calculated using
          20% of Review Cycle.
        </td>
      </tr>
    </table>
    """


def extract_manager_emails(source_members_df):
    """Use the existing manager designation in TEAM_AVAIL."""
    manager_rows = source_members_df[
        source_members_df["Kanban"].fillna("").astype(str).str.strip()
        .str.upper().eq("MANAGERS")
    ]
    emails = manager_rows["Linde_E-Mail"].fillna("").astype(str).str.strip().tolist()
    emails.extend(NO_BREACH_CC)

    unique = {}
    for email in emails:
        if email and email.lower() != "nan":
            unique.setdefault(email.lower(), email)
    return list(unique.values())



# ============================================================
# CONSOLIDATED EMAIL BODY
# ============================================================

def create_consolidated_review_email():

    print("\nPreparing Pending SLA Dashboard email...")

    # Only the dashboard data is filtered to Pending. Reviewer distribution
    # is intentionally extracted from the complete document catalogue.
    pending_df = get_pending_documents(doc_df)
    print(f"Pending documents in dashboard: {len(pending_df)}")

    dashboard_html = build_pending_sla_dashboard(pending_df)

    # Preserve all reviewers from the full catalogue in To.
    to_emails = extract_to_emails(doc_df, members_df)

    # Preserve the current configured CC list and also include all managers
    # designated in TEAM_AVAIL (Kanban == MANAGERS).
    cc_emails = extract_manager_emails(members_df)

    subject = "Document Review - Pending SLA Dashboard"

    body = f"""
<html>
<body style="margin:0;padding:12px;background:#FFFFFF;
             font-family:Arial,sans-serif;font-size:10pt;color:#26384D;">
  <p>Dear Team,</p>
  <p>Please find below the current Pending Document Review SLA Dashboard.
     Only documents with Status = Pending are included.</p>
  {dashboard_html}
  <br>
  <p>Regards,<br>SAP Basis Automation</p>
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

    print("\nPending SLA Dashboard generated successfully.")
    print("Total Pending Documents:", len(pending_df))
    print("Total TO Recipients:", len(to_emails))
    print("Total CC Recipients:", len(cc_emails))


# ============================================================
# MAIN
# ============================================================

outlook = initialize_outlook()

create_consolidated_review_email()

print("\n===================================")
print("PROCESS COMPLETED")
print("===================================")