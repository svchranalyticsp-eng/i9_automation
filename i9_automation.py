
# -*- coding: utf-8 -*-
"""
I-9 Automation Script — GitHub Actions Version
Reads all config from environment variables (set by GitHub Secrets).
No config.py needed — fully self-contained.
"""

import pandas as pd
import numpy as np
import datetime
import os
import shutil
import logging
import sys
import msal
import requests
import base64
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ============================================================
# CONFIGURATION — All from environment variables
# ============================================================
GRAPH_TENANT_ID = os.environ["GRAPH_TENANT_ID"]
GRAPH_CLIENT_ID = os.environ["GRAPH_CLIENT_ID"]
GRAPH_CLIENT_SECRET = os.environ["GRAPH_CLIENT_SECRET"]
GRAPH_SENDER_EMAIL = os.environ["GRAPH_SENDER_EMAIL"]
GOOGLE_SERVICE_ACCOUNT_KEY = os.environ["GOOGLE_SERVICE_ACCOUNT_KEY"]
GOOGLE_DRIVE_FOLDER_ID = os.environ["GOOGLE_DRIVE_FOLDER_ID"]

# Google Drive folder IDs for source reports
# These are the shared Drive folders where reports are uploaded daily
TRACKER_FOLDER_ID = os.environ.get("TRACKER_FOLDER_ID", "")
DAILY_REPORT_FOLDER_ID = os.environ.get("DAILY_REPORT_FOLDER_ID", "")
FUTURE_HIRES_FOLDER_ID = os.environ.get("FUTURE_HIRES_FOLDER_ID", "")
AOR_FOLDER_ID = os.environ.get("AOR_FOLDER_ID", "")

GOOGLE_DRIVE_LINK = (
    "https://drive.google.com/drive/folders/"
    + GOOGLE_DRIVE_FOLDER_ID
)
DAILY_RECIPIENTS = [
    "#HumanResources@dish.com",
    "stephany.guerrero@dish.com"
]

WORKING_DIR = os.path.join(os.getcwd(), "data")
FINAL_REPORT_DIR = os.path.join(WORKING_DIR, "Final Report")
MERIDIAN_REPORT_DIR = os.path.join(WORKING_DIR, "Meridian Report")

# ============================================================
# LOGGING
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("i9_automation.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


# ============================================================
# GOOGLE DRIVE CLIENT
# ============================================================
class GoogleDriveClient:
    """Download and upload files via Google Drive API."""

    SCOPES = [
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive.readonly"
    ]

    def __init__(self, service_account_key_path):
        creds = service_account.Credentials.from_service_account_file(
            service_account_key_path, scopes=self.SCOPES
        )
        self.service = build("drive", "v3", credentials=creds)

    def download_latest_file(self, folder_id, filename_contains, dest_path):
        """
        Download the most recent file from a Drive folder
        whose name contains the given string.
        """
        query = (
            f"'{folder_id}' in parents and trashed=false "
            f"and name contains '{filename_contains}'"
        )
        results = self.service.files().list(
            q=query,
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id, name, modifiedTime)"
        ).execute()
        files = results.get("files", [])

        if not files:
            raise FileNotFoundError(
                f"No file containing '{filename_contains}' "
                f"found in folder {folder_id}"
            )

        file_info = files[0]
        logger.info(
            f"  Downloading: {file_info['name']} "
            f"(modified: {file_info['modifiedTime']})"
        )

        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        request = self.service.files().get_media(fileId=file_info["id"])
        with open(dest_path, "wb") as f:
            f.write(request.execute())

        logger.info(f"  Saved to: {dest_path}")
        return dest_path

    def find_or_create_folder(self, folder_name, parent_id):
        """Find a subfolder by name, or create it."""
        query = (
            f"name='{folder_name}' "
            f"and mimeType='application/vnd.google-apps.folder' "
            f"and '{parent_id}' in parents and trashed=false"
        )
        results = self.service.files().list(
            q=query, fields="files(id, name)"
        ).execute()
        files = results.get("files", [])
        if files:
            return files[0]["id"]
        else:
            metadata = {
                "name": folder_name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_id]
            }
            folder = self.service.files().create(
                body=metadata, fields="id"
            ).execute()
            logger.info(f"  Created Drive folder: {folder_name}")
            return folder["id"]

    def upload_file(self, local_path, parent_folder_id, filename=None):
        """Upload a file to Google Drive."""
        if not filename:
            filename = os.path.basename(local_path)
        metadata = {"name": filename, "parents": [parent_folder_id]}
        media = MediaFileUpload(local_path, resumable=True)
        uploaded = self.service.files().create(
            body=metadata, media_body=media, fields="id"
        ).execute()
        logger.info(f"  Uploaded: {filename} (ID: {uploaded['id']})")
        return uploaded["id"]


# ============================================================
# GRAPH EMAIL CLIENT
# ============================================================
class GraphEmailClient:
    """Send emails via Microsoft Graph API."""

    def __init__(self, tenant_id, client_id, client_secret, sender_email):
        self.sender_email = sender_email
        self.app = msal.ConfidentialClientApplication(
            client_id,
            authority=(
                f"https://login.microsoftonline.com/{tenant_id}"
            ),
            client_credential=client_secret,
        )

    def _get_token(self):
        result = self.app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" in result:
            return result["access_token"]
        raise Exception(
            f"Auth failed: {result.get('error_description', result)}"
        )

    def send_email(self, to_recipients, subject, html_body,
                   attachments=None, importance="high"):
        token = self._get_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        to_list = [
            {"emailAddress": {"address": a.strip()}}
            for a in to_recipients if a.strip()
        ]
        message = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": html_body},
                "toRecipients": to_list,
                "importance": importance,
                "attachments": []
            },
            "saveToSentItems": "true"
        }
        if attachments:
            for fp in attachments:
                if os.path.exists(fp):
                    with open(fp, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode("utf-8")
                    message["message"]["attachments"].append({
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": os.path.basename(fp),
                        "contentBytes": b64
                    })

        url = (
            f"https://graph.microsoft.com/v1.0/users/"
            f"{self.sender_email}/sendMail"
        )
        resp = requests.post(url, headers=headers, json=message)
        if resp.status_code == 202:
            logger.info(f"  Email sent: '{subject}' → {to_recipients}")
        else:
            logger.error(f"  Email failed ({resp.status_code}): {resp.text}")
            raise Exception(f"Email failed: {resp.text}")


# ============================================================
# EMAIL TEMPLATES
# ============================================================
def build_daily_email_html(drive_link):
    return f"""
    <html><body>
    <p>Hi Team,</p>
    <p>Today's I-9 report is listed below.</p>
    <p>As a reminder:</p>
    <ul>
        <li>Employees highlighted in
            <span style='background:#8B0000; color:white; padding:2px 6px;'>
            DARK RED</span> have action items <b>PAST due</b>.
            <ul><li><b>These employees are out of compliance.</b></li></ul>
        </li>
        <li>Employees highlighted in
            <span style='background:#FF0000; padding:2px 6px;'>RED</span>
            have actions <b>due today</b>.
            <ul><li><b>Will be out of compliance if not completed today.
            </b></li></ul>
        </li>
        <li>Employees highlighted in
            <span style='background:#87CEFA; padding:2px 6px;'>BLUE</span>
            have action items due in the <b>next three business days</b>.
        </li>
    </ul>
    <p>Please take notice of employees with action items due tomorrow.
    <b>If reviewing separation for failure to provide documents, submit a
    Service Request in HR Link prior to submitting the term.</b></p>
    <p>Regarding reverifications, the employee's HRBP is responsible for
    all communication and updating HR Operations of any delays.</p>
    <p>Full report on Shared Google Drive:
    <a href="{drive_link}">HERE</a>.</p>
    <p>Questions? Create a Service Request using the I-9 category in
    'Ask HR'.</p>
    <p>Thank you,<br>HR Operations</p>
    </body></html>
    """


def build_wednesday_overdue_html():
    return """
    <html><body>
    <p>Hi Team,</p>
    <p>The following employees have overdue I-9 actions.</p>
    <p>Please take the following actions as applicable:</p>
    <ul>
        <li><b>Missing/mismatch SSN:</b> Update in HR Link.</li>
        <li><b>Reverification overdue:</b> Term at EOD or place on
        unpaid admin leave.</li>
        <li><b>Employee is termed:</b> Enter term date in Tracker.</li>
        <li><b>Cancel hire:</b> Remove from Tracker.</li>
    </ul>
    <p>Feel free to reach out with questions!</p>
    <p>Thanks,<br>HR Operations</p>
    </body></html>
    """


# ============================================================
# MAIN
# ============================================================
def main():
    logger.info("=" * 60)
    logger.info("I-9 Automation — GitHub Actions — Starting")
    logger.info("=" * 60)

    # --- Dates ---
    today = datetime.datetime.today().replace(microsecond=0)
    today_norm = pd.Timestamp(today).normalize()
    yesterday_norm = today_norm - pd.Timedelta(days=1)
    tomorrow_norm = today_norm + pd.Timedelta(days=1)

    todaytext_full = today.strftime("%m.%d.%y %H.%M.%S")
    todayyeartext = today.strftime("%Y")
    todaymonthtext = today.strftime("%B")
    todaymonthyeartext = f"{todaymonthtext} {todayyeartext}"
    todayweekday = today.weekday()
    todaytext_email = today.strftime(" %m/%d/%Y")

    FutureDateEnd = today + datetime.timedelta(days=28)
    DateRangeStart = today + datetime.timedelta(days=-14)
    ThreeWeeksTracker = today + datetime.timedelta(days=30)
    blue_cutoff = today_norm + pd.offsets.BusinessDay(3)

    excluded_locs = [
        'BANGALORE', 'UKRAINE DISH TECHNOLOGIES',
        'HYDERABAD', 'PUNE INDIA'
    ]

    # --- Initialize Clients ---
    logger.info("Initializing clients...")
    drive_client = GoogleDriveClient(GOOGLE_SERVICE_ACCOUNT_KEY)
    email_client = GraphEmailClient(
        GRAPH_TENANT_ID, GRAPH_CLIENT_ID,
        GRAPH_CLIENT_SECRET, GRAPH_SENDER_EMAIL
    )

    # --- Download Source Reports from Google Drive ---
    logger.info("Downloading source reports from Google Drive...")
    os.makedirs(WORKING_DIR, exist_ok=True)

    tracker_path = os.path.join(WORKING_DIR, "tracker_report.xls")
    daily_path = os.path.join(WORKING_DIR, "daily_report.xlsx")
    future_path = os.path.join(WORKING_DIR, "future_hires.xlsx")

    drive_client.download_latest_file(
        TRACKER_FOLDER_ID, "Custom Report", tracker_path
    )
    drive_client.download_latest_file(
        DAILY_REPORT_FOLDER_ID, "I9_ACTIVES", daily_path
    )
    drive_client.download_latest_file(
        FUTURE_HIRES_FOLDER_ID, "Future Hires", future_path
    )

    # --- Load Reports ---
    logger.info("Loading reports...")
    df_2trackerCustom = pd.read_excel(tracker_path, skiprows=0)
    df_1I9Daily = pd.read_excel(daily_path, skiprows=0)
    df_new_hires = pd.read_excel(future_path, skiprows=0)

    logger.info(f"  Tracker: {len(df_2trackerCustom)} records")
    logger.info(f"  Daily: {len(df_1I9Daily)} records")
    logger.info(f"  Future hires: {len(df_new_hires)} records")

    # --- Pre-Format ---
    logger.info("Pre-formatting...")
    df_1I9Daily['Start Date'] = pd.to_datetime(
        df_1I9Daily['Start Date'], errors='coerce'
    )
    df_1I9Daily['Projected Start Date'] = pd.to_datetime(
        df_1I9Daily['Projected Start Date'], errors='coerce'
    )

    df_2trackerCustomOrig = df_2trackerCustom.copy()
    df_2trackerCustom = df_2trackerCustom.loc[
        (df_2trackerCustom['Next Action'] != 'NONE') &
        (df_2trackerCustom['Next Action Due Date'] != 'Enter Term Date')
    ].copy()
    df_2trackerCustom['Next Action Due Date'] = pd.to_datetime(
        df_2trackerCustom['Next Action Due Date'], errors='coerce'
    )
    df_2trackerCustom = df_2trackerCustom.loc[
        df_2trackerCustom['Next Action Due Date'] <= ThreeWeeksTracker
    ]
    df_2trackerCustom.rename(columns={
        'Employee Name': 'Name',
        'Employee Profile Full SSN': 'SSN',
        'Employee ID': 'EID'
    }, inplace=True)

    df_1I9Daily = df_1I9Daily.loc[
        ~df_1I9Daily['Location'].isin(excluded_locs)
    ]
    df_1I9Daily_Filtered = df_1I9Daily.loc[
        (df_1I9Daily['Start Date'] >= DateRangeStart) &
        (df_1I9Daily['Start Date'] <= FutureDateEnd) &
        (df_1I9Daily['Projected Start Date'] <= FutureDateEnd)
    ].copy()

    # --- Merge ---
    logger.info("Merging reports...")
    tracker_eids = (
        df_2trackerCustomOrig['Employee ID']
        .astype(str).str.strip().unique()
    )

    df_MergedI92 = pd.merge(
        df_2trackerCustom, df_1I9Daily, on="EID", how="inner"
    )
    df_MergedI92 = df_MergedI92[[
        'EID', 'Name_x', 'Next Action', 'Start Date_x',
        'Next Action Due Date', 'Department', 'Location',
        'SSN_x', 'Assignment Status', 'Type'
    ]]
    df_MergedI92.rename(columns={
        'Name_x': 'Name', 'SSN_x': 'SSN',
        'Start Date_x': 'Start Date'
    }, inplace=True)

    df_1I9Daily5 = df_1I9Daily_Filtered[
        ~df_1I9Daily_Filtered['EID']
        .astype(str).str.strip().isin(tracker_eids)
    ].copy()
    df_1I9Daily5.rename(
        columns={'Projected Start Date': 'Next Action Due Date'},
        inplace=True
    )
    df_1I9Daily5.insert(3, 'Next Action', 'Create I-9')

    df_master = pd.concat(
        [df_MergedI92, df_1I9Daily5], ignore_index=True
    )

    # --- Reconcile New Hires ---
    logger.info("Reconciling new hires...")
    df_master['EID'] = df_master['EID'].astype(str).str.strip()
    df_new_hires['EID'] = df_new_hires['EID'].astype(str).str.strip()
    df_new_hires['Start Date'] = pd.to_datetime(
        df_new_hires['Start Date'], errors='coerce'
    )
    df_master['Start Date'] = pd.to_datetime(
        df_master['Start Date'], errors='coerce'
    )
    df_master['Next Action Due Date'] = pd.to_datetime(
        df_master['Next Action Due Date'], errors='coerce'
    )

    new_records = []
    for _, row in df_new_hires.iterrows():
        eid = row['EID']
        nh_start = row['Start Date']
        loc = str(row['Location']).upper()
        if pd.isna(nh_start) or loc in excluded_locs:
            continue
        if eid in df_master['EID'].values:
            mask = df_master['EID'] == eid
            if df_master.loc[mask, 'Start Date'].iloc[0] > nh_start:
                df_master.loc[mask, 'Start Date'] = nh_start
                df_master.loc[mask, 'Next Action Due Date'] = nh_start
        else:
            new_records.append({
                'EID': eid, 'Name': row['Name'],
                'Next Action': 'Create I-9',
                'Start Date': nh_start,
                'Department': row['Department'],
                'Location': row['Location'],
                'Next Action Due Date': nh_start,
                'Assignment Status': row['Assignment Status'],
                'Type': row['Type']
            })

    if new_records:
        df_master = pd.concat(
            [df_master, pd.DataFrame(new_records)], ignore_index=True
        )
        logger.info(f"  Added {len(new_records)} new hire records")

    # --- Sort & Categorize ---
    logger.info("Sorting and categorizing...")
    df_master['Next Action Due Date'] = pd.to_datetime(
        df_master['Next Action Due Date'], errors='coerce'
    )
    df_master.sort_values('Next Action Due Date', inplace=True)
    df_master.reset_index(drop=True, inplace=True)

    dates_norm = df_master['Next Action Due Date'].dt.normalize()
    df_darkred_idx = df_master.index[
        dates_norm <= yesterday_norm
    ].tolist()
    df_red_idx = df_master.index[
        dates_norm == today_norm
    ].tolist()
    df_blue_idx = df_master.index[
        (dates_norm >= tomorrow_norm) & (dates_norm <= blue_cutoff)
    ].tolist()
    df_gray_idx = df_master.index[
        df_master['Assignment Status'].isin(
            ['LOA - Unpaid Admin', 'LOA - Block']
        )
    ].tolist()

    logger.info(f"  Total: {len(df_master)}")
    logger.info(f"  Overdue: {len(df_darkred_idx)}")
    logger.info(f"  Due today: {len(df_red_idx)}")
    logger.info(f"  Next 3 days: {len(df_blue_idx)}")
    logger.info(f"  On leave: {len(df_gray_idx)}")

    df_master['Next Action Due Date'] = (
        df_master['Next Action Due Date'].dt.strftime('%m/%d/%Y')
    )
    df_master['Start Date'] = (
        pd.to_datetime(df_master['Start Date']).dt.strftime('%m/%d/%Y')
    )

    # --- Generate Excel ---
    logger.info("Generating Excel report...")
    os.makedirs(FINAL_REPORT_DIR, exist_ok=True)
    report_name = f"I-9 Status Report {todaytext_full}.xlsx"
    output_path = os.path.join(FINAL_REPORT_DIR, report_name)

    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    df_master.to_excel(
        writer, sheet_name="Needs Action",
        index=False, startrow=1, header=False
    )
    wb = writer.book
    ws = writer.sheets['Needs Action']

    hdr = wb.add_format({
        'bold': True, 'align': 'center', 'bg_color': '#D3D3D3',
        'font_size': 8, 'font_name': 'Arial', 'border': 1
    })
    default = wb.add_format({
        'align': 'center', 'font_size': 8, 'font_name': 'Arial'
    })
    fmt_darkred = wb.add_format({
        'bg_color': '#8B0000', 'font_color': 'white',
        'align': 'center', 'font_size': 8, 'font_name': 'Arial'
    })
    fmt_red = wb.add_format({
        'bg_color': '#FF0000', 'align': 'center',
        'font_size': 8, 'font_name': 'Arial'
    })
    fmt_blue = wb.add_format({
        'bg_color': '#87CEFA', 'align': 'center',
        'font_size': 8, 'font_name': 'Arial'
    })
    fmt_gray = wb.add_format({
        'bg_color': '#808080', 'font_color': 'white',
        'align': 'center', 'font_size': 8, 'font_name': 'Arial'
    })

    for c, val in enumerate(df_master.columns):
        ws.write(0, c, val, hdr)
    for i, col in enumerate(df_master.columns):
        w = max(df_master[col].astype(str).str.len().max(), len(col) + 2)
        ws.set_column(i, i, w)
    for r in range(len(df_master)):
        ws.set_row(r + 1, None, default)
    for idx in df_darkred_idx:
        ws.set_row(idx + 1, None, fmt_darkred)
    for idx in df_red_idx:
        ws.set_row(idx + 1, None, fmt_red)
    for idx in df_blue_idx:
        ws.set_row(idx + 1, None, fmt_blue)
    for idx in df_gray_idx:
        ws.set_row(idx + 1, None, fmt_gray)

    ws.protect('rain')
    writer.close()
    logger.info(f"  Report saved: {output_path}")

    # --- Send Daily Email ---
    logger.info("Sending daily email...")
    email_client.send_email(
        to_recipients=DAILY_RECIPIENTS,
        subject="I-9 Status Report",
        html_body=build_daily_email_html(GOOGLE_DRIVE_LINK),
        attachments=[output_path]
    )

    # --- Wednesday Overdue ---
    if todayweekday == 2:
        logger.info("Wednesday — generating overdue report...")
        aor_path = os.path.join(WORKING_DIR, "aor_report.xlsx")
        try:
            drive_client.download_latest_file(
                AOR_FOLDER_ID, "AOR", aor_path
            )
            df_AOR = pd.read_excel(aor_path)
            df_AOR = df_AOR.loc[df_AOR['Unnamed: 4'] != 'SEC_REP']
            df_AOR = df_AOR.loc[
                ~df_AOR['Unnamed: 2'].isin(
                    excluded_locs + ['BANGALORE INDIA']
                )
            ]
            df_AOR.rename(columns={
                'Unnamed: 1': 'Department', 'Unnamed: 2': 'Location',
                'Unnamed: 5': 'HRBP Name', 'Unnamed: 6': 'HRBP EID',
                'Unnamed: 7': 'HRBP Email'
            }, inplace=True)
            df_AOR = df_AOR[[
                'Department', 'Location', 'HRBP EID',
                'HRBP Name', 'HRBP Email'
            ]]

            df_overdue = df_master.iloc[df_darkred_idx].copy()
            df_hrbp = pd.merge(
                df_overdue, df_AOR,
                on=["Department", "Location"], how="left"
            )

            overdue_name = f"I-9 Overdue {todaytext_full}.xlsx"
            overdue_path = os.path.join(FINAL_REPORT_DIR, overdue_name)
            ow = pd.ExcelWriter(overdue_path, engine='xlsxwriter')
            df_hrbp.to_excel(
                ow, sheet_name="Overdue Items",
                index=False, startrow=1, header=False
            )
            owb = ow.book
            ows = ow.sheets['Overdue Items']
            ohdr = owb.add_format({
                'bold': True, 'align': 'center', 'bg_color': '#D3D3D3',
                'font_size': 8, 'font_name': 'Arial', 'border': 1
            })
            odr = owb.add_format({
                'bg_color': '#8B0000', 'font_color': 'white',
                'align': 'center', 'font_size': 8, 'font_name': 'Arial'
            })
            for c, val in enumerate(df_hrbp.columns):
                ows.write(0, c, val, ohdr)
            for i, col in enumerate(df_hrbp.columns):
                w = max(
                    df_hrbp[col].astype(str).str.len().max(),
                    len(col) + 2
                )
                ows.set_column(i, i, w)
            for r in range(len(df_hrbp)):
                ows.set_row(r + 1, None, odr)
            ows.protect('rain')
            ow.close()

            hrbp_emails = [
                e for e in df_hrbp['HRBP Email'].dropna().unique()
                if e.strip()
            ]
            if hrbp_emails:
                email_client.send_email(
                    to_recipients=hrbp_emails,
                    subject=f"I-9 Overdue Items{todaytext_email}",
                    html_body=build_wednesday_overdue_html(),
                    attachments=[overdue_path]
                )
                logger.info(
                    f"  Overdue email sent to {len(hrbp_emails)} HRBPs"
                )
        except Exception as e:
            logger.error(f"  Wednesday overdue failed: {e}")

    # --- Monday Meridian ---
    if todayweekday == 0:
        logger.info("Monday — generating Meridian report...")
        meridian_locs = [
            "ENGLEWOOD (CORP) (MERIDIAN A)", "DENVER 17TH",
            "ENGLEWOOD (100)", "ENGLEWOOD (CORP) (94)",
            "ENGLEWOOD (ETC) (90)", "DENVER 39TH",
            "LITTLETON", "WAH", "WAH6 DIRECT SALES"
        ]
        df_meridian = df_master.loc[
            df_master["Location"].isin(meridian_locs)
        ].copy()
        os.makedirs(MERIDIAN_REPORT_DIR, exist_ok=True)
        mer_name = f"Meridian Status Report {todaytext_full}.xlsx"
        mer_path = os.path.join(MERIDIAN_REPORT_DIR, mer_name)

        mw = pd.ExcelWriter(mer_path, engine='xlsxwriter')
        df_meridian.to_excel(
            mw, sheet_name="Meridian Action Items",
            index=False, startrow=1, header=False
        )
        mwb = mw.book
        mws = mw.sheets['Meridian Action Items']
        mhdr = mwb.add_format({
            'bold': True, 'align': 'center', 'bg_color': '#D3D3D3',
            'font_size': 8, 'font_name': 'Arial', 'border': 1
        })
        mdef = mwb.add_format({
            'align': 'center', 'font_size': 8, 'font_name': 'Arial'
        })
        for c, val in enumerate(df_meridian.columns):
            mws.write(0, c, val, mhdr)
        for i, col in enumerate(df_meridian.columns):
            w = max(
                df_meridian[col].astype(str).str.len().max(),
                len(col) + 2
            )
            mws.set_column(i, i, w)
        for r in range(len(df_meridian)):
            mws.set_row(r + 1, None, mdef)
        mw.close()
        logger.info(f"  Meridian report saved: {mer_path}")

    # --- Archive to Google Drive ---
    logger.info("Archiving to Google Drive...")
    try:
        yr_id = drive_client.find_or_create_folder(
            todayyeartext, GOOGLE_DRIVE_FOLDER_ID
        )
        mo_id = drive_client.find_or_create_folder(
            todaymonthyeartext, yr_id
        )
        drive_client.upload_file(output_path, mo_id, report_name)
        logger.info("  Archive complete")
    except Exception as e:
        logger.error(f"  Archive failed: {e}")

    logger.info("=" * 60)
    logger.info("I-9 Automation completed successfully!")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()

