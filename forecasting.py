from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from email.message import EmailMessage
from pathlib import Path
import mimetypes
import os, sys
import shutil
import copy
import smtplib
import tempfile
import tkinter as tk
from tkinter import messagebox

import pandas as pd
import pyodbc
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.drawing.image import Image
# Add the parent directory ('NCC Automations') to the Python path
# This allows us to import the 'PythonTools' package from there.
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import CREDS, EMAILS

PV_SYST_DB_PATH = Path(
    r"G:\Shared drives\O&M\NCC Automations\Performance Reporting\PVsyst (Josephs Edits).accdb"
)
DEFAULT_DOWNLOADS_DIR = Path.home() / "Downloads"
BASE_OUTPUT_DIR = DEFAULT_DOWNLOADS_DIR / "Forecasting Outputs"
GMAIL_SMTP_HOST = "smtp.gmail.com"
GMAIL_SMTP_PORT = 465
DEFAULT_SENDER_EMAIL = EMAILS['NCC Desk']
DEFAULT_RECIPIENT_EMAILS = [EMAILS['Joseph Lang']]
LOGO = Path(r"G:\Shared drives\O&M\NCC Automations\Icons\narenco_logo-email.jpg")

LILY_CSV_COLUMNS = [
    "recorder_id",
    "Rating_(MW)",
    "Current_max(MW)",
    "date",
    "hour",
    "Expected_MWh",
]


@dataclass(frozen=True)
class PlantConfig:
    name: str
    db_name: str
    rating_mw: float
    output_subdir: str
    output_filename_pattern: str
    workbook_filename_label: str | None = None
    recorder_id: int | None = None


PLANTS: dict[str, PlantConfig] = {
    "BLUEBIRD": PlantConfig(
        name="BLUEBIRD",
        db_name="BLUEBIRD",
        rating_mw=3.0,
        output_subdir="Bluebird",
        output_filename_pattern="Bluebird Solar Forecast {report_date}.xlsx",
        workbook_filename_label="BLUEBIRD",
    ),
    "CARDINAL": PlantConfig(
        name="CARDINAL",
        db_name="CARDINAL",
        rating_mw=6.5,
        output_subdir="Cardinal",
        output_filename_pattern="Cardinal Forecast {report_date}.xlsx",
        workbook_filename_label="CARDINAL",
    ),
    "CHERRY BLOSSOM": PlantConfig(
        name="CHERRY BLOSSOM",
        db_name="CHERRY BLOSSOM",
        rating_mw=10.0,
        output_subdir="Cherry Blossom",
        output_filename_pattern="Cherry Blossom Forecast {report_date}.xlsx",
        workbook_filename_label="CHERRY BLOSSOM",
    ),
    "LILY": PlantConfig(
        name="LILY",
        db_name="LILY",
        rating_mw=70.0,
        output_subdir="Lily",
        output_filename_pattern="ID 2201697 Forecast {report_date}.csv",
        recorder_id=2201697,
    ),
}


def get_pvsyst_connection(db_path: Path = PV_SYST_DB_PATH) -> pyodbc.Connection:
    connection_string = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    return pyodbc.connect(connection_string)


def _normalize_hour_code(value: object) -> int:
    hour = int(value)
    if 0 <= hour <= 23:
        return hour + 1
    return hour


def fetch_pvsyst_hourly(
    plant: PlantConfig,
    start_date: date,
    end_date: date,
    connection: pyodbc.Connection,
) -> pd.DataFrame:
    query = """
        SELECT
            MonthCode,
            DayCode,
            HourCode,
            EGrid_KWH
        FROM PVsystHourly
        WHERE PlantName = ?
    """
    cursor = connection.cursor()
    cursor.execute(query, (plant.db_name,))
    rows = cursor.fetchall()
    columns = [column[0] for column in cursor.description]
    df = pd.DataFrame.from_records(rows, columns=columns)

    if df.empty:
        raise ValueError(f"No PVsyst data found for plant '{plant.db_name}'.")

    df["MonthCode"] = pd.to_numeric(df["MonthCode"], errors="coerce")
    df["DayCode"] = pd.to_numeric(df["DayCode"], errors="coerce")
    df["HourCode"] = pd.to_numeric(df["HourCode"], errors="coerce")
    df["EGrid_KWH"] = pd.to_numeric(df["EGrid_KWH"], errors="coerce").fillna(0)
    df["EGrid_MWH"] = df["EGrid_KWH"] / 1000.0

    df = df.dropna(subset=["MonthCode", "DayCode", "HourCode"]).copy()
    df["MonthCode"] = df["MonthCode"].astype(int)
    df["DayCode"] = df["DayCode"].astype(int)
    df["Hour"] = df["HourCode"].map(_normalize_hour_code).astype(int)

    all_days = pd.date_range(start=start_date, end=end_date, freq="D")
    expanded = pd.DataFrame({"ForecastDate": all_days})
    expanded["MonthCode"] = expanded["ForecastDate"].dt.month
    expanded["DayCode"] = expanded["ForecastDate"].dt.day

    merged = expanded.merge(
        df[["MonthCode", "DayCode", "Hour", "EGrid_MWH"]],
        on=["MonthCode", "DayCode"],
        how="left",
    )

    merged["EGrid_MWH"] = merged["EGrid_MWH"].fillna(0.0)
    merged["ForecastDate"] = pd.to_datetime(merged["ForecastDate"]).dt.date
    merged = merged[(merged["Hour"] >= 1) & (merged["Hour"] <= 24)].copy()
    merged = merged.sort_values(["ForecastDate", "Hour"]).reset_index(drop=True)

    counts = merged.groupby("ForecastDate")["Hour"].nunique()
    incomplete_days = counts[counts != 24]
    if not incomplete_days.empty:
        missing_dates = ", ".join(str(idx) for idx in incomplete_days.index.tolist())
        raise ValueError(
            f"PVsyst query returned incomplete hourly data for {plant.name}: {missing_dates}"
        )

    return merged


def build_excel_report_frame(hourly_df: pd.DataFrame, plant: PlantConfig) -> pd.DataFrame:
    working = hourly_df.copy()
    working["PlantName"] = plant.workbook_filename_label or plant.name

    pivot = (
        working.pivot_table(
            index=["ForecastDate", "PlantName"],
            columns="Hour",
            values="EGrid_MWH",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
        .rename(columns={"ForecastDate": "Date (Eastern Prevailing Time)", "PlantName": "PLANT_NAME"})
    )

    for hour in range(1, 25):
        if hour not in pivot.columns:
            pivot[hour] = 0.0

    ordered_hour_columns = [hour for hour in range(1, 25)]
    pivot = pivot[["Date (Eastern Prevailing Time)", "PLANT_NAME", *ordered_hour_columns]]

    rename_map = {hour: f"HE{hour}" for hour in ordered_hour_columns}
    pivot = pivot.rename(columns=rename_map)

    hour_columns = [f"HE{hour}" for hour in range(1, 25)]

    def find_start(row: pd.Series) -> int:
        non_zero_hours = [hour for hour in range(1, 25) if float(row[f"HE{hour}"]) > 0]
        return non_zero_hours[0] if non_zero_hours else 0

    def find_end(row: pd.Series) -> int:
        non_zero_hours = [hour for hour in range(1, 25) if float(row[f"HE{hour}"]) > 0]
        return non_zero_hours[-1] if non_zero_hours else 0

    pivot["Start"] = pivot.apply(find_start, axis=1)
    pivot["End"] = pivot.apply(find_end, axis=1)

    pivot["Date (Eastern Prevailing Time)"] = pd.to_datetime(
        pivot["Date (Eastern Prevailing Time)"]
    ).dt.strftime("%Y-%m-%d")

    return pivot[["Date (Eastern Prevailing Time)", "PLANT_NAME", *hour_columns, "Start", "End"]]


def build_lily_csv_frame(hourly_df: pd.DataFrame, plant: PlantConfig) -> pd.DataFrame:
    output = hourly_df.copy()
    output["recorder_id"] = plant.recorder_id
    output["Rating_(MW)"] = plant.rating_mw
    output["Current_max(MW)"] = plant.rating_mw
    output["date"] = output["ForecastDate"].map(
        lambda d: f"{d.month}/{d.day}/{d.year}"
    )
    output["hour"] = output["Hour"].astype(int)
    output["Expected_MWh"] = output["EGrid_MWH"].astype(float)

    return output[LILY_CSV_COLUMNS]


def _write_sheet_title_block(
    worksheet,
    title: str,
    plant: PlantConfig,
    report_df: pd.DataFrame,
    submitted_date: date,
) -> None:
    worksheet["A1"] = title
    worksheet["A2"] = "Plant Name:"
    worksheet["B2"] = plant.name
    worksheet["A3"] = "Nameplate quantity (MW):"
    worksheet["B3"] = plant.rating_mw

    he_columns = [f"HE{hour}" for hour in range(1, 25)]
    energy = float(report_df[he_columns].to_numpy().sum())
    max_possible = plant.rating_mw * 24 * len(report_df)
    capacity_factor = energy / max_possible if max_possible else 0

    worksheet["A4"] = "Energy (MWh):"
    worksheet["B4"] = energy
    worksheet["C4"] = max_possible
    worksheet["A5"] = (
        "Day-Ahead Capacity Factor:"
        if "Day-Ahead" in title
        else "Week-Ahead Capacity Factor:"
    )
    worksheet["B5"] = capacity_factor
    worksheet["A6"] = "Submitted:"
    worksheet["B6"] = submitted_date.strftime("%Y-%m-%d")


def apply_excel_formatting(worksheet) -> None:
    worksheet.sheet_view.showGridLines = False

    grey_fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
    bold_font = Font(bold=True)
    thin_side = Side(style="thin", color="000000")
    side_border = Border(
        left=thin_side,
        right=thin_side,
    )

    for cell_ref in ("A1", "A2", "A3", "A4", "A5", "A6"):
        worksheet[cell_ref].font = bold_font

    worksheet["C4"].font = Font(color="FFFFFF")
    worksheet["B4"].number_format = "0.0"
    worksheet["C4"].number_format = "0.0"
    worksheet["B5"].number_format = "0.0%"
    worksheet["B6"].number_format = "yyyy-mm-dd"

    for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row):
        if row[0].value is not None:
            row[0].alignment = Alignment(horizontal="right")
        if len(row) > 1 and row[1].value is not None:
            row[1].alignment = Alignment(horizontal="center")

    for cell in worksheet[8]:
        cell.fill = grey_fill
        cell.font = bold_font

    for row in worksheet.iter_rows(min_row=9, max_row=worksheet.max_row):
        for cell in row[2:26]:
            cell.number_format = "0.0"
            cell.border = side_border


def write_excel_report(
    output_path: Path,
    plant: PlantConfig,
    day_ahead_df: pd.DataFrame,
    week_ahead_df: pd.DataFrame,
    submitted_date: date,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        day_ahead_df.to_excel(
            writer,
            sheet_name="DayAheadForecast",
            index=False,
            startrow=7,
        )
        week_ahead_df.to_excel(
            writer,
            sheet_name="WeekAheadForecast",
            index=False,
            startrow=7,
        )

        workbook = writer.book
        day_ws = workbook["DayAheadForecast"]
        week_ws = workbook["WeekAheadForecast"]

        _write_sheet_title_block(
            day_ws,
            "Business Day-Ahead Solar Energy Forecast",
            plant,
            day_ahead_df,
            submitted_date,
        )
        _write_sheet_title_block(
            week_ws,
            "Week-Ahead Solar Energy Forecast",
            plant,
            week_ahead_df,
            submitted_date,
        )

        # Add logo to both sheets
        if LOGO.exists():
            img = Image(LOGO)

            # Resize image to fit approximately within cells I2:I6.
            # A standard row is ~15-20 pixels high, so 5 rows is ~75-100px.
            # We'll set the height and scale the width to maintain aspect ratio.
            original_height = img.height
            img.height = 90
            img.width = img.width * (90 / original_height)

            # A copy of the image is needed for each sheet it's added to.
            day_ws.add_image(copy.copy(img), "M2")
            week_ws.add_image(copy.copy(img), "M2")
        else:
            print(f"Logo file not found at {LOGO}, skipping embedding.")

        for worksheet in (day_ws, week_ws):
            apply_excel_formatting(worksheet)
            worksheet.column_dimensions["A"].width = 42
            worksheet.column_dimensions["B"].width = 20
            for column_letter in [
                "C",
                "D",
                "E",
                "F",
                "G",
                "H",
                "I",
                "J",
                "K",
                "L",
                "M",
                "N",
                "O",
                "P",
                "Q",
                "R",
                "S",
                "T",
                "U",
                "V",
                "W",
                "X",
                "Y",
                "Z",
                "AA",
                "AB",
            ]:
                worksheet.column_dimensions[column_letter].width = 12


def write_lily_csv(output_path: Path, csv_df: pd.DataFrame) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    csv_df.to_csv(output_path, index=False)


def build_output_path(
    plant: PlantConfig,
    report_date: date,
    base_output_dir: Path = BASE_OUTPUT_DIR,
) -> Path:
    filename = plant.output_filename_pattern.format(
        report_date=report_date.strftime("%m-%d-%Y")
    )
    return base_output_dir / plant.output_subdir / filename


def send_email(
    sender_email: str,
    sender_password: str,
    to_emails: list[str],
    subject: str,
    body: str,
    attachments: list[Path],
    cc_emails: list[str] | None = None,
) -> None:
    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = ", ".join(to_emails)
    if cc_emails:
        message["Cc"] = ", ".join(cc_emails)
    message["Subject"] = subject
    message.set_content(body)

    for attachment in attachments:
        mime_type, _ = mimetypes.guess_type(attachment.name)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"

        with attachment.open("rb") as file_handle:
            message.add_attachment(
                file_handle.read(),
                maintype=maintype,
                subtype=subtype,
                filename=attachment.name,
            )

    with smtplib.SMTP_SSL(GMAIL_SMTP_HOST, GMAIL_SMTP_PORT) as smtp:
        smtp.login(sender_email, sender_password)
        smtp.send_message(message)


def send_forecast_emails(generated_files: dict[str, Path], report_date: date) -> None:
    sender_email = DEFAULT_SENDER_EMAIL
    sender_password = CREDS['shiftsumEmail']
    lily_forecast_recipients = EMAILS['Lily Forecast']
    narenco_forecast_recipients = EMAILS['NARENCO Forecast']
    lily_recipient_emails = (
        [email.strip() for email in lily_forecast_recipients if email.strip()]
        if lily_forecast_recipients
        else DEFAULT_RECIPIENT_EMAILS
    )
    narenco_recipient_emails = (
        [email.strip() for email in narenco_forecast_recipients if email.strip()]
        if narenco_forecast_recipients
        else DEFAULT_RECIPIENT_EMAILS
    )

    if not sender_password:
        print(
            "Skipping email delivery. Set FORECAST_EMAIL_PASSWORD to the Gmail app password "
            f"for {sender_email} to enable sending."
        )
        return

    lily_date = (report_date + timedelta(days=1)).strftime("%m-%d-%Y")
    report_date_label = (report_date + timedelta(days=1)).strftime("%m-%d-%Y")

    root = tk.Tk()
    root.withdraw()
    try:
        send_email(
            sender_email=sender_email,
            sender_password=sender_password,
            to_emails=lily_recipient_emails,
            subject=f"Lily Solar Forecast {lily_date}",
            body=(f"""Good Morning,

Please see attached the Lily Solar Forecasts for {report_date_label}. Please Let the NCC know if there are any questions. 
This email was sent automatically by the forecasting tool.

Thank you,
NCC Automations
"""
            ),
            attachments=[generated_files["LILY"]],
        )

        send_email(
            sender_email=sender_email,
            sender_password=sender_password,
            to_emails=narenco_recipient_emails,
            subject=f"Bluebird, Cardinal, and Cherry Blossom Solar Forecasts {report_date_label}",
            body=(f"""Good Morning,

Please see attached the Bluebird, Cardinal, and Cherry Blossom Solar Forecasts for {report_date_label}. Please Let the NCC know if there are any questions. 
This email was sent automatically by the forecasting tool.

Thank you,
NCC Automations
"""
            ),
            attachments=[
                generated_files["BLUEBIRD"],
                generated_files["CARDINAL"],
                generated_files["CHERRY BLOSSOM"],
            ],
        )
        messagebox.showinfo("Success", "Forecast emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Email Error", f"Failed to send forecast emails: {e}")
    finally:
        root.destroy()


def generate_reports(
    report_date: date | None = None,
    base_output_dir: Path = BASE_OUTPUT_DIR,
) -> dict[str, Path]:
    submitted_date = report_date or date.today()
    day_ahead_start = submitted_date + timedelta(days=1)
    day_ahead_end = submitted_date + timedelta(days=2)
    week_ahead_start = submitted_date + timedelta(days=1)
    week_ahead_end = submitted_date + timedelta(days=10)

    outputs: dict[str, Path] = {}

    with get_pvsyst_connection() as connection:
        for plant_name in ("BLUEBIRD", "CARDINAL", "CHERRY BLOSSOM"):
            plant = PLANTS[plant_name]
            hourly = fetch_pvsyst_hourly(
                plant=plant,
                start_date=week_ahead_start,
                end_date=week_ahead_end,
                connection=connection,
            )
            day_ahead_hourly = hourly[
                hourly["ForecastDate"].isin([day_ahead_start, day_ahead_end])
            ].copy()
            week_ahead_hourly = hourly.copy()

            day_ahead_df = build_excel_report_frame(day_ahead_hourly, plant)
            week_ahead_df = build_excel_report_frame(week_ahead_hourly, plant)

            output_path = build_output_path(plant, day_ahead_start, base_output_dir)
            write_excel_report(
                output_path=output_path,
                plant=plant,
                day_ahead_df=day_ahead_df,
                week_ahead_df=week_ahead_df,
                submitted_date=submitted_date,
            )
            outputs[plant.name] = output_path

        lily = PLANTS["LILY"]
        lily_hourly = fetch_pvsyst_hourly(
            plant=lily,
            start_date=submitted_date,
            end_date=submitted_date + timedelta(days=10),
            connection=connection,
        )
        lily_csv_df = build_lily_csv_frame(lily_hourly, lily)
        lily_output = build_output_path(
            lily,
            submitted_date + timedelta(days=1),
            base_output_dir,
        )
        write_lily_csv(lily_output, lily_csv_df)
        outputs[lily.name] = lily_output

    return outputs


def main() -> None:
    temp_root = Path(tempfile.mkdtemp(prefix="forecasting_reports_"))

    try:
        generated_files = generate_reports(base_output_dir=temp_root)
        send_forecast_emails(generated_files, date.today())
        print("Forecasting reports generated and email delivery attempted.")
        for plant_name, output_path in generated_files.items():
            print(f" - {plant_name}: {output_path.name}")
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


if __name__ == "__main__":
    main()
