import io
import os
import re
import sqlite3
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

import pandas as pd
import pdfplumber
import psycopg2
from openpyxl.styles import Alignment, Font, PatternFill
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# Set these in your cloud host's Environment Variables (or paste token for local testing)
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_TELEGRAM_BOT_TOKEN")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
SQLITE_PATH = "tpddl_records.db"
USE_POSTGRES = bool(DATABASE_URL)


# ==========================================
# 1. HEALTH CHECK SERVER (KEEPS FREE HOSTING ALIVE)
# ==========================================
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(b"TP-DDL Bot is running")

    def log_message(self, format, *args):
        pass  # Silence HTTP access logs


def start_health_server():
    port = int(os.getenv("PORT", "8080"))
    server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()


# ==========================================
# 2. DATABASE LAYER (POSTGRES IN CLOUD / SQLITE LOCALLY)
# ==========================================
def get_conn():
    if USE_POSTGRES:
        db_url = DATABASE_URL.replace("postgres://", "postgresql://", 1)
        return psycopg2.connect(db_url)
    return sqlite3.connect(SQLITE_PATH)


def ph(query: str) -> str:
    """Converts SQLite '?' placeholders to PostgreSQL '%s' when using Postgres."""
    return query.replace("?", "%s") if USE_POSTGRES else query


def init_db():
    pk_col = "SERIAL PRIMARY KEY" if USE_POSTGRES else "INTEGER PRIMARY KEY AUTOINCREMENT"
    conn = get_conn()
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute(f"""
                    CREATE TABLE IF NOT EXISTS service_orders (
                        id {pk_col},
                        user_id BIGINT NOT NULL,
                        username TEXT,
                        upload_date TEXT NOT NULL,
                        upload_timestamp TEXT NOT NULL,
                        source_file TEXT NOT NULL,
                        service_order_no TEXT NOT NULL,
                        print_date TEXT,
                        notification_no TEXT,
                        notification_type TEXT,
                        notification_date TEXT,
                        ca_no TEXT,
                        installation_no TEXT,
                        customer_name TEXT,
                        phone_no TEXT,
                        address TEXT,
                        zone TEXT,
                        pole_no TEXT,
                        ws_no TEXT,
                        sanctioned_load_kw TEXT,
                        billing_class TEXT,
                        rate_category TEXT,
                        old_meter_no TEXT,
                        new_meter_no TEXT,
                        meter_make TEXT,
                        construction_class TEXT,
                        removal_date TEXT,
                        meter_location TEXT,
                        removal_kwh TEXT,
                        removal_kvah TEXT,
                        removal_mdi_kw TEXT,
                        removal_mdi_kva TEXT,
                        install_kwh TEXT,
                        install_kvah TEXT,
                        removed_seals TEXT,
                        installed_seals TEXT,
                        van_no TEXT,
                        agency_name TEXT,
                        staff_code TEXT,
                        fse_name TEXT,
                        employee_code TEXT,
                        cable_marking TEXT,
                        materials_consumed TEXT,
                        additional_remarks TEXT,
                        UNIQUE(user_id, service_order_no)
                    )
                """)
    finally:
        conn.close()


# ==========================================
# 3. TP-DDL PDF PARSER
# ==========================================
def clean_str(val: str) -> str:
    if not val:
        return ""
    return re.sub(r"\s+", " ", val).strip()


def extract_regex(pattern: str, text: str, group: int = 1, flags: int = 0) -> str:
    match = re.search(pattern, text, flags)
    return clean_str(match.group(group)) if match else ""


def parse_materials_page(page) -> str:
    consumed = []
    tables = page.extract_tables() or []

    for table in tables:
        for row in table:
            cells = [clean_str(c) for c in row]
            n = len(cells)
            i = 0
            while i < n:
                cell = cells[i]

                m_3digit = re.match(r"^(\d{3})\s+(.+)$", cell)
                if m_3digit:
                    code, desc = m_3digit.group(1), m_3digit.group(2)
                    inst_val, rem_val = "", ""

                    if i + 1 < n and cells[i + 1] in ("EA", "Mt"):
                        if i + 2 < n:
                            inst_val = cells[i + 2]
                        if i + 3 < n:
                            rem_val = cells[i + 3]
                        i += 4
                    else:
                        if i + 1 < n and cells[i + 1]:
                            parts = cells[i + 1].split()
                            if len(parts) == 2:
                                inst_val, rem_val = parts[0], parts[1]
                            elif len(parts) == 1:
                                inst_val = parts[0]
                        if i + 2 < n and cells[i + 2] and not re.match(r"^\d{3}\s+", cells[i + 2]):
                            if not rem_val:
                                rem_val = cells[i + 2]
                        i += 2

                    if any(ch.isdigit() for ch in inst_val) or any(ch.isdigit() for ch in rem_val):
                        details = []
                        if any(ch.isdigit() for ch in inst_val):
                            details.append(f"Inst: {inst_val}")
                        if any(ch.isdigit() for ch in rem_val):
                            details.append(f"Rem: {rem_val}")
                        consumed.append(f"{code} {desc} ({', '.join(details)})")
                    continue

                if re.match(r"^(10|11|12|13|14|15|16|19)$", cell) and i + 2 < n:
                    code = cell
                    desc = cells[i + 1]
                    inst_val = cells[i + 2]
                    rem_val = cells[i + 3] if (i + 3 < n and cells[i + 3].isdigit()) else ""
                    if desc and (any(ch.isdigit() for ch in inst_val) or any(ch.isdigit() for ch in rem_val)):
                        details = []
                        if any(ch.isdigit() for ch in inst_val):
                            details.append(f"Inst: {inst_val}")
                        if any(ch.isdigit() for ch in rem_val):
                            details.append(f"Rem: {rem_val}")
                        consumed.append(f"BA-{code} {desc} ({', '.join(details)})")
                    i += 4
                    continue

                i += 1

    seen = set()
    unique_consumed = []
    for item in consumed:
        if item not in seen:
            seen.add(item)
            unique_consumed.append(item)

    return "; ".join(unique_consumed)


def parse_tpddl_pdf(pdf_bytes: bytes) -> dict:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page1_text = pdf.pages[0].extract_text() or "" if len(pdf.pages) >= 1 else ""
        page2_text = pdf.pages[1].extract_text() or "" if len(pdf.pages) >= 2 else ""
        page1_tables = pdf.pages[0].extract_tables() or [] if len(pdf.pages) >= 1 else []
        materials_summary = parse_materials_page(pdf.pages[2]) if len(pdf.pages) >= 3 else ""

    service_order_no = extract_regex(r"Service Order No\.?:\s*(\d+)", page1_text)
    if not service_order_no:
        raise ValueError("Could not find 'Service Order No.' — please upload a valid TP-DDL form PDF.")

    print_date = extract_regex(r"Print Date:\s*(\d{2}-\d{2}-\d{4})", page1_text)
    customer_name = extract_regex(r"Name:\s*([^\n]+)", page1_text)

    mobile_remark = extract_regex(r"Do you want to stop paper bill\?\s+YES\s+(\d{10})", page1_text)
    masked_phone = extract_regex(r"Phone:\s*([X\d]+)", page1_text)
    phone_no = mobile_remark if mobile_remark else masked_phone

    address = extract_regex(r"Address:\s*(.+?)\s*Installation No\.?:", page1_text, flags=re.DOTALL)
    installation_no = extract_regex(r"Installation No\.?:\s*(\d+)", page1_text)
    ca_no = extract_regex(r"CA No\.?:\s*(\d+)", page1_text)
    notification_no = extract_regex(r"Notification No:\s*(\d+)", page1_text)
    notification_type = extract_regex(r"Notification Type:\s*([^\n]+)", page1_text)
    notification_date = extract_regex(r"Notification Date:\s*(\d{2}-\d{2}-\d{4})", page1_text)

    billing_class = extract_regex(r"Billing Class:\s*([^\n]+)", page1_text)
    rate_category = extract_regex(r"Rate Category:\s*([^\n]+)", page1_text)
    ws_no = extract_regex(r"WS No\.?:\s*([^\n]+)", page1_text)
    sanctioned_load_kw = extract_regex(r"Sanctioned Load\(kW\):\s*([^\n]+)", page1_text)
    zone = extract_regex(r"Zone:\s*([^\n]+)", page1_text)
    pole_no = extract_regex(r"Pole:\s*([^\n]+)", page1_text)

    van_no = extract_regex(r"Van No:\s*([A-Z0-9-]+(?:\n[A-Z0-9-]+)?)", page1_text).replace(" ", "")
    additional_remarks = extract_regex(
        r"Additional Remarks:\s*(.+?)\s*Seal Remark:", page1_text, flags=re.DOTALL
    )
    cable_marking_match = re.search(r"Cable Marking\s*From\s*(\d+)\s*\S+\s*(\d+)", page1_text)
    cable_marking = (
        f"From {cable_marking_match.group(1)} To {cable_marking_match.group(2)}"
        if cable_marking_match
        else ""
    )

    old_meter_no = ""
    new_meter_no = ""
    meter_make = ""
    construction_class = ""
    removal_date = ""
    meter_location = ""
    removed_seals = []
    installed_seals = []
    readings = {
        "removal_kwh": "",
        "removal_kvah": "",
        "removal_mdi_kw": "",
        "removal_mdi_kva": "",
        "install_kwh": "",
        "install_kvah": "",
    }

    for table in page1_tables:
        for row in table:
            cells = [clean_str(c) for c in row]
            if not cells:
                continue
            label = cells[0]

            if label == "Meter No.":
                old_meter_no = cells[1] if len(cells) > 1 else ""
                new_meter_no = cells[2] if len(cells) > 2 else ""
            elif label == "Construction Class":
                construction_class = cells[1] if len(cells) > 1 else ""
            elif label == "Meter Make":
                meter_make = cells[1] if len(cells) > 1 else ""
            elif label == "Removal Date":
                removal_date = cells[1] if len(cells) > 1 else ""
            elif label == "Meter Location":
                meter_location = cells[1] if len(cells) > 1 else ""

            if len(cells) >= 6:
                rem_cell = cells[-2]
                inst_cell = cells[-1]
                for token in rem_cell.split():
                    if re.match(r"^[A-Z]{2}\d{5,}$", token):
                        removed_seals.append(token)
                for token in inst_cell.split():
                    if re.match(r"^[A-Z]{2}\d{5,}$", token):
                        installed_seals.append(token)

            if label == "Kwh":
                readings["removal_kwh"] = cells[1] if len(cells) > 1 else ""
                readings["install_kwh"] = cells[3] if len(cells) > 3 else ""
            elif label == "Kvah":
                readings["removal_kvah"] = cells[1] if len(cells) > 1 else ""
                readings["install_kvah"] = cells[3] if len(cells) > 3 else ""
            elif label == "MDI (Kw)":
                readings["removal_mdi_kw"] = cells[1] if len(cells) > 1 else ""
            elif label == "MDI (Kva)":
                readings["removal_mdi_kva"] = cells[1] if len(cells) > 1 else ""

    agency_name = extract_regex(
        r"Name:.*?\n.*?Name:\s*([A-Z0-9 &.,()-]+)\nRelationship:", page2_text, flags=re.DOTALL
    )
    staff_code = extract_regex(r"Staff Code:\s*(\d+)", page2_text)
    fse_name = extract_regex(r"Name:\s*([^\n]+)\nEmployee Code:", page2_text)
    employee_code = extract_regex(r"Employee Code:\s*(\d+)", page2_text)

    return {
        "service_order_no": service_order_no,
        "print_date": print_date,
        "notification_no": notification_no,
        "notification_type": notification_type,
        "notification_date": notification_date,
        "ca_no": ca_no,
        "installation_no": installation_no,
        "customer_name": customer_name,
        "phone_no": phone_no,
        "address": address,
        "zone": zone,
        "pole_no": pole_no,
        "ws_no": ws_no,
        "sanctioned_load_kw": sanctioned_load_kw,
        "billing_class": billing_class,
        "rate_category": rate_category,
        "old_meter_no": old_meter_no,
        "new_meter_no": new_meter_no,
        "meter_make": meter_make,
        "construction_class": construction_class,
        "removal_date": removal_date,
        "meter_location": meter_location,
        "removal_kwh": readings["removal_kwh"],
        "removal_kvah": readings["removal_kvah"],
        "removal_mdi_kw": readings["removal_mdi_kw"],
        "removal_mdi_kva": readings["removal_mdi_kva"],
        "install_kwh": readings["install_kwh"],
        "install_kvah": readings["install_kvah"],
        "removed_seals": ", ".join(removed_seals),
        "installed_seals": ", ".join(installed_seals),
        "van_no": van_no,
        "agency_name": agency_name,
        "staff_code": staff_code,
        "fse_name": fse_name,
        "employee_code": employee_code,
        "cable_marking": cable_marking,
        "materials_consumed": materials_summary,
        "additional_remarks": additional_remarks,
    }


# ==========================================
# 4. TELEGRAM BOT HANDLERS
# ==========================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Send your daily TP-DDL Service Order PDFs here. Each user's records are stored separately.\n\n"
        "Commands:\n"
        "/export - Download your complete Excel (.xlsx) sheet with all dates & records\n"
        "/today - Download Excel sheet for today's uploads only\n"
        "/status - View your logged PDF count & date range\n"
        "/delete <ServiceOrderNo> - Remove a specific Service Order\n"
        "/clear - Delete all your saved records"
    )


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    document = update.message.document
    status_msg = await update.message.reply_text(f"Parsing `{document.file_name}`...", parse_mode="Markdown")

    try:
        tg_file = await document.get_file()
        pdf_bytes = bytes(await tg_file.download_as_bytearray())
        data = parse_tpddl_pdf(pdf_bytes)

        now = datetime.now()
        upload_date = now.strftime("%Y-%m-%d")
        upload_ts = now.strftime("%Y-%m-%d %H:%M:%S")

        sql = ph("""
            INSERT INTO service_orders (
                user_id, username, upload_date, upload_timestamp, source_file,
                service_order_no, print_date, notification_no, notification_type, notification_date,
                ca_no, installation_no, customer_name, phone_no, address,
                zone, pole_no, ws_no, sanctioned_load_kw, billing_class, rate_category,
                old_meter_no, new_meter_no, meter_make, construction_class, removal_date, meter_location,
                removal_kwh, removal_kvah, removal_mdi_kw, removal_mdi_kva, install_kwh, install_kvah,
                removed_seals, installed_seals, van_no, agency_name, staff_code, fse_name, employee_code,
                cable_marking, materials_consumed, additional_remarks
            ) VALUES (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?
            )
            ON CONFLICT(user_id, service_order_no) DO UPDATE SET
                upload_date=excluded.upload_date,
                upload_timestamp=excluded.upload_timestamp,
                source_file=excluded.source_file,
                print_date=excluded.print_date,
                notification_no=excluded.notification_no,
                notification_type=excluded.notification_type,
                notification_date=excluded.notification_date,
                ca_no=excluded.ca_no,
                installation_no=excluded.installation_no,
                customer_name=excluded.customer_name,
                phone_no=excluded.phone_no,
                address=excluded.address,
                zone=excluded.zone,
                pole_no=excluded.pole_no,
                ws_no=excluded.ws_no,
                sanctioned_load_kw=excluded.sanctioned_load_kw,
                billing_class=excluded.billing_class,
                rate_category=excluded.rate_category,
                old_meter_no=excluded.old_meter_no,
                new_meter_no=excluded.new_meter_no,
                meter_make=excluded.meter_make,
                construction_class=excluded.construction_class,
                removal_date=excluded.removal_date,
                meter_location=excluded.meter_location,
                removal_kwh=excluded.removal_kwh,
                removal_kvah=excluded.removal_kvah,
                removal_mdi_kw=excluded.removal_mdi_kw,
                removal_mdi_kva=excluded.removal_mdi_kva,
                install_kwh=excluded.install_kwh,
                install_kvah=excluded.install_kvah,
                removed_seals=excluded.removed_seals,
                installed_seals=excluded.installed_seals,
                van_no=excluded.van_no,
                agency_name=excluded.agency_name,
                staff_code=excluded.staff_code,
                fse_name=excluded.fse_name,
                employee_code=excluded.employee_code,
                cable_marking=excluded.cable_marking,
                materials_consumed=excluded.materials_consumed,
                additional_remarks=excluded.additional_remarks
        """)

        params = (
            user.id,
            user.username or user.first_name,
            upload_date,
            upload_ts,
            document.file_name,
            data["service_order_no"],
            data["print_date"],
            data["notification_no"],
            data["notification_type"],
            data["notification_date"],
            data["ca_no"],
            data["installation_no"],
            data["customer_name"],
            data["phone_no"],
            data["address"],
            data["zone"],
            data["pole_no"],
            data["ws_no"],
            data["sanctioned_load_kw"],
            data["billing_class"],
            data["rate_category"],
            data["old_meter_no"],
            data["new_meter_no"],
            data["meter_make"],
            data["construction_class"],
            data["removal_date"],
            data["meter_location"],
            data["removal_kwh"],
            data["removal_kvah"],
            data["removal_mdi_kw"],
            data["removal_mdi_kva"],
            data["install_kwh"],
            data["install_kvah"],
            data["removed_seals"],
            data["installed_seals"],
            data["van_no"],
            data["agency_name"],
            data["staff_code"],
            data["fse_name"],
            data["employee_code"],
            data["cable_marking"],
            data["materials_consumed"],
            data["additional_remarks"],
        )

        conn = get_conn()
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(sql, params)
        finally:
            conn.close()

        await status_msg.edit_text(
            f"Saved Service Order `{data['service_order_no']}`\n"
            f"• **Type:** {data['notification_type']}\n"
            f"• **CA No:** `{data['ca_no']}` | **Meter:** `{data['old_meter_no']}`\n"
            f"• **Customer:** {data['customer_name']}\n"
            f"• **Work Date:** {data['removal_date'] or data['print_date']}\n"
            f"• **Kwh / Kvah:** {data['removal_kwh']} / {data['removal_kvah']}\n\n"
            f"Send /export anytime to download your Excel file.",
            parse_mode="Markdown",
        )

    except Exception as e:
        await status_msg.edit_text(f"Error processing `{document.file_name}`: {e}", parse_mode="Markdown")


def generate_excel_buffer(user_id: int, date_filter: str = None):
    query = """
        SELECT
            upload_date AS "Upload Date",
            upload_timestamp AS "Upload Timestamp",
            service_order_no AS "Service Order No",
            print_date AS "Form Print Date",
            removal_date AS "Removal / Work Date",
            notification_type AS "Activity Type",
            notification_no AS "Notification No",
            notification_date AS "Notification Date",
            ca_no AS "CA No",
            installation_no AS "Installation No",
            customer_name AS "Customer Name",
            phone_no AS "Phone No",
            address AS "Address",
            zone AS "Zone",
            pole_no AS "Pole No",
            ws_no AS "WS No",
            sanctioned_load_kw AS "Sanctioned Load (kW)",
            billing_class AS "Billing Class",
            rate_category AS "Rate Category",
            old_meter_no AS "Old Meter No",
            new_meter_no AS "New Meter No",
            meter_make AS "Meter Make",
            construction_class AS "Construction Class",
            meter_location AS "Meter Location",
            removal_kwh AS "Removal Kwh",
            removal_kvah AS "Removal Kvah",
            removal_mdi_kw AS "Removal MDI (kW)",
            removal_mdi_kva AS "Removal MDI (kVA)",
            install_kwh AS "Install Kwh",
            install_kvah AS "Install Kvah",
            removed_seals AS "Removed Seals",
            installed_seals AS "Installed Seals",
            van_no AS "Van No",
            agency_name AS "Agency / BA Name",
            staff_code AS "Staff Code",
            fse_name AS "FSE Name",
            employee_code AS "Employee Code",
            cable_marking AS "Cable Marking",
            materials_consumed AS "Materials Consumed (Inst/Rem)",
            additional_remarks AS "Additional Remarks",
            source_file AS "Source PDF"
        FROM service_orders
        WHERE user_id = ?
    """
    params = [user_id]
    if date_filter:
        query += " AND upload_date = ?"
        params.append(date_filter)
    query += " ORDER BY upload_date ASC, id ASC"

    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(ph(query), params)
            rows = cur.fetchall()
            cols = [desc[0] for desc in cur.description]
    finally:
        conn.close()

    if not rows:
        return None, 0

    df = pd.DataFrame(rows, columns=cols)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="TPDDL_Service_Orders")
        ws = writer.sheets["TPDDL_Service_Orders"]

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            col_letter = col[0].column_letter
            header_cell = col[0]
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(horizontal="center", vertical="center")

            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col_letter].width = min(max_len + 4, 48)

    buffer.seek(0)
    return buffer, len(df)


async def export_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    excel_buf, count = generate_excel_buffer(user.id)

    if not excel_buf:
        await update.message.reply_text("You have no saved PDF records yet. Upload a PDF first.")
        return

    filename = f"TPDDL_Report_{user.first_name}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    await update.message.reply_document(
        document=excel_buf,
        filename=filename,
        caption=f"Exported {count} total record(s) with dates.",
    )


async def export_today(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    today_str = datetime.now().strftime("%Y-%m-%d")
    excel_buf, count = generate_excel_buffer(user.id, date_filter=today_str)

    if not excel_buf:
        await update.message.reply_text(f"No PDFs uploaded today ({today_str}).")
        return

    filename = f"TPDDL_Today_{user.first_name}_{today_str}.xlsx"
    await update.message.reply_document(
        document=excel_buf,
        filename=filename,
        caption=f"Exported {count} record(s) uploaded on {today_str}.",
    )


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(
                ph("SELECT COUNT(*), MIN(upload_date), MAX(upload_date) FROM service_orders WHERE user_id = ?"),
                (user.id,),
            )
            count, min_date, max_date = cur.fetchone()
    finally:
        conn.close()

    if count == 0:
        await update.message.reply_text("No records saved yet. Send a PDF to begin.")
    else:
        await update.message.reply_text(
            f"Saved Service Orders: {count}\nDate Range: {min_date} to {max_date}\nUse /export to download."
        )


async def delete_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not context.args:
        await update.message.reply_text("Usage: `/delete 000218885530`", parse_mode="Markdown")
        return

    so_no = context.args[0].strip()
    conn = get_conn()
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute(
                    ph("DELETE FROM service_orders WHERE user_id = ? AND service_order_no = ?"),
                    (user.id, so_no),
                )
                deleted = cur.rowcount
    finally:
        conn.close()

    if deleted:
        await update.message.reply_text(f"Deleted Service Order `{so_no}`.", parse_mode="Markdown")
    else:
        await update.message.reply_text(f"Service Order `{so_no}` was not found in your records.", parse_mode="Markdown")


async def clear_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    conn = get_conn()
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute(ph("DELETE FROM service_orders WHERE user_id = ?"), (user.id,))
    finally:
        conn.close()
    await update.message.reply_text("All your saved records have been cleared.")


def main():
    init_db()
    start_health_server()
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("export", export_all))
    app.add_handler(CommandHandler("today", export_today))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("delete", delete_order))
    app.add_handler(CommandHandler("clear", clear_data))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))

    db_mode = "PostgreSQL (Cloud)" if USE_POSTGRES else "SQLite (Local)"
    print(f"TP-DDL Telegram Bot is running [{db_mode}]...")
    app.run_polling()


if __name__ == "__main__":
    main()