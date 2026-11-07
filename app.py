"""
Bulk Email Sender (Gmail + App Password + Excel)
Run with:  streamlit run app.py
"""
import html
import io
import mimetypes
import random
import re
import smtplib
import time
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr

import pandas as pd
import streamlit as st

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465  # SSL
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
PLACEHOLDER_RE = re.compile(r"\{([^{}]+)\}")

st.set_page_config(page_title="Bulk Email Sender", page_icon="✉️", layout="centered")


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def render(template: str, row: pd.Series) -> str:
    """Replace {Column Name} with the row's value (case-insensitive).
    Unknown placeholders are left untouched."""

    def repl(match):
        key = match.group(1).strip().lower()
        for col in row.index:
            if str(col).strip().lower() == key:
                val = row[col]
                return "" if pd.isna(val) else str(val).strip()
        return match.group(0)

    return PLACEHOLDER_RE.sub(repl, template)


def guess_column(columns, keywords):
    """Pick the column whose name contains one of the keywords."""
    for kw in keywords:
        for c in columns:
            if kw in str(c).lower():
                return c
    return None


def build_message(sender_name, sender_email, to_addr, subject, body, attachments):
    msg = EmailMessage()
    msg["From"] = formataddr((sender_name, sender_email)) if sender_name else sender_email
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)
    # simple HTML version (keeps line breaks)
    msg.add_alternative(
        "<html><body style='font-family:Arial,sans-serif;font-size:14px'>"
        + html.escape(body).replace("\n", "<br>")
        + "</body></html>",
        subtype="html",
    )
    for name, data in attachments:
        ctype, _ = mimetypes.guess_type(name)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
    return msg


class _ForceFamily:
    """Temporarily make DNS lookups return only IPv4 or only IPv6 addresses."""

    def __init__(self, family):
        self.family = family

    def __enter__(self):
        import socket
        self._orig = socket.getaddrinfo
        fam = self.family

        def patched(host, port, family=0, type=0, proto=0, flags=0):
            return self._orig(host, port, fam or family, type, proto, flags)

        socket.getaddrinfo = patched

    def __exit__(self, *exc):
        import socket
        socket.getaddrinfo = self._orig


def connect(sender_email, app_password):
    """Try several ways to reach Gmail: IPv4 / IPv6 / auto, SSL 465 / STARTTLS 587."""
    import socket
    import ssl

    ctx = ssl.create_default_context()
    errors = []
    families = [("IPv4", socket.AF_INET), ("IPv6", socket.AF_INET6), ("auto", 0)]
    for fam_name, fam in families:
        for mode in ("ssl465", "starttls587"):
            server = None
            try:
                with _ForceFamily(fam):
                    if mode == "ssl465":
                        server = smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=30, context=ctx)
                    else:
                        server = smtplib.SMTP(SMTP_HOST, 587, timeout=30)
                        server.ehlo()
                        server.starttls(context=ctx)
                        server.ehlo()
                server.login(sender_email, app_password)
                return server
            except smtplib.SMTPAuthenticationError:
                raise
            except Exception as e:
                errors.append(f"{fam_name}/{mode}: {e}")
                try:
                    if server:
                        server.close()
                except Exception:
                    pass
    raise ConnectionError(
        "All connection methods failed. Antivirus email scanning, a firewall or the "
        "network is probably blocking Gmail. Details: " + " | ".join(errors)
    )


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
        ws = writer.sheets["Results"]
        for col_cells in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(width + 2, 10), 50)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Sidebar: Gmail credentials
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("🔐 Gmail login")
    sender_email = st.text_input("Your Gmail address", placeholder="you@gmail.com")
    app_password = st.text_input(
        "Google App Password",
        type="password",
        help="16-character password from Google Account → Security → 2-Step Verification → App passwords.",
    )
    sender_name = st.text_input("Sender display name (optional)", placeholder="Rahul from ABC Pvt Ltd")
    st.caption("Your password is used only for this session and is never saved.")
    st.divider()
    st.caption(
        "Limits: a free Gmail account can send about 500 emails/day "
        "(Google Workspace about 2,000/day)."
    )

st.title("✉️ Bulk Email Sender")

# ----------------------------------------------------------------------------
# Step 1: Upload Excel
# ----------------------------------------------------------------------------
st.subheader("1. Upload Excel sheet")
excel_file = st.file_uploader("Excel file with company emails (.xlsx / .xls)", type=["xlsx", "xls"])

df = None
email_col = name_col = company_col = None

if excel_file is not None:
    try:
        df = pd.read_excel(excel_file, dtype=str)
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
        df = None

if df is not None:
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    st.success(f"Loaded {len(df)} rows.")
    st.dataframe(df.head(10), use_container_width=True)

    cols = list(df.columns)
    none_opt = "(none)"

    guessed_email = guess_column(cols, ["email", "e-mail", "mail"])
    guessed_name = guess_column(cols, ["name", "contact", "person"])
    guessed_company = guess_column(cols, ["company", "organization", "organisation", "firm"])

    c1, c2, c3 = st.columns(3)
    email_col = c1.selectbox(
        "Email column", cols, index=cols.index(guessed_email) if guessed_email else 0
    )
    name_col = c2.selectbox(
        "Name column", [none_opt] + cols,
        index=([none_opt] + cols).index(guessed_name) if guessed_name else 0,
    )
    company_col = c3.selectbox(
        "Company column", [none_opt] + cols,
        index=([none_opt] + cols).index(guessed_company) if guessed_company else 0,
    )

    # Normalised helper columns used for the {name} and {company} placeholders
    work = df.copy()
    if name_col != none_opt:
        work["name"] = work[name_col]
    if company_col != none_opt:
        work["company"] = work[company_col]

    placeholders = ", ".join("{" + c + "}" for c in work.columns)
    st.caption(f"Placeholders you can use in subject/body: {placeholders}")

# ----------------------------------------------------------------------------
# Step 2: Compose
# ----------------------------------------------------------------------------
st.subheader("2. Write your message")
subject_tpl = st.text_input("Subject", value="Hello {name} – Partnership with {company}")
body_tpl = st.text_area(
    "Message body",
    height=220,
    value=(
        "Dear {name},\n\n"
        "I hope you are doing well. I am reaching out to {company} regarding ...\n\n"
        "Please let me know a convenient time to talk.\n\n"
        "Best regards,\nYour Name"
    ),
)

attach_files = st.file_uploader(
    "Attachments (optional – sent with every email)", accept_multiple_files=True
)
attachments = [(f.name, f.getvalue()) for f in attach_files] if attach_files else []
if attachments:
    total_mb = sum(len(d) for _, d in attachments) / 1_048_576
    st.caption(f"{len(attachments)} file(s), {total_mb:.1f} MB per email. Gmail's limit is 25 MB.")
    if total_mb > 24:
        st.warning("Attachments are close to or over Gmail's 25 MB limit.")

# ----------------------------------------------------------------------------
# Step 3: Sending options
# ----------------------------------------------------------------------------
st.subheader("3. Sending options")
d1, d2 = st.columns(2)
delay_range = d1.slider("Delay between emails (seconds)", 1, 60, (5, 12))
max_per_run = d2.number_input("Max emails in this run", min_value=1, max_value=2000, value=100, step=10)
skip_sent = st.checkbox("Skip rows already marked 'Sent' (if Excel has a Status column)", value=True)

# ----------------------------------------------------------------------------
# Preview + test email
# ----------------------------------------------------------------------------
ready = df is not None and sender_email and app_password

if df is not None:
    with st.expander("👀 Preview first email"):
        first = work.iloc[0]
        st.markdown(f"**To:** {first[email_col]}")
        st.markdown(f"**Subject:** {render(subject_tpl, first)}")
        st.text(render(body_tpl, first))

    if st.button("📨 Send a TEST email to myself", disabled=not ready):
        try:
            first = work.iloc[0]
            with connect(sender_email, app_password.replace(" ", "")) as server:
                msg = build_message(
                    sender_name, sender_email, sender_email,
                    "[TEST] " + render(subject_tpl, first),
                    render(body_tpl, first), attachments,
                )
                server.send_message(msg)
            st.success(f"Test email sent to {sender_email}. Check your inbox.")
        except smtplib.SMTPAuthenticationError:
            st.error("Login failed. Check your Gmail address and App Password.")
        except Exception as e:
            st.error(f"Test failed: {e}")

# ----------------------------------------------------------------------------
# Send
# ----------------------------------------------------------------------------
st.divider()
send_clicked = st.button("🚀 Send emails to everyone in the sheet", type="primary", disabled=not ready)
if not ready:
    st.caption("Fill in Gmail login (sidebar) and upload the Excel file to enable sending.")

if send_clicked and ready:
    result = work.copy()
    for col in ("Status", "Sent At", "Error"):
        if col not in result.columns:
            result[col] = ""
        result[col] = result[col].fillna("")

    password = app_password.replace(" ", "")
    try:
        server = connect(sender_email, password)
    except smtplib.SMTPAuthenticationError:
        st.error("Login failed. Use a 16-character Google App Password (not your normal password) "
                 "and make sure 2-Step Verification is on.")
        st.stop()
    except Exception as e:
        st.error(f"Could not connect to Gmail: {e}")
        st.stop()

    progress = st.progress(0.0)
    status_box = st.empty()
    log_box = st.empty()
    log_lines = []

    # decide which rows to process
    todo = []
    seen = set()
    for idx, row in result.iterrows():
        addr = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
        if skip_sent and str(row["Status"]).strip().lower() == "sent":
            continue
        if not EMAIL_RE.match(addr):
            result.at[idx, "Status"] = "Failed"
            result.at[idx, "Error"] = "Invalid or missing email"
            continue
        if addr.lower() in seen:
            result.at[idx, "Status"] = "Skipped"
            result.at[idx, "Error"] = "Duplicate email"
            continue
        seen.add(addr.lower())
        todo.append((idx, addr))
    todo = todo[: int(max_per_run)]

    sent = failed = 0
    aborted = False
    for n, (idx, addr) in enumerate(todo, start=1):
        row = result.loc[idx]
        status_box.info(f"Sending {n}/{len(todo)} → {addr}")
        try:
            msg = build_message(
                sender_name, sender_email, addr,
                render(subject_tpl, row), render(body_tpl, row), attachments,
            )
            try:
                server.send_message(msg)
            except smtplib.SMTPServerDisconnected:
                server = connect(sender_email, password)  # reconnect once
                server.send_message(msg)
            result.at[idx, "Status"] = "Sent"
            result.at[idx, "Sent At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            result.at[idx, "Error"] = ""
            sent += 1
            log_lines.append(f"✅ {addr}")
        except smtplib.SMTPDataError as e:
            # Usually the daily sending limit
            result.at[idx, "Status"] = "Failed"
            result.at[idx, "Error"] = f"{e.smtp_code} {e.smtp_error!r}"[:200]
            failed += 1
            log_lines.append(f"❌ {addr}: {e.smtp_code}")
            if e.smtp_code in (550, 552, 421, 454) and "limit" in str(e.smtp_error).lower():
                st.error("Gmail sending limit reached. Stopping – try again tomorrow.")
                aborted = True
        except Exception as e:
            result.at[idx, "Status"] = "Failed"
            result.at[idx, "Error"] = str(e)[:200]
            failed += 1
            log_lines.append(f"❌ {addr}: {e}")

        progress.progress(n / len(todo))
        log_box.text("\n".join(log_lines[-15:]))
        if aborted:
            break
        if n < len(todo):
            time.sleep(random.uniform(*delay_range))

    try:
        server.quit()
    except Exception:
        pass

    # Keep the original columns first, then the helper columns removed
    drop_cols = [c for c in ("name", "company") if c in result.columns and c not in df.columns]
    result = result.drop(columns=drop_cols)

    status_box.empty()
    st.success(f"Done. Sent: {sent} | Failed: {failed} | Total processed: {len(todo)}")
    st.session_state["result_bytes"] = df_to_excel_bytes(result)
    st.session_state["result_df"] = result

# ----------------------------------------------------------------------------
# Download results
# ----------------------------------------------------------------------------
if "result_bytes" in st.session_state:
    st.subheader("4. Download results")
    st.dataframe(st.session_state["result_df"], use_container_width=True)
    st.download_button(
        "⬇️ Download Excel with Sent/Failed status",
        data=st.session_state["result_bytes"],
        file_name=f"email_results_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
