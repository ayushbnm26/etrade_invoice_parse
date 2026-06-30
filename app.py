from __future__ import annotations

import hashlib
import logging
import secrets as secrets_lib
import sys
import tempfile
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterator

import streamlit as st


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from invoice_processing.config import AppConfig  # noqa: E402
from invoice_processing.emailer import EmailDeliveryError, SMTPConfig, send_email  # noqa: E402
from invoice_processing.processing import InvoiceBatchProcessor  # noqa: E402


PAGE_TITLE = "Invoice Public Workbook Generator"
LOG_FILE_NAME = "invoice_processing.log"
WORKBOOK_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
AUTH_SESSION_KEY = "invoice_processing_authenticated"
SMTP_SECRET_KEYS = ("HOST", "PORT", "USERNAME", "APP_PASSWORD", "FROM_EMAIL", "ADMIN_EMAIL")
USAGE_NOTICE = """
### Important Usage Notice

This tool is designed specifically for Etrade invoice PDFs that follow the supported invoice layout pattern.

It works best when the uploaded PDF has the same or very similar structure, table format, field placement, and text quality as the sample Etrade invoices used during development. A major change in invoice design, column order, address placement, tax layout, scanned-image quality, or page structure can significantly reduce extraction accuracy.

PDF text and table extraction are not guaranteed to be 100% accurate. Scanned-image PDFs or OCR-dependent documents may fail or be less reliable because this app is built around PDF text/table extraction. The generated Excel file must be verified by a human before it is used for finance, accounting, GST, reconciliation, reporting, payment, audit, or any official business process.

Upload only one supported Etrade invoice PDF at a time.
"""

LOGGER = logging.getLogger("invoice_processing.dashboard")


@contextmanager
def _capture_invoice_logs(log_dir: Path) -> Iterator[Path]:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / LOG_FILE_NAME

    logger = logging.getLogger("invoice_processing")
    previous_level = logger.level
    logger.setLevel(logging.INFO)

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(file_handler)

    try:
        yield log_path
    finally:
        logger.removeHandler(file_handler)
        file_handler.close()
        logger.setLevel(previous_level)


def _read_log_text(log_path: Path) -> str:
    if not log_path.exists():
        return ""
    return log_path.read_text(encoding="utf-8", errors="replace")


def _flush_invoice_logs() -> None:
    for handler in logging.getLogger("invoice_processing").handlers:
        handler.flush()


def _get_streamlit_secret(name: str) -> str | None:
    try:
        value = st.secrets.get(name)
    except Exception:
        return None

    if value is None:
        return None

    text = str(value).strip()
    return text or None


def _read_smtp_config_from_secrets() -> tuple[SMTPConfig | None, str | None, str | None]:
    try:
        smtp_secrets = st.secrets.get("SMTP")
    except Exception:
        return None, None, "SMTP secrets are not configured."

    if not smtp_secrets:
        return None, None, "SMTP secrets are not configured."

    missing_keys = [
        key
        for key in SMTP_SECRET_KEYS
        if not str(smtp_secrets.get(key, "")).strip()
    ]
    if missing_keys:
        return None, None, f"SMTP secrets are missing: {', '.join(missing_keys)}."

    try:
        port = int(smtp_secrets["PORT"])
    except (TypeError, ValueError):
        return None, None, "SMTP.PORT must be an integer."

    smtp_config = SMTPConfig(
        host=str(smtp_secrets["HOST"]).strip(),
        port=port,
        username=str(smtp_secrets["USERNAME"]).strip(),
        app_password=str(smtp_secrets["APP_PASSWORD"]).strip(),
        from_email=str(smtp_secrets["FROM_EMAIL"]).strip(),
    )
    return smtp_config, str(smtp_secrets["ADMIN_EMAIL"]).strip(), None


def _render_password_gate() -> bool:
    app_password = _get_streamlit_secret("APP_PASSWORD")
    if not app_password:
        st.error(
            "Application password is not configured. Ask the app owner to set APP_PASSWORD "
            "in Streamlit secrets."
        )
        return False

    if st.session_state.get(AUTH_SESSION_KEY):
        return True

    with st.form("app_password_gate", clear_on_submit=True):
        entered_password = st.text_input("App password", type="password")
        submitted = st.form_submit_button("Continue")

    if submitted:
        if secrets_lib.compare_digest(entered_password, app_password):
            st.session_state[AUTH_SESSION_KEY] = True
            st.rerun()
        else:
            st.error("Incorrect app password.")

    return False


def _upload_key(file_name: str, file_bytes: bytes) -> str:
    digest = hashlib.sha256(file_bytes).hexdigest()
    return f"{Path(file_name).name}:{len(file_bytes)}:{digest}"


def _validate_pdf_upload(file_name: str, file_bytes: bytes) -> str | None:
    if Path(file_name).suffix.lower() != ".pdf":
        return "Please upload one PDF file with a .pdf extension."
    if not file_bytes:
        return "The uploaded PDF is empty."
    if not file_bytes.startswith(b"%PDF"):
        return "The uploaded file does not look like a valid PDF."
    return None


def _clean_exception_message(exc: Exception) -> str:
    message = str(exc).strip() or "No additional details were provided."
    return f"{exc.__class__.__name__}: {message}"


def _exception_summaries(exceptions: list[dict[str, Any]]) -> list[str]:
    summaries: list[str] = []
    for exception in exceptions:
        source_file = exception.get("source_file") or "Uploaded PDF"
        error_type = exception.get("error_type") or "ProcessingError"
        error = exception.get("error") or "No additional details were provided."
        summaries.append(f"{source_file}: {error_type}: {error}")
    return summaries


def _first_existing_public_workbook(paths: list[Path]) -> Path | None:
    for path in paths:
        if path.exists():
            return path
    return None


def _path_attachment_status(label: str, path: Path | None) -> tuple[Path | None, str]:
    if path is None:
        return None, f"- {label}: unavailable - the processor did not produce this artifact."
    if not path.exists():
        return None, f"- {label}: unavailable - expected file was not found at runtime."
    if not path.is_file():
        return None, f"- {label}: unavailable - expected a file but found a non-file path."
    return path, f"- {label}: attached ({path.name})"


def _collect_admin_attachments(
    *,
    pdf_path: Path,
    public_workbook_path: Path | None,
    internal_workbook_path: Path | None,
    log_path: Path | None,
) -> tuple[list[Path], list[str]]:
    attachments: list[Path] = []
    status_lines: list[str] = []

    for label, path in (
        ("Original uploaded PDF", pdf_path),
        ("Public-facing workbook", public_workbook_path),
        ("Detailed/internal workbook", internal_workbook_path),
        ("Runtime log", log_path),
    ):
        attachment_path, status_line = _path_attachment_status(label, path)
        status_lines.append(status_line)
        if attachment_path is not None:
            attachments.append(attachment_path)

    return attachments, status_lines


def _validation_status_lines(validations: list[dict[str, Any]]) -> list[str]:
    if not validations:
        return ["- Unavailable from processor result."]

    lines: list[str] = []
    for validation in validations:
        source_file = validation.get("source_file") or "Uploaded PDF"
        status = validation.get("status") or "UNKNOWN"
        errors = validation.get("errors") or "No validation errors reported."
        lines.append(f"- {source_file}: {status}; {errors}")
    return lines


def _admin_email_body(
    *,
    source_file_name: str,
    processed_at: datetime,
    public_workbook_path: Path | None,
    internal_workbook_path: Path | None,
    validations: list[dict[str, Any]],
    exceptions: list[str],
    attachment_status_lines: list[str],
) -> str:
    public_file_name = public_workbook_path.name if public_workbook_path else "Unavailable"
    internal_file_name = internal_workbook_path.name if internal_workbook_path else "Unavailable"
    exception_lines = [f"- {summary}" for summary in exceptions] or ["- None captured."]

    return "\n".join(
        [
            "Invoice parser run completed.",
            "",
            f"Source PDF filename: {source_file_name}",
            f"Processing timestamp: {processed_at.strftime('%Y-%m-%d %H:%M:%S %Z')}",
            "Public workbook generation succeeded: yes",
            f"Public workbook filename: {public_file_name}",
            f"Internal workbook filename: {internal_file_name}",
            "",
            "Validation status:",
            *_validation_status_lines(validations),
            "",
            "Exception summary:",
            *exception_lines,
            "",
            "Attachment status:",
            *attachment_status_lines,
            "",
            "Human verification required: these files must be checked by a person before any finance, "
            "accounting, GST, reconciliation, reporting, payment, audit, or official business use.",
        ]
    )


def _send_admin_email_notification(
    *,
    smtp_config: SMTPConfig | None,
    admin_email: str | None,
    config_error: str | None,
    source_file_name: str,
    processed_at: datetime,
    pdf_path: Path,
    public_workbook_path: Path | None,
    internal_workbook_path: Path | None,
    log_path: Path | None,
    validations: list[dict[str, Any]],
    exceptions: list[str],
) -> bool:
    if smtp_config is None or admin_email is None:
        LOGGER.warning("Admin email notification skipped: %s", config_error or "SMTP configuration is incomplete.")
        return False

    _flush_invoice_logs()
    attachments, attachment_status_lines = _collect_admin_attachments(
        pdf_path=pdf_path,
        public_workbook_path=public_workbook_path,
        internal_workbook_path=internal_workbook_path,
        log_path=log_path,
    )
    body = _admin_email_body(
        source_file_name=source_file_name,
        processed_at=processed_at,
        public_workbook_path=public_workbook_path,
        internal_workbook_path=internal_workbook_path,
        validations=validations,
        exceptions=exceptions,
        attachment_status_lines=attachment_status_lines,
    )
    subject = f"Invoice parser run - {source_file_name} - {processed_at.strftime('%Y-%m-%d %H:%M:%S UTC')}"

    try:
        send_email(
            smtp_config,
            admin_email,
            subject,
            body,
            attachment_paths=attachments,
        )
    except EmailDeliveryError as exc:
        LOGGER.warning("Admin email notification failed: %s", exc)
        return False

    return True


def _failure_result(
    message: str,
    *,
    log_text: str = "",
    exceptions: list[str] | None = None,
    no_rows: bool = False,
) -> dict[str, Any]:
    return {
        "ok": False,
        "message": message,
        "exceptions": exceptions or [],
        "log_text": log_text,
        "no_rows": no_rows,
    }


def _process_invoice(
    file_name: str,
    file_bytes: bytes,
    progress: Callable[[str], None],
    *,
    smtp_config: SMTPConfig | None,
    admin_email: str | None,
    smtp_config_error: str | None,
) -> dict[str, Any]:
    safe_file_name = Path(file_name).name or "invoice.pdf"
    processed_at = datetime.now(timezone.utc)

    with tempfile.TemporaryDirectory(prefix="invoice_public_workbook_") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        temp_input = temp_dir / "input"
        temp_output = temp_dir / "output"
        temp_logs = temp_dir / "logs"

        progress("Creating an isolated temporary workspace.")
        temp_input.mkdir(parents=True, exist_ok=True)
        temp_output.mkdir(parents=True, exist_ok=True)
        temp_logs.mkdir(parents=True, exist_ok=True)

        progress("Saving the uploaded PDF.")
        pdf_path = temp_input / safe_file_name
        pdf_path.write_bytes(file_bytes)

        config = AppConfig(
            input_dir=temp_input,
            output_dir=temp_output,
            log_dir=temp_logs,
        ).resolved()

        with _capture_invoice_logs(config.log_dir) as log_path:
            try:
                progress("Parsing the invoice and generating the workbook.")
                result = InvoiceBatchProcessor(config).process()

                progress("Preparing the public workbook for download.")
                log_text = _read_log_text(log_path)
                exceptions = _exception_summaries(result.exceptions)
                public_workbook_path = _first_existing_public_workbook(result.public_output_paths)
                internal_workbook_path = result.output_path if result.output_path and result.output_path.exists() else None

                if public_workbook_path is None:
                    message = (
                        "No workbook could be generated. This may be an unsupported or wrong-layout PDF. "
                        "Please try a known supported Etrade invoice PDF."
                    )
                    return _failure_result(
                        message,
                        log_text=log_text,
                        exceptions=exceptions,
                        no_rows=not result.line_items,
                    )

                progress("Sending the admin email notification.")
                email_sent = _send_admin_email_notification(
                    smtp_config=smtp_config,
                    admin_email=admin_email,
                    config_error=smtp_config_error,
                    source_file_name=safe_file_name,
                    processed_at=processed_at,
                    pdf_path=pdf_path,
                    public_workbook_path=public_workbook_path,
                    internal_workbook_path=internal_workbook_path,
                    log_path=log_path,
                    validations=result.validations,
                    exceptions=exceptions,
                )

                return {
                    "ok": True,
                    "download_name": public_workbook_path.name,
                    "workbook_bytes": public_workbook_path.read_bytes(),
                    "email_failed": not email_sent,
                    "no_rows": not result.line_items,
                }
            except Exception as exc:
                LOGGER.exception("Dashboard processing failed")
                no_rows = "No invoice table rows found" in str(exc)
                return _failure_result(
                    (
                        "Processing failed. This may be an unsupported or wrong-layout PDF. "
                        "Please check the file and try a known supported Etrade invoice PDF."
                    ),
                    log_text=_read_log_text(log_path),
                    exceptions=[_clean_exception_message(exc)],
                    no_rows=no_rows,
                )


def _clear_result() -> None:
    st.session_state.pop("dashboard_result", None)


def _render_result(result: dict[str, Any]) -> None:
    if result["ok"]:
        if result.get("no_rows"):
            st.warning("The invoice was parsed, but no invoice item rows were extracted.")

        st.success("Public workbook generated successfully.")
        st.info(
            "Verify the generated workbook manually before using it for finance, accounting, GST, "
            "payment, audit, or any official process."
        )
        if result.get("email_failed"):
            st.warning("Workbook generated, but admin email notification failed.")
        st.download_button(
            "Download Public Workbook",
            data=result["workbook_bytes"],
            file_name=result["download_name"],
            mime=WORKBOOK_MIME,
        )
    else:
        st.error(result["message"])
        if result.get("no_rows"):
            st.warning("No invoice item rows were extracted from this PDF.")


def main() -> None:
    st.set_page_config(page_title=PAGE_TITLE)
    st.title(PAGE_TITLE)

    if not _render_password_gate():
        return

    st.warning(USAGE_NOTICE)

    uploaded_file = st.file_uploader(
        "Upload invoice PDF",
        type=["pdf"],
        accept_multiple_files=False,
    )

    file_bytes = b""
    validation_error = None

    if uploaded_file is None:
        st.session_state.pop("active_upload_key", None)
        _clear_result()
    else:
        file_bytes = uploaded_file.getvalue()
        current_upload_key = _upload_key(uploaded_file.name, file_bytes)
        if st.session_state.get("active_upload_key") != current_upload_key:
            st.session_state["active_upload_key"] = current_upload_key
            _clear_result()
        validation_error = _validate_pdf_upload(uploaded_file.name, file_bytes)

    if validation_error:
        st.error(validation_error)

    process_clicked = st.button(
        "Process Invoice",
        disabled=uploaded_file is None or validation_error is not None,
    )

    if process_clicked and uploaded_file is not None:
        _clear_result()
        smtp_config, admin_email, smtp_config_error = _read_smtp_config_from_secrets()
        with st.status("Processing invoice...", expanded=True) as status:
            result = _process_invoice(
                uploaded_file.name,
                file_bytes,
                status.write,
                smtp_config=smtp_config,
                admin_email=admin_email,
                smtp_config_error=smtp_config_error,
            )
            st.session_state["dashboard_result"] = result
            if result["ok"]:
                status.update(label="Workbook ready.", state="complete")
            else:
                status.update(label="Processing failed.", state="error")

    result = st.session_state.get("dashboard_result")
    if result:
        _render_result(result)


if __name__ == "__main__":
    main()
