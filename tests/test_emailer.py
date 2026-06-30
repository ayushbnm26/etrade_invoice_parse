from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from invoice_processing.emailer import (  # noqa: E402
    EmailAttachment,
    SMTPConfig,
    build_email_message,
    sanitize_email_exception,
)


class EmailerTests(unittest.TestCase):
    def _smtp_config(self) -> SMTPConfig:
        return SMTPConfig(
            host="smtp.gmail.com",
            port=587,
            username="sender@example.com",
            app_password="abcd efgh ijkl mnop",
            from_email="sender@example.com",
        )

    def test_build_email_message_adds_path_and_byte_attachments(self) -> None:
        smtp_config = self._smtp_config()

        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / "invoice.pdf"
            pdf_path.write_bytes(b"%PDF-test")

            message = build_email_message(
                smtp_config,
                "admin@example.com",
                "Invoice run",
                "Body text",
                attachment_paths=[pdf_path],
                attachments=[
                    EmailAttachment(
                        filename="public.xlsx",
                        content=b"workbook-bytes",
                        mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                ],
            )

        self.assertEqual(message["From"], "sender@example.com")
        self.assertEqual(message["To"], "admin@example.com")
        self.assertEqual(message["Subject"], "Invoice run")
        attachments = list(message.iter_attachments())
        self.assertEqual([part.get_filename() for part in attachments], ["invoice.pdf", "public.xlsx"])
        self.assertEqual(attachments[0].get_content_type(), "application/pdf")
        self.assertEqual(
            attachments[1].get_content_type(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    def test_sanitize_email_exception_redacts_credentials(self) -> None:
        smtp_config = self._smtp_config()
        error = RuntimeError("login failed for sender@example.com using abcd efgh ijkl mnop")

        sanitized = sanitize_email_exception(error, smtp_config)

        self.assertNotIn("sender@example.com", sanitized)
        self.assertNotIn("abcd efgh ijkl mnop", sanitized)
        self.assertIn("[redacted]", sanitized)


if __name__ == "__main__":
    unittest.main()
