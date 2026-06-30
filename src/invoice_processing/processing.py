from __future__ import annotations

from datetime import datetime
import logging
import traceback
from pathlib import Path

from invoice_processing.config import AppConfig
from invoice_processing.exceptions import ConfigurationError, ExportError
from invoice_processing.exporters.excel import ExcelExporter
from invoice_processing.models import BatchResult, ParsedInvoice
from invoice_processing.parsers.invoice_parser import InvoiceParser


LOGGER = logging.getLogger(__name__)


class InvoiceBatchProcessor:
    def __init__(
        self,
        config: AppConfig,
        parser: InvoiceParser | None = None,
        exporter: ExcelExporter | None = None,
    ) -> None:
        self.config = config
        self.parser = parser or InvoiceParser()
        self.exporter = exporter or ExcelExporter()

    def process(self) -> BatchResult:
        self._validate_input_dir()
        pdf_files = sorted(self.config.input_dir.glob("*.pdf"))
        if not pdf_files:
            raise ConfigurationError(f"No PDF files found in: {self.config.input_dir}")

        result = BatchResult()
        LOGGER.info("Found %s PDF file(s) in %s", len(pdf_files), self.config.input_dir)

        for pdf_path in pdf_files:
            try:
                LOGGER.info("Parsing %s", pdf_path.name)
                parsed_invoice = self.parser.parse(pdf_path)
            except Exception as exc:
                LOGGER.exception("Failed to parse %s", pdf_path.name)
                result.exceptions.append(self._exception_record(pdf_path, exc))
                if self.config.fail_fast:
                    break
                continue

            result.headers.append(parsed_invoice.header)
            result.summaries.append(parsed_invoice.summary)
            result.line_items.extend(parsed_invoice.line_items)
            result.validations.append(parsed_invoice.validation)

            try:
                self._export_public_workbook(pdf_path, parsed_invoice, result)
            except Exception as exc:
                LOGGER.exception("Failed to create public workbook for %s", pdf_path.name)
                result.exceptions.append(self._exception_record(pdf_path, exc))
                if self.config.fail_fast:
                    break

        output_path = self.config.output_file
        if output_path is None:
            raise ConfigurationError("Output file path was not configured")

        result.output_path = self.exporter.export(result, output_path)
        LOGGER.info("Excel output created at %s", result.output_path)
        return result

    def _validate_input_dir(self) -> None:
        if not self.config.input_dir.exists():
            raise ConfigurationError(f"Input folder does not exist: {self.config.input_dir}")
        if not self.config.input_dir.is_dir():
            raise ConfigurationError(f"Input path is not a folder: {self.config.input_dir}")

    def _export_public_workbook(
        self,
        pdf_path: Path,
        parsed_invoice: ParsedInvoice,
        result: BatchResult,
    ) -> None:
        public_output_path = self.config.output_dir / "public_workbooks" / f"{pdf_path.stem}_public.xlsx"
        try:
            output_path = self.exporter.export_public_invoice(parsed_invoice, public_output_path)
        except ExportError as exc:
            if "Permission denied" not in str(exc):
                raise

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback_output_path = (
                self.config.output_dir / "public_workbooks" / f"{pdf_path.stem}_public_{timestamp}.xlsx"
            )
            LOGGER.warning(
                "Public workbook %s is locked; writing fallback file %s",
                public_output_path,
                fallback_output_path,
            )
            output_path = self.exporter.export_public_invoice(parsed_invoice, fallback_output_path)

        result.public_output_paths.append(output_path)
        LOGGER.info("Public Excel output created at %s", output_path)

    def _exception_record(self, pdf_path: Path, exc: Exception) -> dict[str, str]:
        return {
            "source_file": pdf_path.name,
            "error_type": exc.__class__.__name__,
            "error": str(exc),
            "traceback": traceback.format_exc(),
        }
