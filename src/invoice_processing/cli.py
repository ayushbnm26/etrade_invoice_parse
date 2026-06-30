from __future__ import annotations

import argparse
import logging
from pathlib import Path

from invoice_processing.config import AppConfig, DEFAULT_INPUT_DIR, DEFAULT_LOG_DIR, DEFAULT_OUTPUT_DIR
from invoice_processing.exceptions import InvoiceProcessingError
from invoice_processing.logging_config import setup_logging
from invoice_processing.processing import InvoiceBatchProcessor


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Parse invoice PDFs and export structured Excel reports.")
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT_DIR, help="Folder containing invoice PDFs.")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help="Folder for Excel output.")
    parser.add_argument("--output-file", type=Path, default=None, help="Optional output workbook name or path.")
    parser.add_argument("--log-dir", type=Path, default=DEFAULT_LOG_DIR, help="Folder for runtime logs.")
    parser.add_argument("--log-level", default="INFO", help="Python log level, for example INFO or DEBUG.")
    parser.add_argument("--fail-fast", action="store_true", help="Stop after the first PDF-level parsing failure.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = AppConfig(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        log_dir=args.log_dir,
        output_file=args.output_file,
        fail_fast=args.fail_fast,
        log_level=args.log_level,
    ).resolved()

    setup_logging(config.log_dir, config.log_level)

    try:
        result = InvoiceBatchProcessor(config).process()
    except InvoiceProcessingError as exc:
        logging.getLogger(__name__).error("%s", exc)
        return 2
    except Exception as exc:
        logging.getLogger(__name__).exception("Unexpected failure: %s", exc)
        return 99

    print(f"Excel created: {result.output_path}")
    print(f"Public workbooks created: {len(result.public_output_paths)}")
    print(f"Invoices parsed: {len(result.headers)}")
    print(f"Line items extracted: {len(result.line_items)}")
    print(f"PDF exceptions captured: {len(result.exceptions)}")
    return 0 if result.headers else 1


if __name__ == "__main__":
    raise SystemExit(main())
