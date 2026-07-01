from __future__ import annotations

import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from invoice_processing.parsers.common import (  # noqa: E402
    compact_alnum,
    compact_code,
    dedupe_repeated_phrase,
    parse_amount,
    parse_int,
)
from invoice_processing.models import PdfDocumentData, PdfPageData  # noqa: E402
from invoice_processing.parsers.invoice_parser import InvoiceParser  # noqa: E402


class CommonParserTests(unittest.TestCase):
    def test_parse_amount_repairs_broken_decimal(self) -> None:
        self.assertEqual(parse_amount("1,007.2\n8"), 1007.28)
        self.assertEqual(parse_amount("-1,048.96"), 1048.96)

    def test_parse_int_returns_positive_integer(self) -> None:
        self.assertEqual(parse_int("-8"), 8)

    def test_compact_codes(self) -> None:
        self.assertEqual(compact_alnum("548\n016\n077"), "548016077")
        self.assertEqual(compact_code("PB/26\n-\n27/02\n8"), "PB/26-27/028")

    def test_dedupe_repeated_phrase(self) -> None:
        self.assertEqual(
            dedupe_repeated_phrase("ETRADE MARKETING PRIVATE LIMITED ETRADE MARKETING PRIVATE LIMITED"),
            "ETRADE MARKETING PRIVATE LIMITED",
        )


class HeaderParserTests(unittest.TestCase):
    def test_system_ref_no_stops_at_next_header_label(self) -> None:
        cases = [
            ("System Ref No : 42000104035 Credit Note Date : 02-Jun-2026", "42000104035"),
            ("System Ref No : 42000103534 Credit Note Date : 02-Jun-2026", "42000103534"),
            ("System Ref No : 30001158056 Invoice Date : 21-May-2026", "30001158056"),
        ]

        parser = InvoiceParser()
        for text, expected in cases:
            with self.subTest(text=text):
                document = PdfDocumentData(
                    path=Path("test.pdf"),
                    page_count=1,
                    pages=[PdfPageData(page_number=1, text=text, words=[], tables=[])],
                )

                self.assertEqual(parser._parse_header(document)["system_ref_no"], expected)


if __name__ == "__main__":
    unittest.main()
