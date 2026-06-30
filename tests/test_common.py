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


if __name__ == "__main__":
    unittest.main()
