class InvoiceProcessingError(Exception):
    """Base exception for expected invoice processing failures."""


class ConfigurationError(InvoiceProcessingError):
    """Raised when runtime configuration is invalid."""


class PdfReadError(InvoiceProcessingError):
    """Raised when a PDF cannot be read or inspected."""


class InvoiceParseError(InvoiceProcessingError):
    """Raised when an invoice does not match the supported parsing rules."""


class ExportError(InvoiceProcessingError):
    """Raised when output cannot be written."""
