"""XLSX Reader Package.

A simple Excel file reader for educational purposes.
"""

from .excel_processor import (
    PaymentTerm,
    TermComparison,
    compare_payment_terms,
    get_qb_payment_terms,
    process_payment_terms,
    read_payment_terms,
)

__version__ = "0.1.0"
__all__ = [
    "PaymentTerm",
    "TermComparison",
    "read_payment_terms",
    "get_qb_payment_terms",
    "compare_payment_terms",
    "process_payment_terms",
]
