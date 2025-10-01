"""Script to run payment terms comparison between Excel and QuickBooks."""

import sys

from xlsx_reader.excel_processor import process_payment_terms

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python run_comparison.py <path_to_excel_file>")
        print("Example: python run_comparison.py C:\\datasets\\paymentterms.xlsx")
        sys.exit(1)

    excel_file = sys.argv[1]

    print(f"Starting payment terms comparison for: {excel_file}\n")

    try:
        result = process_payment_terms(excel_file)
        print("\n=== Summary ===")
        print("Completed successfully!")
        print(f"- Matching terms: {result.matching_count}")
        print(f"- Terms with same ID but different names: {len(result.same_id_diff_name)}")
        print(f"- Terms only in Excel (added to QB): {len(result.only_in_excel)}")
        print(f"- Terms only in QB: {len(result.only_in_qb)}")

    except Exception as e:
        print(f"\nError: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
