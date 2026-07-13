#!/usr/bin/env python3
"""
Tool tra cứu thông tin từ file Excel
"""

import sys
import os
import glob
import pandas as pd


def find_excel_files(pattern: str, search_dir: str = None):
    """Tìm các file Excel theo pattern"""
    if search_dir is None:
        search_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Tìm trong workspace và các thư mục con
    patterns = [
        os.path.join(search_dir, "**", f"*{pattern}*.xlsx"),
        os.path.join(search_dir, "**", f"*{pattern}*.xls"),
    ]

    files = []
    for p in patterns:
        files.extend(glob.glob(p, recursive=True))

    return sorted(set(files))


def search_in_excel(file_path: str, search_term: str):
    """Tìm kiếm thông tin trong file Excel"""
    results = []

    try:
        xlsx = pd.ExcelFile(file_path)

        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)

            # Tìm các dòng có chứa từ khóa
            for idx, row in df.iterrows():
                row_str = " ".join([str(v) for v in row if pd.notna(v)])
                if search_term.lower() in row_str.lower():
                    # Lấy thông tin dòng
                    row_data = []
                    for col_idx, val in enumerate(row):
                        if pd.notna(val):
                            row_data.append(f"Col {col_idx}: {val}")

                    results.append({"sheet": sheet_name, "row": idx, "data": row_data})
    except Exception as e:
        print(f"Lỗi đọc file {file_path}: {e}")

    return results


def main():
    if len(sys.argv) < 3:
        print("Cách dùng: python tra_cuu.py <tên_file_excel> <thông_tin_tìm_kiếm>")
        print("Ví dụ: python tra_cuu.py qd3176 NGAY_VAO")
        sys.exit(1)

    file_pattern = sys.argv[1]
    search_term = sys.argv[2]

    # Tìm các file Excel
    print(f"=== Tìm file Excel có chứa '{file_pattern}' ===\n")
    files = find_excel_files(file_pattern)

    if not files:
        print(f"Không tìm thấy file nào có chứa '{file_pattern}'")
        sys.exit(1)

    print(f"Tìm thấy {len(files)} file(s):\n")
    for f in files:
        print(f"  - {f}")
    print()

    # Tìm kiếm trong từng file
    for file_path in files:
        print(f"\n{'=' * 80}")
        print(f"File: {os.path.basename(file_path)}")
        print("=" * 80)

        results = search_in_excel(file_path, search_term)

        if not results:
            print(f"Không tìm thấy '{search_term}' trong file này")
            continue

        print(f"\nTìm thấy {len(results)} kết quả:\n")

        for r in results:
            print(f"--- Sheet: {r['sheet']}, Row: {r['row']} ---")
            for line in r["data"]:
                print(f"  {line}")
            print()


if __name__ == "__main__":
    main()
