#!/usr/bin/env python3
import os
import sys
import argparse
import csv
import shutil

def list_excel_files(folder):
    """
    Returns a set of Excel filenames (ending with .xlsx) in the given folder.
    """
    return {f for f in os.listdir(folder)
            if f.lower().endswith('.xlsx') and os.path.isfile(os.path.join(folder, f))}

def get_common_files(folder1, folder2):
    """
    Returns a list of filenames that exist in both folders.
    """
    files1 = list_excel_files(folder1)
    files2 = list_excel_files(folder2)
    return list(files1.intersection(files2))

def main():
    parser = argparse.ArgumentParser(
        description="Compare Excel file sizes in two folders and list the top files by absolute and percentage differences. "
                    "Then copy the larger version of each file into separate destination folders."
    )
    parser.add_argument("folder1", help="Path to the first folder")
    parser.add_argument("folder2", help="Path to the second folder")
    parser.add_argument("--abs-count", type=int, default=10,
                        help="Number of files to display for absolute differences (default: 10)")
    parser.add_argument("--rel-count", type=int, default=10,
                        help="Number of files to display for relative differences (default: 10)")
    parser.add_argument("--csv-abs-out", default="absolute_comparison.csv",
                        help="CSV file to save the absolute comparison results (default: absolute_comparison.csv)")
    parser.add_argument("--csv-rel-out", default="relative_comparison.csv",
                        help="CSV file to save the relative comparison results (default: relative_comparison.csv)")
    parser.add_argument("--abs-dest", default="absolute_top",
                        help="Destination folder to copy files with highest absolute differences (default: absolute_top)")
    parser.add_argument("--rel-dest", default="relative_top",
                        help="Destination folder to copy files with highest relative differences (default: relative_top)")
    args = parser.parse_args()

    folder1 = args.folder1
    folder2 = args.folder2

    # Check if both source folders exist.
    for d in [folder1, folder2]:
        if not os.path.isdir(d):
            print(f"Folder not found: {d}")
            sys.exit(1)

    common_files = get_common_files(folder1, folder2)
    if not common_files:
        print("No common Excel files found.")
        sys.exit(0)

    results = []
    for fname in common_files:
        path1 = os.path.join(folder1, fname)
        path2 = os.path.join(folder2, fname)
        size1 = os.path.getsize(path1)
        size2 = os.path.getsize(path2)
        diff = abs(size1 - size2)
        # Calculate percentage difference relative to the larger file.
        if max(size1, size2) > 0:
            percentage_diff = (diff / max(size1, size2)) * 100
        else:
            percentage_diff = 0
        # Determine the "original" file (the one with the larger size)
        if size1 >= size2:
            original_path = path1
        else:
            original_path = path2
        results.append({
            "filename": fname,
            "folder1_size": size1,
            "folder2_size": size2,
            "size_diff": diff,
            "percentage_diff": percentage_diff,
            "original_path": original_path
        })

    # Sort by absolute size difference and percentage difference.
    abs_sorted = sorted(results, key=lambda x: x["size_diff"], reverse=True)
    rel_sorted = sorted(results, key=lambda x: x["percentage_diff"], reverse=True)

    abs_top = abs_sorted[:args.abs_count]
    rel_top = rel_sorted[:args.rel_count]

    # Print results to console.
    print("\nTop files by absolute size difference:")
    print("Filename\tFolder1 Size\tFolder2 Size\tAbsolute Difference")
    for r in abs_top:
        print(f'{r["filename"]}\t{r["folder1_size"]}\t{r["folder2_size"]}\t{r["size_diff"]}')

    print("\nTop files by percentage size difference:")
    print("Filename\tFolder1 Size\tFolder2 Size\t% Difference")
    for r in rel_top:
        print(f'{r["filename"]}\t{r["folder1_size"]}\t{r["folder2_size"]}\t{r["percentage_diff"]:.2f}%')

    # Write the absolute difference results to a CSV file.
    abs_fieldnames = ["filename", "folder1_size", "folder2_size", "size_diff", "percentage_diff"]
    with open(args.csv_abs_out, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=abs_fieldnames)
        writer.writeheader()
        for r in abs_top:
            # Filter out keys not in abs_fieldnames.
            row = {k: r[k] for k in abs_fieldnames}
            writer.writerow(row)
    print(f"\nAbsolute difference results saved to {args.csv_abs_out}.")

    # Write the relative difference results to a CSV file.
    rel_fieldnames = ["filename", "folder1_size", "folder2_size", "percentage_diff", "size_diff"]
    with open(args.csv_rel_out, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=rel_fieldnames)
        writer.writeheader()
        for r in rel_top:
            # Filter out keys not in rel_fieldnames.
            row = {k: r[k] for k in rel_fieldnames}
            writer.writerow(row)
    print(f"Relative difference results saved to {args.csv_rel_out}.")

    # Create destination directories if they do not exist.
    if not os.path.exists(args.abs_dest):
        os.makedirs(args.abs_dest)
    if not os.path.exists(args.rel_dest):
        os.makedirs(args.rel_dest)

    # Copy the original files from the absolute top list.
    for entry in abs_top:
        try:
            shutil.copy2(entry["original_path"], args.abs_dest)
            print(f'Copied {entry["filename"]} to {args.abs_dest}')
        except Exception as e:
            print(f"Error copying {entry['filename']} to {args.abs_dest}: {e}")

    # Copy the original files from the relative top list.
    for entry in rel_top:
        try:
            shutil.copy2(entry["original_path"], args.rel_dest)
            print(f'Copied {entry["filename"]} to {args.rel_dest}')
        except Exception as e:
            print(f"Error copying {entry['filename']} to {args.rel_dest}: {e}")

if __name__ == "__main__":
    main()
