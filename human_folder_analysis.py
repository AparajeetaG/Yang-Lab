#!/usr/bin/env python
"""
Human folder detailed analysis

What it does (same as your original, plus a Subject_Status column and two small fixes):
- Scans the entire folder tree under a given "Human" root.
- Counts folders/files, DICOM (.dcm/.ima), DAT (.dat), and NIfTI (.nii/.nii.gz).
- Tracks file-type frequency globally and per main subfolder.
- Writes an Excel workbook with:
  * Overview (global + per-main-subfolder summary)
  * File_Types (top 20 extensions globally)
  * One sheet per main subfolder with one row per directory and hierarchy columns
  * All_Folders (entire directory listing)
- Adds Subject_Status to each subfolder sheet:
  Processed / Not Processed / Unknown (see rules below).

Run:
    python human_folder_analysis.py --root "C:\\Users\\<you>\\Box\\Data\\Human" --out "C:\\Users\\<you>\\Documents"

Requirements:
    pip install pandas openpyxl
"""

import os
import re
import argparse
from datetime import datetime
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

# ----------------------------
# Defaults (can be overridden via CLI)
# ----------------------------
BOX_HUMAN_ROOT = r"C:\Users\GuhaA2\Box\Data\Human"
OUTPUT_FOLDER = r"C:\Users\guhaa2\Documents"


def _now_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def is_nifti_filename(fname: str) -> bool:
    """
    True only for .nii or .nii.gz (case-insensitive).
    """
    s = ''.join(Path(fname).suffixes).lower()
    return s.endswith('.nii') or s.endswith('.nii.gz')


def scan_human_folder_complete(human_root: str):
    """
    Comprehensive scan of Human folder.
    Returns a nested dictionary with all stats.
    """

    print("\nPhase 1: Scanning Human folder structure...")
    print("-" * 50)

    # Container for all results
    human_data = {
        'overview': {},
        'subfolders': {},
        'all_folders': [],
        'all_files_by_type': Counter(),
        'folder_tree': {}  # reserved (not used)
    }

    # Discover immediate subfolders of Human
    try:
        human_subfolders = [
            f for f in os.listdir(human_root)
            if os.path.isdir(os.path.join(human_root, f))
        ]
        print(f"Found {len(human_subfolders)} main subfolders in Human")
        for sf in human_subfolders[:10]:
            print(f"  - {sf}")
        if len(human_subfolders) > 10:
            print(f"  ... and {len(human_subfolders)-10} more")
    except Exception as e:
        print(f"Error reading Human folder: {e}")
        return None

    # Initialize per-subfolder data structure
    for subfolder in human_subfolders:
        human_data['subfolders'][subfolder] = {
            'path': os.path.join(human_root, subfolder),
            'all_folders': [],
            'file_counts': Counter(),
            'depth_counts': Counter(),
            'total_files': 0,
            'total_folders': 0,
            'dicom_files': 0,
            'dat_files': 0,
            'folder_details': []
        }

    # Walk the entire tree
    total_folders = 0
    total_files = 0

    for root, dirs, files in os.walk(human_root):
        total_folders += 1
        total_files += len(files)

        # Relative path from Human root
        rel_path = os.path.relpath(root, human_root)

        # Determine main subfolder and depth
        if rel_path == '.':
            current_subfolder = None
            depth = 0
        else:
            path_parts = rel_path.split(os.sep)
            current_subfolder = path_parts[0]
            depth = len(path_parts)

        # Count file types at this folder
        file_types = Counter()
        dicom_count = 0
        dat_count = 0
        nifti_count = 0

        for file in files:
            ext = os.path.splitext(file)[1].lower()
            file_types[ext] += 1
            human_data['all_files_by_type'][ext] += 1

            if ext in ('.dcm', '.ima'):
                dicom_count += 1
            if ext == '.dat':
                dat_count += 1
            if is_nifti_filename(file):
                nifti_count += 1

        folder_info = {
            'path': rel_path,
            'full_path': root,
            'depth': depth,
            'num_subfolders': len(dirs),
            'num_files': len(files),
            'file_types': dict(file_types),
            'dicom_files': dicom_count,
            'dat_files': dat_count,
            'nifti_files': nifti_count,
            'subfolder_names': dirs[:10] if dirs else []
        }

        human_data['all_folders'].append(folder_info)

        # Accumulate into that main subfolder (if applicable)
        if current_subfolder and current_subfolder in human_data['subfolders']:
            sf_data = human_data['subfolders'][current_subfolder]
            sf_data['folder_details'].append(folder_info)
            sf_data['total_folders'] += 1
            sf_data['total_files'] += len(files)
            sf_data['dicom_files'] += dicom_count
            sf_data['dat_files'] += dat_count
            sf_data['depth_counts'][depth] += 1
            for ext, count in file_types.items():
                sf_data['file_counts'][ext] += count

        # Progress every 100 folders
        if total_folders % 100 == 0:
            print(f"  Processed {total_folders} folders...", end='\r')

    # Overview stats
    human_data['overview'] = {
        'total_folders': total_folders,
        'total_files': total_files,
        'main_subfolders': len(human_subfolders),
        'total_dicom': sum(f['dicom_files'] for f in human_data['all_folders']),
        'total_dat': sum(f['dat_files'] for f in human_data['all_folders']),
        'total_nifti': sum(f['nifti_files'] for f in human_data['all_folders']),
        'max_depth': max([f['depth'] for f in human_data['all_folders']]) if human_data['all_folders'] else 0
    }

    print(f"\n  Total folders scanned: {total_folders}")
    print(f"  Total files found: {total_files}")

    return human_data


def create_overview_sheet(human_data):
    """
    Build the overview DataFrame and the (global) top-20 file-type summary.
    """
    print("\nCreating overview sheet...")

    overview_data = []

    # Global summary row
    overview_data.append({
        'Category': 'HUMAN FOLDER TOTAL',
        'Total_Folders': human_data['overview']['total_folders'],
        'Total_Files': human_data['overview']['total_files'],
        'DICOM_Files': human_data['overview']['total_dicom'],
        'DAT_Files': human_data['overview']['total_dat'],
        'NIFTI_Files': human_data['overview']['total_nifti'],
        'Main_Subfolders': human_data['overview']['main_subfolders'],
        'Max_Depth': human_data['overview']['max_depth']
    })

    # One row per main subfolder
    for subfolder_name, sf_data in human_data['subfolders'].items():
        overview_data.append({
            'Category': subfolder_name,
            'Total_Folders': sf_data['total_folders'],
            'Total_Files': sf_data['total_files'],
            'DICOM_Files': sf_data['dicom_files'],
            'DAT_Files': sf_data['dat_files'],
            'NIFTI_Files': sum(f['nifti_files'] for f in sf_data['folder_details']),
            # Keep same logic as original to preserve output:
            'Main_Subfolders': sf_data['folder_details'][0]['num_subfolders'] if sf_data['folder_details'] else 0,
            'Max_Depth': max(sf_data['depth_counts'].keys()) if sf_data['depth_counts'] else 0
        })

    df_overview = pd.DataFrame(overview_data)

    # (Global) top 20 file types
    file_type_summary = pd.DataFrame([
        {'File_Extension': ext, 'Count': count}
        for ext, count in human_data['all_files_by_type'].most_common(20)
    ])

    return df_overview, file_type_summary


def create_subfolder_sheet(subfolder_name, sf_data):
    """
    Build detailed DataFrame for a specific main subfolder.
    Adds a Subject_Status column:
      - "Processed" if any .mat exists under that subject
      - "Not Processed" if (.dat or .ima) exist but no .mat
      - "Unknown" otherwise
    """
    if not sf_data['folder_details']:
        return pd.DataFrame()

    # --- Aggregate per-subject extension counts across the entire sub-tree ---
    subject_ext_counts = defaultdict(Counter)
    for folder in sf_data['folder_details']:
        # Parse path parts
        if folder['path'] == subfolder_name:
            parts = [subfolder_name]
        else:
            parts = folder['path'].split(os.sep)

        # Subject is Level_2 = parts[1] if present
        subject = parts[1] if len(parts) > 1 else None
        if subject:
            for ext, c in folder['file_types'].items():
                subject_ext_counts[subject][ext] += c

    # Decide status for each subject
    subject_status = {}
    for subj, cnts in subject_ext_counts.items():
        mat = cnts.get('.mat', 0)
        dat = cnts.get('.dat', 0)
        ima = cnts.get('.ima', 0)
        if mat > 0:
            status = 'Processed'
        elif (dat + ima) > 0:
            status = 'Not Processed'
        else:
            status = 'Unknown'
        subject_status[subj] = status

    # --- Build the row-wise table for this subfolder ---
    rows = []
    for folder in sf_data['folder_details']:
        # Parse path to get hierarchy columns
        if folder['path'] == subfolder_name:
            path_parts = [subfolder_name]
        else:
            path_parts = folder['path'].split(os.sep)

        subj = path_parts[1] if len(path_parts) > 1 else ''
        status = subject_status.get(subj, '') if subj else ''

        rows.append({
            'Relative_Path': folder['path'],
            'Level_1': path_parts[0] if len(path_parts) > 0 else '',
            'Level_2': path_parts[1] if len(path_parts) > 1 else '',
            'Level_3': path_parts[2] if len(path_parts) > 2 else '',
            'Level_4': path_parts[3] if len(path_parts) > 3 else '',
            'Deeper': os.sep.join(path_parts[4:]) if len(path_parts) > 4 else '',
            'Depth': folder['depth'],
            'Direct_Subfolders': folder['num_subfolders'],
            'Total_Files': folder['num_files'],
            'DICOM': folder['dicom_files'],
            'DAT': folder['dat_files'],
            'NIFTI': folder['nifti_files'],
            'File_Types': ', '.join([
                f"{ext}({count})" for ext, count in
                sorted(folder['file_types'].items(),
                       key=lambda x: x[1], reverse=True)[:5]
            ]),
            'Subfolder_Names': ', '.join(folder['subfolder_names'][:5]),
            'Subject_Status': status  # <-- new column
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(['Depth', 'Relative_Path'])
    return df


def save_human_analysis_excel(human_data, output_folder: str):
    """
    Save the Excel with:
      - Overview
      - File_Types
      - One sheet per main subfolder
      - All_Folders
    """
    timestamp = _timestamp()
    output_file = os.path.join(output_folder, f"Human_Detailed_Analysis_{timestamp}.xlsx")

    # We know we will add All_Folders; count it in the announced number
    expected_sheets = 2 + len(human_data['subfolders']) + 1
    print(f"\nSaving Excel with {expected_sheets} sheets...")
    print("-" * 50)

    sheet_count = 0
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet 1: Overview
        df_overview, df_file_types = create_overview_sheet(human_data)
        df_overview.to_excel(writer, sheet_name='Overview', index=False)
        sheet_count += 1

        # Sheet 2: File Type Summary
        df_file_types.to_excel(writer, sheet_name='File_Types', index=False)
        sheet_count += 1

        # Per-main-subfolder sheets
        for subfolder_name, sf_data in human_data['subfolders'].items():
            sheet_name = re.sub(r'[\\/*?:\[\]]', '_', subfolder_name)[:31]
            df_subfolder = create_subfolder_sheet(subfolder_name, sf_data)
            if not df_subfolder.empty:
                df_subfolder.to_excel(writer, sheet_name=sheet_name, index=False)
                sheet_count += 1
                print(f"  Created sheet: {sheet_name} ({len(df_subfolder)} folders)")
            else:
                print(f"  Skipped empty: {sheet_name}")

        # All_Folders sheet
        all_folders_df = pd.DataFrame(human_data['all_folders'])
        if not all_folders_df.empty:
            cols_to_keep = [
                'path', 'depth', 'num_subfolders', 'num_files',
                'dicom_files', 'dat_files', 'nifti_files'
            ]
            all_folders_df = all_folders_df[cols_to_keep].sort_values(['depth', 'path'])
            all_folders_df.to_excel(writer, sheet_name='All_Folders', index=False)
            sheet_count += 1

    print(f"\nExcel saved: {output_file}")
    print(f"Total sheets created: {sheet_count}")
    return output_file


def main(human_root: str, output_folder: str):
    """
    Orchestrates the run: prints headers, scans, summarizes, writes Excel.
    """
    print("=" * 80)
    print("HUMAN FOLDER DETAILED ANALYSIS")
    print("=" * 80)
    print(f"Start time: {_now_str()}")
    print(f"Analyzing: {human_root}")
    print("=" * 80)

    # Check root exists
    if not os.path.exists(human_root):
        print(f"ERROR: Human folder not found at: {human_root}")
        return None

    # Scan
    human_data = scan_human_folder_complete(human_root)
    if not human_data:
        print("ERROR: Failed to scan Human folder")
        return None

    # Summary block (same fields/wording as before)
    print("\n" + "=" * 80)
    print("HUMAN FOLDER SUMMARY")
    print("=" * 80)
    print(f"Total folders: {human_data['overview']['total_folders']}")
    print(f"Total files: {human_data['overview']['total_files']}")
    print(f"Main subfolders: {human_data['overview']['main_subfolders']}")
    print(f"DICOM files: {human_data['overview']['total_dicom']}")
    print(f"DAT files: {human_data['overview']['total_dat']}")
    print(f"NIFTI files: {human_data['overview']['total_nifti']}")
    print(f"Maximum depth: {human_data['overview']['max_depth']}")

    # Show the first 10 main subfolders with quick stats
    print("\nMain subfolders in Human:")
    for sf_name, sf_data in list(human_data['subfolders'].items())[:10]:
        print(f"  {sf_name}: {sf_data['total_folders']} folders, {sf_data['total_files']} files")

    # Save Excel
    excel_file = save_human_analysis_excel(human_data, output_folder)

    print("\n" + "=" * 80)
    print("ANALYSIS COMPLETE")
    print("=" * 80)
    print(f"Excel file: {excel_file}")
    print(f"End time: {_now_str()}")
    return excel_file


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Human folder detailed analysis")
    parser.add_argument("--root", type=str, default=BOX_HUMAN_ROOT,
                        help="Path to the Human root folder (default is the value in the script).")
    parser.add_argument("--out", type=str, default=OUTPUT_FOLDER,
                        help="Folder where the Excel file will be written (default is the value in the script).")
    args = parser.parse_args()
    main(args.root, args.out)
