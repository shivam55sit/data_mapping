import pandas as pd
import os
import glob

# ── CONFIGURATION ───────────────────────────────────────────────────────────
# Reference file (tablet-recorded) — defines the standard 220 columns
REFERENCE_FILE = "Demonstration Multipurpose School.xls"

# Manual input files to standardise (add more files here as you get them)
MANUAL_FILES = [
    "Vishwa Bharathi, Excel1.xlsx",
    # "Another School.xlsx",
]

# Output folder for standardised files
OUTPUT_DIR = "output"

# ── VCH MAP (Sheet 0 — Vision / Clinical / History) ────────────────────────
# Manual column name → Standard column name
VCH_MAP = {
    # Identity & Demographics
    "ID Number": "Child Unique Code",
    "Name of the child": "Student Name",
    "Name of the Father/Guardian": "Parent Or Guardians Name",
    "Gender  Male:0, Female:1": "Gender",
    "Age": "Age",
    "Date": "Date Of Examination",
    "Name of the VT/Optom": "Examiner Name",
    "Additional information": "Remarks",

    # Vision Acuity
    "OD Unaided Distance": "Va OD Unaided",
    "OD Pinhole Distance": "Va OD Pin Hole",
    "OS Unaided Distance": "Va OS Unaided",
    "OS Pinhole Distance": "Va OS Pin Hole",
    "OD Aided Distance": "Va OD Aided Glasses",
    "OS Aided Distance": "Va OS Aided Glasses",
    "Present glasses (0=No, 1=Yes)": "Using Glasses",

    # Objective Refraction
    "OD Spherical  Objective refraction": "Sph Or RE",
    "OD Cylindrical Objective refraction": "Cyl Or RE",
    "OD Axis Objective refraction": "Axis Or RE",
    "OS Spherical Objective refraction": "Sph Or LE",
    "OS Cylindrical Objective refraction": "Cyl Or LE",
    "OS Axis Objective refraction": "Axis Or LE",

    # Acceptance
    "OD Spherical Acceptance": "Sph Acceptance RE",
    "OD Cylindircal Acceptance": "Cyl Acceptance RE",
    "OD Axis Acceptance": "Axis Acceptance RE",
    "OD VA Acceptance": "C2 Va Accptance RE",
    "OS Spherical Acceptance": "Sph Acceptance LE",
    "OS Cylindircal  Acceptance": "Cyl Acceptance LE",
    "OS Axis  Acceptance": "Axis Acceptance LE",
    "OS VA  Acceptance": "C2 Va Accptance LE",

    # Lensometry
    "OD Spherical Lensometry": "Sph Lensometry RE",
    "OD Cylindircal Lensometry": "Cyl Lensometry RE",
    "OD Axis Lensometry": "Axis Lensometry RE",
    "OS Spherical Lensometry": "Sph Lensometry LE",
    "OS Cylindircal Lensometry": "Cyl Lensometry LE",
    "OS Axis Lensometry": "Axis Lensometry LE",

    # Slit Lamp — Lids
    "Lids OD": "Slit Lids OD",
    "Specify.2": "C6 Lids Others RE",
    "Lids OS": "Slit Lids OS",
    "Specify.3": "C6 Lids Others LE",

    # Slit Lamp — Conjunctiva
    "Conjunctiva OD": "Slit Conjunctiva OD C6",
    "Specify.4": "C6 Conjunctiva Others RE",
    "Conjunctiva OS": "Slit Conjunctiva OS C6",
    "Specify.5": "C6 Conjunctiva Others LE",

    # Slit Lamp — Cornea
    "Cornea OD": "Slit Cornea OD C6",
    "Specify.6": "C6 Cornea Others RE",
    "Cornea OS": "Slit Cornea OS C6",
    "Specify.7": "C6 Cornea Others LE",

    # Slit Lamp — Anterior Chamber
    "Anterior Chamber OD": "Slit AC Status OD C6",
    "Specify.8": "Slit AC Status OD C6 Other",
    "Anterior Chamber OS": "Slit AC Status OS C6",
    "Specify.9": "Slit AC Status OS C6 Other",

    # Slit Lamp — Iris/Pupil
    "Iris/Pupil OD": "Slit Pupil Status OD C6",
    "Specify.10": "Slit Pupil Status OD C6 Other",
    "Iris/Pupil OS": "Slit Pupil Status OS C6",
    "Specify.11": "Slit Pupil Status OS C6 Other",

    # Slit Lamp — Lens
    "lens OD": "Slit Lens Status OD C6",
    "Specify.12": "C6 Lens RE Others",
    "lens OS": "Slit Lens Status OS C6",
    "Specify.13": "C6 Lens LE Others",

    # Red Reflex
    "Redd reflex OD": "Red Reflex OD C7",
    "Specify.14": "C7 Red Reflex Others RE",
    "Redd reflex OS": "Red Reflex OS C7",
    "Specify.15": "C7 Red Reflex Others LE",

    # Diagnosis / Impression
    "Diagnosis OD": "Right Impression",
    "Diagnosis OS": "Left Impression",

    # Cause of Visual Impairment
    "Cause of Visual impairement OD": "Right Refractive Error",
    "Cause of Visual impairement OS": "Left Refractive Error",

    # Cover Test / Squint / Nystagmus
    "Cover Test": "SS Squint",
    "Nystagmus": "Shaking Eye Ball",

    # Fundus
    "Fundus OD": "Right Retinal Evaluation",
    "Fundus OS": "Left Retinal Evaluation",

    # Topography
    "Topography OD": "Right Asymmetrical Topography",
    "Topography OS": "Left Asymmetrical Topography",

    # Action / Referral
    "Action": "Advise For Next Module",
    "Name of the TC/ City centre/VC": "C11 Vision Center Name",

    # Date of last eye check (also available from GHS)
    "Date of last eye Check up": "Last Eye Check B1",
}

# ── GHS MAP (Sheet 1 — General Health Screening questionnaire) ──────────────
# Keys are PREFIX strings — the actual Excel headers may be longer.
# The rename_by_prefix() helper matches any column that starts with a key.
GHS_MAP = {
    "Child Id": "Child Unique Code",
    "Name of the child": "Student Name",
    "Name of the Father/Guardian": "Parent Or Guardians Name",
    "Date of examination": "Clinical Examination Date",
    "1.When was the last time you got your eyes checked?": "Last Eye Check B1",
    "2. Do you have the habit of rubbbing eyes?": "Rub Eyes B2",
    "3.How frequency do you rub your eyes?": "Rub Freq B3",
    "4. Why do you rub your eyes?": "Rub Reason B4",
    "5.when you rub your eyes, which of the following do you usually use?": "Rub Method B5",
    "6. Do you think eye rubbing can damage the eyes?": "Think Damage Rub B6",
    "8. Do you frequently encounter redness in your eye?": "Eye Redness B8",
    "9. If Yes, how frequently?": "Redness Frequency B9",
    "10.Do you feel any itching/scratching sensation in your eyes?": "Eye Itching B10",
    "11.If Yes, How frequently?": "Itching Frequency B11",
    "12. Do you ever have difficulty completing school/work": "Difficulty School Work B12",
    "13. Do you have trouble paying attention": "Attention School B13",
    "14.Do you miss school": "School Absent B14",
    "15.Do you have the habit of sleeping with your face down?": "Sleep Face Down B15",
    "16. Do you have episodes of sneezing": "Sneezing Breath B16",
    "17.How frequently?": "Sneeze Freq B17",
    "18. Do you have any history of recurrent itching": "Skin Itch Rash B18",
    "19.Are you on medications": "Medication Allergy B19",
    "20. If yes, do you recall what medications you are taking": "Medication Name B20",
    "21. Are you left/right handed?": "Handedness B21",
    "22.Haveyou undergone any eye surgeries before?": "Eye Surgery B22",
}


def get_standard_columns(reference_file):
    """Read the standard column order from the tablet-recorded reference file."""
    ref_df = pd.read_excel(reference_file, nrows=0)  # just headers
    return list(ref_df.columns)


def clean_columns(df):
    """Strip whitespace from column names."""
    df.columns = df.columns.str.strip()
    return df


def normalize_columns(df):
    """
    Multi-line column headers — extract only the first line.
    """
    df.columns = [col.split("\n")[0].strip() for col in df.columns]
    return df


def rename_by_prefix(df, col_map):
    """
    Rename columns using prefix matching.
    For each map key, if a column starts with that key text, rename it.
    Skips pandas dedup suffixes like '.1', '.2' to avoid false matches.
    """
    rename_dict = {}
    # Sort keys longest-first so more specific keys match before shorter ones
    sorted_keys = sorted(col_map.keys(), key=len, reverse=True)
    for actual_col in df.columns:
        for map_key in sorted_keys:
            if actual_col == map_key:
                rename_dict[actual_col] = col_map[map_key]
                break
            remaining = actual_col[len(map_key):]
            if actual_col.startswith(map_key) and remaining and not remaining[0].isalnum() and not remaining.startswith('.'):
                rename_dict[actual_col] = col_map[map_key]
                break
    df = df.rename(columns=rename_dict)
    return df


def process_file(input_file, standard_columns):
    """Process a single manual Excel file and return a standardised DataFrame."""
    print(f"\n{'='*60}")
    print(f"[→] Processing: {input_file}")
    print(f"{'='*60}")

    # Read sheets (header is in row index 1 for manual files)
    vch = pd.read_excel(input_file, sheet_name=0, header=1)
    ghs = pd.read_excel(input_file, sheet_name=1, header=1)

    # Clean and normalise column names
    vch = clean_columns(vch)
    vch = normalize_columns(vch)
    ghs = clean_columns(ghs)
    ghs = normalize_columns(ghs)

    print(f"  VCH: {vch.shape[0]} rows, {vch.shape[1]} columns")
    print(f"  GHS: {ghs.shape[0]} rows, {ghs.shape[1]} columns")

    # Rename: exact match for VCH, prefix match for GHS (GHS has longer headers)
    vch = vch.rename(columns=VCH_MAP)
    ghs = rename_by_prefix(ghs, GHS_MAP)

    # Log which columns were successfully mapped
    vch_mapped = [v for v in set(VCH_MAP.values()) if v in vch.columns]
    ghs_mapped = [v for v in set(GHS_MAP.values()) if v in ghs.columns]
    print(f"  VCH mapped columns: {len(vch_mapped)}")
    print(f"  GHS mapped columns: {len(ghs_mapped)}")

    # ── Merge VCH + GHS on student ID ────────────────────────────────────
    # Convert Child Unique Code to string in both frames to avoid type mismatch
    if "Child Unique Code" in vch.columns:
        vch["Child Unique Code"] = vch["Child Unique Code"].apply(
            lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if pd.notna(x) else ""
        )
    if "Child Unique Code" in ghs.columns:
        ghs["Child Unique Code"] = ghs["Child Unique Code"].apply(
            lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if pd.notna(x) else ""
        )

    if "Child Unique Code" in vch.columns and "Child Unique Code" in ghs.columns:
        # Drop overlapping non-key columns from GHS to avoid _x/_y suffixes
        ghs_only_cols = [c for c in ghs.columns if c not in vch.columns or c == "Child Unique Code"]
        ghs_for_merge = ghs[ghs_only_cols]

        merged = pd.merge(vch, ghs_for_merge, on="Child Unique Code", how="outer")
        print(f"  Merged: {merged.shape[0]} rows (outer join on Child Unique Code)")
    else:
        # Fallback: if no ID column, just use VCH data
        print("  [WARN] Cannot merge — 'Child Unique Code' missing. Using VCH data only.")
        merged = vch

    # ── Drop duplicate columns from merge (keep first) ─────────────────
    merged = merged.loc[:, ~merged.columns.duplicated(keep='first')]

    # ── Build final output with all standard columns ─────────────────────
    final_df = pd.DataFrame()
    mapped_count = 0
    blank_count = 0

    for col in standard_columns:
        if col in merged.columns:
            final_df[col] = merged[col].values
            mapped_count += 1
        else:
            final_df[col] = ""
            blank_count += 1

    # Auto-generate Reg Id
    final_df["Reg Id"] = range(1, len(final_df) + 1)

    print(f"  Output: {len(final_df)} rows, {len(standard_columns)} columns")
    print(f"  Columns with data: {mapped_count} | Blank columns: {blank_count}")

    return final_df


def main():
    try:
        # ── Get standard column order from reference file ────────────────
        if not os.path.exists(REFERENCE_FILE):
            print(f"[ERROR] Reference file not found: {REFERENCE_FILE}")
            return

        standard_columns = get_standard_columns(REFERENCE_FILE)
        print(f"[✓] Standard format: {len(standard_columns)} columns from '{REFERENCE_FILE}'")

        # ── Create output directory ──────────────────────────────────────
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # ── Process each manual file ─────────────────────────────────────
        for input_file in MANUAL_FILES:
            if not os.path.exists(input_file):
                print(f"\n[SKIP] File not found: {input_file}")
                continue

            final_df = process_file(input_file, standard_columns)

            # Save output
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            output_file = os.path.join(OUTPUT_DIR, f"{base_name}_standardised.xlsx")
            final_df.to_excel(output_file, index=False)
            print(f"  [✓] Saved: {output_file}")

        print(f"\n{'='*60}")
        print("[✓] All files processed successfully!")
        print(f"{'='*60}")

    except Exception as e:
        import traceback
        print(f"[ERROR] {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()