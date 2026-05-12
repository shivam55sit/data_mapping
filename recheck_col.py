import pandas as pd


ref_excel = pd.read_excel(r'C:\Users\shivam.prajapati\Desktop\Harshitha\Vishwa Bharathi, Excel.xlsx')

encoded_excel = pd.read_excel(r"C:\Users\shivam.prajapati\Downloads\Vishwa Bharathi, Excel_standardised_decoded.xlsx")


def fill_col(ref_df, encoded_df, ref_col_name, encoded_col_name):
    """
    Fill blank/missing values in encoded_df using values from ref_df,
    matched on 'Child Unique Code'.

    The column names can differ between the two files. For example, the
    reference file may call it 'Age (Completed Years)' while the encoded
    file calls it 'Age'.

    For every row in encoded_df where `encoded_col_name` is blank/NaN:
      1. Get the 'Child Unique Code' for that row.
      2. Look up the same 'Child Unique Code' in ref_df.
      3. Copy the value of `ref_col_name` from ref_df into encoded_df.

    Parameters
    ----------
    ref_df : pd.DataFrame
        The reference Excel DataFrame containing the source values.
    encoded_df : pd.DataFrame
        The encoded/standardised Excel DataFrame to fill.
    ref_col_name : str
        The column name in the reference Excel to fetch values from.
    encoded_col_name : str
        The column name in the encoded Excel to fill values into.

    Returns
    -------
    pd.DataFrame
        The updated encoded_df with blanks filled for the given column.
    """
    # Validate that 'Child Unique Code' exists in both DataFrames
    if 'Child Unique Code' not in ref_df.columns:
        print("Error: 'Child Unique Code' column not found in the reference Excel.")
        return encoded_df

    if 'Child Unique Code' not in encoded_df.columns:
        print("Error: 'Child Unique Code' column not found in the encoded Excel.")
        return encoded_df

    # Validate that the source column exists in the reference Excel
    if ref_col_name not in ref_df.columns:
        print(f"Error: Column '{ref_col_name}' not found in the reference Excel.")
        return encoded_df

    # If the target column doesn't exist in encoded_df at all, create it as empty
    if encoded_col_name not in encoded_df.columns:
        encoded_df[encoded_col_name] = pd.NA

    # Normalize 'Child Unique Code' to string in both DataFrames for reliable matching
    ref_df['Child Unique Code'] = ref_df['Child Unique Code'].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x).strip() if pd.notna(x) else ""
    )
    encoded_df['Child Unique Code'] = encoded_df['Child Unique Code'].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x).strip() if pd.notna(x) else ""
    )

    # Build a lookup dict: Child Unique Code -> ref_col_name value from ref_df
    lookup = {}
    for _, row in ref_df.iterrows():
        code = row['Child Unique Code']
        value = row[ref_col_name]
        if code and pd.notna(value) and str(value).strip() != "":
            lookup[code] = value

    # Fill blank/NaN values in encoded_df using the lookup
    filled_count = 0
    for idx, row in encoded_df.iterrows():
        cell_value = row[encoded_col_name]
        is_blank = pd.isna(cell_value) or str(cell_value).strip() == ""

        if is_blank:
            code = row['Child Unique Code']
            if code in lookup:
                encoded_df.at[idx, encoded_col_name] = lookup[code]
                filled_count += 1

    print(f"Filled {filled_count} blank values in '{encoded_col_name}' (from '{ref_col_name}') out of {len(encoded_df)} total rows.")
    return encoded_df


# ── Run ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    ref_col_name = input("Enter the column name in REFERENCE Excel: ").strip()
    encoded_col_name = input("Enter the column name in ENCODED Excel: ").strip()
    encoded_excel = fill_col(ref_excel, encoded_excel, ref_col_name, encoded_col_name)

    # Save the updated encoded Excel
    output_path = r"C:\Users\shivam.prajapati\Downloads\Vishwa Bharathi, Excel_standardised_decoded.xlsx"
    encoded_excel.to_excel(output_path, index=False)
    print(f"Updated file saved to: {output_path}")
