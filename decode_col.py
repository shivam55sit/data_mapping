import pandas as pd
import numpy as np


decode = {
    'Rub Eyes B2': {0: 'No', 1: 'Yes'},



'Rub Freq B3': {1: 'Multiple times a day', 2: 'Few days/week', 3: 'Few days a month', 4: 'Few episodes per month', 5: 'Do not keep count', -98: 'Not applicable'},



    'Rub Reason B4': {1: 'Vision blurry', 2: 'Itching of eyes', 3: 'Pain in the eyes', 4: 'Others, Specify', -98: 'Not Applicable'},



    'Rub Method B5': {1: 'The back of your hand', 2: 'Your fingertipd', 3: 'Your Knuckles', 4: 'The base of your thumb', -98: 'Not Aplicable'},



    'Think Damage Rub B6': {1: 'Definitely Yes', 2: 'Probably Yes', 3: 'Not sure', 4: 'Prabably No', 5: 'Definitely yes', -98: 'Not Applicable'},



 

    'Eye Redness B8': {0: 'No', 1: 'Yes'},



    'Redness Frequency B9': {1: 'Everyday', 2: 'Once in 2-3 days (very frequently)', 3: 'Once in 2-3 weeks', 4: 'Few episodes per year', -98: 'Not Applicable'},



    'Eye Itching B10': {0: 'No', 1: 'yes'},



    'Itching Frequency B11': {1: 'Everyday', 2: 'Once in 2-3 days (very frequently)', 3: 'Once in 2-3 weeks', 4: 'Few episodes per year', -98: 'Not Applicable'},



    'Difficulty School Work B12': {1: 'Very frequently', 2: 'Frequently', 3: 'Occasionally', 4: 'Rarely', 5: 'Never', -98: 'Not Applicable'},



    'Attention School B13': {1: 'Very frequently', 2: 'Frequently', 3: 'Occasionally', 4: 'Rarely', 5: 'Never', -98: 'Not Applicable'},



    'School Absent B14': {1: 'Very frequently', 2: 'Frequently', 3: 'Occasionally', 4: 'Rarely', 5: 'Never', -98: 'Not Applicable'},



    'Sleep Face Down B15': {0: 'No', 1: 'Yes', -99: "Don't know"},



    'Sneezing Breath B16': {0: 'No', 1: 'Yes', -99: "Don't know"},



    'Sneeze Freq B17': {1: 'Once or twice a year', 2: 'Once amonth3.Less than 2 times per month', 3: 'More than 2 times per month', 5: 'Cannot remember bu it happens', -98: 'Not Apllicable'},



    'Skin Itch Rash B18': {0: 'No', 1: 'yes'},



    'Medication Allergy B19': {0: 'No', 1: 'Yes'},



    'Medication Name B20': {-98: 'not Appliable', -99: "don'y know"},



    'Handedness B21': {1: 'Left handed', 2: 'Right Handed'},

 'Eye Surgery B22': {0: 'No', 1: 'Yes'},

}


def decode_columns(excel_path, sheet_name=0, header=1, output_path=None):
    """
    Read an Excel file and replace encoded numeric values with their
    actual text labels for every column defined in the `decode` dictionary.

    Parameters
    ----------
    excel_path : str
        Path to the input Excel file.
    sheet_name : int or str, optional
        Sheet to read (default 0).
    header : int, optional
        Row number to use as column header (default 1).
    output_path : str or None, optional
        If provided, save the decoded DataFrame to this path.
        If None, saves to '<original_name>_decoded.xlsx'.

    Returns
    -------
    pd.DataFrame
        DataFrame with encoded numbers replaced by text values.
    """
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)

    # Cast to string and strip whitespace (some headers may be NaN/float)
    df.columns = df.columns.astype(str).str.strip()

    replaced_count = 0

    for col_name, value_map in decode.items():
        # Try exact match first
        if col_name in df.columns:
            df[col_name] = df[col_name].astype(object)
            df[col_name] = df[col_name].replace(value_map)
            replaced_count += 1
            print(f"  [✓] Decoded: {col_name[:60]}...")
        else:
            # Try prefix match — the Excel header may be longer than the key
            matched = [c for c in df.columns if isinstance(c, str) and c.startswith(col_name)]
            if matched:
                for m in matched:
                    df[m] = df[m].astype(object)
                    df[m] = df[m].replace(value_map)
                    replaced_count += 1
                    print(f"  [✓] Decoded (prefix): {m[:60]}...")
            else:
                print(f"  [–] Column not found: {col_name[:60]}...")

    print(f"\n  Total columns decoded: {replaced_count}/{len(decode)}")

    # Save output
    if output_path is None:
        import os
        base, ext = os.path.splitext(excel_path)
        output_path = f"{base}_decoded.xlsx"

    df.to_excel(output_path, index=False)
    print(f"  [✓] Saved decoded file: {output_path}")

    return df


if __name__ == "__main__":
    input_file = r"C:\Users\shivam.prajapati\Downloads\Vishwa Bharathi, Excel1_standardised.xlsx"
    decoded_df = decode_columns(input_file)