# FINAL WORKING - COLLECT 
# FINAL WORKING - COLLECT ALL 60HP20200P + ALL 60HP20210P PER BLOCK
import pandas as pd
from pathlib import Path
import numpy as np
# ────────────────────────────────────────────────
# CONFIGURATION
# ────────────────────────────────────────────────
# EXCEL_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\CARL-san\Documents\AI 60HP NEW LINE\ANALIZATION\HP 60_1.xlsx"
# OUTPUT_CSV = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\CARL-san\Documents\AI 60HP NEW LINE\ANALIZATION\60 HP PYTHON.csv"
EXCEL_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\CARL-san\Documents\AI 60HP NEW LINE\ANALIZATION\80HP.xlsx"
OUTPUT_CSV = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\CARL-san\Documents\AI 60HP NEW LINE\ANALIZATION\80 HP PYTHON.csv"


PREFIXES = [
    '80HP10572P',
    '80HP20027P',
    '80HP20210P',
    '80HP20220P',
    '80HP20240P',
    '80HP20260P',
    '80HP20300P',
    '80HP20501P',
    '80HP20502P',
    '80HP20503P',
    '80HP20504P',
    '80HP20505P',
    '80HP20506P',
    '80HP20507P',
    '80HP20508P',
    '80HP20509P',
    '80HP20510P',
    '80HP20511P',
    '80HP20512P',
    '80HP20513P',
    '80HP20515P',
    '80HP20516P',
    '80HP20540P',
    '80HP20541P',
    '80HP20542P',
    '80HP20545P',
    '80HP20550P',
    '80HP20562P',
    '80HP20572P',
    '80HP20602P',
    '80HP20603P',
    '80HP20612P',
    '80HP20613P',
    '80HP20614P',
    '80HP20632P',
    '80HP20633P',
    '80HP20639P',
    '80HP20652P',
    '80HP20662P',
    '80HP20700P',
    '80HP20710P',
    '80HP20740P',
    '80HP20750P',
    '80HP20760P',
    '80HP20780P',
    '80HP20791P',
    '80HP20802P',
    '80HP20803P',
    '80HP20812P',
    '80HP20820P',
    '80HP20850P',
    '80HP20852P',
    '80HP20860P',
    '80HP20996P',
    '80HP29500P',
    '80HP29501P',
    '80HP29540P',
    '80HP40041P',
    '80HP40045P',
    '80HP40046P',
    '80HP40063P',
    '80HP40068P',
    '80HP40073P',
    '80HP40074P',
    '80HP40075P',
    '80HP40076P',
    '80HP40081P',
    '80HP40090P',
    '80HP40091P',
    '80HP40100P',
    '80HP40101P',
    '80HP40572P',
    '80HP40630P',
    '80HP40701P',
    '80HP40750P',
    '80HP40800P',
    '80HP40801P',
    '80HP40995P',
    '80HP40997P',
    '80HP40998P',
    '80HP40999P',
    '80HP49701P',
    '80HP50005P',
    '80HPB4001P'

]

# PREFIXES = [
#     "60HP20200P",
#     "60HP20210P",
#     "60HP20220P",
#     "60HP20502P",
#     "60HP20503P",
#     "60HP20504P",
#     "60HP20512P",
#     "60HP20513P",
#     "60HP20562P",
#     "60HP20572P",
#     "60HP20602P",
#     "60HP20603P",
#     "60HP20632P",
#     "60HP20633P",
#     "60HP20662P",
#     "60HP20672P",
#     "60HP20681P",
#     "60HP20700P",
#     "60HP20760P",
#     "60HP20770P",
#     "60HP20780P",
#     "60HP20802P",
#     "60HP20803P",
#     "60HP20812P",
#     "60HP40004P",
#     "60HP40041P",
#     "60HP40045P",
#     "60HP40046P",
#     "60HP40068P",
#     "60HP40074P",
#     "60HP40076P",
#     "60HP40090P",
#     "60HP40091P",
#     "60HP40100P",
#     "60HP40101P",
#     "60HP40572P",
#     "60HP40630P",
#     "60HP40701P",
#     "60HP40800P",
#     "60HP40801P",
#     "60HP40995P",
#     "60HP40997P",
#     "60HP40998P",
#     "60HP40999P",
#     "60HP50005P",
#     "60HPB4001P"
# ]


# ────────────────────────────────────────────────
def main():
    print("=== Extractor: ALL LISTED 60HP PREFIXES ===\n")
    print("Reading Excel file...\n")
    try:
        df = pd.read_excel(EXCEL_PATH, header=None, dtype=str)
    except Exception as e:
        print("ERROR reading file:")
        print(str(e))
        return
    # Clean data
    df = df.fillna('').astype(str).apply(lambda col: col.str.strip())
    LABEL_ROW = 0
    DATA_START_ROW = 1
    # Basic validation
    row0 = df.iloc[LABEL_ROW].str.lower().str.strip()
    if not any("item" in str(x) for x in row0):
        print("Warning: Row 0 may not contain column labels")
        print("Row 0 sample:", df.iloc[0].tolist()[:8], "...\n")
    header_data_row = df.iloc[DATA_START_ROW]
    block_start_cols = header_data_row[header_data_row.str.startswith(tuple(PREFIXES))].index.tolist()
    if not block_start_cols:
        print("No blocks starting with listed prefixes found in row 1!")
        print("Row 1 sample:", header_data_row.tolist()[:12], "...")
        return
    print(f"Found {len(block_start_cols)} block(s)\n")
    # ── Collect all matching rows ─────────────────────────────────────────
    all_rows = {prefix: [] for prefix in PREFIXES}
    for col_start in block_start_cols:
        block = df.iloc[DATA_START_ROW:, col_start : col_start + 4].copy()
        if block.empty:
            continue
        block.columns = ['Item', 'Item Description', 'Material', 'Material Description']
        # Remove fully empty rows
        block = block.replace('', pd.NA).dropna(how='all')
        # Filter for each prefix
        for prefix in PREFIXES:
            matching = block[block['Item'].str.startswith(prefix, na=False)]
            if not matching.empty:
                all_rows[prefix].append(matching)
    # Check if any data was found
    has_data = any(all_rows[prefix] for prefix in PREFIXES)
    if not has_data:
        print("No rows found for any of the listed prefixes")
        return
    # ── Combine into final structure ──────────────────────────────────────
    df_list = []
    counts = {}
    max_len = 0
    for prefix in PREFIXES:
        if all_rows[prefix]:
            df_prefix = pd.concat(all_rows[prefix], ignore_index=True)
        else:
            df_prefix = pd.DataFrame(columns=['Item', 'Item Description', 'Material', 'Material Description'])
        count = len(df_prefix)
        counts[prefix] = count
        max_len = max(max_len, count)
        # Rename columns to include prefix
        df_prefix = df_prefix.rename(columns={
            'Item': f'{prefix}_Item',
            'Item Description': f'{prefix}_Item_Description',
            'Material': f'{prefix}_Material',
            'Material Description': f'{prefix}_Material_Description'
        })
        # Reindex to max_len
        df_prefix = df_prefix.reindex(range(max_len))
        df_list.append(df_prefix)
    # Final horizontal concatenation
    final_df = pd.concat(df_list, axis=1)
    # Clean up
    final_df = final_df.replace({pd.NA: '', np.nan: ''})
    # ── Save ──────────────────────────────────────────────────────────────
    try:
        final_df.to_csv(OUTPUT_CSV, index=False, encoding='utf-8-sig')
        print("╔════════════════════════════════════════════════════════════╗")
        print("║ SUCCESS - ALL LISTED 60HP PREFIXES extracted        ║")
        print("╚════════════════════════════════════════════════════════════╝")
        print(f"Saved to: {OUTPUT_CSV}")
        print(f"Total rows: {len(final_df):,}")
        for prefix, count in counts.items():
            print(f"Items for {prefix}: {count:,}")
        print("Columns:", list(final_df.columns))
        print("Done!\n")
    except Exception as e:
        print("Failed to save:", str(e))
        fallback = Path.home() / "Desktop" / "60HP_ALL_PREFIXES.csv"
        final_df.to_csv(fallback, index=False, encoding='utf-8-sig')
        print(f"SAVED TO DESKTOP: {fallback}")
if __name__ == "__main__":
    main()