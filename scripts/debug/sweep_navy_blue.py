"""Sweep all account tabs: mark navy blue (BU/Head/Type/TCP/Narration/
Confidence/Reason) on every row that is either
  - a dual-head Beneficiary Master case (already done, but extend to the
    REASON/CONFIDENCE columns too), or
  - CONFIDENCE = "Low", or
  - currently colored orange (AI-classified by rag_classifier.py)
"""
import classify_transactions as ct
from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS, MASTER_SHEET_ID

ACCOUNT_TABS = [
    "YES BANK - 0264",
    "YES BANK - 2477",
    "YES BANK - 0490",
    "YES BANK - 0377",
    "YES BANK - 2314",
    "YES BANK - 2457",
]

ORANGE = {"red": 1.0, "green": 0.5, "blue": 0.0}


def color_is_orange(color):
    if not color:
        return False
    return (
        round(color.get("red", 0), 2) == 1.0
        and round(color.get("green", 0), 2) == 0.5
        and round(color.get("blue", 0), 2) == 0.0
    )


def main():
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(MASTER_SHEET_ID)

    for tab in ACCOUNT_TABS:
        ws = ss.worksheet(tab)
        hdr = ws.row_values(1)
        col_indices = {name: i + 1 for i, name in enumerate(hdr)}
        n_rows = ws.row_count
        col_letter_end = _col_letter(len(hdr))

        meta = ss.fetch_sheet_metadata(params={
            "ranges": [f"{tab}!A1:{col_letter_end}{n_rows}"],
            "fields": "sheets.data.rowData.values.userEnteredFormat.textFormat.foregroundColor,sheets.data.rowData.values.formattedValue",
        })
        row_data = meta["sheets"][0]["data"][0].get("rowData", [])

        i_conf = hdr.index("CONFIDENCE")
        i_head = hdr.index("HEAD")

        navy_rows = []
        for r_idx, row in enumerate(row_data[1:], start=2):
            values = row.get("values", [])
            if i_conf >= len(values) or i_head >= len(values):
                continue
            conf_val = values[i_conf].get("formattedValue", "").strip()
            head_fmt = values[i_head].get("userEnteredFormat", {}).get("textFormat", {}).get("foregroundColor", {})
            is_low_conf = conf_val.lower() == "low"
            is_orange = color_is_orange(head_fmt)
            if is_low_conf or is_orange:
                navy_rows.append(r_idx)

        if navy_rows:
            ct._mark_dual_head_rows(ws, navy_rows, col_indices)

        print(f"{tab}: {len(navy_rows)} rows marked navy blue (low-confidence or AI)")


def _col_letter(n):
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


if __name__ == "__main__":
    main()
