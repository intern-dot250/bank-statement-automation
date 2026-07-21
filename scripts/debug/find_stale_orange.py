"""Find rows with stale orange (AI) text-color formatting that have since
been manually corrected (CONFIDENCE = High), and clear the formatting."""
from upload_to_sheets import get_gspread_client, DEFAULT_CREDENTIALS, MASTER_SHEET_ID

ACCOUNT_TABS = [
    "YES BANK - 0264",
    "YES BANK - 2477",
    "YES BANK - 0490",
    "YES BANK - 0377",
    "YES BANK - 2314",
    "YES BANK - 2457",
]

AI_ORANGE = {"red": 1.0, "green": 0.5, "blue": 0.0}


def color_is_orange(color):
    if not color:
        return False
    r = round(color.get("red", 0), 2)
    g = round(color.get("green", 0), 2)
    b = round(color.get("blue", 0), 2)
    return r == 1.0 and g == 0.5 and b == 0.0


def main():
    client = get_gspread_client(DEFAULT_CREDENTIALS)
    ss = client.open_by_key(MASTER_SHEET_ID)

    for tab in ACCOUNT_TABS:
        ws = ss.worksheet(tab)
        hdr = ws.row_values(1)
        n_cols = len(hdr)
        n_rows = ws.row_count
        col_letter_end = _col_letter(n_cols)

        meta = ss.fetch_sheet_metadata(params={
            "ranges": [f"{tab}!A1:{col_letter_end}{n_rows}"],
            "fields": "sheets.data.rowData.values.userEnteredFormat.textFormat.foregroundColor,sheets.data.rowData.values.formattedValue",
        })
        row_data = meta["sheets"][0]["data"][0].get("rowData", [])

        i_head = hdr.index("HEAD")
        i_conf = hdr.index("CONFIDENCE")

        stale_rows = []
        for r_idx, row in enumerate(row_data[1:], start=2):
            values = row.get("values", [])
            if i_head >= len(values) or i_conf >= len(values):
                continue
            head_fmt = values[i_head].get("userEnteredFormat", {}).get("textFormat", {}).get("foregroundColor", {})
            conf_val = values[i_conf].get("formattedValue", "")
            if color_is_orange(head_fmt) and conf_val.strip() == "High":
                stale_rows.append(r_idx)

        print(f"{tab}: {len(stale_rows)} stale orange rows -> {stale_rows}")


def _col_letter(n):
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


if __name__ == "__main__":
    main()
