import os
import re

import numpy as np
import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename


# =========================
# Global configuration
# =========================

# Directory where Workday source Excel files are located
SOURCE_DIR = r"TEMP"

# Directory where formatted output files will be written
OUTPUT_DIR = r"TEMP"


class OutPutClass:
    """
    Stores header metadata from the Workday export.

    Only the Effective Date is required, which is used to name
    the output Excel file.
    """
    def __init__(self):
        self.EffectiveDate = ""


def SelectOneExcelFile():
    """
    Opens a file picker allowing the user to select a single
    Workday Costing Allocations Excel file.
    """
    Tk().withdraw()  # Hide the root Tk window

    filename = askopenfilename(
        title="Select Workday Costing Allocations Excel file",
        initialdir=SOURCE_DIR,
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    return filename


def ReadInFirst15Lines(filename):
    """
    Reads the header section at the top of the Workday export
    and extracts the Effective Date.
    """
    df = pd.read_excel(filename, skiprows=1, nrows=15, header=None)
    output = OutPutClass()

    for _, row in df.iterrows():
        if pd.isna(row[0]):
            continue

        key = str(row[0]).strip().lower().replace(" ", "").replace("_", "")
        value = "" if pd.isna(row[1]) else str(row[1]).strip()

        if key == "effectivedate":
            output.EffectiveDate = value

    return output


def CreateOutPutFileName(outputData: OutPutClass):

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    raw_date = (outputData.EffectiveDate or "").strip()

    if raw_date == "":
        formatted_date = "unknowndate"
    else:
        try:
            date_obj = pd.to_datetime(raw_date)
            formatted_date = date_obj.strftime("%m_%d_%Y")
        except Exception:
            formatted_date = (
                str(raw_date)
                .replace("/", "_")
                .replace("\\", "_")
                .replace(":", "_")
            )

    filename = f"{formatted_date} Costing Allocations.xlsx"
    return os.path.join(OUTPUT_DIR, filename)


def _extract_cc_number(cost_center_cell) -> str:
    """
    Extracts the short Cost Center code (e.g. CC0075)
    from a longer Workday cost center string.
    """
    s = "" if pd.isna(cost_center_cell) else str(cost_center_cell).strip()
    match = re.search(r"\bCC\d+\b", s, flags=re.IGNORECASE)
    return match.group(0).upper() if match else ""


def _split_worker_name(worker_cell) -> tuple[str, str]:
    """
    Splits the Workday Worker field into (first_name, last_name).

    Handles:
      - "Last, First"
      - "First Middle Last"
    """
    s = "" if pd.isna(worker_cell) else str(worker_cell).strip()
    if not s:
        return ("", "")

    if "," in s:
        parts = [p.strip() for p in s.split(",") if p.strip()]
        last = parts[0] if len(parts) >= 1 else ""
        first = parts[1].split()[0] if len(parts) >= 2 else ""
        return (first, last)

    tokens = s.split()
    if len(tokens) == 1:
        return (tokens[0], "")
    return (tokens[0], tokens[-1])


def _parse_fte(fte_cell) -> float:
    """
    Converts Workday FTE values (e.g. '50.%') into a decimal (0.5)
    so Excel can display them as percentages.
    """
    if pd.isna(fte_cell):
        return np.nan

    s = str(fte_cell).strip().replace("%", "")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s == "":
        return np.nan

    val = float(s)
    return val / 100.0 if val > 1 else val


def _format_mmddyyyy(date_cell) -> str:
    """
    Converts an Excel date cell into mm/dd/yyyy format.
    Returns blank if missing or invalid.
    """
    if pd.isna(date_cell) or str(date_cell).strip() == "":
        return ""

    dt = pd.to_datetime(date_cell, errors="coerce")
    if pd.isna(dt):
        return ""

    return dt.strftime("%m/%d/%Y")


def _extract_budget_number(row: pd.Series) -> str:
    """
    Payroll budget logic:

    - Budget number may be in Program, Grant, or Gift
    - Prefer Program > Grant > Gift
    - Return number only (first token)
    """
    for col in ["Program", "Grant", "Gift"]:
        if col in row.index:
            v = "" if pd.isna(row[col]) else str(row[col]).strip()
            if v:
                return v.split()[0]
    return ""


def ReadInCostingAllocationsFile(filename):
    """
    Reads the Workday table (header row is always Excel row 17)
    and transforms it into the required output format.
    """
    df = pd.read_excel(filename, header=16)
    df = df.dropna(how="all").reset_index(drop=True)

    required = ["Worker", "Title", "FTE", "Start Date", "End Date", "Distribution Percent"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        messagebox.showerror(
            "column name error",
            "The excel file layout has changed.\n\n"
            "Missing columns:\n\n"
            + "\n".join(missing)
        )
        return None

    out = pd.DataFrame()

    # Handle duplicate Cost Center columns
    cc_org_col = "Cost Center" if "Cost Center" in df.columns else None
    cc_alloc_col = "Cost Center.1" if "Cost Center.1" in df.columns else None

    if cc_org_col is None and cc_alloc_col is None:
        messagebox.showerror(
            "column name error",
            "No Cost Center column found."
        )
        return None

    def _pick_cost_center(row):
        alloc = _extract_cc_number(row[cc_alloc_col]) if cc_alloc_col else ""
        org = _extract_cc_number(row[cc_org_col]) if cc_org_col else ""
        return alloc if alloc else org

    out["Cost Center"] = df.apply(_pick_cost_center, axis=1)

    first_last = df["Worker"].apply(_split_worker_name)
    out["Last Name"] = first_last.apply(lambda t: t[1])
    out["First Name"] = first_last.apply(lambda t: t[0])

    out["Title"] = df["Title"].fillna("").astype(str).str.strip()
    out["FTE"] = df["FTE"].apply(_parse_fte)
    out["Start Date"] = df["Start Date"].apply(_format_mmddyyyy)
    out["End Date"] = df["End Date"].apply(_format_mmddyyyy)

    dist = pd.to_numeric(df["Distribution Percent"], errors="coerce")
    out["Distribution Percent"] = np.where(dist > 1, dist / 100, dist)

    out["Budget"] = df.apply(_extract_budget_number, axis=1)

    # Remove rows with no identifying information
    out = out[
        ~(
            (out["Cost Center"] == "") &
            (out["First Name"] == "") &
            (out["Last Name"] == "")
        )
    ]

    # Enforce final column order
    out = out[
        [
            "Cost Center",
            "Last Name",
            "First Name",
            "Title",
            "FTE",
            "Start Date",
            "End Date",
            "Distribution Percent",
            "Budget",
        ]
    ]

    # Sort output: Cost Center first, then Last Name
    out = out.sort_values(
        by=["Cost Center", "Last Name"],
        ascending=[True, True],
        kind="mergesort"
    ).reset_index(drop=True)

    return out


def WriteCostingAllocationsToExcel(df_out, output_filename):
    """
    Writes the output dataframe to Excel and applies formatting.
    """
    with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, sheet_name="Costing Allocations", index=False)

        worksheet = writer.sheets["Costing Allocations"]
        workbook = writer.book

        percent_fmt = workbook.add_format({"num_format": "0.00%"})

        col_widths = {
            "Cost Center": 12,
            "Last Name": 18,
            "First Name": 18,
            "Title": 35,
            "FTE": 10,
            "Start Date": 12,
            "End Date": 12,
            "Distribution Percent": 20,
            "Budget": 22,
        }

        for col, width in col_widths.items():
            idx = df_out.columns.get_loc(col)
            worksheet.set_column(idx, idx, width)

        worksheet.set_column(
            df_out.columns.get_loc("FTE"),
            df_out.columns.get_loc("FTE"),
            col_widths["FTE"],
            percent_fmt,
        )

        worksheet.set_column(
            df_out.columns.get_loc("Distribution Percent"),
            df_out.columns.get_loc("Distribution Percent"),
            col_widths["Distribution Percent"],
            percent_fmt,
        )


def main():
    """
    Main program flow.
    """
    file_to_open = SelectOneExcelFile()
    if not file_to_open:
        return

    header_data = ReadInFirst15Lines(file_to_open)
    output_file_path = CreateOutPutFileName(header_data)

    if os.path.exists(output_file_path):
        messagebox.showerror(
            "file already exists",
            f"The output file already exists:\n\n{output_file_path}"
        )
        return

    df_out = ReadInCostingAllocationsFile(file_to_open)
    if df_out is None:
        return

    WriteCostingAllocationsToExcel(df_out, output_file_path)

    messagebox.showinfo("Done", f"Output created:\n\n{output_file_path}")


if __name__ == "__main__":
    main()
