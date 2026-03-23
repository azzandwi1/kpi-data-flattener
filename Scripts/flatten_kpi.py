import re
import pandas as pd

REGIONS = [
    "BALI", "INDOTIM", "JABODETABEK", "JAWA",
    "KALIMANTAN", "SULAWESI", "SUMATERA", "NUSRA",
]

GROUP_MAP = {
    "JAWA": "JAWABALI",
    "BALI": "JAWABALI",
    "KALIMANTAN": "KALNUSRA",
    "NUSRA": "KALNUSRA",
    "INDOTIM": "SULTIM",
    "SULAWESI": "SULTIM",
}

CATEGORY_BY_SHEET = {
    "NON COD": "NON COD",
    "COD": "COD",
    "KK": "KK",
    "OS NON COD": "NON COD",
    "OS COD": "COD",
    "OS KK": "KK",
    "OTP NON COD": "NON COD",
    "OTP COD": "COD",
    "OTP KK": "KK",
}
OS_OVER_EXC_KPI = "OS Over exc Shipment Hold dan Proses Return Balik"
OS_KPI_ALIASES = {"OS NON COD", "OS COD", "OS KK"}

PIC_SUFFIX_BY_REGION = {
    ("JAWA", "INDRA CHOIRUL HALIM"): "INDRA CHOIRUL HALIM - JAWA",
    ("BALI", "INDRA CHOIRUL HALIM"): "INDRA CHOIRUL HALIM - BALI",
}
KALNUSRA_COMBINE_PICS = {"AJENG", "ANDHIKA WAHYU"}
KALNUSRA_MEMBERS = {"KALIMANTAN", "NUSRA"}
SULTIM_COMBINE_PICS = {"REHULINA"}
SULTIM_MEMBERS = {"INDOTIM", "SULAWESI"}
PIC_BY_REGION = {
    "BALI": {"INDRA CHOIRUL HALIM"},
    "INDOTIM": {"REHULINA", "YUDHISTIRA"},
    "JABODETABEK": {"ANGGITA", "BAYU MAULANA", "DILLA"},
    "JAWA": {
        "FAIZAHTUL HASANAH",
        "IKHSANUL PUTRA SANDY",
        "INDRA CHOIRUL HALIM",
        "LUTHFI KURNIAWAN",
        "TEGUH MAULANA AZMI",
    },
    "KALIMANTAN": {"AJENG", "ANDHIKA DWI S", "ANDHIKA WAHYU"},
    "NUSRA": {"AJENG", "ANDHIKA WAHYU"},
    "SULAWESI": {"REHULINA", "RITFAL", "SISWANDO"},
    "SUMATERA": {"ANGGA", "BINSAR", "KIKY", "LUKMAN"},
}
# In sheet STUCK, only process pivots before this Excel column boundary.
# Boundary between pivot #3 and #4 is at column BC.
STUCK_MAX_COL_EXCLUSIVE_IDX = 54  # BC (1-based 55) in 0-based indexing


def should_combine_pic(region, pic):
    return (
        (region in KALNUSRA_MEMBERS and pic in KALNUSRA_COMBINE_PICS)
        or (region in SULTIM_MEMBERS and pic in SULTIM_COMBINE_PICS)
    )


def normalize_kpi_title(kpi_title):
    text = str(kpi_title).strip()
    if text.upper() in OS_KPI_ALIASES:
        return OS_OVER_EXC_KPI
    # Handle readable variants like:
    # "OS Over exc Shipment Hold dan Proses Return Balik (COD)"
    # while keeping category sourced from sheet name.
    if re.match(
        r"^OS Over exc Shipment Hold dan Proses Return Balik\s*\((COD|NON COD|KK)\)$",
        text,
        flags=re.IGNORECASE,
    ):
        return OS_OVER_EXC_KPI
    return text


def split_kpi_and_category_from_title(title, default_category=None):
    text = str(title).strip()
    m = re.search(r"\((COD|NON COD|KK)\)\s*$", text, flags=re.IGNORECASE)
    if m:
        category = m.group(1).upper()
        kpi = re.sub(r"\s*\((COD|NON COD|KK)\)\s*$", "", text, flags=re.IGNORECASE).strip()
        return kpi, category
    return text, default_category


def infer_otp_kpi_title(df, header_row_idx, header_col_idx, fallback_title):
    # Scan banner rows above a pivot and detect OTP KPI title.
    row_start = max(0, header_row_idx - 10)
    row_end = header_row_idx
    col_end = min(df.shape[1], header_col_idx + 16)
    patterns = ("ON TIME PERFORMANCE SLA INTERNAL", "ON TIME PERFORMANCE SLA PUBLISH")

    for r in range(row_start, row_end):
        for c in range(header_col_idx, col_end):
            cell = df.iloc[r, c]
            if pd.isna(cell):
                continue
            text = str(cell).strip()
            if not text:
                continue
            text_upper = text.upper()
            if any(p in text_upper for p in patterns):
                return text
    return fallback_title


def find_header_rows(df):
    header_positions = []
    for i in range(len(df)):
        row = df.iloc[i]
        matches = row.astype(str).str.strip().str.lower() == "row labels"
        if matches.any():
            for col_idx in matches[matches].index.tolist():
                header_positions.append((i, col_idx))
    return header_positions


def find_region_header(df):
    for i in range(len(df)):
        row = df.iloc[i]
        for col_idx, val in row.items():
            if isinstance(val, str) and val.strip().upper() == "REGION":
                return (i, col_idx)
    return None


def find_all_region_headers(df):
    headers = []
    for i in range(len(df)):
        row = df.iloc[i]
        for col_idx, val in row.items():
            if isinstance(val, str) and val.strip().upper() == "REGION":
                headers.append((i, col_idx))
    return headers


def find_grand_total_col(header_row, start_idx):
    for j in range(start_idx + 1, len(header_row)):
        label = header_row.iloc[j]
        if pd.isna(label):
            continue
        if str(label).strip().lower() == "grand total":
            return j
    for j in range(start_idx + 1, len(header_row)):
        label = header_row.iloc[j]
        if pd.isna(label):
            continue
        return j
    return None


def is_region(label):
    return label.upper() in REGIONS


def find_otp_qty_cols(df, header_row_idx, header_col_idx, force_offsets=False):
    header_row = df.iloc[header_row_idx]
    on_time_col = None
    over_time_col = None
    total_col = None

    # Strategy:
    # - Find the first two occurrences of "AWB (Qty)" to the right of Region: ON TIME and OVER TIME
    # - Find "Total AWB (Qty)" for total column (optional)
    awb_qty_cols = []
    for j in range(header_col_idx + 1, len(header_row)):
        label = header_row.iloc[j]
        if pd.isna(label):
            continue
        label_str = str(label).strip().upper()
        if label_str == "AWB (QTY)":
            awb_qty_cols.append(j)
        if label_str == "TOTAL AWB (QTY)":
            total_col = j

    if not force_offsets and len(awb_qty_cols) >= 2:
        on_time_col = awb_qty_cols[0]
        over_time_col = awb_qty_cols[1]
    else:
        # Fallback to fixed offsets based on observed layout:
        # Region in header_col_idx, ON TIME (Qty) at +1, OVER TIME (Qty) at +3
        if header_col_idx + 1 < len(header_row):
            on_time_col = header_col_idx + 1
        if header_col_idx + 3 < len(header_row):
            over_time_col = header_col_idx + 3

    return on_time_col, over_time_col, total_col


def infer_stuck_kpi(df, header_row_idx, header_col_idx, default_title):
    # Look for the ON/OVER SLA label above the pivot to determine KPI.
    search_start = max(0, header_row_idx - 6)
    for r in range(header_row_idx - 1, search_start - 1, -1):
        cell = df.iloc[r, header_col_idx]
        if pd.isna(cell):
            continue
        text = str(cell).strip().upper()
        if "ON/OVER SLA" in text or "ON/OVER" in text:
            if "EXC SHIPMENT HOLD" in text:
                return "OS Stuck Exc Shipment Hold"
            if "SHIPMENT HOLD" in text:
                return "OS Stuck Shipment Hold"
            if "INCOMING" in text:
                return "OS Stuck Incoming"
    return default_title


def flatten_pivot(input_path, output_path, week_label="Week 3 Jan"):
    xl = pd.ExcelFile(input_path)
    rows = []
    combined_rows = {}

    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None)

        kpi_title = df.iloc[1, 2] if len(df) > 1 and df.shape[1] > 2 else None
        if pd.isna(kpi_title):
            kpi_title = sheet
        kpi_title = normalize_kpi_title(kpi_title)

        sheet_upper = str(sheet).strip().upper()
        kategori = CATEGORY_BY_SHEET.get(sheet_upper, None)
        is_stuck_sheet = sheet_upper in {"STUCK", "OS STUCK"}
        is_otp_category_sheet = sheet_upper in {"OTP COD", "OTP NON COD", "OTP KK"}

        if is_otp_category_sheet:
            region_headers = sorted(find_all_region_headers(df), key=lambda x: (x[1], x[0]))
            for header_row_idx, header_col_idx in region_headers:
                pivot_title_raw = infer_otp_kpi_title(df, header_row_idx, header_col_idx, sheet)
                pivot_kpi_title, pivot_kategori = split_kpi_and_category_from_title(
                    pivot_title_raw, default_category=kategori
                )
                pivot_kpi_title = normalize_kpi_title(pivot_kpi_title)

                on_time_col, over_time_col, total_col = find_otp_qty_cols(
                    df, header_row_idx, header_col_idx, force_offsets=False
                )
                # These OTP sheets should behave like Tiket Investigasi style:
                # ON TIME / OVER TIME only, no TOTAL OS requirement.
                if on_time_col is None and (header_col_idx + 1) < df.shape[1]:
                    on_time_col = header_col_idx + 1
                if over_time_col is None and (header_col_idx + 3) < df.shape[1]:
                    over_time_col = header_col_idx + 3

                if on_time_col is None and over_time_col is None:
                    continue

                found_otp = {}
                region_vals = {}
                current_region = None

                for r in range(header_row_idx + 1, len(df)):
                    label = df.iloc[r, header_col_idx]
                    if pd.isna(label):
                        continue
                    label_str = str(label).strip()
                    if label_str == "":
                        continue
                    if label_str.lower() == "grand total":
                        break

                    if is_region(label_str):
                        current_region = label_str.upper()
                        on_region = 0
                        over_region = 0
                        if on_time_col is not None:
                            val = df.iloc[r, on_time_col]
                            if not pd.isna(val):
                                try:
                                    on_region = float(val)
                                except Exception:
                                    on_region = val
                        if over_time_col is not None:
                            val = df.iloc[r, over_time_col]
                            if not pd.isna(val):
                                try:
                                    over_region = float(val)
                                except Exception:
                                    over_region = val
                        region_vals[current_region] = (on_region, over_region)
                        continue

                    if current_region is None:
                        continue

                    key = (current_region, label_str.upper())

                    if on_time_col is not None:
                        val = df.iloc[r, on_time_col]
                        if not pd.isna(val):
                            try:
                                val = float(val)
                            except Exception:
                                pass
                            found_otp[(key, "ON TIME")] = found_otp.get((key, "ON TIME"), 0) + val

                    if over_time_col is not None:
                        val = df.iloc[r, over_time_col]
                        if not pd.isna(val):
                            try:
                                val = float(val)
                            except Exception:
                                pass
                            found_otp[(key, "OVER TIME")] = found_otp.get((key, "OVER TIME"), 0) + val

                found_pic_pairs = {(r, p) for ((r, p), _) in found_otp.keys()}

                for region in REGIONS:
                    for pic in sorted(PIC_BY_REGION.get(region, set())):
                        output_region = GROUP_MAP.get(region, region)
                        output_pic = PIC_SUFFIX_BY_REGION.get((region, pic), pic)
                        if (region, pic) in found_pic_pairs:
                            on_time = found_otp.get(((region, pic), "ON TIME"), 0)
                            over_time = found_otp.get(((region, pic), "OVER TIME"), 0)
                        else:
                            on_time, over_time = region_vals.get(region, (0, 0))
                        total_os = 0

                        if should_combine_pic(region, pic):
                            key = (output_region, pivot_kpi_title, output_pic, pivot_kategori)
                            current = combined_rows.get(key, {"ON TIME": 0, "OVER TIME": 0, "TOTAL OS": 0})
                            current["ON TIME"] += on_time
                            current["OVER TIME"] += over_time
                            current["TOTAL OS"] += total_os
                            combined_rows[key] = current
                        else:
                            rows.append({
                                "Week": week_label,
                                "Region": output_region,
                                "KPI INDICES": pivot_kpi_title,
                                "PIC": output_pic,
                                "Kategori": pivot_kategori,
                                "ON TIME": on_time,
                                "OVER TIME": over_time,
                                "TOTAL OS": total_os,
                            })
            continue

        if sheet_upper == "TIKET INVEST":
            # Two pivots: left = Tiket Investigasi Terupload (On SLA/Over SLA),
            # right = Tingkat Investigasi Kalah (Kalah/Menang)
            region_headers = find_all_region_headers(df)
            if region_headers:
                region_headers = sorted(region_headers, key=lambda x: (x[1], x[0]))
                left_pos = region_headers[0]
                right_pos = region_headers[-1] if len(region_headers) > 1 else None

                def process_tiket_pivot(header_row_idx, header_col_idx, kpi_name, on_off_offsets):
                    current_region = None
                    region_vals = {}
                    for r in range(header_row_idx + 1, len(df)):
                        label = df.iloc[r, header_col_idx]
                        if pd.isna(label):
                            continue
                        label_str = str(label).strip()
                        if label_str == "":
                            continue
                        if label_str.lower() == "grand total":
                            break

                        if is_region(label_str):
                            current_region = label_str.upper()
                            # Capture region totals directly from the region row
                            on_col = header_col_idx + on_off_offsets[0]
                            off_col = header_col_idx + on_off_offsets[1]

                            on_val = 0
                            off_val = 0
                            if on_col < df.shape[1]:
                                val = df.iloc[r, on_col]
                                if not pd.isna(val):
                                    try:
                                        on_val = float(val)
                                    except Exception:
                                        on_val = val
                            if off_col < df.shape[1]:
                                val = df.iloc[r, off_col]
                                if not pd.isna(val):
                                    try:
                                        off_val = float(val)
                                    except Exception:
                                        off_val = val
                            region_vals[current_region] = (on_val, off_val)
                            continue

                        if current_region is None:
                            continue

                        # If PIC rows exist, capture them too
                        key = (current_region, label_str.upper())
                        on_col = header_col_idx + on_off_offsets[0]
                        off_col = header_col_idx + on_off_offsets[1]

                        if on_col < df.shape[1]:
                            val = df.iloc[r, on_col]
                            if not pd.isna(val):
                                try:
                                    val = float(val)
                                except Exception:
                                    pass
                                found_otp[(key, "ON TIME")] = found_otp.get((key, "ON TIME"), 0) + val

                        if off_col < df.shape[1]:
                            val = df.iloc[r, off_col]
                            if not pd.isna(val):
                                try:
                                    val = float(val)
                                except Exception:
                                    pass
                                found_otp[(key, "OVER TIME")] = found_otp.get((key, "OVER TIME"), 0) + val

                    return region_vals

                # Use found_otp for ON/OFF storage (no TOTAL OS)
                found_otp = {}

                # Left pivot: On SLA (col D) and Over SLA (col F)
                left_region_vals = process_tiket_pivot(left_pos[0], left_pos[1], "Tiket Investigasi Terupload", (1, 3))
                for region in REGIONS:
                    for pic in sorted(PIC_BY_REGION.get(region, set())):
                        output_region = GROUP_MAP.get(region, region)
                        output_pic = PIC_SUFFIX_BY_REGION.get((region, pic), pic)
                        if (region, pic) in {(r, p) for ((r, p), _) in found_otp.keys()}:
                            on_time = found_otp.get(((region, pic), "ON TIME"), 0)
                            over_time = found_otp.get(((region, pic), "OVER TIME"), 0)
                        else:
                            on_time, over_time = left_region_vals.get(region, (0, 0))
                        total_os = 0
                        if should_combine_pic(region, pic):
                            key = (output_region, "Tiket Investigasi Terupload", output_pic, kategori)
                            current = combined_rows.get(key, {"ON TIME": 0, "OVER TIME": 0, "TOTAL OS": 0})
                            current["ON TIME"] += on_time
                            current["OVER TIME"] += over_time
                            combined_rows[key] = current
                        else:
                            rows.append({
                                "Week": week_label,
                                "Region": output_region,
                                "KPI INDICES": "Tiket Investigasi Terupload",
                                "PIC": output_pic,
                                "Kategori": kategori,
                                "ON TIME": on_time,
                                "OVER TIME": over_time,
                                "TOTAL OS": total_os,
                            })

                # Right pivot: Kalah (col O) and Menang (col Q)
                if right_pos:
                    found_otp = {}
                    # In right pivot, "Menang" should map to ON TIME and "Kalah" to OVER TIME.
                    right_region_vals = process_tiket_pivot(right_pos[0], right_pos[1], "Tingkat Investigasi Kalah", (2, 1))
                    for region in REGIONS:
                        for pic in sorted(PIC_BY_REGION.get(region, set())):
                            output_region = GROUP_MAP.get(region, region)
                            output_pic = PIC_SUFFIX_BY_REGION.get((region, pic), pic)
                            if (region, pic) in {(r, p) for ((r, p), _) in found_otp.keys()}:
                                on_time = found_otp.get(((region, pic), "ON TIME"), 0)
                                over_time = found_otp.get(((region, pic), "OVER TIME"), 0)
                            else:
                                on_time, over_time = right_region_vals.get(region, (0, 0))
                            total_os = 0
                            if should_combine_pic(region, pic):
                                key = (output_region, "Tingkat Investigasi Kalah", output_pic, kategori)
                                current = combined_rows.get(key, {"ON TIME": 0, "OVER TIME": 0, "TOTAL OS": 0})
                                current["ON TIME"] += on_time
                                current["OVER TIME"] += over_time
                                combined_rows[key] = current
                            else:
                                rows.append({
                                    "Week": week_label,
                                    "Region": output_region,
                                    "KPI INDICES": "Tingkat Investigasi Kalah",
                                    "PIC": output_pic,
                                    "Kategori": kategori,
                                    "ON TIME": on_time,
                                    "OVER TIME": over_time,
                                    "TOTAL OS": total_os,
                                })

            continue

        header_positions = find_header_rows(df)
        if not header_positions:
            continue

        # For sheets with multiple pivots, use only the left pivot (smallest col),
        # except STUCK which should process all pivots separately.
        if not is_stuck_sheet:
            header_positions = sorted(header_positions, key=lambda x: (x[1], x[0]))
            header_positions = [header_positions[0]]
        else:
            # Keep only first 3 pivots (left side), i.e. before column BC.
            header_positions = [hp for hp in header_positions if hp[1] < STUCK_MAX_COL_EXCLUSIVE_IDX]

        # Special handling for OTP: use left pivot (Region header) for ON/OFF TIME
        # and right pivot (Row Labels header) for TOTAL OS.
        otp_title = isinstance(kpi_title, str) and "OTP (OP - PU)" in kpi_title.upper()
        left_pos = None
        right_pos = None
        if not is_stuck_sheet and otp_title:
            left_pos = find_region_header(df)
            all_positions = find_header_rows(df)
            if all_positions:
                all_positions = sorted(all_positions, key=lambda x: (x[1], x[0]))
                right_pos = all_positions[-1]
            if left_pos and right_pos:
                header_positions = [left_pos, right_pos] if left_pos != right_pos else [left_pos]

        otp_found_otp = {}
        otp_found_total = {}

        for header_row_idx, header_col_idx in header_positions:
            found = {}
            found_otp = {}
            header_row = df.iloc[header_row_idx]
            gt_col_idx = find_grand_total_col(header_row, header_col_idx)

            is_otp_sheet = sheet_upper in {"PU API", "PU MANUAL"} or otp_title

            on_time_col = over_time_col = total_col = None
            if is_otp_sheet:
                # Force fixed offsets for the left OTP pivot to avoid picking right-pivot columns.
                force_offsets = otp_title and left_pos is not None and (header_row_idx, header_col_idx) == left_pos
                on_time_col, over_time_col, total_col = find_otp_qty_cols(
                    df, header_row_idx, header_col_idx, force_offsets=force_offsets
                )
                # If OTP columns not found, fall back to grand total
                if on_time_col is None and over_time_col is None:
                    is_otp_sheet = False

            if gt_col_idx is None and not is_otp_sheet:
                continue

            current_region = None

            for r in range(header_row_idx + 1, len(df)):
                label = df.iloc[r, header_col_idx]
                if pd.isna(label):
                    continue
                label_str = str(label).strip()
                if label_str == "":
                    continue
                if label_str.lower() == "grand total":
                    break

                if is_region(label_str):
                    current_region = label_str.upper()
                    continue

                if current_region is None:
                    continue

                if is_otp_sheet:
                    key = (current_region, label_str.upper())
                    is_left = (left_pos is not None and (header_row_idx, header_col_idx) == left_pos)
                    is_right = (right_pos is not None and (header_row_idx, header_col_idx) == right_pos)
                    # ON TIME (left pivot only)
                    if on_time_col is not None and is_left:
                        val = df.iloc[r, on_time_col]
                        if not pd.isna(val):
                            try:
                                val = float(val)
                            except Exception:
                                pass
                            found_otp[(key, "ON TIME")] = found_otp.get((key, "ON TIME"), 0) + val
                    # OVER TIME (left pivot only)
                    if over_time_col is not None and is_left:
                        val = df.iloc[r, over_time_col]
                        if not pd.isna(val):
                            try:
                                val = float(val)
                            except Exception:
                                pass
                            found_otp[(key, "OVER TIME")] = found_otp.get((key, "OVER TIME"), 0) + val
                    # TOTAL OS from right pivot (grand total only)
                    if is_right and gt_col_idx is not None:
                        val = df.iloc[r, gt_col_idx]
                        if not pd.isna(val):
                            try:
                                val = float(val)
                            except Exception:
                                pass
                            found[key] = found.get(key, 0) + val
                else:
                    val = df.iloc[r, gt_col_idx]
                    if pd.isna(val):
                        continue
                    try:
                        val = float(val)
                    except Exception:
                        pass

                    key = (current_region, label_str.upper())
                    found[key] = found.get(key, 0) + val

            pivot_kpi_title = kpi_title
            if is_stuck_sheet:
                pivot_kpi_title = normalize_kpi_title(
                    infer_stuck_kpi(df, header_row_idx, header_col_idx, kpi_title)
                )

            if is_otp_sheet and otp_title:
                # accumulate across left/right pivots, append rows after all pivots processed
                otp_found_otp.update(found_otp)
                for key, val in found.items():
                    otp_found_total[key] = otp_found_total.get(key, 0) + val
            else:
                # Use fixed PIC list per region so zero values still appear
                for region in REGIONS:
                    for pic in sorted(PIC_BY_REGION.get(region, set())):
                        output_region = GROUP_MAP.get(region, region)
                        output_pic = PIC_SUFFIX_BY_REGION.get((region, pic), pic)

                        total_os = found.get((region, pic), 0)
                        if should_combine_pic(region, pic):
                            key = (output_region, pivot_kpi_title, output_pic, kategori)
                            current = combined_rows.get(key, {"ON TIME": "", "OVER TIME": "", "TOTAL OS": 0})
                            current["TOTAL OS"] += total_os
                            combined_rows[key] = current
                        else:
                            rows.append({
                                "Week": week_label,
                                "Region": output_region,
                                "KPI INDICES": pivot_kpi_title,
                                "PIC": output_pic,
                                "Kategori": kategori,
                                "ON TIME": "",
                                "OVER TIME": "",
                                "TOTAL OS": total_os,
                            })

        if otp_title and otp_found_otp:
            for region in REGIONS:
                for pic in sorted(PIC_BY_REGION.get(region, set())):
                    output_region = GROUP_MAP.get(region, region)
                    output_pic = PIC_SUFFIX_BY_REGION.get((region, pic), pic)
                    on_time = otp_found_otp.get(((region, pic), "ON TIME"), 0)
                    over_time = otp_found_otp.get(((region, pic), "OVER TIME"), 0)
                    total_os = otp_found_total.get((region, pic), 0)
                    if should_combine_pic(region, pic):
                        key = (output_region, kpi_title, output_pic, kategori)
                        current = combined_rows.get(key, {"ON TIME": 0, "OVER TIME": 0, "TOTAL OS": 0})
                        current["ON TIME"] += on_time
                        current["OVER TIME"] += over_time
                        current["TOTAL OS"] += total_os
                        combined_rows[key] = current
                    else:
                        rows.append({
                            "Week": week_label,
                            "Region": output_region,
                            "KPI INDICES": kpi_title,
                            "PIC": output_pic,
                            "Kategori": kategori,
                            "ON TIME": on_time,
                            "OVER TIME": over_time,
                            "TOTAL OS": total_os,
                        })

    for (region, kpi, pic, kategori), vals in combined_rows.items():
        rows.append({
            "Week": week_label,
            "Region": region,
            "KPI INDICES": kpi,
            "PIC": pic,
            "Kategori": kategori,
            "ON TIME": vals.get("ON TIME", ""),
            "OVER TIME": vals.get("OVER TIME", ""),
            "TOTAL OS": vals.get("TOTAL OS", ""),
        })

    out_df = pd.DataFrame(rows)
    columns = ["Week", "Region", "KPI INDICES", "PIC", "Kategori", "ON TIME", "OVER TIME", "TOTAL OS"]
    for c in columns:
        if c not in out_df.columns:
            out_df[c] = None

    out_df = out_df[columns]
    out_df.to_excel(output_path, index=False)


if __name__ == "__main__":
    flatten_pivot(
        input_path=r"d:\SAPX\Data Analisis\KPI_Data\Data OS Week 4 Jan 2026.xlsx",
        output_path=r"d:\SAPX\Data Analisis\KPI_Data\Exports\Combined Data Week 4 Jan 2026.xlsx",
        week_label="Week 4 Jan",
    )
