import re
from pathlib import Path
from statistics import median

import pandas as pd
import win32com.client


# ============================================================
# SETTINGS
# ============================================================

YEAR = 2026
TOP_N = 20

RPA_FILE = "RPA.csv"
SELLOUT_FILE = "sellout.csv"
PRICING_FILE = "pricing.csv"

OUTPUT_DIR = Path(__file__).resolve().parent
OUTPUT_CSV = OUTPUT_DIR / "weekly_insight_output.csv"
OUTPUT_TXT = OUTPUT_DIR / "weekly_insight_top20.txt"
# ============================================================
# HELPERS
# ============================================================

def clean_text(x):
    if x is None:
        return ""
    return " ".join(str(x).strip().split())


def normalize_sku(x):
    return clean_text(x).upper()


def parse_number(x):
    if x is None:
        return None

    if isinstance(x, (int, float)):
        return float(x)

    text = clean_text(x)

    if text == "":
        return None

    text = (
        text.replace(",", "")
        .replace("$", "")
        .replace("%", "")
    )

    try:
        return float(text)
    except Exception:
        return None


def normalize_account(account):
    original = clean_text(account)
    x = original.lower()

    online_subset = "(online sales)" in x

    x = re.sub(
        r"\s*\(online sales\)\s*",
        "",
        x,
        flags=re.IGNORECASE
    ).strip()

    simple = re.sub(r"[^a-z0-9]", "", x)

    mapping = {
        "mht": "BBY",
        "bby": "BBY",
        "bestbuy": "BBY",
        "target": "Target",
        "amazon": "Amazon",
        "costco": "Costco",
        "samsclub": "Sams Club",
        "walmart": "Walmart",
        "bjs": "BJs",
        "total": "TOTAL",
    }

    normalized = mapping.get(simple)

    if normalized is None:
        normalized = clean_text(
            re.sub(
                r"\s*\(online sales\)\s*",
                "",
                original,
                flags=re.IGNORECASE
            )
        )

    return normalized, online_subset


def classify_level(
    account,
    material_group,
    material,
    online_subset
):
    if online_subset:
        return "Online subset"

    a = clean_text(account).lower()
    g = clean_text(material_group).lower()
    m = clean_text(material).lower()

    if (
        a == "total"
        and g == "total"
        and m == "total"
    ):
        return "Company Total"

    if (
        a == "total"
        and g != "total"
        and m == "total"
    ):
        return "Category Total"

    if (
        a == "total"
        and m != "total"
    ):
        return "SKU Total"

    if (
        a != "total"
        and m == "total"
    ):
        return "Account Total"

    if (
        a != "total"
        and m != "total"
    ):
        return "Account x SKU"

    return "Other"


def find_col(headers, name):
    target = name.lower()

    for i, header in enumerate(headers):
        if clean_text(header).lower() == target:
            return i

    return None


# ============================================================
# CONNECT TO OPEN EXCEL FILES
# ============================================================

excel = win32com.client.GetActiveObject(
    "Excel.Application"
)


def find_workbook(filename):
    for wb in excel.Workbooks:
        if wb.Name.lower() == filename.lower():
            return wb

    raise RuntimeError(
        f"{filename} is not open in Excel."
    )


def read_full_sheet(filename):
    wb = find_workbook(filename)
    ws = wb.Worksheets.Item(1)

    used = ws.UsedRange

    last_row = (
        used.Row
        + used.Rows.Count
        - 1
    )

    last_col = (
        used.Column
        + used.Columns.Count
        - 1
    )

    values = ws.Range(
        ws.Cells(1, 1),
        ws.Cells(last_row, last_col)
    ).Value2

    return [
        list(row)
        for row in values
    ]


print()
print("=" * 72)
print("WEEKLY SELLOUT INSIGHT ENGINE")
print("=" * 72)

print("Reading RPA...")
rpa_raw = read_full_sheet(
    RPA_FILE
)

print("Reading historical sellout...")
sellout_raw = read_full_sheet(
    SELLOUT_FILE
)

print("Reading master pricing...")
pricing_raw = read_full_sheet(
    PRICING_FILE
)


# ============================================================
# 1. RPA
# ============================================================

rpa_headers = [
    clean_text(x)
    for x in rpa_raw[0]
]


latest_week = -1
sellout_col = None


for i, header in enumerate(
    rpa_headers
):
    match = re.fullmatch(
        r"Sellout\s*WK\[(\d+)\]",
        header,
        flags=re.IGNORECASE
    )

    if match:
        week = int(
            match.group(1)
        )

        if week > latest_week:
            latest_week = week
            sellout_col = i


if sellout_col is None:
    raise RuntimeError(
        "Could not find latest Sellout WK[x] in RPA."
    )


forecast_col = None


for i, header in enumerate(
    rpa_headers
):
    h = header.lower()

    if (
        "-1w" in h
        and "sellout" in h
        and "fcst" in h
        and f"wk[{latest_week}]" in h
    ):
        forecast_col = i
        break


if forecast_col is None:
    raise RuntimeError(
        f"Could not find -1W forecast for W{latest_week}."
    )


print(
    f"Analysis week: W{latest_week}"
)

# Preliminary report:
# historical actual only through prior week.
actual_cutoff = (
    latest_week - 1
)


rpa_records = []


for row in rpa_raw[1:]:
    if len(row) < 3:
        continue

    original_account = clean_text(
        row[0]
    )

    material_group = clean_text(
        row[1]
    )

    material = normalize_sku(
        row[2]
    )

    if original_account == "":
        continue

    actual = parse_number(
        row[sellout_col]
    )

    forecast = parse_number(
        row[forecast_col]
    )

    if actual is None:
        actual = 0

    if forecast is None:
        forecast = 0

    (
        account,
        online_subset
    ) = normalize_account(
        original_account
    )

    level = classify_level(
        original_account,
        material_group,
        material,
        online_subset
    )

    rpa_records.append({
        "Account": account,
        "OriginalAccount": original_account,
        "OnlineSubset": online_subset,
        "Level": level,
        "MaterialGroup": material_group,
        "Material": material,
        "Actual": actual,
        "Forecast": forecast,
    })


rpa_df = pd.DataFrame(
    rpa_records
)


# Online sales rows are subsets,
# not extra incremental sales.
rpa_df = rpa_df[
    rpa_df["OnlineSubset"] == False
].copy()


# MHT and BBY now both normalize to BBY.
rpa_df = (
    rpa_df
    .groupby(
        [
            "Account",
            "Level",
            "MaterialGroup",
            "Material"
        ],
        as_index=False
    )
    .agg({
        "Actual": "sum",
        "Forecast": "sum"
    })
)


rpa_df["Gap"] = (
    rpa_df["Forecast"]
    - rpa_df["Actual"]
)


rpa_df["Achievement"] = (
    rpa_df["Actual"]
    / rpa_df[
        "Forecast"
    ].replace(
        0,
        pd.NA
    )
)


top20 = (
    rpa_df
    .sort_values(
        "Gap",
        ascending=False
    )
    .head(TOP_N)
    .copy()
)


# ============================================================
# 2. HISTORICAL SELLOUT
# ============================================================

# sellout file structure:
#
# Account
# Item
# Category
# 202601 = W1
# 202602 = W2
# ...
#
# Category contains:
# 2026_SELL OUT (Org.)
# 2026_Ch. Inv. (Org.)
# 2026_WOS (Org.)
#
# We only use SELL OUT here.

SELLOUT_HEADER_ROW = 1
sellout_header_index = SELLOUT_HEADER_ROW - 1

sellout_headers = [
    clean_text(x)
    for x in sellout_raw[
        sellout_header_index
    ]
]


sellout_account_col = find_col(
    sellout_headers,
    "Account"
)

sellout_item_col = find_col(
    sellout_headers,
    "Item"
)

sellout_category_col = find_col(
    sellout_headers,
    "Category"
)


if (
    sellout_account_col is None
    or sellout_item_col is None
    or sellout_category_col is None
):
    raise RuntimeError(
        "Could not find Account / Item / Category "
        "in sellout file."
    )


sellout_week_cols = {}


for i, header in enumerate(
    sellout_headers
):
    match = re.fullmatch(
        rf"{YEAR}(\d{{2}})",
        header
    )

    if not match:
        continue

    week = int(
        match.group(1)
    )

    # W32+ in this file may be forecast.
    # Do not use them as historical actual.
    if week <= actual_cutoff:
        sellout_week_cols[
            week
        ] = i


print(
    f"Historical actual sellout: "
    f"W1-W{actual_cutoff}"
)


sellout_lookup = {}


for row in sellout_raw[
    sellout_header_index + 1:
]:
    if len(row) <= max(
        sellout_account_col,
        sellout_item_col,
        sellout_category_col
    ):
        continue

    original_account = clean_text(
        row[
            sellout_account_col
        ]
    )

    sku = normalize_sku(
        row[
            sellout_item_col
        ]
    )

    category = clean_text(
        row[
            sellout_category_col
        ]
    )

    if (
        category.lower()
        != f"{YEAR}_sell out (org.)".lower()
    ):
        continue

    (
        account,
        online_subset
    ) = normalize_account(
        original_account
    )

    if online_subset:
        continue

    if sku == "":
        continue

    key = (
        account,
        sku
    )

    if key not in sellout_lookup:
        sellout_lookup[
            key
        ] = {}

    for (
        week,
        col
    ) in sellout_week_cols.items():

        if col >= len(row):
            continue

        units = parse_number(
            row[col]
        )

        if units is None:
            continue

        # This also combines BBY + MHT
        # after normalization.
        sellout_lookup[
            key
        ][week] = (
            sellout_lookup[
                key
            ].get(
                week,
                0
            )
            + units
        )


print(
    f"Historical account/SKU series loaded: "
    f"{len(sellout_lookup)}"
)
print("SELL OUT DEBUG:")
print("Series count:", len(sellout_lookup))
print("Target B400F:", ("Target", "HW-B400F/ZA") in sellout_lookup)
print("TOTAL B400F:", ("TOTAL", "HW-B400F/ZA") in sellout_lookup)
print("Target B53WF:", ("Target", "HW-B53WF/ZA") in sellout_lookup)
print("First 10 keys:", list(sellout_lookup.keys())[:10])
# ============================================================
# 3. NATIONAL MASTER PRICING
# ============================================================

# Pricing file has NO account dimension.
#
# We match:
# SKU + Year + Week
#
# Confirmed structure:
# Row 19 = header
#
# Important columns in that header include:
# SKU
# Category
# Year
# W1
# W2
# ...

PRICING_HEADER_ROW = 19


pricing_headers = [
    clean_text(x)
    for x in pricing_raw[
        PRICING_HEADER_ROW - 1
    ]
]


pricing_sku_col = find_col(
    pricing_headers,
    "SKU"
)

pricing_category_col = find_col(
    pricing_headers,
    "Category"
)

pricing_year_col = find_col(
    pricing_headers,
    "Year"
)


if pricing_sku_col is None:
    raise RuntimeError(
        "MASTER pricing: SKU column not found."
    )

if pricing_category_col is None:
    raise RuntimeError(
        "MASTER pricing: Category column not found."
    )

if pricing_year_col is None:
    raise RuntimeError(
        "MASTER pricing: Year column not found."
    )


pricing_week_cols = {}


for i, header in enumerate(
    pricing_headers
):
    match = re.fullmatch(
        r"W(\d+)",
        header,
        flags=re.IGNORECASE
    )

    if match:
        week = int(
            match.group(1)
        )

        pricing_week_cols[
            week
        ] = i


if len(pricing_week_cols) == 0:
    raise RuntimeError(
        "MASTER pricing: no W1/W2/... columns found."
    )


# key:
# SKU
#
# value:
# {
#   week: national promo price
# }

price_lookup = {}


for row in pricing_raw[
    PRICING_HEADER_ROW:
]:

    if len(row) <= max(
        pricing_sku_col,
        pricing_category_col,
        pricing_year_col
    ):
        continue

    sku = normalize_sku(
        row[
            pricing_sku_col
        ]
    )

    category = clean_text(
        row[
            pricing_category_col
        ]
    )

    year = parse_number(
        row[
            pricing_year_col
        ]
    )

    if sku == "":
        continue

    # Only use the final national promo price row.
    if category.lower() != "promo price":
        continue

    if year is None:
        continue

    if int(year) != YEAR:
        continue

    weekly_prices = {}

    for (
        week,
        col
    ) in pricing_week_cols.items():

        if col >= len(row):
            continue

        price = parse_number(
            row[col]
        )

        if price is None:
            continue

        weekly_prices[
            week
        ] = price

    price_lookup[
        sku
    ] = weekly_prices


print(
    f"MASTER pricing loaded: "
    f"{len(price_lookup)} "
    f"{YEAR} SKU price histories"
)


current_price_count = sum(
    latest_week in weekly_prices
    for weekly_prices in price_lookup.values()
)


print(
    f"W{latest_week} prices available for: "
    f"{current_price_count} SKUs"
)


# ============================================================
# 4. PROMO / EVENT CALENDAR
# ============================================================

# Key promo period names live above the pricing header
# and line up with W1/W2/... columns.
#
# Example:
# KDP
# Prime Day
# Labor Day
#
# We KEEP these weeks.
# They are context, not bad outliers.

event_by_week = {}


for row_index in range(
    PRICING_HEADER_ROW - 2,
    -1,
    -1
):

    row = pricing_raw[
        row_index
    ]

    candidate_events = {}

    for (
        week,
        col
    ) in pricing_week_cols.items():

        if col >= len(row):
            continue

        value = clean_text(
            row[col]
        )

        if value == "":
            continue

        # Ignore cells that are basically dates/numbers.
        if re.fullmatch(
            r"[\d/\-~. ]+",
            value
        ):
            continue

        lower = value.lower()

        # Ignore labels that are clearly not promo names.
        if lower in {
            "sound device",
            "year",
            "jan",
            "feb",
            "mar",
            "apr",
            "may",
            "jun",
            "jul",
            "aug",
            "sep",
            "oct",
            "nov",
            "dec",
        }:
            continue

        candidate_events[
            week
        ] = value

    # A useful promo-calendar row should
    # contain at least two event labels.
    if len(candidate_events) >= 2:
        event_by_week = (
            candidate_events
        )
        break


print(
    f"Promo/event labels detected: "
    f"{len(event_by_week)} weeks"
)


# ============================================================
# 5. TOP 20 ANALYSIS
# ============================================================

output_records = []


for _, row in top20.iterrows():

    account = row[
        "Account"
    ]

    level = row[
        "Level"
    ]

    material_group = row[
        "MaterialGroup"
    ]

    sku = normalize_sku(
        row[
            "Material"
        ]
    )

    actual = float(
        row[
            "Actual"
        ]
    )

    forecast = float(
        row[
            "Forecast"
        ]
    )

    gap = float(
        row[
            "Gap"
        ]
    )

    if forecast != 0:
        achievement = (
            actual
            / forecast
        )
    else:
        achievement = None


    result = {
        "Week": f"W{latest_week}",
        "Account": account,
        "Level": level,
        "MaterialGroup": material_group,
        "Material": sku,

        "Actual": actual,
        "Forecast": forecast,
        "Gap": gap,
        "Achievement": achievement,

        "CurrentPrice": None,
        "CurrentEvent": event_by_week.get(
            latest_week,
            ""
        ),

        "HistoricalWeeksFound": 0,
        "SamePriceWeeks": 0,

        "SamePriceMedian": None,
        "SamePriceMin": None,
        "SamePriceMax": None,

        "VsSamePriceMedian": None,

        "NormalSamePriceMedian": None,
        "EventSamePriceMedian": None,

        "PerformanceFlag": "",
        "Insight": ""
    }


    # --------------------------------------------------------
    # TOTAL rows
    # --------------------------------------------------------

    if sku.lower() == "total":

        result[
            "PerformanceFlag"
        ] = "GAP PRIORITY"

        if achievement is not None:

            result[
                "Insight"
            ] = (
                f"{account} {material_group} total: "
                f"W{latest_week} sellout "
                f"{actual:,.0f} "
                f"vs -1W forecast "
                f"{forecast:,.0f}, "
                f"gap {gap:,.0f}, "
                f"achievement "
                f"{achievement:.1%}."
            )

        else:

            result[
                "Insight"
            ] = (
                f"{account} {material_group} total: "
                f"W{latest_week} "
                f"gap {gap:,.0f}."
            )

        output_records.append(
            result
        )

        continue


    # --------------------------------------------------------
    # CURRENT NATIONAL MASTER PRICE
    # --------------------------------------------------------

    sku_prices = price_lookup.get(
        sku,
        {}
    )

    current_price = sku_prices.get(
        latest_week
    )

    result[
        "CurrentPrice"
    ] = current_price


    # --------------------------------------------------------
    # HISTORICAL SELLOUT SERIES
    # --------------------------------------------------------

    # SKU Total should compare to TOTAL sellout.
    #
    # Account x SKU should compare to that account.
if level == "SKU Total":

    # Build national SKU history by summing
    # the SKU across all retailer accounts.
    historical_units = {}

    for (hist_account, hist_sku), week_data in sellout_lookup.items():

        if hist_sku != sku:
            continue

        # Avoid any true TOTAL row if one happens to exist,
        # so we do not double count.
        if hist_account == "TOTAL":
            continue

        for week, units in week_data.items():

            historical_units[week] = (
                historical_units.get(week, 0)
                + units
            )

else:

    historical_units = sellout_lookup.get(
        (
            account,
            sku
        ),
        {}
    )

    result[
        "HistoricalWeeksFound"
    ] = len(
        historical_units
    )


    # --------------------------------------------------------
    # NO CURRENT MASTER PRICE
    # --------------------------------------------------------

    if current_price is None:

        result[
            "PerformanceFlag"
        ] = "NO MASTER PRICE"

        result[
            "Insight"
        ] = (
            f"{account} {sku}: "
            f"W{latest_week} sellout "
            f"{actual:,.0f}, "
            f"forecast {forecast:,.0f}, "
            f"gap {gap:,.0f}. "
            f"No W{latest_week} national "
            f"master Promo Price found."
        )

        output_records.append(
            result
        )

        continue


    # --------------------------------------------------------
    # NO HISTORICAL SELLOUT MATCH
    # --------------------------------------------------------

    if len(
        historical_units
    ) == 0:

        result[
            "PerformanceFlag"
        ] = "NO SELLOUT HISTORY"

        result[
            "Insight"
        ] = (
            f"{account} {sku}: "
            f"W{latest_week} sellout "
            f"{actual:,.0f} "
            f"at national master price "
            f"${current_price:,.2f}, "
            f"but no historical "
            f"account/SKU sellout series "
            f"was matched."
        )

        output_records.append(
            result
        )

        continue


    # --------------------------------------------------------
    # JOIN HISTORY BY WEEK
    # --------------------------------------------------------

    # sellout:
    # 202601 -> W1
    #
    # pricing:
    # W1 -> W1
    #
    # So the shared join key is the integer week number.

    comparisons = []


    for week in range(
        1,
        actual_cutoff + 1
    ):

        historical_price = (
            sku_prices.get(
                week
            )
        )

        historical_sellout = (
            historical_units.get(
                week
            )
        )

        if historical_price is None:
            continue

        if historical_sellout is None:
            continue

        # Compare only weeks with the same
        # NATIONAL master price.
        if abs(
            historical_price
            - current_price
        ) > 0.01:
            continue

        comparisons.append({
            "Week": week,
            "Units": historical_sellout,
            "Price": historical_price,
            "Event": event_by_week.get(
                week,
                ""
            )
        })


    result[
        "SamePriceWeeks"
    ] = len(
        comparisons
    )


    # --------------------------------------------------------
    # NO PRIOR SAME-PRICE WEEK
    # --------------------------------------------------------

    if len(
        comparisons
    ) == 0:

        result[
            "PerformanceFlag"
        ] = "NO SAME-PRICE HISTORY"

        result[
            "Insight"
        ] = (
            f"{account} {sku}: "
            f"W{latest_week} sellout "
            f"{actual:,.0f}, "
            f"forecast {forecast:,.0f}, "
            f"gap {gap:,.0f}, "
            f"national master price "
            f"${current_price:,.2f}. "
            f"Historical sellout was found, "
            f"but no prior actual week used "
            f"the same national master price."
        )

        output_records.append(
            result
        )

        continue
    # --------------------------------------------------------
    # SAME-PRICE DISTRIBUTION
    # --------------------------------------------------------

    same_price_units = [
        x["Units"]
        for x in comparisons
    ]

    same_price_median = median(
        same_price_units
    )

    same_price_min = min(
        same_price_units
    )

    same_price_max = max(
        same_price_units
    )


    result[
        "SamePriceMedian"
    ] = same_price_median

    result[
        "SamePriceMin"
    ] = same_price_min

    result[
        "SamePriceMax"
    ] = same_price_max


    if same_price_median != 0:

        deviation = (
            actual
            / same_price_median
            - 1
        )

    else:

        deviation = None


    result[
        "VsSamePriceMedian"
    ] = deviation


    # --------------------------------------------------------
    # KEEP EVENT WEEKS AS CONTEXT
    # --------------------------------------------------------

    normal_units = []

    event_units = []


    for x in comparisons:

        if clean_text(
            x["Event"]
        ) == "":

            normal_units.append(
                x["Units"]
            )

        else:

            event_units.append(
                x["Units"]
            )


    if normal_units:

        result[
            "NormalSamePriceMedian"
        ] = median(
            normal_units
        )


    if event_units:

        result[
            "EventSamePriceMedian"
        ] = median(
            event_units
        )


    # --------------------------------------------------------
    # PERFORMANCE FLAG
    # --------------------------------------------------------

    if len(
        comparisons
    ) < 3:

        flag = "LOW CONFIDENCE"


    elif deviation is None:

        flag = "CHECK"


    elif deviation <= -0.20:

        flag = "WEAK"


    elif deviation >= 0.20:

        flag = "STRONG"


    else:

        flag = "IN LINE"


    result[
        "PerformanceFlag"
    ] = flag


    # --------------------------------------------------------
    # INSIGHT TEXT
    # --------------------------------------------------------

    if deviation is None:

        deviation_text = (
            "could not be meaningfully compared with"
        )

    elif deviation < 0:

        deviation_text = (
            f"is {abs(deviation):.0%} below"
        )

    else:

        deviation_text = (
            f"is {abs(deviation):.0%} above"
        )


    current_event = clean_text(
        event_by_week.get(
            latest_week,
            ""
        )
    )


    event_sentence = ""


    if current_event != "":

        event_sentence = (
            f" Current national promo period: "
            f"{current_event}."
        )


    context = ""


    if result[
        "NormalSamePriceMedian"
    ] is not None:

        context += (
            f" Normal/non-key-event "
            f"same-price median: "
            f"{result['NormalSamePriceMedian']:,.0f}."
        )


    if result[
        "EventSamePriceMedian"
    ] is not None:

        context += (
            f" Key-event same-price median: "
            f"{result['EventSamePriceMedian']:,.0f}."
        )


    result[
        "Insight"
    ] = (
        f"{account} {sku}: "
        f"W{latest_week} sellout "
        f"{actual:,.0f} "
        f"vs -1W forecast "
        f"{forecast:,.0f} "
        f"(gap {gap:,.0f}, "
        f"{achievement:.1%} achievement). "
        f"At national master price "
        f"${current_price:,.2f}, "
        f"current sellout "
        f"{deviation_text} "
        f"the historical same-price median "
        f"of {same_price_median:,.0f} "
        f"across {len(comparisons)} actual weeks "
        f"(range "
        f"{same_price_min:,.0f}-"
        f"{same_price_max:,.0f})."
        f"{context}"
        f"{event_sentence}"
    )


    output_records.append(
        result
    )


# ============================================================
# OUTPUT
# ============================================================

output = pd.DataFrame(
    output_records
)


output.insert(
    0,
    "Rank",
    range(
        1,
        len(output) + 1
    )
)


output.to_csv(
    OUTPUT_CSV,
    index=False,
    encoding="utf-8-sig"
)


with open(
    OUTPUT_TXT,
    "w",
    encoding="utf-8"
) as file:

    file.write(
        f"W{latest_week} "
        f"WEEKLY INSIGHT TOP {TOP_N}\n"
    )

    file.write(
        "=" * 85
        + "\n\n"
    )

    file.write(
        "Price source: national MASTER Promo Price. "
        "Retailer-specific extra discounts "
        "are NOT yet included.\n"
    )

    file.write(
        "Key promo/event weeks are preserved "
        "as historical context "
        "and are not automatically deleted "
        "as outliers.\n\n"
    )


    for _, row in output.iterrows():

        file.write(
            f"{int(row['Rank'])}. "
            f"{row['PerformanceFlag']}\n"
        )

        file.write(
            row["Insight"]
            + "\n\n"
        )


print()
print("=" * 72)
print("DONE")
print("=" * 72)

print()
print("Created:")
print(OUTPUT_CSV)
print(OUTPUT_TXT)

print()
input("Press Enter to close...")