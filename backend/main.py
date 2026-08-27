from __future__ import annotations

import re
import uuid
from copy import copy
from pathlib import Path
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# BASIC CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path("/tmp/outputs")

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


app = FastAPI(
    title="Survey Spreadsheet Validator",
    version="1.0.0",
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://upcensus-detector-zph2.vercel.app",
        "http://localhost:5173",
        "http://127.0.0.1:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# COLUMN ALIASES
# ============================================================

LAT_ALIASES = {
    "location [latitude]",
    "location latitude",
    "latitude",
    "lat",
}

LON_ALIASES = {
    "location [longitude]",
    "location longitude",
    "longitude",
    "long",
    "lng",
}

MOBILE_ALIASES = {
    "mobile number",
    "mobile no",
    "mobile no.",
    "mobile",
    "contact number",
    "contact no",
    "contact no.",
    "phone number",
    "phone no",
    "phone",
}

OUTLET_ALIASES = {
    "outlet name",
    "shop name",
    "shop",
    "store name",
    "outlet",
}

IMAGE_HINTS = (
    "photo",
    "image",
    "picture",
    "pic",
    "screenshot",
)


# ============================================================
# TEXT / HEADER NORMALIZATION
# ============================================================

def clean_text(value: Any) -> str:
    """
    Convert any cell value into a clean string.

    Handles:
    - None
    - NaN
    - Excel blank cells
    - whitespace
    """

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    return str(value).strip()


def normalized_column_name(value: Any) -> str:
    """
    Aggressively normalizes Excel column headers.

    Example:

        Mobile
        Number

    becomes:

        mobile number

    Also handles:
    - \\n
    - \\r
    - \\t
    - non-breaking spaces
    - BOM
    - zero-width Unicode characters
    - punctuation
    - multiple spaces
    """

    if value is None:
        return ""

    text = str(value)

    # Remove invisible Unicode characters
    text = (
        text.replace("\ufeff", "")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\xa0", " ")
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("\t", " ")
    )

    text = text.lower().strip()

    # Convert punctuation / symbols into spaces.
    #
    # This means:
    # Mobile-Number
    # Mobile/Number
    # Mobile.Number
    #
    # all become:
    # mobile number
    text = re.sub(
        r"[^a-z0-9\u0900-\u097f]+",
        " ",
        text,
    )

    # Remove duplicate spaces
    text = re.sub(
        r"\s+",
        " ",
        text,
    ).strip()

    return text


# ============================================================
# COLUMN FINDER
# ============================================================

def find_column(
    df: pd.DataFrame,
    aliases: set[str],
) -> str | None:
    """
    Robustly find a column even when Excel contains:

        Mobile
        Number

    or:

        Mobile Number
        Mobile-No
        Mobile No.
        Contact Number
        Phone Number
    """

    normalized_aliases = {
        normalized_column_name(alias)
        for alias in aliases
    }

    # --------------------------------------------------------
    # Exact normalized match.
    # --------------------------------------------------------

    for column in df.columns:

        normalized = normalized_column_name(
            column
        )

        if normalized in normalized_aliases:
            return column

    # --------------------------------------------------------
    # Token / semantic fallback.
    # --------------------------------------------------------

    for column in df.columns:

        normalized = normalized_column_name(
            column
        )

        words = set(
            normalized.split()
        )

        if (
            "mobile" in words
            and "number" in words
        ):
            return column

        if (
            "mobile" in words
            and "no" in words
        ):
            return column

        if (
            "contact" in words
            and "number" in words
        ):
            return column

        if (
            "phone" in words
            and "number" in words
        ):
            return column

    # --------------------------------------------------------
    # Substring fallback.
    # --------------------------------------------------------

    for column in df.columns:

        normalized = normalized_column_name(
            column
        )

        if (
            "mobile" in normalized
            and (
                "number" in normalized
                or "no" in normalized
            )
        ):
            return column

        if (
            "contact" in normalized
            and "number" in normalized
        ):
            return column

        if (
            "phone" in normalized
            and "number" in normalized
        ):
            return column

    return None


# ============================================================
# IMAGE COLUMN DETECTION
# ============================================================

def find_image_columns(
    df: pd.DataFrame,
) -> list[str]:
    """
    Detect all image/photo related columns.
    """

    columns: list[str] = []

    for column in df.columns:

        name = normalized_column_name(column)

        if any(
            hint in name
            for hint in IMAGE_HINTS
        ):
            columns.append(column)

    return columns


# ============================================================
# INPUT FILE READER
# ============================================================

def read_input(
    path: Path,
) -> pd.DataFrame:
    """
    Read Excel / CSV / TSV and automatically detect the actual
    survey header row.

    This specifically handles Excel headers such as:

        Mobile
        Number

    which pandas may read as "Mobile\\nNumber".

    It also handles:
    - blank/title rows before the header
    - line breaks
    - tabs
    - non-breaking spaces
    - hidden Unicode characters
    - CSV / TSV files
    """

    suffix = path.suffix.lower()

    # --------------------------------------------------------
    # Read without assuming the first row is the header.
    # --------------------------------------------------------

    if suffix in {".xlsx", ".xls"}:

        raw_df = pd.read_excel(
            path,
            dtype=str,
            header=None,
        ).fillna("")

    elif suffix in {".csv", ".tsv"}:

        separator = "\t" if suffix == ".tsv" else ","

        raw_df = pd.read_csv(
            path,
            sep=separator,
            dtype=str,
            header=None,
            keep_default_na=False,
        ).fillna("")

    else:
        raise ValueError(
            "Unsupported file type."
        )

    if raw_df.empty:
        raise ValueError(
            "The uploaded spreadsheet is empty."
        )

    # --------------------------------------------------------
    # Header search normalization.
    # --------------------------------------------------------

    def header_search_text(value: Any) -> str:

        if value is None:
            return ""

        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass

        text = str(value)

        text = (
            text.replace("\ufeff", " ")
            .replace("\u200b", " ")
            .replace("\u200c", " ")
            .replace("\u200d", " ")
            .replace("\xa0", " ")
            .replace("\r", " ")
            .replace("\n", " ")
            .replace("\t", " ")
        )

        text = text.lower().strip()

        text = re.sub(
            r"[^a-z0-9]+",
            " ",
            text,
        )

        return re.sub(
            r"\s+",
            " ",
            text,
        ).strip()

    # --------------------------------------------------------
    # Score the first 50 rows.
    #
    # We don't simply search for "Mobile Number", because
    # some survey sheets have title/description rows before
    # the real header.
    # --------------------------------------------------------

    best_header_row: int | None = None
    best_score = -1

    rows_to_check = min(
        50,
        len(raw_df),
    )

    for row_index in range(rows_to_check):

        values = [
            header_search_text(value)
            for value in raw_df.iloc[
                row_index
            ].tolist()
        ]

        non_empty = [
            value
            for value in values
            if value
        ]

        if not non_empty:
            continue

        row_text = " | ".join(non_empty)

        score = 0

        # Strong identifiers for this survey.
        if any(
            "response id" == value
            or "response id" in value
            for value in non_empty
        ):
            score += 5

        if any(
            "mobile number" in value
            or (
                "mobile" in value
                and "number" in value
            )
            for value in non_empty
        ):
            score += 10

        if any(
            "outlet name" in value
            or "shop name" in value
            for value in non_empty
        ):
            score += 8

        if any(
            "latitude" in value
            for value in non_empty
        ):
            score += 5

        if any(
            "longitude" in value
            for value in non_empty
        ):
            score += 5

        if any(
            "collector name" in value
            for value in non_empty
        ):
            score += 3

        if any(
            "pin code" in value
            or "pincode" in value
            for value in non_empty
        ):
            score += 3

        # Prefer rows that actually look like a wide survey
        # header rather than a random description row.
        if len(non_empty) >= 8:
            score += 2

        if score > best_score:
            best_score = score
            best_header_row = row_index

    # --------------------------------------------------------
    # Header not found.
    # --------------------------------------------------------

    if best_header_row is None or best_score < 10:

        raise ValueError(
            "Could not detect the survey header row. "
            "The spreadsheet must contain a 'Mobile Number' "
            "column/header."
        )

    # --------------------------------------------------------
    # Build clean, unique column names.
    # --------------------------------------------------------

    raw_headers = raw_df.iloc[
        best_header_row
    ].tolist()

    headers: list[str] = []
    used_headers: dict[str, int] = {}

    for index, value in enumerate(raw_headers):

        header = clean_text(value)

        if not header:
            header = f"Unnamed Column {index + 1}"

        # Preserve the original visible header, but make
        # duplicate headers unique so pandas can reference them.
        if header in used_headers:

            used_headers[header] += 1

            header = (
                f"{header} "
                f"({used_headers[header]})"
            )

        else:
            used_headers[header] = 1

        headers.append(header)

    # --------------------------------------------------------
    # Data begins immediately after detected header row.
    # --------------------------------------------------------

    df = raw_df.iloc[
        best_header_row + 1:
    ].copy()

    df.columns = headers

    # Remove completely blank rows.
    df = df[
        ~df.apply(
            lambda row: all(
                clean_text(value) == ""
                for value in row
            ),
            axis=1,
        )
    ].reset_index(drop=True)

    # --------------------------------------------------------
    # Debug output.
    # --------------------------------------------------------

    print(
        "\n========================================"
    )
    print(
        "SURVEY HEADER DETECTION"
    )
    print(
        "========================================"
    )
    print(
        "Detected header row:",
        best_header_row + 1,
    )
    print(
        "Header detection score:",
        best_score,
    )

    for index, column in enumerate(df.columns):

        normalized = normalized_column_name(
            column
        )

        print(
            index,
            repr(column),
            "=>",
            repr(normalized),
        )

    print(
        "========================================\n"
    )

    return df


# ============================================================
# MOBILE NUMBER
# ============================================================

def mobile_digits(
    value: Any,
) -> str:

    text = clean_text(value)

    # Remove all non-digit characters.
    digits = re.sub(
        r"\D",
        "",
        text,
    )

    # Handle:
    #
    # +91XXXXXXXXXX
    # 91XXXXXXXXXX
    #
    # Convert to:
    #
    # XXXXXXXXXX

    if (
        len(digits) == 12
        and digits.startswith("91")
    ):
        digits = digits[2:]

    return digits


def valid_indian_mobile(
    value: Any,
) -> bool:

    digits = mobile_digits(value)

    return bool(
        re.fullmatch(
            r"[6-9]\d{9}",
            digits,
        )
    )


def suspicious_mobile(
    value: Any,
) -> bool:

    digits = mobile_digits(value)

    if len(digits) != 10:
        return False

    # --------------------------------------------------------
    # All same digit
    #
    # 0000000000
    # 1111111111
    # 2222222222
    # etc.
    # --------------------------------------------------------

    if len(set(digits)) == 1:
        return True

    # --------------------------------------------------------
    # Common fake / test numbers
    # --------------------------------------------------------

    fake_patterns = {
        "0123456789",
        "1234567890",
        "9876543210",
        "0987654321",
    }

    if digits in fake_patterns:
        return True

    return False


# ============================================================
# SHOP NAME NORMALIZATION
# ============================================================

def normalized_shop(
    value: Any,
) -> str:

    text = clean_text(value).lower()

    # Remove punctuation
    text = re.sub(
        r"[^a-z0-9\u0900-\u097f]+",
        " ",
        text,
    )

    # Multiple spaces
    text = re.sub(
        r"\s+",
        " ",
        text,
    ).strip()

    return text


# ============================================================
# IMAGE REFERENCE
# ============================================================

def image_reference(
    value: Any,
) -> str:

    text = clean_text(value)

    if not text:
        return ""

    # Normalize case
    text = text.lower().strip()

    # Remove surrounding spaces
    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text


# ============================================================
# SCREENSHOT DETECTION
# ============================================================

def image_reason_label(
    value: Any,
) -> str:

    text = clean_text(value)

    if not text:
        return ""

    # Remove the interviewer note from the displayed reason.
    # Example:
    # "Photo of the Shop :Note : Interviewer ..."
    # becomes:
    # "Photo of the Shop"
    text = re.split(
        r"\s*:\s*note\b.*$",
        text,
        maxsplit=1,
        flags=re.IGNORECASE | re.DOTALL,
    )[0].strip()

    return text.rstrip(" :")


def looks_like_screenshot(
    value: Any,
) -> bool:

    text = clean_text(value).lower()

    if not text:
        return False

    screenshot_terms = (
        "screenshot",
        "screen shot",
        "screen_shot",
        "screen-shot",
        "snip_",
        "snipping",
        "snip",
        "capture_",
        "capture-",
        "screen-record",
        "screen_record",
        "screen record",
        "screenrecord",
    )

    return any(
        term in text
        for term in screenshot_terms
    )


# ============================================================
# NUMERIC VALUE
# ============================================================

def numeric_value(
    value: Any,
) -> float | None:

    text = clean_text(value)

    if not text:
        return None

    try:

        return float(
            text.replace(",", "")
        )

    except ValueError:

        return None


# ============================================================
# EXCEL FORMATTING
# ============================================================

def format_worksheet(
    worksheet: Any,
) -> None:

    # Freeze first row
    worksheet.freeze_panes = "A2"

    # Filter
    if worksheet.max_row > 1:
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

    # Header formatting
    for cell in worksheet[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF",
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1D4ED8",
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    # Header height
    worksheet.row_dimensions[1].height = 42

    # Data alignment
    for row in worksheet.iter_rows(
        min_row=2,
    ):

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    # Auto width
    for column_cells in worksheet.columns:

        column_letter = get_column_letter(
            column_cells[0].column
        )

        max_length = 0

        # Limit scanning for performance
        for cell in column_cells[:300]:

            value = (
                ""
                if cell.value is None
                else str(cell.value)
            )

            max_length = max(
                max_length,
                len(value),
            )

        width = max(
            12,
            min(
                max_length + 2,
                45,
            ),
        )

        worksheet.column_dimensions[
            column_letter
        ].width = width


# ============================================================
# MAIN VALIDATION
# ============================================================

def create_output(
    df: pd.DataFrame,
    output_path: Path,
) -> dict[str, Any]:

    # --------------------------------------------------------
    # FIND IMPORTANT COLUMNS
    # --------------------------------------------------------

    mobile_col = find_column(
        df,
        MOBILE_ALIASES,
    )

    outlet_col = find_column(
        df,
        OUTLET_ALIASES,
    )

    lat_col = find_column(
        df,
        LAT_ALIASES,
    )

    lon_col = find_column(
        df,
        LON_ALIASES,
    )

    image_cols = find_image_columns(df)

    # --------------------------------------------------------
    # Detection diagnostic
    # --------------------------------------------------------

    print(
        "\n========================================"
    )
    print(
        "IMPORTANT COLUMN DETECTION"
    )
    print(
        "========================================"
    )
    print(
        "Mobile Number:",
        repr(mobile_col),
    )
    print(
        "Outlet Name:",
        repr(outlet_col),
    )
    print(
        "Latitude:",
        repr(lat_col),
    )
    print(
        "Longitude:",
        repr(lon_col),
    )
    print(
        "Image Columns:",
        [repr(column) for column in image_cols],
    )
    print(
        "========================================\n"
    )

    # --------------------------------------------------------
    # Validate required columns
    # --------------------------------------------------------

    if mobile_col is None:

        detected_columns = [
            str(column)
            for column in df.columns
        ]

        raise ValueError(
            "Could not find the 'Mobile Number' column. "
            "Detected columns: "
            + ", ".join(detected_columns[:80])
        )

    if outlet_col is None:

        raise ValueError(
            "Could not find the 'Outlet Name' column. "
            "Please check the spreadsheet header."
        )

    if lat_col is None:

        raise ValueError(
            "Could not find the Latitude column."
        )

    if lon_col is None:

        raise ValueError(
            "Could not find the Longitude column."
        )

    # ========================================================
    # MOBILE DUPLICATES
    # ========================================================

    mobile_series = df[
        mobile_col
    ].map(mobile_digits)

    mobile_counts = (
        mobile_series[
            mobile_series != ""
        ]
        .value_counts()
    )

    duplicate_mobiles = set(
        mobile_counts[
            mobile_counts > 1
        ].index
    )


    # ========================================================
    # IMAGE DUPLICATES
    # ========================================================

    image_occurrences: dict[
        str,
        int,
    ] = {}

    for column in image_cols:

        for value in df[column]:

            reference = image_reference(
                value
            )

            if reference:

                image_occurrences[
                    reference
                ] = (
                    image_occurrences.get(
                        reference,
                        0,
                    )
                    + 1
                )

    duplicate_images = {
        reference
        for reference, count
        in image_occurrences.items()
        if count > 1
    }

    # ========================================================
    # RESULT CONTAINERS
    # ========================================================

    correct_rows: list[
        dict[str, Any]
    ] = []

    discarded_rows: list[
        dict[str, Any]
    ] = []

    # ========================================================
    # COUNTERS
    # ========================================================

    counters = {
        "Mobile Duplicates": 0,
        "Invalid Mobile": 0,
        "Suspicious Mobile": 0,
        "Missing Mobile Number": 0,
            "Duplicate Images": 0,
        "Screenshots": 0,
        "Missing Location": 0,
        "Invalid Coordinates": 0,
    }

    # ========================================================
    # PROCESS EVERY ROW
    # ========================================================

    for _, row in df.iterrows():

        reasons: list[str] = []

        # ----------------------------------------------------
        # Mobile
        # ----------------------------------------------------

        mobile_value = clean_text(
            row[mobile_col]
        )

        mobile = mobile_digits(
            row[mobile_col]
        )

        if not mobile_value:

            reasons.append(
                "Missing Mobile Number"
            )

            counters[
                "Missing Mobile Number"
            ] += 1

        elif not valid_indian_mobile(
            row[mobile_col]
        ):

            reasons.append(
                "Invalid Indian Mobile Number"
            )

            counters[
                "Invalid Mobile"
            ] += 1

        if suspicious_mobile(
            row[mobile_col]
        ):

            reasons.append(
                "Suspicious Mobile Number"
            )

            counters[
                "Suspicious Mobile"
            ] += 1

        if (
            mobile
            and mobile in duplicate_mobiles
        ):

            reasons.append(
                "Duplicate Mobile Number"
            )

            counters[
                "Mobile Duplicates"
            ] += 1


        # ----------------------------------------------------
        # Latitude / Longitude
        # ----------------------------------------------------

        lat = numeric_value(
            row[lat_col]
        )

        lon = numeric_value(
            row[lon_col]
        )

        if (
            lat is None
            or lon is None
        ):

            reasons.append(
                "Latitude/Longitude Missing"
            )

            counters[
                "Missing Location"
            ] += 1

        else:

            invalid_latitude = not (
                -90 <= lat <= 90
            )

            invalid_longitude = not (
                -180 <= lon <= 180
            )

            if (
                invalid_latitude
                or invalid_longitude
            ):

                reasons.append(
                    "Invalid Latitude/Longitude"
                )

                counters[
                    "Invalid Coordinates"
                ] += 1

        # ----------------------------------------------------
        # Images
        # ----------------------------------------------------

        for image_col in image_cols:

            reference = image_reference(
                row[image_col]
            )

            if (
                reference
                and reference in duplicate_images
            ):

                reasons.append(
                    f"Duplicate Image Reference: {image_reason_label(image_col)}"
                )

                counters[
                    "Duplicate Images"
                ] += 1

            if (
                reference
                and looks_like_screenshot(
                    reference
                )
            ):

                reasons.append(
                    f"Screenshot Image: {image_reason_label(image_col)}"
                )

                counters[
                    "Screenshots"
                ] += 1

        # ----------------------------------------------------
        # Preserve original data
        # ----------------------------------------------------

        base = {
            str(column): row[column]
            for column in df.columns
        }

        # Remove duplicate reasons
        unique_reasons = list(
            dict.fromkeys(reasons)
        )

        # ----------------------------------------------------
        # Correct or Discard
        # ----------------------------------------------------

        if unique_reasons:

            base[
                "Discard Reason"
            ] = "; ".join(
                unique_reasons
            )

            discarded_rows.append(
                base
            )

        else:

            correct_rows.append(
                base
            )

    # ========================================================
    # DATAFRAMES
    # ========================================================

    original_columns = [
        str(column)
        for column in df.columns
    ]

    correct_df = pd.DataFrame(
        correct_rows,
        columns=original_columns,
    )

    discard_df = pd.DataFrame(
        discarded_rows,
        columns=[
            *original_columns,
            "Discard Reason",
        ],
    )

    # ========================================================
    # CREATE EXCEL
    # ========================================================

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:

        correct_df.to_excel(
            writer,
            index=False,
            sheet_name="Correct Data",
        )

        discard_df.to_excel(
            writer,
            index=False,
            sheet_name="Discard Data",
        )

        workbook = writer.book

        # ----------------------------------------------------
        # Format Correct Data
        # ----------------------------------------------------

        correct_sheet = workbook[
            "Correct Data"
        ]

        format_worksheet(
            correct_sheet
        )

        # ----------------------------------------------------
        # Format Discard Data
        # ----------------------------------------------------

        discard_sheet = workbook[
            "Discard Data"
        ]

        format_worksheet(
            discard_sheet
        )

        # ----------------------------------------------------
        # Highlight Discard Reason
        # ----------------------------------------------------

        if discard_sheet.max_column > 0:

            discard_reason_column = (
                discard_sheet.max_column
            )

            for row_number in range(
                2,
                discard_sheet.max_row + 1,
            ):

                cell = discard_sheet.cell(
                    row=row_number,
                    column=discard_reason_column,
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="FEE2E2",
                )

                cell.font = Font(
                    color="991B1B",
                    bold=True,
                )

        # ----------------------------------------------------
        # Highlight Correct Data status
        # ----------------------------------------------------

        # Add green visual formatting to Correct Data rows.
        for row_number in range(
            2,
            correct_sheet.max_row + 1,
        ):

            for cell in correct_sheet[
                row_number
            ]:

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="F0FDF4",
                )

    # ========================================================
    # RESPONSE
    # ========================================================

    return {
        "total_rows": len(df),
        "correct_rows": len(
            correct_df
        ),
        "discarded_rows": len(
            discard_df
        ),
        "output_filename": output_path.name,
        "checks": counters,
        "detected_columns": {
            "Mobile Number": str(
                mobile_col
            ),
            "Outlet Name": str(
                outlet_col
            ),
            "Latitude": str(
                lat_col
            ),
            "Longitude": str(
                lon_col
            ),
            "Image Columns": [
                str(column)
                for column in image_cols
            ],
        },
    }


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health() -> dict[str, str]:

    return {
        "status": "ok"
    }


# ============================================================
# VALIDATE FILE
# ============================================================

@app.post("/validate")
async def validate(
    file: UploadFile = File(...),
) -> dict[str, Any]:

    original_name = (
        file.filename
        or "survey_data.xlsx"
    )

    suffix = (
        Path(original_name)
        .suffix
        .lower()
    )

    # --------------------------------------------------------
    # Validate extension
    # --------------------------------------------------------

    allowed_extensions = {
        ".xlsx",
        ".xls",
        ".csv",
        ".tsv",
    }

    if suffix not in allowed_extensions:

        raise HTTPException(
            status_code=400,
            detail=(
                "Only XLSX, XLS, CSV and TSV "
                "files are supported."
            ),
        )

    # --------------------------------------------------------
    # Create temporary paths
    # --------------------------------------------------------

    job_id = uuid.uuid4().hex

    input_path = (
        OUTPUT_DIR
        / f"{job_id}{suffix}"
    )

    output_name = (
        f"Validated_"
        f"{Path(original_name).stem}_"
        f"{job_id[:8]}.xlsx"
    )

    output_path = (
        OUTPUT_DIR
        / output_name
    )

    # --------------------------------------------------------
    # Process
    # --------------------------------------------------------

    try:

        file_content = await file.read()

        input_path.write_bytes(
            file_content
        )

        df = read_input(
            input_path
        )

        if df.empty:

            raise ValueError(
                "The uploaded spreadsheet "
                "contains no records."
            )

        result = create_output(
            df,
            output_path,
        )

        return result

    except HTTPException:
        raise

    except Exception as exc:

        if output_path.exists():
            output_path.unlink()

        raise HTTPException(
            status_code=400,
            detail=str(exc),
        ) from exc

    finally:

        if input_path.exists():
            input_path.unlink()


# ============================================================
# DOWNLOAD OUTPUT
# ============================================================

@app.get(
    "/download/{filename}"
)
def download(
    filename: str,
) -> FileResponse:

    # Security:
    # only use the filename, not any supplied path.

    safe_name = Path(
        filename
    ).name

    path = (
        OUTPUT_DIR
        / safe_name
    )

    if (
        not path.exists()
        or path.suffix.lower() != ".xlsx"
    ):

        raise HTTPException(
            status_code=404,
            detail="Output file not found.",
        )

    return FileResponse(
        path=path,
        filename=safe_name,
        media_type=(
            "application/vnd."
            "openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )