import base64
from dataclasses import dataclass
import os
from pathlib import Path
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

try:
    import pandas as pd
except ModuleNotFoundError:
    pd = None

try:
    import streamlit as st
except ModuleNotFoundError:
    st = None


ROOT = Path(__file__).resolve().parent
WORKBOOK_PATH = ROOT / "Client Income  by Business Sector - Summary V3.xlsx"
ACTIVE_WORKBOOK_PATH = ROOT / "ACTIVE POLICIES AONPASS CARS.xlsx"
LOGO_PATH = ROOT / "logo.png"
STREAMLIT_RUNTIME_ENABLED = os.environ.get("CLIENT_DASHBOARD_STREAMLIT") == "1"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"main": MAIN_NS, "r": REL_NS, "pkg": PKG_REL_NS}


def cache_data(*args, **kwargs):
    def decorator(func):
        return func

    if st is not None and STREAMLIT_RUNTIME_ENABLED:
        return st.cache_data(*args, **kwargs)
    return decorator


@dataclass
class SheetData:
    name: str
    max_row: int
    max_column: int
    rows: dict[int, dict[int, object]]


@dataclass
class SheetAnalysis:
    name: str
    header_row: int | None
    headers: list[str]
    records: list[dict[str, object]]
    detected_columns: dict[str, str | None]
    max_row: int
    max_column: int


def text_value(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def numeric_value(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = text_value(value).replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_header(value: object, index: int) -> str:
    text = text_value(value)
    return text or f"column_{index}"


def unique_headers(values: list[object]) -> list[str]:
    counts: dict[str, int] = {}
    headers: list[str] = []
    for index, value in enumerate(values, start=1):
        base = normalize_header(value, index)
        counts[base] = counts.get(base, 0) + 1
        suffix = f"_{counts[base]}" if counts[base] > 1 else ""
        headers.append(f"{base}{suffix}")
    return headers


def format_number(value: int | float) -> str:
    return f"{value:,.0f}"


def column_index_from_ref(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref)
    if not match:
        return 0

    result = 0
    for char in match.group(1):
        result = (result * 26) + (ord(char) - ord("A") + 1)
    return result


def parse_cell_ref(ref: str) -> tuple[int, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        return 0, 0
    return int(match.group(2)), column_index_from_ref(match.group(1))


def parse_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("main:si", NS):
        parts = [node.text or "" for node in item.findall(".//main:t", NS)]
        strings.append("".join(parts))
    return strings


def decode_cell(cell: ET.Element, shared_strings: list[str]) -> object:
    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        parts = [node.text or "" for node in cell.findall(".//main:t", NS)]
        return "".join(parts)

    value_node = cell.find("main:v", NS)
    if value_node is None:
        return None

    raw = value_node.text or ""
    if cell_type == "s":
        if raw.isdigit():
            index = int(raw)
            if 0 <= index < len(shared_strings):
                return shared_strings[index]
        return raw
    if cell_type == "b":
        return raw == "1"
    return raw


def parse_sheet(zf: zipfile.ZipFile, target: str, sheet_name: str, shared_strings: list[str]) -> SheetData:
    root = ET.fromstring(zf.read(target))
    rows: dict[int, dict[int, object]] = {}
    max_row = 0
    max_column = 0

    for row_node in root.findall(".//main:sheetData/main:row", NS):
        row_index = int(row_node.attrib.get("r", "0"))
        cells: dict[int, object] = {}
        for cell in row_node.findall("main:c", NS):
            ref = cell.attrib.get("r", "")
            _, column_index = parse_cell_ref(ref)
            if column_index == 0:
                continue
            cells[column_index] = decode_cell(cell, shared_strings)
            max_column = max(max_column, column_index)
        if cells:
            rows[row_index] = cells
            max_row = max(max_row, row_index)

    return SheetData(
        name=sheet_name,
        max_row=max_row,
        max_column=max_column,
        rows=rows,
    )


def load_workbook(path: Path) -> list[SheetData]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = parse_shared_strings(zf)

        workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map: dict[str, str] = {}
        for rel in rel_root.findall("pkg:Relationship", NS):
            rel_id = rel.attrib["Id"]
            target = rel.attrib["Target"]
            rel_map[rel_id] = posixpath.normpath(posixpath.join("xl", target))

        sheets: list[SheetData] = []
        for sheet in workbook_root.findall("main:sheets/main:sheet", NS):
            sheet_name = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
            target = rel_map[rel_id]
            sheets.append(parse_sheet(zf, target, sheet_name, shared_strings))

        return sheets


def iter_rows(sheet: SheetData, min_row: int = 1, max_row: int | None = None):
    last_row = min(max_row or sheet.max_row, sheet.max_row)
    for row_index in range(min_row, last_row + 1):
        sparse = sheet.rows.get(row_index, {})
        yield row_index, [sparse.get(col) for col in range(1, sheet.max_column + 1)]


def detect_header_row(sheet: SheetData) -> int | None:
    first_row = next((row for _, row in iter_rows(sheet, 1, 1)), [])
    first_row_text = [text_value(item).lower() for item in first_row if text_value(item)]
    header_keywords = ("client", "policy", "premium", "brokerage", "servicer")
    if any(any(keyword in cell for keyword in header_keywords) for cell in first_row_text):
        return 1

    best_row = None
    best_score = -1
    for row_index, row in iter_rows(sheet, 1, min(15, sheet.max_row)):
        values = [text_value(item) for item in row]
        non_empty = [item for item in values if item]
        if len(non_empty) < 2:
            continue

        string_like = sum(1 for item in non_empty if numeric_value(item) is None)
        unique = len(set(non_empty))
        score = (string_like * 2) + unique + len(non_empty)
        if score > best_score:
            best_score = score
            best_row = row_index

    return best_row


def find_first_header(headers: list[str], tokens: tuple[str, ...]) -> str | None:
    for token in tokens:
        for header in headers:
            if token in header.lower():
                return header
    return None


def build_records(sheet: SheetData, header_row: int) -> tuple[list[str], list[dict[str, object]]]:
    header_values = next(row for _, row in iter_rows(sheet, header_row, header_row))
    headers = unique_headers(list(header_values))

    records: list[dict[str, object]] = []
    for _, row in iter_rows(sheet, header_row + 1, sheet.max_row):
        if not any(text_value(item) for item in row):
            continue
        records.append({header: row[index] for index, header in enumerate(headers)})

    return headers, records


def detect_columns(headers: list[str]) -> dict[str, str | None]:
    sector_header = find_first_header(headers, ("sector", "business sector"))
    policy_header = find_first_header(headers, ("policy type", "policy"))
    return {
        "client_name": find_first_header(headers, ("client name",)),
        "client_number": find_first_header(headers, ("client number",)),
        "policy_type": policy_header,
        "servicer": find_first_header(headers, ("servicer",)),
        "premium": find_first_header(headers, ("premium",)),
        "brokerage": find_first_header(headers, ("brokerage",)),
        "grouping": sector_header or policy_header or find_first_header(headers, ("servicer",)),
    }


def analyze_sheet(sheet: SheetData) -> SheetAnalysis:
    header_row = detect_header_row(sheet)
    if header_row is None:
        return SheetAnalysis(
            name=sheet.name,
            header_row=None,
            headers=[],
            records=[],
            detected_columns={},
            max_row=sheet.max_row,
            max_column=sheet.max_column,
        )

    headers, records = build_records(sheet, header_row)
    return SheetAnalysis(
        name=sheet.name,
        header_row=header_row,
        headers=headers,
        records=records,
        detected_columns=detect_columns(headers),
        max_row=sheet.max_row,
        max_column=sheet.max_column,
    )


@cache_data(show_spinner=False)
def load_analyses(workbook_path_text: str) -> list[SheetAnalysis]:
    workbook_path = Path(workbook_path_text)
    sheets = load_workbook(workbook_path)
    return [analyze_sheet(sheet) for sheet in sheets]


def non_empty_unique_values(records: list[dict[str, object]], header: str | None) -> list[str]:
    if not header:
        return []
    return sorted({text_value(record.get(header)) for record in records if text_value(record.get(header))})


def filter_records(
    records: list[dict[str, object]],
    policy_header: str | None,
    servicer_header: str | None,
    selected_policy: str,
    selected_servicer: str,
) -> list[dict[str, object]]:
    filtered: list[dict[str, object]] = []
    for record in records:
        policy = text_value(record.get(policy_header)) if policy_header else ""
        servicer = text_value(record.get(servicer_header)) if servicer_header else ""
        if policy_header and selected_policy != "All" and policy != selected_policy:
            continue
        if servicer_header and selected_servicer != "All" and servicer != selected_servicer:
            continue
        filtered.append(record)
    return filtered


def client_key(record: dict[str, object], client_name_header: str | None, client_number_header: str | None) -> str:
    if client_name_header:
        name = text_value(record.get(client_name_header))
        if name:
            return name

    if client_number_header:
        number = text_value(record.get(client_number_header))
        if number:
            return number

    return ""


def total_clients(records: list[dict[str, object]], client_name_header: str | None, client_number_header: str | None) -> int:
    return len(
        {
            client_key(record, client_name_header, client_number_header)
            for record in records
            if client_key(record, client_name_header, client_number_header)
        }
    )


def clients_by_policy_type(
    records: list[dict[str, object]],
    policy_header: str | None,
    client_name_header: str | None,
    client_number_header: str | None,
) -> list[tuple[str, int]]:
    if not policy_header:
        return []

    grouped: dict[str, set[str]] = {}
    for record in records:
        policy_type = text_value(record.get(policy_header)) or "Unspecified"
        key = client_key(record, client_name_header, client_number_header)
        if not key:
            continue
        grouped.setdefault(policy_type, set()).add(key)

    return sorted(
        ((policy_type, len(clients)) for policy_type, clients in grouped.items()),
        key=lambda item: item[1],
        reverse=True,
    )


@cache_data(show_spinner=False)
def aonpass_motor_private_clients(workbook_path_text: str) -> int:
    workbook_path = Path(workbook_path_text)
    if not workbook_path.exists():
        return 0

    analyses = load_analyses(str(workbook_path))
    for analysis in analyses:
        member_id_header = find_first_header(analysis.headers, ("member_id",))
        status_header = find_first_header(analysis.headers, ("status",))
        if not member_id_header:
            continue

        member_ids: set[str] = set()
        for record in analysis.records:
            status = text_value(record.get(status_header)).lower() if status_header else "active"
            member_id = text_value(record.get(member_id_header))
            if member_id and (not status_header or status == "active"):
                member_ids.add(member_id)
        if member_ids:
            return len(member_ids)

    return 0


def build_preview_rows(records: list[dict[str, object]], headers: list[str], limit: int | None = None) -> list[dict[str, str]]:
    preview_headers = [header for header in headers if not header.startswith("column_")]
    rows: list[dict[str, str]] = []
    source_records = records if limit is None else records[:limit]
    for record in source_records:
        rows.append({header: text_value(record.get(header)) for header in preview_headers})
    return rows


def image_data_uri(path: Path) -> str | None:
    if not path.exists():
        return None
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def render_preview_table(rows: list[dict[str, str]]) -> None:
    if pd is None:
        st.dataframe(rows, width="stretch", height=520, hide_index=True)
        return

    dataframe = pd.DataFrame(rows)
    styled = (
        dataframe.style.hide(axis="index")
        .set_table_styles(
            [
                {
                    "selector": "thead th",
                    "props": [
                        ("background-color", "#121212"),
                        ("color", "#ffffff"),
                        ("font-weight", "700"),
                        ("border-bottom", "2px solid #d71920"),
                    ],
                },
                {
                    "selector": "tbody tr:nth-child(odd)",
                    "props": [
                        ("background-color", "#ffffff"),
                        ("color", "#121212"),
                    ],
                },
                {
                    "selector": "tbody tr:nth-child(even)",
                    "props": [
                        ("background-color", "#ffd8dc"),
                        ("color", "#121212"),
                    ],
                },
                {
                    "selector": "tbody td",
                    "props": [
                        ("border-bottom", "1px solid #f0b2b7"),
                    ],
                },
            ]
        )
    )
    st.dataframe(styled, width="stretch", height=520)


def render_style() -> None:
    st.markdown(
        """
        <style>
        :root {
            --minet-red: #d71920;
            --minet-red-deep: #a10f15;
            --ink: #121212;
            --paper: #ffffff;
            --mist: #f5f5f5;
            --line: rgba(18, 18, 18, 0.12);
        }
        .stApp {
            background: linear-gradient(180deg, #ffffff 0%, #f7f7f7 100%);
            color: var(--ink);
        }
        section[data-testid="stSidebar"] {
            background: #0f0f10;
            border-right: 3px solid var(--minet-red);
        }
        section[data-testid="stSidebar"] * {
            color: #ffffff !important;
        }
        .hero {
            padding: 1.35rem 1.5rem;
            border-radius: 20px;
            background: linear-gradient(135deg, #0a0a0b 0%, #151517 62%, #8f0f16 100%);
            border-left: 8px solid var(--minet-red);
            color: white;
            box-shadow: 0 18px 45px rgba(0, 0, 0, 0.18);
            margin-bottom: 1rem;
        }
        .hero p {
            margin: 0.3rem 0 0;
            opacity: 0.92;
        }
        .hero-brand {
            display: flex;
            align-items: center;
            gap: 1rem;
        }
        .hero-logo {
            width: 200px;
            max-width: 100%;
            height: 120px;
            object-fit: contain;
            background: #ffffff;
            border-radius: 14px;
            padding: 0.45rem 0.7rem;
        }
        .hero-copy h1 {
            margin: 0.2rem 0 0.1rem !important;
            color: #ffffff !important;
        }
        .eyebrow {
            font-size: 0.78rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            opacity: 0.8;
        }
        [data-testid="stMetric"] {
            background: var(--paper);
            border: 1px solid var(--line);
            border-top: 6px solid var(--minet-red);
            border-radius: 18px;
            padding: 0.8rem;
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.08);
        }
        .stApp h2,
        .stApp h3 {
            color: var(--ink) !important;
        }
        [data-testid="stMetricLabel"] p,
        [data-testid="stMetricLabel"] label {
            color: #5a5a5a !important;
            opacity: 1 !important;
            font-weight: 600 !important;
        }
        [data-testid="stMetricValue"],
        [data-testid="stMetricValue"] * {
            color: var(--ink) !important;
        }
        div[data-testid="stDataFrame"] {
            background: var(--paper);
            border: 1px solid var(--line);
            border-radius: 18px;
            padding: 0.35rem;
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.06);
        }
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }
        .stButton button,
        .stDownloadButton button {
            background: var(--minet-red) !important;
            color: #ffffff !important;
            border: 1px solid var(--minet-red-deep) !important;
            border-radius: 999px !important;
            font-weight: 700 !important;
        }
        .stButton button:hover,
        .stDownloadButton button:hover {
            background: #111111 !important;
            color: #ffffff !important;
            border-color: #111111 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard() -> None:
    st.set_page_config(
        page_title="Car Insurance Database",
        page_icon=":bar_chart:",
        layout="wide",
    )
    render_style()

    if not WORKBOOK_PATH.exists():
        st.error(f"Workbook not found: {WORKBOOK_PATH}")
        return

    analyses = load_analyses(str(WORKBOOK_PATH))
    if not analyses:
        st.error("No sheets could be read from the workbook.")
        return

    analysis = next((item for item in analyses if item.records), analyses[0])

    if analysis.header_row is None:
        st.error("A header row could not be detected in the selected sheet.")
        return

    detected = analysis.detected_columns
    policy_options = non_empty_unique_values(analysis.records, detected.get("policy_type"))
    servicer_options = non_empty_unique_values(analysis.records, detected.get("servicer"))

    selected_policy = st.sidebar.selectbox(
        "Policy Type",
        options=["All"] + policy_options,
        index=0,
    ) if policy_options else "All"
    selected_servicer = st.sidebar.selectbox(
        "Servicer",
        options=["All"] + servicer_options,
        index=0,
    ) if servicer_options else "All"

    filtered_records = filter_records(
        analysis.records,
        detected.get("policy_type"),
        detected.get("servicer"),
        selected_policy,
        selected_servicer,
    )
    total_client_count = total_clients(
        filtered_records,
        detected.get("client_name"),
        detected.get("client_number"),
    )
    policy_client_counts = dict(
        clients_by_policy_type(
            filtered_records,
            detected.get("policy_type"),
            detected.get("client_name"),
            detected.get("client_number"),
        )
    )
    aonpass_private_count = aonpass_motor_private_clients(str(ACTIVE_WORKBOOK_PATH))
    include_aonpass_private = (
        selected_policy in ("All", "MOTOR PRIVATE")
        and selected_servicer == "All"
    )
    if include_aonpass_private and aonpass_private_count:
        total_client_count += aonpass_private_count
        policy_client_counts["MOTOR PRIVATE"] = policy_client_counts.get("MOTOR PRIVATE", 0) + aonpass_private_count

    sorted_policy_client_counts = sorted(
        policy_client_counts.items(),
        key=lambda item: item[1],
        reverse=True,
    )
    logo_uri = image_data_uri(LOGO_PATH)
    logo_markup = (
        f'<img src="{logo_uri}" alt="Logo" class="hero-logo">'
        if logo_uri
        else ""
    )

    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-brand">
                {logo_markup}
                <div class="hero-copy">
                    <h1>Car Insurance Database</h1>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric_items = [("Total Clients", total_client_count)] + sorted_policy_client_counts
    metric_columns = st.columns(len(metric_items)) if metric_items else []
    for column, (label, value) in zip(metric_columns, metric_items):
        with column:
            st.metric(label, format_number(value))

    preview_rows = build_preview_rows(filtered_records, analysis.headers, limit=None)
    if preview_rows:
        render_preview_table(preview_rows)
    else:
        st.write("No records match the current filters.")


def is_streamlit_runtime() -> bool:
    if os.environ.get("CLIENT_DASHBOARD_STREAMLIT") == "1":
        return True

    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
    except Exception:
        return False

    return get_script_run_ctx() is not None


def main() -> None:
    if st is None:
        print("Streamlit is not installed in this environment.")
        print("Install it with: python -m pip install streamlit")
        print("Then run: python run_dashboard.py")
        return

    if not is_streamlit_runtime():
        print("This file is a Streamlit app.")
        print("Run it with: python run_dashboard.py")
        return

    render_dashboard()


if __name__ == "__main__":
    main()

