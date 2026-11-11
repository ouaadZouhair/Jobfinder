
import streamlit as st
import pandas as pd
import importlib
import local_scrapers

from jobspy import scrape_jobs
from io import BytesIO
from datetime import datetime
from urllib.parse import urlparse

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


# =====================================================
# 1. APPLICATION CONFIGURATION
# =====================================================

st.set_page_config(
    page_title="Morocco Job Finder",
    page_icon="💼",
    layout="wide"
)

st.title("💼 Morocco Job Finder")
st.caption(
    "Search jobs across Moroccan cities and export "
    "your results to Excel."
)


# =====================================================
# 2. CONFIGURATION DATA
# =====================================================

CITIES = [
    "Casablanca",
    "Rabat",
    "Tanger",
    "Marrakech",
    "Fès",
    "Agadir",
    "Kénitra",
    "Meknès",
    "Oujda",
    "Tétouan",
    "El Jadida",
    "Safi",
    "Salé",
    "Mohammedia",
    "Béni Mellal",
    "Nador",
    "Settat",
    "Khouribga",
    "Laâyoune",
    "Dakhla"
]

PLATFORMS = {
    "Indeed": "indeed",
    "LinkedIn": "linkedin",
    "Google Jobs": "google",
    "Glassdoor": "glassdoor",
    "ZipRecruiter": "zip_recruiter",
    "Bayt": "bayt",
    "Emploi.ma": "emploi_ma",
    "MarocAnnonces": "marocannonces",
    "ReKrute.com": "rekrute",
    "MarocEmploi.net": "marocemploi",
}

COLUMNS = [
    "site",
    "title",
    "company",
    "location",
    "date_posted",
    "job_type",
    "min_amount",
    "max_amount",
    "currency",
    "job_url",
    "search_keyword",
    "search_city"
]

COLUMN_NAMES = {
    "site": "Platform",
    "title": "Job Title",
    "company": "Company",
    "location": "Location",
    "date_posted": "Date Posted",
    "job_type": "Job Type",
    "min_amount": "Minimum Salary",
    "max_amount": "Maximum Salary",
    "currency": "Currency",
    "job_url": "Application Link",
    "search_keyword": "Search Keyword",
    "search_city": "Search City"
}


# =====================================================
# 3. HELPER FUNCTIONS
# =====================================================

def parse_keywords(text):
    """Convert comma/newline-separated text into keywords."""

    keywords = [
        item.strip()
        for item in text.replace(",", "\n").splitlines()
        if item.strip()
    ]

    return list(dict.fromkeys(keywords))


def prepare_jobs(dataframes):
    """Combine, deduplicate and sort job results."""

    df = pd.concat(dataframes, ignore_index=True)

    for column in COLUMNS:
        if column not in df.columns:
            df[column] = None

    df = df[COLUMNS].copy()

    # Convert dates, preserving valid timestamps.
    # UTC normalization avoids mixed-timezone errors.
    df["date_posted"] = pd.to_datetime(
        df["date_posted"],
        errors="coerce",
        utc=True
    )

    # Normalize links before removing duplicates
    df["job_url"] = df["job_url"].astype("string").str.strip()

    # Sort newest first
    df = df.sort_values(
        by="date_posted",
        ascending=False,
        na_position="last",
        kind="stable"
    )

    # Remove duplicates with the same URL
    has_url = df["job_url"].notna() & df["job_url"].ne("")

    with_url = df.loc[has_url].drop_duplicates(
        subset=["job_url"],
        keep="first"
    )

    without_url = df.loc[~has_url].drop_duplicates(
        subset=["site", "title", "company", "location"],
        keep="first"
    )

    df = pd.concat(
        [with_url, without_url],
        ignore_index=True
    )

    # Re-sort after combining
    df = df.sort_values(
        by="date_posted",
        ascending=False,
        na_position="last",
        kind="stable"
    )

    return df.reset_index(drop=True)


def display_jobs(df):
    """Prepare a readable DataFrame for the UI."""

    result = df.copy()

    # JobSpy usually returns date-only values.
    # Do not invent a publication time.
    result["date_posted"] = (
        result["date_posted"]
        .dt.strftime("%Y-%m-%d")
        .fillna("Unknown")
    )

    return result.rename(columns=COLUMN_NAMES)


def safe_excel_value(value):
    """Convert values safely for Excel."""

    if value is None:
        return None

    if isinstance(value, (list, dict, tuple, set)):
        value = str(value)

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, str):
        # Prevent untrusted scraped text from becoming formulas.
        if value.startswith(("=", "+", "-", "@", "\t", "\r", "\n")):
            return "'" + value

    return value


# =====================================================
# 4. EXCEL EXPORT
# =====================================================

def export_excel(df):
    """Create a styled Excel workbook in memory."""

    wb = Workbook()
    ws = wb.active
    ws.title = "Morocco Jobs"

    headers = [COLUMN_NAMES[col] for col in COLUMNS]
    ws.append(headers)

    export_df = df.copy()

    # Keep publication dates as dates, not fabricated times.
    export_df["date_posted"] = (
        export_df["date_posted"].dt.date
    )

    for record in export_df[COLUMNS].itertuples(
        index=False,
        name=None
    ):
        ws.append([
            safe_excel_value(value)
            for value in record
        ])

    ws.freeze_panes = "A2"

    # Create a real Excel table
    if ws.max_row > 1:
        table = Table(
            displayName="MoroccoJobs",
            ref=f"A1:L{ws.max_row}"
        )

        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        ws.add_table(table)

    widths = [
        18, 42, 30, 30, 18, 20,
        20, 20, 14, 60, 30, 24
    ]

    for index, width in enumerate(widths, start=1):
        letter = get_column_letter(index)
        ws.column_dimensions[letter].width = width

    # Style headers
    for cell in ws[1]:
        cell.font = Font(
            bold=True,
            color="FFFFFF",
            size=11
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[1].height = 30

    # Format date cells
    for row in range(2, ws.max_row + 1):
        date_cell = ws.cell(row, 5)
        date_cell.number_format = "yyyy-mm-dd"

        # Salary number formats
        for col in (7, 8):
            ws.cell(row, col).number_format = "#,##0.00"

        # Clickable application links
        link_cell = ws.cell(row, 10)
        url = link_cell.value

        if isinstance(url, str):
            parsed = urlparse(url)

            if (
                parsed.scheme in ("http", "https")
                and parsed.netloc
            ):
                link_cell.hyperlink = url
                link_cell.style = "Hyperlink"

    output = BytesIO()
    wb.save(output)

    return output.getvalue()


# =====================================================
# 5. SESSION STATE
# =====================================================

if "jobs" not in st.session_state:
    st.session_state.jobs = pd.DataFrame(columns=COLUMNS)

if "errors" not in st.session_state:
    st.session_state.errors = []


# =====================================================
# 6. SIDEBAR - SEARCH FORM
# =====================================================

with st.sidebar:
    st.header("⚙️ Search Configuration")

    with st.form("job_search_form"):

        st.subheader("Job Keywords")

        keywords_text = st.text_area(
            "Enter job titles or keywords",
            value=(
                "Laravel Developer\n"
                "React Developer\n"
                "Python Developer\n"
                "Full Stack Developer"
            ),
            height=150,
            help=(
                "Enter one keyword per line "
                "or separate keywords with commas."
            )
        )

        keywords = parse_keywords(keywords_text)

        st.caption(f"{len(keywords)} keywords selected")

        st.divider()

        st.subheader("📍 Moroccan Cities")

        cities = st.multiselect(
            "Choose cities",
            options=CITIES,
            default=["Casablanca", "Rabat"]
        )

        st.divider()

        st.subheader("🌐 Job Platforms")

        selected_platforms = st.multiselect(
            "Choose platforms",
            options=list(PLATFORMS.keys()),
            default=["Indeed"]
        )

        st.caption(
            "Platform availability and Morocco coverage vary."
        )

        st.divider()

        st.subheader("📅 Publication Date")

        period = st.selectbox(
            "Search jobs published within",
            options=[
                "Any time",
                "Last 24 hours",
                "Last 3 days",
                "Last 7 days",
                "Last 14 days",
                "Last 30 days"
            ],
            index=3
        )

        PERIOD_HOURS = {
            "Any time": None,
            "Last 24 hours": 24,
            "Last 3 days": 72,
            "Last 7 days": 168,
            "Last 14 days": 336,
            "Last 30 days": 720
        }

        hours_old = PERIOD_HOURS[period]

        st.divider()

        st.subheader("Search Limits")

        limit = st.number_input(
            "Results per keyword, city and platform",
            min_value=1,
            max_value=100,
            value=20,
            step=5
        )

        total_searches = (
            len(keywords)
            * len(cities)
            * len(selected_platforms)
        )

        st.info(
            f"Planned searches: {total_searches}\n\n"
            f"Maximum requested results: "
            f"{total_searches * limit}"
        )

        search = st.form_submit_button(
            "🔍 Search Jobs",
            type="primary",
            use_container_width=True
        )


# =====================================================
# 7. SEARCH JOBS
# =====================================================

if search:
    # Streamlit reruns this script while imported modules remain cached.
    # Refresh local source definitions before dispatching a new search.
    local_scrapers = importlib.reload(local_scrapers)

    if not keywords:
        st.warning("Please enter at least one keyword.")

    elif not cities:
        st.warning("Please select at least one city.")

    elif not selected_platforms:
        st.warning("Please select at least one platform.")

    else:
        results = []
        errors = []
        blocked_local_sources = set()

        total = (
            len(keywords)
            * len(cities)
            * len(selected_platforms)
        )

        completed = 0

        progress = st.progress(0)
        status = st.empty()

        for keyword in keywords:
            for city in cities:
                for platform_name in selected_platforms:

                    platform = PLATFORMS[platform_name]

                    if platform in blocked_local_sources:
                        completed += 1
                        progress.progress(completed / total)
                        continue

                    status.info(
                        f"Searching '{keyword}' "
                        f"on {platform_name} "
                        f"in {city}..."
                    )

                    try:
                        if platform in local_scrapers.LOCAL_SOURCES:
                            jobs = local_scrapers.scrape_local_jobs(
                                source=platform,
                                keyword=keyword,
                                city=city,
                                results_wanted=int(limit),
                                hours_old=hours_old,
                            )
                        else:
                            params = {
                                "site_name": [platform],
                                "search_term": keyword,
                                "location": f"{city}, Morocco",
                                "country_indeed": "Morocco",
                                "results_wanted": int(limit)
                            }

                            if hours_old is not None:
                                params["hours_old"] = hours_old

                            jobs = scrape_jobs(**params)

                        if not jobs.empty:
                            jobs = jobs.copy()

                            jobs["search_keyword"] = keyword
                            jobs["search_city"] = city

                            results.append(jobs)

                    except Exception as error:
                        status_code = getattr(
                            getattr(error, "response", None),
                            "status_code",
                            None,
                        )
                        if (
                            platform in local_scrapers.LOCAL_SOURCES
                            and status_code in (403, 429)
                        ):
                            blocked_local_sources.add(platform)
                            if platform == "emploi_ma" and status_code == 403:
                                errors.append(
                                    "Emploi.ma requires browser verification "
                                    "(HTTP 403). Search directly at "
                                    "https://www.emploi.ma/recherche-jobs-maroc. "
                                    "Remaining Emploi.ma searches were skipped."
                                )
                            else:
                                errors.append(
                                    f"{platform_name}: HTTP {status_code} "
                                    "from the site. Remaining searches on "
                                    "this platform were skipped."
                                )
                        else:
                            errors.append(
                                f"{platform_name} / {city} / "
                                f"{keyword}: {str(error)}"
                            )

                    completed += 1
                    progress.progress(completed / total)

        status.empty()

        st.session_state.errors = errors

        if results:
            df = prepare_jobs(results)

            st.session_state.jobs = df

            st.success(
                f"Successfully collected "
                f"{len(df)} unique job offers!"
            )

        else:
            st.session_state.jobs = pd.DataFrame(
                columns=COLUMNS
            )

            st.warning(
                "No jobs collected. Try another "
                "keyword, city, platform or date range."
            )


# =====================================================
# 8. RESULTS DASHBOARD
# =====================================================

df = st.session_state.jobs

if not df.empty:

    st.subheader("📊 Search Results")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Total Jobs",
        len(df)
    )

    col2.metric(
        "Companies",
        df["company"].nunique()
    )

    col3.metric(
        "Platforms",
        df["site"].nunique()
    )

    col4.metric(
        "Locations",
        df["location"].nunique()
    )

    st.divider()

    # Additional results filters
    filter1, filter2 = st.columns(2)

    available_sites = sorted(
        df["site"].dropna().astype(str).unique()
    )

    with filter1:
        filter_sites = st.multiselect(
            "Filter by platform",
            options=available_sites,
            default=available_sites
        )

    with filter2:
        text_filter = st.text_input(
            "Filter results by title or company",
            placeholder="e.g. Laravel"
        )

    filtered = df.copy()

    filtered = filtered[
        filtered["site"].astype(str).isin(filter_sites)
    ]

    if text_filter.strip():
        mask = (
            filtered["title"]
            .fillna("")
            .astype(str)
            .str.contains(
                text_filter,
                case=False,
                regex=False
            )
            |
            filtered["company"]
            .fillna("")
            .astype(str)
            .str.contains(
                text_filter,
                case=False,
                regex=False
            )
        )

        filtered = filtered[mask]

    filtered = filtered.sort_values(
        "date_posted",
        ascending=False,
        na_position="last",
        kind="stable"
    )

    st.caption(
        f"Showing {len(filtered)} jobs | "
        "Sorted by publication date: newest first"
    )

    st.dataframe(
        display_jobs(filtered),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Application Link": st.column_config.LinkColumn(
                "Apply / View Job",
                display_text="Open Job"
            )
        }
    )

    st.divider()

    # Excel export
    st.subheader("📥 Export Results")

    export_col1, export_col2 = st.columns(2)

    with export_col1:
        st.download_button(
            label="Download Filtered Excel",
            data=export_excel(filtered),
            file_name=(
                "morocco_jobs_"
                + datetime.now().strftime("%Y%m%d_%H%M")
                + ".xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
            disabled=filtered.empty
        )

    with export_col2:
        st.download_button(
            label="Download All Jobs (Excel)",
            data=export_excel(df),
            file_name="morocco_all_jobs.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

else:
    st.info(
        "Configure your search in the sidebar "
        "and click Search Jobs."
    )


# =====================================================
# 9. ERROR REPORTING
# =====================================================

if st.session_state.errors:

    with st.expander(
        f"⚠️ Search Errors "
        f"({len(st.session_state.errors)})"
    ):
        for error in st.session_state.errors:
            st.write(error)


# =====================================================
# 10. FOOTER
# =====================================================

st.divider()

st.caption(
    "Morocco Job Finder | Python + JobSpy + Streamlit"
)
