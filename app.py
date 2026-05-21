import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

PAGE_TITLE = "Filename Extractor"
PAGE_ICON = ":file_folder:"
OUTPUT_FILENAME = "filenames.xlsx"
EXCEL_SHEET = "Filenames"

COLUMNS = ["Filename", "Extension", "Type", "Path", "Source"]


# ---------------------------------------------------------------------------
# Domain helpers
# ---------------------------------------------------------------------------

def _file_type(suffix: str, is_dir: bool = False) -> str:
    if is_dir:
        return "Folder"
    return {
        "PDF": "PDF Document",
        "CSV": "CSV Spreadsheet",
        "TXT": "Text File",
        "XLSX": "Excel Workbook",
        "XLS": "Excel Workbook (Legacy)",
        "DOCX": "Word Document",
        "DOC": "Word Document (Legacy)",
        "PPTX": "PowerPoint Presentation",
        "ZIP": "ZIP Archive",
    }.get(suffix, f"{suffix} File" if suffix else "Unknown")


def _row(name: str, path: str, source: str, is_dir: bool = False) -> dict:
    ext = Path(name).suffix.lstrip(".").upper()
    return {
        "Filename": name,
        "Extension": ext if not is_dir else "",
        "Type": _file_type(ext, is_dir),
        "Path": path,
        "Source": source,
    }


def extract_from_zip(uploaded_file) -> list[dict]:
    rows = []
    try:
        with zipfile.ZipFile(uploaded_file) as zf:
            for info in zf.infolist():
                p = Path(info.filename)
                rows.append(_row(
                    name=p.name or str(p),
                    path=info.filename,
                    source=uploaded_file.name,
                    is_dir=info.is_dir(),
                ))
    except zipfile.BadZipFile:
        st.warning(f"`{uploaded_file.name}` is not a valid ZIP archive — skipped.")
    return rows


def extract_from_files(uploaded_files) -> list[dict]:
    rows = []
    for f in uploaded_files:
        ext = Path(f.name).suffix.lstrip(".").upper()
        if ext == "ZIP":
            rows.extend(extract_from_zip(f))
        else:
            rows.append(_row(name=f.name, path=f.name, source=f.name))
    return rows


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def to_excel(rows: list[dict]) -> bytes:
    df = pd.DataFrame(rows, columns=COLUMNS)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=EXCEL_SHEET)
        ws = writer.sheets[EXCEL_SHEET]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# UI components
# ---------------------------------------------------------------------------

def render_summary(df: pd.DataFrame) -> None:
    total = len(df)
    folders = (df["Type"] == "Folder").sum()
    files = total - folders
    type_counts = df[df["Type"] != "Folder"]["Extension"].value_counts()

    cols = st.columns(3)
    cols[0].metric("Total Entries", total)
    cols[1].metric("Files", files)
    cols[2].metric("Folders", folders)

    if not type_counts.empty:
        with st.expander("Breakdown by extension"):
            st.bar_chart(type_counts)


def render_table(df: pd.DataFrame) -> None:
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Path": st.column_config.TextColumn(width="large"),
        },
    )


def render_download(rows: list[dict]) -> None:
    excel_bytes = to_excel(rows)
    st.download_button(
        label="Download Excel",
        data=excel_bytes,
        file_name=OUTPUT_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )


# ---------------------------------------------------------------------------
# App entry point
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title=PAGE_TITLE, page_icon=PAGE_ICON, layout="wide")
    st.title(f"Filename Extractor")
    st.caption(
        "Upload any files or ZIP archives (representing folders). "
        "All filenames are extracted and exported to Excel."
    )

    uploaded_files = st.file_uploader(
        label="Drop files here",
        accept_multiple_files=True,
        help="Accepts PDF, CSV, TXT, Excel, Word, ZIP, and any other file type.",
    )

    if not uploaded_files:
        st.info("Upload one or more files to get started. ZIP files are treated as folders.")
        return

    with st.spinner("Extracting filenames…"):
        rows = extract_from_files(uploaded_files)

    if not rows:
        st.warning("No filenames could be extracted from the uploaded files.")
        return

    df = pd.DataFrame(rows, columns=COLUMNS)
    st.success(f"Extracted **{len(df)}** entr{'y' if len(df) == 1 else 'ies'}.")

    render_summary(df)
    render_table(df)
    render_download(rows)


if __name__ == "__main__":
    main()
