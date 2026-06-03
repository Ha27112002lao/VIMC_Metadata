#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit UI for importing VIMC Business Glossary Excel files into OpenMetadata.

How to run:
    streamlit run streamlit_vimc_openmetadata_app.py

Required files in the same folder:
    - streamlit_vimc_openmetadata_app.py
    - import_vimc_glossary_3_excels_with_dq.py

Main flow:
    1) User enters OpenMetadata host/IP and access token.
    2) User uploads one or more Excel files.
    3) The app imports terms into OpenMetadata Glossary.
    4) Optional: publish Data Quality metadata/results using the existing DQ logic.
"""

from __future__ import annotations

import contextlib
import io
import os
import tempfile
import traceback
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

import streamlit as st

# Import the existing business logic from your CLI script.
# Keep import_vimc_glossary_3_excels_with_dq.py in the same folder as this Streamlit file.
try:
    from import_vimc_glossary_3_excels_with_dq import (  # type: ignore
        DEFAULT_DQ_DATABASE,
        DEFAULT_DQ_DATABASE_DISPLAY,
        DEFAULT_DQ_SCHEMA,
        DEFAULT_DQ_SERVICE,
        DEFAULT_DQ_SERVICE_TYPE,
        DEFAULT_DQ_TABLE,
        DEFAULT_DQ_TEST_SUITE,
        DEFAULT_GLOSSARY,
        DEFAULT_SHEET,
        OpenMetadataClient,
        collect_dq_rows,
        import_excel_file,
        run_data_quality_publication,
        safe_name,
    )
except Exception as import_error:  # pragma: no cover - shown in Streamlit UI
    OpenMetadataClient = None  # type: ignore
    _IMPORT_ERROR = import_error
else:
    _IMPORT_ERROR = None


st.set_page_config(
    page_title="VIMC → OpenMetadata Importer",
    page_icon="📘",
    layout="wide",
)


class StreamlitLogBuffer(io.StringIO):
    """Capture print logs and periodically render them in Streamlit."""

    def __init__(self, placeholder):
        super().__init__()
        self.placeholder = placeholder

    def write(self, text: str) -> int:
        result = super().write(text)
        current_text = self.getvalue()
        # Avoid re-rendering on every tiny whitespace write.
        if "\n" in text or len(current_text) < 3000:
            self.placeholder.code(current_text[-12000:] or "Đang chuẩn bị...", language="text")
        return result

    def flush(self) -> None:
        current_text = self.getvalue()
        self.placeholder.code(current_text[-12000:] or "Đang xử lý...", language="text")


def normalize_host(raw_host: str) -> str:
    host = (raw_host or "").strip()
    if not host:
        return ""
    if not host.startswith(("http://", "https://")):
        host = "http://" + host
    return host.rstrip("/")


def save_uploaded_excels(uploaded_files: Sequence, target_dir: Path) -> List[Path]:
    """Save Streamlit UploadedFile objects to disk so openpyxl can read them."""
    saved_paths: List[Path] = []
    for uploaded in uploaded_files:
        original_name = Path(uploaded.name).name
        suffix = Path(original_name).suffix.lower()
        if suffix not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
            raise ValueError(f"File không phải Excel được hỗ trợ: {original_name}")

        safe_stem = safe_name(Path(original_name).stem, fallback="uploaded_excel", max_len=80)
        target_path = target_dir / f"{safe_stem}{suffix}"
        counter = 2
        while target_path.exists():
            target_path = target_dir / f"{safe_stem}_{counter}{suffix}"
            counter += 1

        with open(target_path, "wb") as f:
            f.write(uploaded.getbuffer())
        saved_paths.append(target_path)
    return saved_paths


def run_import_flow(
    *,
    host: str,
    token: str,
    excel_paths: Sequence[Path],
    glossary_display: str,
    requested_sheet: str,
    all_sheets: bool,
    dry_run: bool,
    sleep_seconds: float,
    header_row: Optional[int],
    data_start_row: Optional[int],
    name_mode: str,
    ungrouped_label: str,
    timeout: int,
    retries: int,
    enable_data_quality: bool,
    dq_strict: bool,
    dq_service_name: str,
    dq_service_type: str,
    dq_database_name: str,
    dq_database_display: str,
    dq_schema: str,
    dq_table: str,
    dq_test_suite: str,
) -> Tuple[int, int, Optional[Tuple[int, int]]]:
    """Run the same logic as the CLI script, but from Streamlit."""
    if OpenMetadataClient is None:
        raise RuntimeError(
            "Không import được file import_vimc_glossary_3_excels_with_dq.py. "
            "Hãy đặt file đó cùng thư mục với streamlit_vimc_openmetadata_app.py."
        )

    glossary_name = safe_name(glossary_display, fallback=DEFAULT_GLOSSARY)

    print("=== VIMC/OpenMetadata Streamlit Import ===")
    print(f"OpenMetadata host: {host}")
    print(f"Glossary: {glossary_name}")
    print(f"Requested sheet: {requested_sheet}")
    print(f"All sheets: {all_sheets}")
    print(f"Name mode: {name_mode}")
    print(f"Dry run: {dry_run}")
    print(f"Data Quality: {'enabled' if enable_data_quality else 'disabled'}")
    print("Uploaded files:")
    for path in excel_paths:
        print(f"- {path.name}")

    client = OpenMetadataClient(
        host=host or "http://localhost:8585",
        token=token,
        timeout=timeout,
        retries=retries,
    )

    if not dry_run:
        client.health_check()

    client.upsert_glossary(
        name=glossary_name,
        display_name=glossary_display,
        description="Business Glossary imported automatically from Streamlit Excel upload.",
        dry_run=dry_run,
    )

    total_groups = 0
    total_terms = 0

    for excel_path in excel_paths:
        groups, terms = import_excel_file(
            client=client,
            excel_path=excel_path,
            glossary_name=glossary_name,
            requested_sheet=requested_sheet,
            all_sheets=all_sheets,
            dry_run=dry_run,
            sleep_seconds=sleep_seconds,
            header_row=header_row,
            data_start_row=data_start_row,
            name_mode=name_mode,
            ungrouped_label=ungrouped_label,
        )
        total_groups += groups
        total_terms += terms

    print(f"Tổng group terms: {total_groups}")
    print(f"Tổng business/indicator terms: {total_terms}")

    dq_summary: Optional[Tuple[int, int]] = None
    if enable_data_quality:
        try:
            dq_rows = collect_dq_rows(
                excel_paths=list(excel_paths),
                glossary_name=glossary_name,
                requested_sheet=requested_sheet,
                all_sheets=all_sheets,
                header_row=header_row,
                data_start_row=data_start_row,
                ungrouped_label=ungrouped_label,
            )
            dq_summary = run_data_quality_publication(
                client=client,
                rows=dq_rows,
                service_name=safe_name(dq_service_name, fallback=DEFAULT_DQ_SERVICE),
                service_type=dq_service_type,
                database_name=safe_name(dq_database_name, fallback=DEFAULT_DQ_DATABASE),
                database_display_name=dq_database_display,
                schema_name=safe_name(dq_schema, fallback=DEFAULT_DQ_SCHEMA),
                table_name=safe_name(dq_table, fallback=DEFAULT_DQ_TABLE),
                test_suite_name=safe_name(dq_test_suite, fallback=DEFAULT_DQ_TEST_SUITE),
                dry_run=dry_run,
            )
        except Exception as exc:
            print(f"\n⚠️ Glossary đã import xong nhưng Data Quality bị lỗi: {exc}")
            print("Gợi ý: tắt 'Bật Data Quality' nếu chỉ cần import Glossary, hoặc kiểm tra version/API Data Quality của OpenMetadata.")
            if dq_strict:
                raise

    if not dry_run:
        print("\nKiểm tra trên OpenMetadata UI:")
        print(f"- Govern → Glossary → {glossary_name}")
        if enable_data_quality:
            print(f"- Observability → Data Quality → Test Suites → {safe_name(dq_test_suite, fallback=DEFAULT_DQ_TEST_SUITE)}")
            print(f"- Explore → Databases → {dq_database_display}")

    return total_groups, total_terms, dq_summary


st.title("📘 VIMC Excel → OpenMetadata")
st.caption("Upload Excel trên Streamlit, nhập IP/token OpenMetadata, rồi import tự động vào Glossary.")

if _IMPORT_ERROR is not None:
    st.error("Không import được module xử lý OpenMetadata.")
    st.code(str(_IMPORT_ERROR), language="text")
    st.info("Hãy đặt file `import_vimc_glossary_3_excels_with_dq.py` cùng thư mục với file Streamlit này rồi chạy lại.")
    st.stop()

with st.sidebar:
    st.header("1) Kết nối OpenMetadata")
    host_input = st.text_input(
        "OpenMetadata IP/Host",
        value=os.getenv("OPENMETADATA_HOST", "http://192.168.74.12:30085"),
        help="Có thể nhập dạng 192.168.74.12:30085 hoặc http://192.168.74.12:30085",
    )
    # st.text_area does not support type="password" in Streamlit.
    # Use st.text_input for a masked one-line JWT/PAT token.
    token_input = st.text_input(
        "Access token / JWT token",
        value=os.getenv("OM_JWT_TOKEN", ""),
        type="password",
        help="Lấy token trong OpenMetadata rồi dán vào đây. Token không được lưu vĩnh viễn bởi app.",
    )

    st.divider()
    st.header("2) Cấu hình import")
    glossary_display = st.text_input("Glossary name", value=DEFAULT_GLOSSARY)
    requested_sheet = st.text_input("Sheet mặc định", value=DEFAULT_SHEET)
    all_sheets = st.checkbox("Import tất cả sheet", value=False)
    dry_run = st.checkbox("Dry-run: chỉ kiểm tra, không ghi OpenMetadata", value=False)

    with st.expander("Cấu hình nâng cao", expanded=False):
        name_mode = st.selectbox("Name mode", options=["ordered", "original"], index=0)
        ungrouped_label = st.text_input("Group mặc định nếu thiếu group", value="Khác")
        header_row_raw = st.number_input("Ép dòng header; 0 = tự dò", min_value=0, max_value=500, value=0, step=1)
        data_start_row_raw = st.number_input("Ép dòng bắt đầu dữ liệu; 0 = tự dò", min_value=0, max_value=1000, value=0, step=1)
        sleep_seconds = st.number_input("Sleep giữa API calls", min_value=0.0, max_value=5.0, value=0.0, step=0.05)
        timeout = st.number_input("API timeout giây", min_value=5, max_value=300, value=45, step=5)
        retries = st.number_input("Retry lỗi 429/5xx", min_value=0, max_value=10, value=2, step=1)

    st.divider()
    st.header("3) Data Quality")
    enable_data_quality = st.checkbox(
        "Bật Data Quality metadata/result",
        value=False,
        help="Nên tắt nếu server OpenMetadata của bạn chưa ổn định API Data Quality. Import Glossary vẫn chạy bình thường.",
    )
    dq_strict = st.checkbox("DQ lỗi thì dừng toàn bộ", value=False, disabled=not enable_data_quality)
    with st.expander("Cấu hình Data Quality", expanded=False):
        dq_service_name = st.text_input("DQ service name", value=DEFAULT_DQ_SERVICE, disabled=not enable_data_quality)
        dq_service_type = st.text_input("DQ service type", value=DEFAULT_DQ_SERVICE_TYPE, disabled=not enable_data_quality)
        dq_database_name = st.text_input("DQ database entity name", value=DEFAULT_DQ_DATABASE, disabled=not enable_data_quality)
        dq_database_display = st.text_input("DQ database display name", value=DEFAULT_DQ_DATABASE_DISPLAY, disabled=not enable_data_quality)
        dq_schema = st.text_input("DQ schema", value=DEFAULT_DQ_SCHEMA, disabled=not enable_data_quality)
        dq_table = st.text_input("DQ table", value=DEFAULT_DQ_TABLE, disabled=not enable_data_quality)
        dq_test_suite = st.text_input("DQ test suite", value=DEFAULT_DQ_TEST_SUITE, disabled=not enable_data_quality)

st.subheader("Upload Excel")
uploaded_files = st.file_uploader(
    "Chọn file Excel Business Glossary",
    type=["xlsx", "xlsm", "xltx", "xltm"],
    accept_multiple_files=True,
)

col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    run_button = st.button("🚀 Import lên OpenMetadata", type="primary", use_container_width=True)
with col2:
    st.write("")
    st.write("")

st.markdown(
    """
**Luồng sử dụng:** vào OpenMetadata lấy access token → nhập IP/token ở sidebar → upload Excel → bấm import → vào OpenMetadata kiểm tra lại ở `Govern → Glossary`.
"""
)

if uploaded_files:
    st.success(f"Đã chọn {len(uploaded_files)} file Excel: " + ", ".join(file.name for file in uploaded_files))
else:
    st.info("Chưa có file Excel. Hãy upload file trước khi import.")

log_placeholder = st.empty()

if run_button:
    host = normalize_host(host_input)
    token = token_input.strip()

    if not dry_run and not host:
        st.error("Thiếu OpenMetadata host/IP.")
        st.stop()
    if not dry_run and not token:
        st.error("Thiếu access token/JWT token.")
        st.stop()
    if not uploaded_files:
        st.error("Bạn cần upload ít nhất 1 file Excel.")
        st.stop()

    header_row = int(header_row_raw) if int(header_row_raw) > 0 else None
    data_start_row = int(data_start_row_raw) if int(data_start_row_raw) > 0 else None

    with tempfile.TemporaryDirectory(prefix="vimc_streamlit_excel_") as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)
        try:
            excel_paths = save_uploaded_excels(uploaded_files, tmp_dir)
        except Exception as exc:
            st.error(f"Không lưu/đọc được file upload: {exc}")
            st.stop()

        log_buffer = StreamlitLogBuffer(log_placeholder)
        try:
            with st.spinner("Đang import vào OpenMetadata..."):
                with contextlib.redirect_stdout(log_buffer), contextlib.redirect_stderr(log_buffer):
                    groups, terms, dq_summary = run_import_flow(
                        host=host,
                        token=token,
                        excel_paths=excel_paths,
                        glossary_display=glossary_display,
                        requested_sheet=requested_sheet,
                        all_sheets=all_sheets,
                        dry_run=dry_run,
                        sleep_seconds=float(sleep_seconds),
                        header_row=header_row,
                        data_start_row=data_start_row,
                        name_mode=name_mode,
                        ungrouped_label=ungrouped_label,
                        timeout=int(timeout),
                        retries=int(retries),
                        enable_data_quality=enable_data_quality,
                        dq_strict=dq_strict,
                        dq_service_name=dq_service_name,
                        dq_service_type=dq_service_type,
                        dq_database_name=dq_database_name,
                        dq_database_display=dq_database_display,
                        dq_schema=dq_schema,
                        dq_table=dq_table,
                        dq_test_suite=dq_test_suite,
                    )
        except Exception as exc:
            log_buffer.write("\n\n❌ Lỗi chi tiết:\n")
            log_buffer.write(traceback.format_exc())
            log_buffer.flush()
            st.error(f"Import thất bại: {exc}")
            st.stop()
        finally:
            log_buffer.flush()

    st.success("Import hoàn tất." if not dry_run else "Dry-run hoàn tất, chưa ghi vào OpenMetadata.")
    m1, m2, m3 = st.columns(3)
    m1.metric("Group terms", groups)
    m2.metric("Business/indicator terms", terms)
    if dq_summary:
        dq_success, dq_failed = dq_summary
        m3.metric("DQ Success / Failed", f"{dq_success} / {dq_failed}")
    else:
        m3.metric("Data Quality", "Tắt hoặc không publish")

    st.markdown("### Kiểm tra lại trên OpenMetadata")
    st.markdown(f"- `Govern → Glossary → {safe_name(glossary_display, fallback=DEFAULT_GLOSSARY)}`")
    if enable_data_quality:
        st.markdown(f"- `Observability → Data Quality → Test Suites → {safe_name(dq_test_suite, fallback=DEFAULT_DQ_TEST_SUITE)}`")
        st.markdown(f"- `Explore → Databases → {dq_database_display}`")
