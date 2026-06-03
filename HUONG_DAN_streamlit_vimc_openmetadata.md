# Hướng dẫn chạy Streamlit: Upload Excel lên OpenMetadata

## 1. File cần đặt cùng thư mục

Đặt 2 file Python trong cùng một folder:

```text
streamlit_vimc_openmetadata_app.py
import_vimc_glossary_3_excels_with_dq.py
```

Có thể đặt thêm file này để cài thư viện:

```text
requirements_streamlit_vimc.txt
```

## 2. Cài thư viện

Mở PowerShell tại thư mục chứa file rồi chạy:

```powershell
pip install -r requirements_streamlit_vimc.txt
```

Nếu chưa có Streamlit:

```powershell
pip install streamlit openpyxl requests
```

## 3. Chạy UI Streamlit

Cách 1:

```powershell
streamlit run streamlit_vimc_openmetadata_app.py
```

Cách 2: double click file:

```text
run_streamlit_vimc.bat
```

## 4. Quy trình sử dụng trên UI

### Bước 1: Lấy token OpenMetadata

Vào OpenMetadata, lấy access token/JWT token của user hoặc bot account có quyền tạo/sửa Glossary.

Trên sidebar Streamlit, nhập:

```text
OpenMetadata IP/Host: http://192.168.74.12:30085
Access token / JWT token: <token của bạn>
```

### Bước 2: Upload Excel

Ở phần **Upload Excel**, chọn file Excel Business Glossary.

Mặc định app dùng:

```text
Glossary name: VIMC_Business_Glossary
Sheet mặc định: Logic Tổng hợp
```

Nếu file không có sheet `Logic Tổng hợp`, code gốc sẽ tự chọn sheet có cấu trúc giống glossary nhất.

### Bước 3: Import

Bấm:

```text
Import lên OpenMetadata
```

App sẽ:

1. Kết nối OpenMetadata bằng IP/token.
2. Tạo hoặc cập nhật Glossary.
3. Đọc Excel.
4. Tự dò header và các cột quan trọng.
5. Tạo source term theo tên file.
6. Tạo group term.
7. Tạo business/indicator terms.
8. In log xử lý ngay trên Streamlit.

### Bước 4: Kiểm tra trên OpenMetadata

Sau khi app báo thành công, vào:

```text
Govern → Glossary → VIMC_Business_Glossary
```

Nếu bật Data Quality, kiểm tra thêm:

```text
Observability → Data Quality → Test Suites
Explore → Databases → VIMC Technical term
```

## 5. Các tùy chọn quan trọng

### Dry-run

Bật **Dry-run** nếu chỉ muốn kiểm tra file Excel và log import mà chưa ghi vào OpenMetadata.

### Import tất cả sheet

Bật **Import tất cả sheet** nếu muốn đọc toàn bộ workbook thay vì chỉ sheet mặc định.

### Data Quality

Mặc định UI để **Data Quality tắt** để tránh lỗi nếu version OpenMetadata chưa hỗ trợ endpoint Data Quality tương ứng.

Nếu muốn publish Data Quality metadata/result, bật:

```text
Bật Data Quality metadata/result
```

Nếu server báo lỗi Data Quality, Glossary vẫn có thể đã import thành công. Khi đó kiểm tra ở:

```text
Govern → Glossary
```

## 6. Cách hoạt động kỹ thuật

File Streamlit không viết lại logic import. Nó gọi lại các function trong file:

```text
import_vimc_glossary_3_excels_with_dq.py
```

Các function chính được dùng:

```text
OpenMetadataClient
import_excel_file
collect_dq_rows
run_data_quality_publication
safe_name
```

Nhờ vậy, logic cũ vẫn được giữ nguyên, nhưng người dùng không cần chạy command line nữa.
