# Hướng dẫn chạy `import_vimc_glossary_3_excels_with_dq.py`

## 1. Mục tiêu file

File Python này giữ nguyên chức năng cũ: đọc Excel Business Glossary, tự dò header/cột, tạo Glossary và Glossary Terms trên OpenMetadata.

Phần mới được bổ sung: sau khi import Glossary xong, chương trình tự tạo một cây metadata kỹ thuật để phục vụ Observability/Data Quality:

```text
Database Service: VIMC_Technical_term_Service
└── Database: VIMC_Technical_term
    displayName: VIMC Technical term
    └── Schema: glossary_quality
        └── Table: imported_glossary_terms
```

Sau đó chương trình tạo Test Suite `VIMC_Technical_term_DQ`, tạo các Test Case Data Quality mặc định và publish kết quả test vào OpenMetadata.

> Lưu ý: vì đầu vào là Excel, không phải database thật có connector đang chạy, chương trình tự tính metric từ dữ liệu Excel rồi ghi kết quả test vào OpenMetadata. Nghĩa là kết quả vẫn hiện trong Observability > Data Quality, nhưng không phải do OpenMetadata query trực tiếp database vật lý.

## 2. Chuẩn bị môi trường

Cài thư viện Python:

```powershell
pip install requests openpyxl
```

Thiết lập OpenMetadata host và token:

```powershell
$env:OPENMETADATA_HOST='http://192.168.74.12:30085'
$env:OM_JWT_TOKEN='PASTE_TOKEN_CUA_BAN'
```

Token cần có quyền tạo/sửa:

- Glossary
- Glossary Terms
- Database Service
- Database
- Database Schema
- Table
- Data Quality Test Suite
- Data Quality Test Case
- Test Case Result

## 3. Chạy thử không ghi OpenMetadata

```powershell
python .\import_vimc_glossary_3_excels_with_dq.py --dry-run --excel 'Business Glossary Demo.xlsx'
```

Lệnh này chỉ đọc Excel, in kế hoạch import Glossary, in kế hoạch tạo Database kỹ thuật và in các test Data Quality sẽ được publish.

## 4. Chạy thật

```powershell
python .\import_vimc_glossary_3_excels_with_dq.py --excel 'Business Glossary Demo.xlsx'
```

Nếu có nhiều file:

```powershell
python .\import_vimc_glossary_3_excels_with_dq.py --excel 'file1.xlsx' --excel 'file2.xlsx'
```

Nếu muốn chỉ chạy chức năng cũ, bỏ qua Data Quality:

```powershell
python .\import_vimc_glossary_3_excels_with_dq.py --excel 'Business Glossary Demo.xlsx' --skip-data-quality
```

Nếu muốn script trả lỗi khi phần Data Quality lỗi:

```powershell
python .\import_vimc_glossary_3_excels_with_dq.py --excel 'Business Glossary Demo.xlsx' --dq-strict
```

## 5. Các tham số Data Quality quan trọng

| Tham số | Mặc định | Ý nghĩa |
|---|---:|---|
| `--skip-data-quality` | false | Tắt phần tạo Database và publish Data Quality |
| `--dq-service-name` | `VIMC_Technical_term_Service` | Database Service kỹ thuật |
| `--dq-service-type` | `CustomDatabase` | Service type dùng khi tạo Database Service |
| `--dq-database-name` | `VIMC_Technical_term` | Entity name của Database |
| `--dq-database-display` | `VIMC Technical term` | Tên hiển thị trên UI |
| `--dq-schema` | `glossary_quality` | Schema kỹ thuật |
| `--dq-table` | `imported_glossary_terms` | Table kỹ thuật để gắn test cases |
| `--dq-test-suite` | `VIMC_Technical_term_DQ` | Test Suite hiển thị trong Data Quality |

Nếu OpenMetadata version của bạn không hỗ trợ `CustomDatabase`, thử đổi:

```powershell
python .\import_vimc_glossary_3_excels_with_dq.py --excel 'Business Glossary Demo.xlsx' --dq-service-type Postgres
```

Tuy nhiên với connector thật như Postgres/MySQL, server có thể yêu cầu `connection` đầy đủ. Khi đó cần dùng service đã cấu hình sẵn hoặc chỉnh payload trong hàm `upsert_database_service()` theo connector thực tế.

## 6. Metric Data Quality mặc định trong code

Chương trình tạo và publish 10 test cases:

1. `vimc_row_count_positive`  
   Kiểm tra số dòng import > 0.

2. `vimc_column_source_file_exists`  
   Kiểm tra table kỹ thuật có cột `source_file`.

3. `vimc_column_sheet_name_exists`  
   Kiểm tra table kỹ thuật có cột `sheet_name`.

4. `vimc_column_display_name_exists`  
   Kiểm tra table kỹ thuật có cột `display_name`.

5. `vimc_column_definition_exists`  
   Kiểm tra table kỹ thuật có cột `definition`.

6. `vimc_column_source_key_exists`  
   Kiểm tra table kỹ thuật có cột `source_key`.

7. `vimc_display_name_not_null`  
   Kiểm tra mọi term import đều có tên hiển thị.

8. `vimc_definition_missing_count_eq_0`  
   Kiểm tra mọi term đều có định nghĩa/mô tả. Test này có thể fail nếu Excel còn dòng thiếu định nghĩa.

9. `vimc_source_key_unique`  
   Kiểm tra khóa truy vết `source_key` là duy nhất.

10. `vimc_display_name_length_1_255`  
    Kiểm tra độ dài tên hiển thị nằm trong 1-255 ký tự.

## 7. Cách thức hoạt động bên trong

### Bước 1: Đọc Excel

- Dò sheet theo `--sheet`, mặc định `Logic Tổng hợp`.
- Nếu không thấy sheet mặc định, tự chọn sheet có header giống glossary nhất.
- Dò header trong 30 dòng đầu.
- Map cột Excel vào các field logic như `code`, `group`, `indicator`, `definition`, `source_name`, `calculation`, v.v.

### Bước 2: Import Glossary

- Tạo/cập nhật Glossary `VIMC_Business_Glossary`.
- Tạo term cấp file Excel.
- Tạo group terms theo cột nhóm/parent.
- Tạo business terms theo từng dòng dữ liệu.
- Description của từng term chứa file nguồn, sheet, dòng Excel, thứ tự import và các cột có dữ liệu.

### Bước 3: Tạo Database kỹ thuật cho Data Quality

- Tạo/cập nhật Database Service.
- Tạo/cập nhật Database với displayName `VIMC Technical term`.
- Tạo/cập nhật Schema `glossary_quality`.
- Tạo/cập nhật Table `imported_glossary_terms` với các cột kỹ thuật đại diện cho dữ liệu Excel đã import.

### Bước 4: Tính metric trong Python

Script đọc lại các dòng Excel hợp lệ và tính:

- tổng số dòng
- số display_name bị thiếu
- số definition bị thiếu
- số source_key bị trùng
- min/max độ dài display_name
- tình trạng tồn tại cột kỹ thuật

### Bước 5: Publish lên OpenMetadata Data Quality

- Tạo/cập nhật executable Test Suite.
- Tạo/cập nhật Test Case cho từng metric.
- Ghi Test Case Result với trạng thái `Success` hoặc `Failed`.

## 8. Kiểm tra trên UI

Sau khi chạy thật:

1. Vào `Govern → Glossary → VIMC_Business_Glossary` để kiểm tra terms.
2. Vào `Explore → Databases`, tìm display name `VIMC Technical term`.
3. Vào `Observability → Data Quality → Test Suites → VIMC_Technical_term_DQ` để xem test cases và kết quả.

## 9. Lỗi thường gặp

### Lỗi 401

Token sai hoặc hết hạn. Cập nhật lại:

```powershell
$env:OM_JWT_TOKEN='TOKEN_MOI'
```

### Lỗi 403

Token không đủ quyền tạo/sửa entity hoặc test case.

### Lỗi 400 ở Database Service

OpenMetadata server không chấp nhận `serviceType=CustomDatabase` hoặc yêu cầu connection config. Cách xử lý:

- thử `--dq-service-type Postgres`; hoặc
- tạo Database Service thật trên UI trước rồi chỉnh code để dùng service đó; hoặc
- bổ sung `connection` đúng schema connector trong hàm `upsert_database_service()`.

### Data Quality fail ở `vimc_definition_missing_count_eq_0`

Điều này có nghĩa file Excel có term thiếu định nghĩa/mô tả. Đây là fail nghiệp vụ bình thường, không phải lỗi code.
