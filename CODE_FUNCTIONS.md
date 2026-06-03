# Mô tả chức năng code `import_vimc_glossary_3_excels.py`

## 1. Mục tiêu

Script dùng để import dữ liệu Business Glossary từ Excel lên OpenMetadata. Code đã được thiết kế lại để không phụ thuộc vào thứ tự cột cố định, phù hợp hơn khi cần import nhiều file Excel khác nhau.

## 2. Luồng xử lý tổng quát

```text
Đọc tham số CLI
→ Kiểm tra file Excel
→ Kết nối OpenMetadata
→ Tạo/cập nhật Glossary
→ Mở từng workbook
→ Chọn sheet cần import
→ Tự dò header và mapping cột
→ Duyệt từng dòng dữ liệu theo đúng thứ tự Excel
→ Tạo parent term cho file/sheet
→ Tạo group terms
→ Tạo business/indicator terms
→ In tổng kết
```

## 3. Cấu trúc hierarchy tạo trên OpenMetadata

```text
<Glossary>
└── <Parent term đại diện file hoặc file_sheet>
    ├── <Group term 1>
    │   ├── <Business term 1>
    │   ├── <Business term 2>
    │   └── ...
    ├── <Group term 2>
    │   └── ...
    └── ...
```

Ví dụ:

```text
VIMC_Business_Glossary
└── Business_Glossary_Demo
    ├── G00001_Chung
    │   ├── T00001_GEN0001
    │   └── T00002_GEN0002
    └── G00002_San_luong
        ├── T00011_VOL0001
        └── T00012_VOL0002
```

## 4. Các nhóm hàm chính

## 4.1. Text helpers

### `clean(value)`

Chuẩn hóa giá trị từ Excel/API thành string sạch:

- `None` → `""`
- Boolean → `"1"` hoặc `"0"`
- Float nguyên → bỏ `.0`
- Chuẩn hóa xuống dòng `\r\n`, `\r` thành `\n`
- Trim khoảng trắng đầu/cuối

### `strip_accents(text)`

Bỏ dấu tiếng Việt để tạo entity name an toàn cho OpenMetadata.

Ví dụ:

```text
"Sản lượng" → "San luong"
"Đơn vị" → "Don vi"
```

### `normalize_text(text)`

Chuẩn hóa text để so khớp header:

- Bỏ dấu
- Lowercase
- Đổi ký tự đặc biệt thành khoảng trắng
- Gom nhiều khoảng trắng thành một

### `is_meaningful(value)`

Kiểm tra một ô có phải giá trị hợp lệ hay không. Bỏ qua các giá trị rỗng hoặc lỗi Excel như:

```text
#N/A, #VALUE!, #REF!, N/A, NULL, NONE, -, ---
```

### `safe_name(text, fallback, max_len)`

Tạo OpenMetadata entity name an toàn:

- Bỏ dấu tiếng Việt
- Chỉ giữ chữ, số, `_`, `-`
- Không cho bắt đầu bằng số
- Cắt độ dài tối đa

### `split_synonyms(text)`

Tách synonyms theo dấu phẩy, chấm phẩy, xuống dòng hoặc `|`.

## 4.2. Khai báo alias cột

Code dùng `FIELD_SPECS` để định nghĩa các field logic và danh sách tên cột có thể nhận diện.

Ví dụ:

```python
FieldSpec("code", "Mã", ("code", "ma", "ma chi tieu", "term code", "id"), 4)
FieldSpec("group", "Nhóm/Parent", ("group", "nhom chi tieu", "parent", "category"), 4)
FieldSpec("definition", "Định nghĩa/Mô tả", ("definition", "dinh nghia", "description", "mo ta"), 4)
```

Nhờ đó Excel không cần đúng thứ tự cột. Chỉ cần tên cột gần đúng là script có thể map được.

## 4.3. Dò header và mapping cột

### `field_match_score(header_text, field)`

Tính điểm khớp giữa text header và một field logic.

Ưu tiên:

1. Khớp chính xác.
2. Header chứa nguyên cụm alias.
3. Header chứa alias một phần.
4. Alias dài hơn được ưu tiên để tránh nhầm `Nhóm chỉ tiêu` thành `Chỉ tiêu`.

### `row_alias_score(ws, row_idx, max_col)`

Chấm điểm một dòng có khả năng là header.

Điểm đặc biệt:

- Dòng hiện tại được ưu tiên mạnh hơn dòng bên dưới.
- Vẫn xét dòng bên dưới để hỗ trợ file có 2 dòng header Anh/Việt.
- Các field quan trọng như `code`, `group`, `indicator`, `definition` được cộng thêm điểm.

### `detect_header(ws, forced_header_row, forced_data_start_row)`

Tự xác định:

- `header_start`
- `header_end`
- `data_start`

Với file mẫu VIMC:

```text
header_start = 3
header_end   = 4
data_start   = 5
```

Nếu tự dò không đúng, người dùng có thể truyền:

```bash
--header-row 3 --data-start-row 5
```

### `build_column_map(ws, header_start, header_end)`

Tạo mapping:

```python
{
  "code": 2,
  "group": 3,
  "indicator": 7,
  "definition": 9,
  ...
}
```

Mapping này thay thế hoàn toàn cách hard-code cột cố định trước đây.

## 4.4. Đọc dòng dữ liệu

### `row_has_term(row_values, field_to_col)`

Một dòng được import nếu có ít nhất một tín hiệu term hợp lệ:

- `code`
- `term_name`
- `display_name`
- `indicator`
- `indicator_new`
- `definition`

Nếu Excel quá đơn giản và header không nhận diện được, hàm fallback sẽ import dòng có ít nhất 2 ô có dữ liệu hợp lệ.

### `collect_sheet_rows(...)`

Thực hiện:

1. Dò header.
2. Tạo column mapping.
3. Duyệt từ dòng data start tới cuối sheet.
4. Bỏ qua dòng rỗng/dòng lỗi.
5. Trả về danh sách `(excel_row, row_values)`.

## 4.5. Build description cho term

### `build_description(...)`

Tạo Markdown description cho mỗi Glossary Term.

Nội dung gồm:

- Nguồn file
- Sheet
- Dòng Excel
- Thứ tự import
- Các field đã map như Mã, Nhóm, Chỉ tiêu, Định nghĩa, Công thức, Ghi chú
- Các cột bổ sung chưa map được nhưng có dữ liệu

Cách này giúp không mất thông tin khi Excel có thêm cột mới.

## 4.6. OpenMetadata API client

### `OpenMetadataClient.normalize_api_base(host)`

Chuẩn hóa host về format:

```text
http://host:port/api/v1
```

Người dùng có thể nhập:

```text
http://host:port
http://host:port/api
http://host:port/api/v1
```

### `OpenMetadataClient.request(method, path, json_body)`

Gửi request tới OpenMetadata, có:

- Timeout
- Retry cho lỗi `429` và `5xx`
- Thông báo gợi ý khi gặp `401`, `403`, `400`, `404`

### `upsert_glossary(...)`

Tạo/cập nhật Glossary qua endpoint:

```text
PUT /api/v1/glossaries
```

Payload chính:

```json
{
  "name": "VIMC_Business_Glossary",
  "displayName": "VIMC_Business_Glossary",
  "description": "...",
  "mutuallyExclusive": false
}
```

### `upsert_term(...)`

Tạo/cập nhật Glossary Term qua endpoint:

```text
PUT /api/v1/glossaryTerms
```

Payload chính:

```json
{
  "name": "T00001_GEN0001",
  "glossary": "VIMC_Business_Glossary",
  "displayName": "TH",
  "description": "...",
  "parent": "VIMC_Business_Glossary.Business_Glossary_Demo.G00001_Chung",
  "synonyms": ["..."]
}
```

## 4.7. Tạo tên term đúng thứ tự

### `make_ordered_name(prefix, order, raw_name, fallback, name_mode)`

Nếu `--name-mode ordered`, tạo tên dạng:

```text
G00001_Chung
T00001_GEN0001
T00002_GEN0002
```

Mục đích: OpenMetadata thường sort theo name, nên prefix giúp term hiển thị theo đúng thứ tự dòng Excel.

Nếu `--name-mode original`, dùng tên gốc:

```text
Chung
GEN0001
GEN0002
```

### `make_unique(name, used, fallback_prefix)`

Đảm bảo không trùng entity name trong cùng parent.

Nếu trùng:

```text
GEN0001
GEN0001_DUP2
GEN0001_DUP3
```

## 4.8. Import từng sheet/file

### `import_excel_sheet(...)`

Xử lý một sheet:

1. Dò header và column mapping.
2. Tạo parent term cho file/sheet.
3. Duyệt từng dòng theo đúng thứ tự Excel.
4. Tạo group term khi gặp group mới.
5. Tạo business/indicator term dưới group tương ứng.
6. Tổng kết số group và term.

### `import_excel_file(...)`

Xử lý một workbook:

- Nếu có sheet được yêu cầu, import sheet đó.
- Nếu không có sheet mặc định `Logic Tổng hợp`, tự chọn sheet phù hợp nhất.
- Nếu dùng `--all-sheets`, import toàn bộ sheets.

## 5. Vì sao bản này import được nhiều file Excel hơn?

Bản cũ phụ thuộc vào thứ tự cột cố định, ví dụ:

```python
"code": 2,
"group": 3,
"indicator": 7,
"definition": 9,
```

Bản mới chuyển sang cơ chế:

```text
Tìm header → nhận diện alias → map field logic → đọc dòng
```

Vì vậy file Excel có thể đổi thứ tự cột hoặc dùng header tiếng Việt/tiếng Anh khác nhau mà vẫn import được.

## 6. Giới hạn hiện tại

- Script xử lý tốt nhất với file Excel dạng bảng, có header rõ ràng.
- Nếu file có nhiều bảng trong cùng một sheet, nên tách riêng hoặc dùng `--header-row` và `--data-start-row`.
- OpenMetadata không có trường sort order riêng cho Glossary Term trong payload import này; do đó script dùng prefix `G/T + số thứ tự` để đảm bảo thứ tự hiển thị ổn định.
- File `.xls` cũ chưa được hỗ trợ trực tiếp. Nên lưu lại thành `.xlsx`.

## 7. Gợi ý mở rộng sau này

- Thêm file cấu hình YAML/JSON để tự khai báo alias cột theo từng dự án.
- Thêm option import related terms/tags/owners/reviewers.
- Thêm log ra file CSV để audit dòng nào thành công/thất bại.
- Thêm chế độ xóa hoặc archive term không còn trong Excel.
