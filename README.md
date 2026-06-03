# VIMC Excel Glossary Importer cho OpenMetadata

Script `import_vimc_glossary_3_excels.py` dùng để đọc file Excel Business Glossary và đẩy lên OpenMetadata dưới dạng Glossary/Glossary Terms.

Bản này đã được chỉnh để **giữ nguyên cách chạy cũ** nhưng linh hoạt hơn:

- Không còn hard-code vị trí cột cố định như B = Code, C = Group.
- Tự dò header tiếng Anh/tiếng Việt trong các dòng đầu của Excel.
- Tự nhận diện các cột phổ biến: `Code/Mã`, `Group/Nhóm chỉ tiêu/Parent`, `Indicator/Chỉ tiêu`, `Definition/Định nghĩa`, `Unit/ĐVT`, `Application`, `API/File`, `Source`, `Formula`, `Note`, v.v.
- Nếu không có sheet `Logic Tổng hợp`, script tự chọn sheet có cấu trúc giống glossary nhất.
- Mặc định thêm prefix thứ tự `G00001_...`, `T00001_...` vào **entity name** để OpenMetadata hiển thị/sắp xếp term theo đúng thứ tự trong Excel. `displayName` vẫn giữ tên gốc để người dùng đọc dễ.

## 1. Cài đặt môi trường

Yêu cầu:

- Python 3.9 trở lên
- Kết nối tới OpenMetadata
- Token có quyền tạo/sửa Glossary và Glossary Term

Cài thư viện:

```bash
pip install requests openpyxl
```

## 2. Chuẩn bị biến môi trường

### PowerShell trên Windows

```powershell
$env:OPENMETADATA_HOST='http://192.168.74.12:30085'
$env:OM_JWT_TOKEN='PASTE_TOKEN_CUA_BAN'
```

### Bash/Linux/macOS

```bash
export OPENMETADATA_HOST='http://192.168.74.12:30085'
export OM_JWT_TOKEN='PASTE_TOKEN_CUA_BAN'
```

`OPENMETADATA_HOST` có thể truyền ở dạng:

- `http://host:port`
- `http://host:port/api`
- `http://host:port/api/v1`

Script sẽ tự chuẩn hóa về `/api/v1`.

## 3. Kiểm tra trước bằng dry-run

Nên chạy dry-run trước để xem script đọc đúng sheet, đúng header, đúng số term hay chưa.

```bash
python import_vimc_glossary_3_excels.py --dry-run --excel "Business Glossary Demo.xlsx"
```

Kết quả dry-run sẽ in ra:

- Glossary sẽ tạo/cập nhật
- Sheet được đọc
- Header rows và data start row
- Mapping cột Excel sang field logic
- Cấu trúc parent/group/term sẽ được import
- Tổng số group terms và business/indicator terms

## 4. Import một file Excel lên OpenMetadata

```bash
python import_vimc_glossary_3_excels.py --excel "Business Glossary Demo.xlsx"
```

Nếu không truyền `--excel`, script vẫn dùng default cũ:

```bash
python import_vimc_glossary_3_excels.py
```

Default file hiện tại là:

```text
Business Glossary Demo.xlsx
```

## 5. Import nhiều file Excel

```bash
python import_vimc_glossary_3_excels.py \
  --excel "Business Glossary 1.xlsx" \
  --excel "Business Glossary 2.xlsx" \
  --excel "Business Glossary 3.xlsx" \
  --glossary VIMC_Business_Glossary
```

Mỗi file sẽ được tạo thành một parent term riêng dưới Glossary.

## 6. Import file Excel có tên sheet khác

Mặc định script tìm sheet `Logic Tổng hợp`.

Nếu file của bạn có sheet khác, có thể truyền:

```bash
python import_vimc_glossary_3_excels.py --excel "my_glossary.xlsx" --sheet "Glossary"
```

Nếu không truyền `--sheet` và file không có `Logic Tổng hợp`, script sẽ tự chọn sheet có header giống glossary nhất.

## 7. Import tất cả sheets trong workbook

```bash
python import_vimc_glossary_3_excels.py --excel "my_glossary.xlsx" --all-sheets
```

Mỗi sheet sẽ được tạo thành một parent term riêng theo format:

```text
<Tên file>_<Tên sheet>
```

## 8. Format hierarchy trên OpenMetadata

Mặc định script tạo cấu trúc:

```text
VIMC_Business_Glossary
└── Business_Glossary_Demo
    ├── G00001_Chung
    │   ├── T00001_GEN0001
    │   ├── T00002_GEN0002
    │   └── ...
    ├── G00002_San_luong
    │   ├── T00011_VOL0001
    │   └── ...
    └── ...
```

Trong đó:

- Level 1: Glossary trên OpenMetadata
- Level 2: Parent term đại diện cho file Excel/sheet
- Level 3: Group/Parent term từ cột `Group`, `Nhóm chỉ tiêu`, `Parent`, `Category`, v.v.
- Level 4: Business term/indicator từ cột `Code`, `Mã`, `Name`, `Term`, `Indicator`, `Chỉ tiêu`, v.v.

## 9. Giữ tên term gốc, không thêm prefix thứ tự

Mặc định `--name-mode ordered` để đảm bảo term sort theo thứ tự Excel.

Nếu muốn giữ entity name như code gốc (`GEN0001`, `VOL0001`, ...), chạy:

```bash
python import_vimc_glossary_3_excels.py --excel "Business Glossary Demo.xlsx" --name-mode original
```

Lưu ý: với `original`, thứ tự hiển thị trên OpenMetadata có thể phụ thuộc cách OpenMetadata sort entity name.

## 10. File Excel cần có format như thế nào?

Script không yêu cầu thứ tự cột cố định. Chỉ cần file có một vài cột nhận diện được, ví dụ:

| Ý nghĩa     | Tên cột có thể dùng                                                                 |
| ----------- | ----------------------------------------------------------------------------------- |
| Mã term     | `Code`, `Mã`, `Mã chỉ tiêu`, `Term Code`, `ID`                                      |
| Tên term    | `Name`, `Term`, `Glossary Term`, `Business Term`, `Chỉ tiêu`, `Indicator`, `Metric` |
| Nhóm/parent | `Group`, `Nhóm chỉ tiêu`, `Parent`, `Category`, `Domain`                            |
| Mô tả       | `Definition`, `Định nghĩa`, `Description`, `Mô tả`, `Business Definition`           |
| Đơn vị      | `Unit`, `ĐVT`, `Đơn vị tính`, `UOM`                                                 |
| Công thức   | `Calculation`, `Formula`, `Công thức`, `Logic`                                      |
| Ghi chú     | `Note`, `Ghi chú`, `Remark`, `Comment`                                              |
| Synonyms    | `Synonyms`, `Alias`, `Từ đồng nghĩa`                                                |

Nếu header đặc biệt khó tự dò, dùng:

```bash
python import_vimc_glossary_3_excels.py \
  --excel "my_file.xlsx" \
  --header-row 3 \
  --data-start-row 5
```

## 11. Các option quan trọng

```text
--excel              Đường dẫn file Excel. Có thể dùng nhiều lần.
--glossary           Tên OpenMetadata Glossary. Default: VIMC_Business_Glossary.
--sheet              Tên sheet cần import. Default: Logic Tổng hợp.
--all-sheets         Import tất cả sheets.
--host               OpenMetadata host. Có thể dùng OPENMETADATA_HOST thay thế.
--token              JWT/PAT token. Có thể dùng OM_JWT_TOKEN thay thế.
--dry-run            Chỉ đọc file và in kế hoạch import, không gọi API.
--sleep              Nghỉ giữa các API call, ví dụ 0.05.
--header-row         Ép dòng header nếu tự dò không đúng.
--data-start-row     Ép dòng bắt đầu dữ liệu.
--name-mode          ordered hoặc original. Default: ordered.
--ungrouped-label    Tên group khi dòng không có Group/Parent. Default: Khác.
--timeout            Timeout mỗi API call. Default: 45 giây.
--retries            Retry cho lỗi 429/5xx. Default: 2.
```

## 12. Lỗi thường gặp

### 401 Unauthorized

Token thiếu hoặc sai.

Kiểm tra lại:

```bash
echo $OM_JWT_TOKEN
```

Hoặc PowerShell:

```powershell
$env:OM_JWT_TOKEN
```

### 403 Forbidden

Token không đủ quyền tạo/sửa Glossary hoặc Glossary Terms.

### 400 Bad Request

Thường do:

- Tên term bị trùng trong cùng parent.
- Parent FQN không hợp lệ.
- Version OpenMetadata có thay đổi schema.

Script đã tự xử lý duplicate bằng suffix `_DUP2`, `_DUP3`, ... trong cùng parent.

### Không tìm thấy dữ liệu để import

Chạy dry-run và xem mapping cột. Nếu header bị dò sai, truyền `--header-row` và `--data-start-row`.

## 13. Quy trình khuyến nghị

1. Đặt script và Excel cùng thư mục.
2. Chạy dry-run.
3. Kiểm tra mapping cột và tổng số term.
4. Chạy import thật.
5. Vào OpenMetadata UI: `Govern → Glossary → VIMC_Business_Glossary`.
