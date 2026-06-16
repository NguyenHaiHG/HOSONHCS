# HƯỚNG DẪN HỆ THỐNG LƯU TRỮ TỔ - TƯƠNG TỰ FORM1

## Tổng quan
Đã cấu hình Form2 với hệ thống lưu trữ tổ tương tự như hệ thống Customers trong Form1:
1. ✅ cbdt1-5 chỉ giới hạn lựa chọn, không tự động chọn giá trị
2. ✅ Mỗi tổ được lưu thành 1 file JSON riêng trong folder "To"
3. ✅ Tự động load dữ liệu tổ khi khởi động Form2
4. ✅ Tự động lưu khi bấm btn03to hoặc btntaoto
5. ✅ Tự động xóa file JSON khi xóa tổ trong dgv2

## Cấu trúc File JSON

### Folder lưu trữ
```
D:\C#\HOSONHCS\bin\Debug\To\
```

### Tên file
Format: `{TenToTruong}_{Xa}_{Thon}.json`

Ví dụ:
- `NguyenVanA_XinMan_Lung.json`
- `TranThiB_DongVan_Pac.json`

Nếu trùng tên: `NguyenVanA_XinMan_Lung_1.json`, `_2.json`, ...

### Nội dung file JSON
```json
{
  "Pgd": "Mèo Vạc",
  "Xa": "Xín Mần",
  "Thon": "Lũng",
  "Totruong": "Nguyễn Văn A",
  "Chuongtrinh": "Hộ nghèo",
  "NgayXuat": "2024-12-20T10:30:00",
  "SoThanhVien": 5,

  "Kh1": "Nguyễn Văn B",
  "Kh2": "Trần Thị C",
  "Kh3": "Lê Văn D",
  "Kh4": "Phạm Thị E",
  "Kh5": "Hoàng Văn F",

  "Tien1": "10.000.000",
  "Tien2": "15.000.000",
  "Tien3": "12.000.000",
  "Tien4": "18.000.000",
  "Tien5": "20.000.000",

  "Md1": "Nông nghiệp",
  "Md2": "Chăn nuôi",
  "Md3": "Thương mại",
  "Md4": "Dịch vụ",
  "Md5": "Sản xuất",

  "Time1": "36 tháng",
  "Time2": "48 tháng",
  "Time3": "60 tháng",
  "Time4": "36 tháng",
  "Time5": "48 tháng",

  "Dt1": "Hộ nghèo",
  "Dt2": "Hộ nghèo",
  "Dt3": "Hộ nghèo",
  "Dt4": "Hộ nghèo",
  "Dt5": "Hộ nghèo",

  "_fileName": "NguyenVanA_XinMan_Lung.json"
}
```

## Class ToData (HOSONHCS\ToData.cs)

```csharp
public class ToData
{
    // Thông tin chung (7 fields)
    public string Pgd { get; set; }
    public string Xa { get; set; }
    public string Thon { get; set; }
    public string Totruong { get; set; }
    public string Chuongtrinh { get; set; }
    public DateTime NgayXuat { get; set; }
    public int SoThanhVien { get; set; }

    // Thông tin 5 tổ viên (25 fields)
    public string Kh1...Kh5 { get; set; }      // Họ tên
    public string Tien1...Tien5 { get; set; }   // Số tiền
    public string Md1...Md5 { get; set; }       // Phương án
    public string Time1...Time5 { get; set; }   // Thời hạn
    public string Dt1...Dt5 { get; set; }       // Đối tượng

    // Metadata
    public string _fileName { get; set; }
}
```

## Các phương thức mới trong Form2

### 1. Quản lý file
```csharp
private string GetToFolderPath()           // Đường dẫn folder "To"
private void EnsureToFolder()               // Tạo folder nếu chưa tồn tại
private void LoadToFromFiles()              // Load tất cả tổ từ JSON
private void SaveToDataToFile(ToData)       // Lưu tổ vào file JSON
private void DeleteToDataFile(ToData)       // Xóa file JSON của tổ
```

### 2. Chuyển đổi dữ liệu
```csharp
private ExportHistory ConvertToDataToExportHistory(ToData)
private ToData ConvertExportHistoryToToData(ExportHistory)
```

### 3. Utility
```csharp
private string MakeFileSystemSafeForTo(string)  // Tạo tên file an toàn
```

## Luồng hoạt động

### A. Khởi động Form2
```
Form2_Load()
    ↓
LoadXinManData()        // Load dữ liệu PGD/Xã/Thôn
    ↓
LoadToFromFiles()       // Load tất cả tổ từ folder "To"
    ├─ Đọc tất cả file *.json
    ├─ Deserialize thành ToData
    ├─ Convert thành ExportHistory
    └─ Add vào exportHistories (hiển thị trong dgv2)
    ↓
LoadFormState()         // Load trạng thái form
```

### B. Tạo tổ mới (btn03to - chế độ Add)
```
Btn03to_Click()
    ├─ Thu thập dữ liệu từ form
    ├─ Tạo ExportHistory mới
    ├─ Add vào exportHistories
    ├─ Convert thành ToData
    ├─ SaveToDataToFile()
    │   ├─ Tạo tên file: {Totruong}_{Xa}_{Thon}.json
    │   ├─ Serialize thành JSON
    │   └─ Lưu vào folder "To"
    └─ Hiển thị thông báo
```

### C. Cập nhật tổ (btn03to - chế độ Edit)
```
Btn03to_Click()
    ├─ Cập nhật currentEditingHistory
    ├─ dgv2.Refresh()
    ├─ Convert thành ToData
    ├─ SaveToDataToFile()
    │   ├─ Lấy _fileName từ ToData
    │   ├─ Ghi đè lên file cũ
    │   └─ Lưu JSON
    └─ Hiển thị thông báo
```

### D. Tạo tổ từ tổ đã có (btntaoto)
```
BtnTaoTo_Click()
    ├─ Thu thập dữ liệu từ form
    ├─ Tạo ExportHistory mới
    ├─ Add vào exportHistories
    ├─ Convert thành ToData (không có _fileName)
    ├─ SaveToDataToFile()
    │   ├─ Tạo tên file mới
    │   └─ Lưu JSON
    └─ Hiển thị thông báo
```

### E. Xóa tổ (btnxoa)
```
BtnXoa_Click()
    ├─ Lấy danh sách ExportHistory đã chọn
    ├─ Foreach ExportHistory:
    │   ├─ Convert thành ToData
    │   ├─ DeleteToDataFile()
    │   │   └─ Xóa file JSON
    │   └─ Remove khỏi exportHistories
    └─ Reset chế độ Edit nếu cần
```

### F. Click vào dgv2 để edit
```
Dgv2_CellClick()
    ├─ Lấy ExportHistory từ row
    ├─ Set isEditMode = true
    ├─ currentEditingHistory = history
    ├─ Load TOÀN BỘ thông tin lên form
    │   ├─ Thông tin chung
    │   ├─ Họ tên 5 tổ viên
    │   ├─ Số tiền
    │   ├─ Phương án
    │   ├─ Thời hạn
    │   └─ Đối tượng
    ├─ btn03to.Text = "Cập nhật tổ"
    └─ Hiển thị thông báo
```

## So sánh với Form1

| Tính năng | Form1 (Customers) | Form2 (To) |
|-----------|-------------------|------------|
| **Folder lưu trữ** | `/Customers/` | `/To/` |
| **Class dữ liệu** | `Customer` | `ToData` |
| **Tên file** | `{Hoten}.json` | `{Totruong}_{Xa}_{Thon}.json` |
| **Load khi khởi động** | `LoadCustomersFromFiles()` | `LoadToFromFiles()` |
| **Lưu file** | `SaveCustomerToFile()` | `SaveToDataToFile()` |
| **Xóa file** | `DeleteCustomerFiles()` | `DeleteToDataFile()` |
| **DataGridView** | `dgv` (Customers) | `dgv2` (ExportHistory) |
| **Binding List** | `BindingList<Customer>` | `BindingList<ExportHistory>` |

## Sửa đổi cbdt1-5

### Trước (tự động chọn):
```csharp
if (doiTuongList.Contains(currentValues[0]))
    cbdt1.Text = currentValues[0];
else if (doiTuongList.Count > 0)
    cbdt1.SelectedIndex = 0;  // ❌ Tự động chọn
```

### Sau (chỉ giới hạn):
```csharp
if (doiTuongList.Contains(currentValues[0]))
    cbdt1.Text = currentValues[0];
else
    cbdt1.Text = "";  // ✅ Để trống, người dùng tự chọn
```

## File đã thay đổi

### 1. **HOSONHCS\ToData.cs** (MỚI)
- Class lưu trữ thông tin tổ
- 32 properties + constructor

### 2. **HOSONHCS\Form2.cs**
- Thêm biến: `savedToList`
- Thêm 8 phương thức quản lý file:
  - `GetToFolderPath()`, `EnsureToFolder()`, `LoadToFromFiles()`
  - `SaveToDataToFile()`, `DeleteToDataFile()`
  - `ConvertToDataToExportHistory()`, `ConvertExportHistoryToToData()`
  - `MakeFileSystemSafeForTo()`
- Cập nhật `Form2_Load()`: Load dữ liệu tổ
- Cập nhật `Btn03to_Click()`: Lưu vào file JSON
- Cập nhật `BtnTaoTo_Click()`: Lưu vào file JSON
- Cập nhật `BtnXoa_Click()`: Xóa file JSON
- Cập nhật `UpdateDoiTuongComboBoxes()`: Không tự động chọn

### 3. **HOSONHCS\Models.cs**
- Xóa class ToData (đã chuyển sang ToData.cs)

## Kiểm tra
✅ Build thành công  
✅ cbdt1-5 chỉ giới hạn, không tự động chọn  
✅ File JSON được tạo khi bấm btn03to/btntaoto  
✅ File JSON được cập nhật khi edit  
✅ File JSON được xóa khi xóa tổ  
✅ Dữ liệu được load lại khi khởi động Form2  
✅ Không động chạm đến Form1  

## Lưu ý sử dụng
- Folder "To" được tạo tự động ở `bin\Debug\To\`
- Mỗi tổ = 1 file JSON riêng
- File được tự động lưu khi:
  - Bấm btn03to (chế độ Add hoặc Edit)
  - Bấm btntaoto
- File được tự động xóa khi:
  - Bấm btnxoa và chọn dòng trong dgv2
- File được tự động load khi:
  - Khởi động Form2
- cbdt1-5 chỉ giới hạn lựa chọn dựa trên cbctr
  - Không tự động chọn giá trị
  - Người dùng phải click và chọn

## Demo

### Tạo tổ mới:
1. Nhập thông tin tổ
2. Chọn chương trình (cbctr) → cbdt1-5 hiện danh sách
3. Chọn đối tượng cho từng tổ viên (cbdt1-5)
4. Bấm "Xuất tổ" (btn03to)
5. → File `NguyenVanA_XinMan_Lung.json` được tạo

### Chỉnh sửa tổ:
1. Click vào dòng trong dgv2
2. Thông tin hiện đầy đủ trên form
3. Chỉnh sửa
4. Bấm "Cập nhật tổ" (btn03to)
5. → File JSON được cập nhật

### Sao chép tổ:
1. Click vào dòng trong dgv2
2. Chỉnh sửa (ví dụ: đổi tên tổ trưởng)
3. Bấm "Tạo tổ mới" (btntaoto)
4. → File JSON mới được tạo

### Xóa tổ:
1. Chọn dòng trong dgv2
2. Bấm "Xóa" (btnxoa)
3. → File JSON bị xóa
