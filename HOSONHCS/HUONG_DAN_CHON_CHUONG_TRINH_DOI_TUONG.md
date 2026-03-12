# HƯỚNG DẪN CHỨC NĂNG CHỌN CHƯƠNG TRÌNH - ĐỐI TƯỢNG

## Tổng quan
Đã cấu hình Form2 để tự động cập nhật danh sách đối tượng (cbdt1-5) dựa trên chương trình (cbctr) được chọn.

## Logic ánh xạ Chương trình → Đối tượng

| STT | Chương trình (cbctr) | Đối tượng cho phép (cbdt1-5) |
|-----|----------------------|------------------------------|
| 1 | **Hộ nghèo** | Chỉ cho chọn: **Hộ nghèo** |
| 2 | **Hộ cận nghèo** | Chỉ cho chọn: **Hộ cận nghèo** |
| 3 | **Hộ mới thoát nghèo** | Chỉ cho chọn: **Hộ mới thoát nghèo** |
| 4 | **Hộ gia đình Sản xuất kinh doanh tại vùng khó khăn** | Chỉ cho chọn: **Hộ GĐ SXKD VKK** |
| 5 | **Hỗ trợ tạo việc làm duy trì và mở rộng việc làm** | Cho chọn 2 tùy chọn:<br>- **Người lao động**<br>- **NLĐ là người DTTS** |
| 6 | **Cấp nước sạch và vệ sinh môi trường nông thôn** | Chỉ cho chọn: **HGĐ cư trú tại VNT** |

## Cách hoạt động

### 1. Khi chọn chương trình (cbctr):
- Event `Cbctr_SelectedIndexChanged` được kích hoạt
- Phân tích tên chương trình để xác định nhóm đối tượng phù hợp
- Gọi `UpdateDoiTuongComboBoxes()` để cập nhật cbdt1-5

### 2. Cập nhật cbdt1-5:
- Xóa tất cả items cũ
- Thêm items mới dựa trên ánh xạ
- Cố gắng giữ giá trị cũ nếu nó vẫn hợp lệ
- Nếu giá trị cũ không còn hợp lệ, chọn tùy chọn đầu tiên

### 3. Khi không chọn chương trình:
- Gọi `ResetDoiTuongComboBoxes()` để xóa sạch cbdt1-5

## Các phương thức mới

### `Cbctr_SelectedIndexChanged()`
- Event handler cho cbctr
- Xác định danh sách đối tượng dựa trên chương trình
- Sử dụng `IndexOf()` để tìm kiếm không phân biệt hoa thường

### `UpdateDoiTuongComboBoxes(List<string> doiTuongList)`
- Nhận danh sách đối tượng cần hiển thị
- Cập nhật cbdt1-5 với danh sách mới
- Cố gắng giữ giá trị hiện tại nếu hợp lệ

### `ResetDoiTuongComboBoxes()`
- Xóa sạch cbdt1-5 (Items và Text)
- Được gọi khi không chọn chương trình

## Ví dụ sử dụng

### Ví dụ 1: Chọn "Hộ nghèo"
```
Bước 1: Chọn cbctr = "Hộ nghèo"
Bước 2: cbdt1-5 tự động chỉ hiển thị 1 tùy chọn: "Hộ nghèo"
Bước 3: Chọn "Hộ nghèo" cho cbdt1, cbdt2, cbdt3, cbdt4, cbdt5
```

### Ví dụ 2: Chọn "Hỗ trợ tạo việc làm"
```
Bước 1: Chọn cbctr = "Hỗ trợ tạo việc làm duy trì và mở rộng việc làm"
Bước 2: cbdt1-5 tự động hiển thị 2 tùy chọn:
        - "Người lao động"
        - "NLĐ là người DTTS"
Bước 3: Có thể chọn khác nhau cho mỗi cbdt (ví dụ cbdt1="Người lao động", cbdt2="NLĐ là người DTTS")
```

### Ví dụ 3: Đổi chương trình giữa chừng
```
Bước 1: Chọn cbctr = "Hộ nghèo"
Bước 2: cbdt1 = "Hộ nghèo", cbdt2 = "Hộ nghèo"
Bước 3: Đổi cbctr = "Hộ cận nghèo"
Bước 4: cbdt1-5 tự động cập nhật thành "Hộ cận nghèo"
        (Giá trị cũ "Hộ nghèo" không còn hợp lệ nên được thay thế)
```

## Tính năng thông minh

### 1. Giữ giá trị cũ nếu hợp lệ
```csharp
// Nếu đang chọn "Người lao động" trong chương trình "Hỗ trợ tạo việc làm"
// Và đổi sang chương trình "Hộ nghèo"
// → cbdt1-5 sẽ tự động chuyển sang "Hộ nghèo" (vì "Người lao động" không hợp lệ)
```

### 2. Tìm kiếm linh hoạt
- Sử dụng `IndexOf()` với `StringComparison.OrdinalIgnoreCase`
- Không cần khớp chính xác tên chương trình
- Ví dụ: "Hộ nghèo ABC" vẫn được nhận dạng là "Hộ nghèo"

### 3. Xử lý ngoại lệ
- Phân biệt "Hộ nghèo" với "Hộ cận nghèo" và "Hộ mới thoát nghèo"
- Sử dụng điều kiện phủ định để tránh nhầm lẫn:
  ```csharp
  if (selectedChuongTrinh.IndexOf("Hộ nghèo", ...) >= 0 &&
      selectedChuongTrinh.IndexOf("cận", ...) < 0 &&
      selectedChuongTrinh.IndexOf("thoát", ...) < 0)
  ```

## Lưu ý kỹ thuật

### Thứ tự kiểm tra
1. "Hộ nghèo" (loại trừ "cận" và "thoát")
2. "Hộ cận nghèo"
3. "Hộ mới thoát nghèo"
4. "Sản xuất kinh doanh" hoặc "SXKD"
5. "việc làm"
6. "nước sạch" hoặc "vệ sinh môi trường"

### Xử lý khi load dữ liệu tổ
- Khi click vào dgv2 để load tổ, giá trị cbctr và cbdt1-5 được load
- Event `Cbctr_SelectedIndexChanged` tự động kích hoạt
- cbdt1-5 sẽ được cập nhật để chỉ hiển thị các tùy chọn hợp lệ
- Giá trị đã lưu sẽ được giữ lại (vì nó hợp lệ với chương trình)

## File được thay đổi
- `HOSONHCS\Form2.cs`:
  - Đăng ký event handler: `cbctr.SelectedIndexChanged += Cbctr_SelectedIndexChanged`
  - Thêm 3 phương thức mới:
    - `Cbctr_SelectedIndexChanged()`: Xử lý khi chọn chương trình
    - `UpdateDoiTuongComboBoxes(List<string> doiTuongList)`: Cập nhật cbdt1-5
    - `ResetDoiTuongComboBoxes()`: Reset cbdt1-5

## Kiểm tra
Build thành công ✓
Event handler đã được đăng ký ✓
Logic ánh xạ đã được implement ✓
