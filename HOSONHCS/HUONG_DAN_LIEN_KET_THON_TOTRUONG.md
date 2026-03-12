# HƯỚNG DẪN LIÊN KẾT THÔN - TỔ TRƯỞNG (FORM2)

## Tổng quan
Đã cấu hình Form2 để tạo liên kết 2 chiều giữa cbThon2 và cbTotruong:
1. **Chọn Thôn** → Chỉ hiển thị tổ trưởng của thôn đó
2. **Chọn Tổ trưởng** → Tự động chọn thôn tương ứng

## Cơ chế hoạt động

### 1. Khi chọn Thôn (cbThon2)
```
User chọn thôn "Lũng"
    ↓
CbThon2_SelectedIndexChanged()
    ↓
Tìm village "Lũng" trong xinmanModel
    ↓
Xóa cbTotruong.Items
    ↓
Load CHỈ các tổ trưởng của thôn "Lũng"
    ├─ Nguyễn Văn A
    ├─ Trần Thị B
    └─ Lê Văn C
```

**Kết quả**: cbTotruong chỉ hiển thị tổ trưởng của thôn "Lũng"

### 2. Khi chọn Tổ trưởng (cbTotruong)
```
User chọn tổ trưởng "Nguyễn Văn A"
    ↓
CbTotruong_SelectedIndexChanged()
    ↓
Tìm thôn nào chứa "Nguyễn Văn A"
    ├─ Duyệt commune.villages
    ├─ Duyệt commune.associations.villages
    └─ Tìm thấy: thôn "Lũng"
    ↓
Tự động set cbThon2.Text = "Lũng"
```

**Kết quả**: cbThon2 tự động chọn thôn "Lũng"

## Tránh vòng lặp vô hạn

### Vấn đề
```
cbThon2 change → cbTotruong update → cbThon2 change → ...
                                      ↑_______________|
                                      (Vòng lặp vô hạn!)
```

### Giải pháp: Biến flag `isUpdatingThonTotruong`

```csharp
private bool isUpdatingThonTotruong = false;

// Trong CbThon2_SelectedIndexChanged:
if (isUpdatingThonTotruong) return;  // Bỏ qua nếu đang update

isUpdatingThonTotruong = true;
try {
    cbTotruong.Items.Clear();  // Update cbTotruong
    // ...
} finally {
    isUpdatingThonTotruong = false;
}

// Trong CbTotruong_SelectedIndexChanged:
if (isUpdatingThonTotruong) return;  // Bỏ qua nếu đang update

isUpdatingThonTotruong = true;
try {
    cbThon2.Text = foundVillage.name;  // Update cbThon2
} finally {
    isUpdatingThonTotruong = false;
}
```

## Các phương thức

### `CbThon2_SelectedIndexChanged()`
**Chức năng**: Load tổ trưởng của thôn được chọn

**Logic**:
1. Check flag `isUpdatingThonTotruong`
2. Lấy Xã và Thôn đã chọn
3. Tìm Village tương ứng (trong communes.villages hoặc associations.villages)
4. Set flag = true
5. Xóa cbTotruong.Items
6. Load groups của village vào cbTotruong
7. Reset flag = false

**Tìm kiếm thông minh**:
- Tìm trong `commune.villages` trước
- Nếu không thấy, tìm trong `commune.associations[].villages`

### `CbTotruong_SelectedIndexChanged()` (MỚI)
**Chức năng**: Tự động chọn thôn khi chọn tổ trưởng

**Logic**:
1. Check flag `isUpdatingThonTotruong`
2. Lấy Xã và Tổ trưởng đã chọn
3. Tìm Village chứa tổ trưởng này
   - Duyệt `commune.villages[].groups`
   - Duyệt `commune.associations[].villages[].groups`
4. Nếu tìm thấy village:
   - Set flag = true
   - Set `cbThon2.Text = foundVillage.name`
   - Reset flag = false

## Ví dụ thực tế

### Dữ liệu mẫu (xinman.json)
```json
{
  "pgd": "Mèo Vạc",
  "communes": [
    {
      "name": "Xín Mần",
      "villages": [
        {
          "name": "Lũng",
          "groups": ["Nguyễn Văn A", "Trần Thị B"]
        },
        {
          "name": "Pắc",
          "groups": ["Lê Văn C", "Phạm Thị D"]
        }
      ]
    }
  ]
}
```

### Kịch bản 1: Chọn Thôn → Lọc Tổ trưởng
```
1. User chọn Xã: "Xín Mần"
2. User chọn Thôn: "Lũng"
   → cbTotruong chỉ hiển thị: ["Nguyễn Văn A", "Trần Thị B"]
   → KHÔNG hiển thị: "Lê Văn C", "Phạm Thị D" (vì thuộc thôn "Pắc")
```

### Kịch bản 2: Chọn Tổ trưởng → Tự động chọn Thôn
```
1. User chọn Xã: "Xín Mần"
2. User chọn Tổ trưởng: "Lê Văn C"
   → cbThon2 TỰ ĐỘNG chọn: "Pắc"
   → cbTotruong cập nhật danh sách: ["Lê Văn C", "Phạm Thị D"]
```

### Kịch bản 3: Đổi Thôn
```
1. Đang chọn: Thôn = "Lũng", Tổ trưởng = "Nguyễn Văn A"
2. User đổi Thôn thành: "Pắc"
   → cbTotruong cập nhật: ["Lê Văn C", "Phạm Thị D"]
   → cbTotruong.Text = "" (xóa tổ trưởng cũ vì không còn hợp lệ)
```

### Kịch bản 4: Đổi Tổ trưởng
```
1. Đang chọn: Thôn = "Lũng", Tổ trưởng = "Nguyễn Văn A"
2. User chọn Tổ trưởng: "Lê Văn C"
   → cbThon2 TỰ ĐỘNG đổi thành: "Pắc"
   → cbTotruong cập nhật danh sách: ["Lê Văn C", "Phạm Thị D"]
```

## Luồng dữ liệu

```
┌─────────────────────────────────────────────────────────┐
│                      xinman.json                         │
│  PGD → Commune → Village → Groups                       │
└─────────────────────────────────────────────────────────┘
                      ↓
        ┌─────────────┴─────────────┐
        ↓                           ↓
   cbXa2 (Xã)                  xinmanModel
        ↓                           ↓
   ┌────┴────┐               ┌──────┴──────┐
   ↓         ↓               ↓             ↓
cbThon2 ←→ cbTotruong   villages[]    groups[]
   │         │
   │         │ (Liên kết 2 chiều)
   │         │
   └────┬────┘
        ↓
   isUpdatingThonTotruong
   (Ngăn vòng lặp)
```

## So sánh với Form1

| Tính năng | Form1 | Form2 | Liên kết? |
|-----------|-------|-------|-----------|
| **ComboBox Thôn** | `cbthon` | `cbThon2` | ❌ Không |
| **ComboBox Tổ** | `cbto` | `cbTotruong` | ❌ Không |
| **Event handler** | Riêng Form1 | Riêng Form2 | ❌ Không |
| **Dữ liệu xinman** | Dùng chung | Dùng chung | ✅ Có |

**Lưu ý**: Chỉ dữ liệu xinman.json được dùng chung, còn lại Form1 và Form2 hoàn toàn độc lập.

## File đã thay đổi

### HOSONHCS\Form2.cs
**Thêm biến**:
- `private bool isUpdatingThonTotruong = false;`

**Đăng ký event**:
- `cbTotruong.SelectedIndexChanged += CbTotruong_SelectedIndexChanged;`

**Cập nhật phương thức**:
- `CbThon2_SelectedIndexChanged()`: Thêm flag check, tìm kiếm thông minh

**Thêm phương thức mới**:
- `CbTotruong_SelectedIndexChanged()`: Tự động chọn thôn

## Kiểm tra
✅ Build thành công  
✅ Chọn thôn → Lọc tổ trưởng  
✅ Chọn tổ trưởng → Tự động chọn thôn  
✅ Không có vòng lặp vô hạn  
✅ Không động chạm đến Form1  

## Lưu ý
- Tìm kiếm không phân biệt hoa thường (`StringComparison.OrdinalIgnoreCase`)
- Tìm trong cả `communes.villages` và `associations.villages`
- Sử dụng flag để tránh vòng lặp
- An toàn với try-catch
- Không ảnh hưởng đến dữ liệu lưu trữ (Customers, To)
