# SỬA "Hộ GĐ vùng NT" THÀNH "HGĐ cư trú tại VNT"

## Vấn đề
Chương trình "Cấp nước sạch và vệ sinh môi trường nông thôn" vẫn hiển thị đối tượng cũ:
- ❌ "Hộ GĐ vùng NT" (sai)
- ✅ "HGĐ cư trú tại VNT" (đúng)

## Nguyên nhân
Có 2 nơi xử lý logic cbChuongtrinh → cbDoituong:
1. **Form1.cs** - Logic cbChuongtrinh_SelectedIndexChanged
2. **Form2.cs** - Logic Cbctr_SelectedIndexChanged

Form2 đã đúng từ trước, nhưng Form1 vẫn dùng text cũ.

## Giải pháp

### File: HOSONHCS\Form1.cs
**Dòng**: 2616

**Trước**:
```csharp
if (normalized.Contains("cap nuoc sach") ||
    normalized.Contains("ve sinh moi truong") ||
    (normalized.Contains("nuoc sach") && normalized.Contains("nong thon")))
{
    cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;
    cbDoituong.Enabled = false;
    cbDoituong.Text = "Hộ GĐ vùng NT";  // ❌ SAI
    return;
}
```

**Sau**:
```csharp
if (normalized.Contains("cap nuoc sach") ||
    normalized.Contains("ve sinh moi truong") ||
    (normalized.Contains("nuoc sach") && normalized.Contains("nong thon")))
{
    cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;
    cbDoituong.Enabled = false;
    cbDoituong.Text = "HGĐ cư trú tại VNT";  // ✅ ĐÚNG
    return;
}
```

## Kiểm tra

### Form1
```
cbChuongtrinh: "Cấp nước sạch và vệ sinh môi trường nông thôn"
    ↓
cbDoituong: "HGĐ cư trú tại VNT" ✅
```

### Form2
```
cbctr: "Cấp nước sạch và vệ sinh môi trường nông thôn"
    ↓
cbdt1-5: ["HGĐ cư trú tại VNT"] ✅
```

## Danh sách đầy đủ ánh xạ Chương trình → Đối tượng

| STT | Chương trình | Đối tượng (Form1) | Đối tượng (Form2) |
|-----|--------------|-------------------|-------------------|
| 1 | Hộ nghèo | Hộ nghèo | Hộ nghèo |
| 2 | Hộ cận nghèo | Hộ cận nghèo | Hộ cận nghèo |
| 3 | Hộ mới thoát nghèo | Hộ mới thoát nghèo | Hộ mới thoát nghèo |
| 4 | Hộ GĐ SXKD VKK | Hộ GĐ SXKD VKK | Hộ GĐ SXKD VKK |
| 5 | Hỗ trợ tạo việc làm | Người lao động<br>NLĐ là người DTTS | Người lao động<br>NLĐ là người DTTS |
| 6 | **Cấp nước sạch** | **HGĐ cư trú tại VNT** ✅ | **HGĐ cư trú tại VNT** ✅ |

## Lịch sử thay đổi

### Tên đối tượng qua các phiên bản
1. **Ban đầu**: "Hộ GĐ vùng NT"
2. **Cập nhật**: "HGĐ cư trú tại VNT"

### Vị trí thay đổi
- ✅ Form2.cs (dòng 1677) - Đã đúng từ trước
- ✅ Form1.cs (dòng 2616) - Vừa sửa

## File đã thay đổi

### HOSONHCS\Form1.cs
- Dòng 2616: `cbDoituong.Text = "HGĐ cư trú tại VNT";`

### HOSONHCS\Form2.cs
- Không cần thay đổi (đã đúng)

## Build
✅ Build thành công  
✅ Form1 hiển thị "HGĐ cư trú tại VNT"  
✅ Form2 hiển thị "HGĐ cư trú tại VNT"  

## Lưu ý cho dữ liệu cũ
Nếu có file JSON cũ đã lưu với "Hộ GĐ vùng NT":
1. File vẫn load được bình thường
2. Khi chọn lại chương trình, sẽ tự động cập nhật thành "HGĐ cư trú tại VNT"
3. Lưu lại → File JSON sẽ có giá trị mới

## Cách kiểm tra
1. Mở Form1
2. Chọn cbChuongtrinh: "Cấp nước sạch và vệ sinh môi trường nông thôn"
3. Kiểm tra cbDoituong → Phải hiện "HGĐ cư trú tại VNT"

4. Mở Form2
5. Chọn cbctr: "Cấp nước sạch và vệ sinh môi trường nông thôn"
6. Kiểm tra cbdt1-5 → Dropdown chỉ có 1 lựa chọn: "HGĐ cư trú tại VNT"
