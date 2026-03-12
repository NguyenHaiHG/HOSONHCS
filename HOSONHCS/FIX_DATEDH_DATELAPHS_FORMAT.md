# SỬA LỖI dateDH VÀ dateLaphs + FORMAT {{ngaylaphs}}

## Vấn đề đã sửa

### 1. ❌ dateDH không điền {{ngaydenhan}} dù tick hay không
**Nguyên nhân**: Điều kiện kiểm tra sai
```csharp
// SAI - Dòng 2039 (cũ)
if (dateDH != null && dateDH.Checked && dateDH.Format != DateTimePickerFormat.Custom)
```

**Vấn đề**: 
- Khi DateTimePicker có checkbox và **unchecked**, nó tự động set `Format = DateTimePickerFormat.Custom`
- Điều kiện `dateDH.Format != DateTimePickerFormat.Custom` → luôn false khi unchecked
- → Không bao giờ đọc được giá trị!

**Giải pháp**:
```csharp
// ĐÚNG - Dòng 2039 (mới)
if (dateDH != null && dateDH.Checked)
```

---

### 2. ❌ dateLaphs bị ngược (tick = biến mất, bỏ tick = hiện)
**Nguyên nhân**: dateLaphs có `ShowCheckBox = true` trong Designer

**Hành vi sai**:
- Tick checkbox → Ngày biến mất trong file Word
- Bỏ tick → Ngày hiện ra

**Giải pháp**: Bỏ checkbox (dateLaphs là trường BẮT BUỘC)
```csharp
// Constructor - Dòng ~193
if (dateLaphs != null)
{
    dateLaphs.ShowCheckBox = false;  // BỎ checkbox
    dateLaphs.Format = DateTimePickerFormat.Custom;
    dateLaphs.CustomFormat = "dd/MM/yyyy";
}
```

---

### 3. ✨ {{ngaylaphs}} hiển thị "Ngày...tháng...năm..."
**Yêu cầu**: Thay đổi format từ "15/03/2024" → "Ngày 15 tháng 03 năm 2024"

**Giải pháp**:

#### A. Thêm helper method
```csharp
// HOSONHCS\Form1.cs - Sau UpdateComputedFields()
/// <summary>
/// Format DateTime thành chuỗi "Ngày...tháng...năm..."
/// Ví dụ: 15/03/2024 → "Ngày 15 tháng 03 năm 2024"
/// </summary>
private string FormatDateToNgayThangNam(DateTime date)
{
    if (date == DateTime.MinValue) return "";
    return $"Ngày {date.Day:D2} tháng {date.Month:D2} năm {date.Year}";
}
```

#### B. Cập nhật replacement
```csharp
// HOSONHCS\Form1.cs - Dòng 619
// TRƯỚC:
{ "{{ngaylaphs}}", c.Ngaylaphs == DateTime.MinValue ? "" : c.Ngaylaphs.ToString("dd/MM/yyyy") }

// SAU:
{ "{{ngaylaphs}}", FormatDateToNgayThangNam(c.Ngaylaphs) }
```

---

## Chi tiết các thay đổi

### File: HOSONHCS\Form1.cs

#### 1. Constructor (dòng ~193)
**Thêm cấu hình cho dateLaphs**:
```csharp
// dateLaphs: BẮT BUỘC - BỎ checkbox
if (dateLaphs != null)
{
    dateLaphs.ShowCheckBox = false;
    dateLaphs.Format = DateTimePickerFormat.Custom;
    dateLaphs.CustomFormat = "dd/MM/yyyy";
}
```

#### 2. ReadForm (dòng 2039)
**Sửa điều kiện kiểm tra dateDH**:
```csharp
// BỎ điều kiện: && dateDH.Format != DateTimePickerFormat.Custom
DateTime ngaydenhan = DateTime.MinValue;
if (dateDH != null && dateDH.Checked)
{
    ngaydenhan = dateDH.Value.Date;
}
```

#### 3. Helper Method (sau UpdateComputedFields)
**Thêm FormatDateToNgayThangNam()**:
```csharp
private string FormatDateToNgayThangNam(DateTime date)
{
    if (date == DateTime.MinValue) return "";
    return $"Ngày {date.Day:D2} tháng {date.Month:D2} năm {date.Year}";
}
```

#### 4. ReplacePlaceholdersInWord (dòng 619)
**Đổi format cho {{ngaylaphs}}**:
```csharp
{ "{{ngaylaphs}}", FormatDateToNgayThangNam(c.Ngaylaphs) }
```

---

## Ví dụ format {{ngaylaphs}}

### Input
```csharp
dateLaphs.Value = new DateTime(2024, 3, 15);
```

### Output trong file Word

**TRƯỚC** (dd/MM/yyyy):
```
15/03/2024
```

**SAU** (Ngày...tháng...năm...):
```
Ngày 15 tháng 03 năm 2024
```

### Các trường hợp khác
| Input Date | Output |
|------------|--------|
| 01/01/2024 | Ngày 01 tháng 01 năm 2024 |
| 25/12/2023 | Ngày 25 tháng 12 năm 2023 |
| DateTime.MinValue | "" (rỗng) |

---

## So sánh trước/sau

### dateDH (Ngày đến hạn)

| Hành động | Trước (Sai) | Sau (Đúng) |
|-----------|-------------|-----------|
| Tick checkbox | ❌ Không điền | ✅ Điền {{ngaydenhan}} |
| Bỏ tick | ❌ Không điền | ✅ Không điền (đúng) |

### dateLaphs (Ngày lập hồ sơ)

| Trạng thái | Trước (Sai) | Sau (Đúng) |
|-----------|-------------|-----------|
| Có checkbox | ✅ Có | ❌ Không có (bỏ) |
| Tick → Hiện ngày | ❌ Biến mất | ✅ Luôn hiện |
| Bỏ tick → Ẩn ngày | ✅ Hiện (sai) | ✅ Không có tick |
| Format | dd/MM/yyyy | **Ngày...tháng...năm...** |

### {{ngaylaphs}}

| Input | Format cũ | Format mới |
|-------|-----------|-----------|
| 15/03/2024 | 15/03/2024 | **Ngày 15 tháng 03 năm 2024** |
| 01/01/2024 | 01/01/2024 | **Ngày 01 tháng 01 năm 2024** |

---

## Danh sách DateTimePicker sau khi sửa

| Control | Placeholder | ShowCheckBox | Format Output |
|---------|-------------|--------------|---------------|
| `dateLaphs` | `{{ngaylaphs}}` | ❌ false | **Ngày...tháng...năm...** |
| `dateDH` | `{{ngaydenhan}}` | ✅ true | dd/MM/yyyy |
| `dateNgaycapCCCD` | `{{ngaycap}}` | ❌ false | dd/MM/yyyy |
| `dateNgaysinh` | `{{ngaysinh}}` | ❌ false | dd/MM/yyyy |
| `datendhcccd` | `{{thoihancccd}}` | ❌ false | dd/MM/yyyy |
| `datentk1-3` | `{{namsinh1-3}}` | ✅ true | Tùy mẫu |

---

## Kiểm tra

### Test Case 1: dateDH với checkbox
```
1. Tick dateDH
2. Chọn ngày: 31/12/2024
3. Xuất Word
   → {{ngaydenhan}} = "31/12/2024" ✅

4. Bỏ tick dateDH
5. Xuất Word
   → {{ngaydenhan}} = "" (rỗng) ✅
```

### Test Case 2: dateLaphs không checkbox
```
1. Chọn ngày: 15/03/2024
2. Xuất Word
   → {{ngaylaphs}} = "Ngày 15 tháng 03 năm 2024" ✅

3. Không thể bỏ chọn (không có checkbox) ✅
```

### Test Case 3: Format Ngày tháng năm
```
Input: 05/06/2024
Output: "Ngày 05 tháng 06 năm 2024" ✅

Input: 25/12/2023
Output: "Ngày 25 tháng 12 năm 2023" ✅
```

---

## File đã thay đổi

### HOSONHCS\Form1.cs
1. **Constructor** (dòng ~193): Set `dateLaphs.ShowCheckBox = false`
2. **ReadForm** (dòng 2039): Bỏ điều kiện `dateDH.Format != DateTimePickerFormat.Custom`
3. **FormatDateToNgayThangNam** (mới): Helper method format "Ngày...tháng...năm..."
4. **ReplacePlaceholdersInWord** (dòng 619): Dùng `FormatDateToNgayThangNam()` cho {{ngaylaphs}}

---

## Build
✅ Build thành công  
✅ dateDH hoạt động đúng  
✅ dateLaphs không còn checkbox  
✅ {{ngaylaphs}} hiển thị "Ngày...tháng...năm..."  

## Lưu ý
- Chỉ {{ngaylaphs}} dùng format "Ngày...tháng...năm..."
- Các placeholder ngày khác vẫn dùng "dd/MM/yyyy"
- Nếu muốn áp dụng format này cho placeholder khác, gọi `FormatDateToNgayThangNam(date)`
