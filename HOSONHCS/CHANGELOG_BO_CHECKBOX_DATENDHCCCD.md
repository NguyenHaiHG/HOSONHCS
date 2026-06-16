# CẬP NHẬT: BỎ CHECKBOX DATENDHCCCD

## Tổng quan
Đã bỏ checkbox của DateTimePicker `datendhcccd` (Thời hạn CCCD) trong Form1.

## Lý do
- Thời hạn CCCD là trường **BẮT BUỘC**
- Không cần checkbox để bỏ chọn
- Luôn phải nhập giá trị

## Các thay đổi

### 1. Constructor Form1 - Bỏ ShowCheckBox
**File**: `HOSONHCS\Form1.cs` (dòng 199-203)

**Trước**:
```csharp
if (datendhcccd != null) 
{ 
    datendhcccd.ShowCheckBox = true; 
    datendhcccd.Checked = false;  // Thời hạn CCCD: optional
}
```

**Sau**:
```csharp
// datendhcccd: BỎ checkbox (bắt buộc nhập)
if (datendhcccd != null) 
{ 
    datendhcccd.ShowCheckBox = false; 
}
```

### 2. PopulateForm - Load dữ liệu
**File**: `HOSONHCS\Form1.cs` (dòng 2194-2213)

**Trước**:
```csharp
if (datendhcccd != null) 
{
    datendhcccd.ShowCheckBox = true;
    if (c.Thoihancccd == DateTime.MinValue)
    {
        datendhcccd.Checked = false;
    }
    else
    {
        datendhcccd.Format = DateTimePickerFormat.Custom;
        datendhcccd.CustomFormat = "dd/MM/yyyy";
        datendhcccd.Checked = true;
        datendhcccd.Value = thoihancccd;
    }
}
```

**Sau**:
```csharp
if (datendhcccd != null) 
{
    // Bỏ checkbox - datendhcccd bắt buộc nhập
    if (c.Thoihancccd == DateTime.MinValue)
    {
        // Nếu chưa có dữ liệu, set ngày mặc định
        datendhcccd.Format = DateTimePickerFormat.Custom;
        datendhcccd.CustomFormat = "dd/MM/yyyy";
        datendhcccd.Value = DateTime.Now;
    }
    else
    {
        datendhcccd.Format = DateTimePickerFormat.Custom;
        datendhcccd.CustomFormat = "dd/MM/yyyy";
        datendhcccd.Value = thoihancccd;
    }
}
```

### 3. ReadForm - Đọc dữ liệu
**File**: `HOSONHCS\Form1.cs` (dòng 2045-2056)

**Trước**:
```csharp
DateTime thoihancccd = DateTime.MinValue;
if (datendhcccd != null && datendhcccd.Checked)  // ← Kiểm tra Checked
{
    thoihancccd = datendhcccd.Value.Date;

    if (thoihancccd < DateTime.Today)
    {
        throw new Exception($"CCCD đã hết hạn...");
    }
}
```

**Sau**:
```csharp
DateTime thoihancccd = DateTime.MinValue;
if (datendhcccd != null)  // ← BỎ kiểm tra Checked
{
    thoihancccd = datendhcccd.Value.Date;

    if (thoihancccd < DateTime.Today)
    {
        throw new Exception($"CCCD đã hết hạn...");
    }
}
```

## Hành vi mới

### Khi mở form mới
- datendhcccd hiển thị ngày hiện tại
- Không có checkbox bên cạnh
- Luôn có giá trị

### Khi load customer chưa có Thoihancccd
```csharp
if (c.Thoihancccd == DateTime.MinValue)
{
    datendhcccd.Value = DateTime.Now;  // Set ngày mặc định
}
```

### Khi load customer đã có Thoihancccd
```csharp
else
{
    datendhcccd.Value = c.Thoihancccd;  // Load từ dữ liệu
}
```

### Khi lưu customer
- Luôn lưu giá trị từ `datendhcccd.Value`
- Không còn kiểm tra `datendhcccd.Checked`
- Validation: CCCD không được hết hạn

## So sánh với các DateTimePicker khác

| Control | ShowCheckBox | Bắt buộc? | Mặc định |
|---------|--------------|-----------|----------|
| `dateNgaycapCCCD` | false | ✅ Có | Ngày hiện tại |
| `dateNgaysinh` | false | ✅ Có | Ngày hiện tại |
| `datendhcccd` | **false** (mới) | ✅ Có | Ngày hiện tại |
| `dateDH` | true | ❌ Không | Unchecked |
| `datentk1/2/3` | true | ❌ Không | Unchecked |

## Validation

### CCCD hết hạn
```csharp
if (thoihancccd < DateTime.Today)
{
    throw new Exception(
        $"CCCD đã hết hạn ngày {thoihancccd:dd/MM/yyyy}.\n\n" +
        "Không thể tạo hồ sơ với CCCD hết hạn.\n\n" +
        "Vui lòng cập nhật CCCD mới."
    );
}
```

### Luôn có giá trị
- Không còn kiểm tra `Thoihancccd == DateTime.MinValue` khi lưu
- Luôn lưu giá trị từ DateTimePicker

## File đã thay đổi

### HOSONHCS\Form1.cs
1. Constructor (dòng ~199): `datendhcccd.ShowCheckBox = false;`
2. PopulateForm (dòng ~2194): Bỏ logic Checked, luôn set Value
3. ReadForm (dòng ~2046): Bỏ kiểm tra `datendhcccd.Checked`

### Form2 (Xác nhận)
✅ Đã kiểm tra - chương trình "Cấp nước sạch" → "HGĐ cư trú tại VNT" (đúng)
- Không cần thay đổi gì trong Form2

## Kiểm tra
✅ Build thành công  
✅ Bỏ checkbox datendhcccd  
✅ Luôn có giá trị mặc định  
✅ Validation CCCD hết hạn vẫn hoạt động  
✅ Load/Save customer hoạt động bình thường  

## Lưu ý
- Chỉ thay đổi Form1
- Form2 không bị ảnh hưởng
- Logic validation CCCD hết hạn vẫn giữ nguyên
- Dữ liệu cũ (có Thoihancccd = DateTime.MinValue) sẽ được set thành ngày hiện tại khi load
