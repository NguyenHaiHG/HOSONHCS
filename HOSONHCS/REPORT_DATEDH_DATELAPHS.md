# BÁO CÁO: KIỂM TRA dateDH VÀ dateLaphs

## Kết quả kiểm tra

### ✅ dateDH (Ngày đến hạn)
**Control**: `dateDH` (DateTimePicker)  
**Field trong Customer**: `Ngaydenhan`  
**Placeholder**: `{{ngaydenhan}}`

#### Đọc từ form (ReadForm)
```csharp
// HOSONHCS\Form1.cs - Dòng 2038-2043
DateTime ngaydenhan = DateTime.MinValue;
if (dateDH != null && dateDH.Checked && dateDH.Format != DateTimePickerFormat.Custom)
{
    // Ngày đến hạn KHÔNG validate ngày tương lai - cho phép nhập bất kỳ ngày nào
    ngaydenhan = dateDH.Value.Date;
}
```

#### Lưu vào Customer
```csharp
// HOSONHCS\Form1.cs - Dòng 2090
Ngaydenhan = ngaydenhan,
```

#### Replace placeholder
```csharp
// HOSONHCS\Form1.cs - Dòng 604
{ "{{ngaydenhan}}", c.Ngaydenhan == DateTime.MinValue ? "" : c.Ngaydenhan.ToString("dd/MM/yyyy") },
```

#### Trạng thái
- ✅ Đã link đến `{{ngaydenhan}}`
- ✅ Format: `dd/MM/yyyy`
- ✅ Có ShowCheckBox (optional)
- ✅ Cho phép ngày tương lai
- ✅ Nếu không check → lưu `DateTime.MinValue` → placeholder = ""

---

### ✅ dateLaphs (Ngày lập hồ sơ)
**Control**: `dateLaphs` (DateTimePicker)  
**Field trong Customer**: `Ngaylaphs`  
**Placeholder**: `{{ngaylaphs}}`

#### Đọc từ form (ReadForm)
```csharp
// HOSONHCS\Form1.cs - Dòng 2035-2036
DateTime ngaylaphs = dateLaphs.Value.Date;
if (ngaylaphs > DateTime.Today) ngaylaphs = DateTime.Today;
```

#### Lưu vào Customer
```csharp
// HOSONHCS\Form1.cs - Dòng 2089
Ngaylaphs = ngaylaphs,
```

#### Replace placeholder
```csharp
// HOSONHCS\Form1.cs - Dòng 619
{ "{{ngaylaphs}}", c.Ngaylaphs == DateTime.MinValue ? "" : c.Ngaylaphs.ToString("dd/MM/yyyy") },
```

#### Trạng thái
- ✅ Đã link đến `{{ngaylaphs}}`
- ✅ Format: `dd/MM/yyyy`
- ✅ Bắt buộc nhập (không có checkbox)
- ✅ Validate: không cho phép ngày tương lai
- ✅ Nếu chọn ngày tương lai → tự động set thành `DateTime.Today`

---

## Tổng hợp tất cả DateTimePicker trong Form1

| Control | Field | Placeholder | Optional? | Validate tương lai? |
|---------|-------|-------------|-----------|---------------------|
| `dateNgaycapCCCD` | `Ngaycap` | `{{ngaycap}}` | ❌ Bắt buộc | ✅ Có (≤ hôm nay) |
| `dateNgaysinh` | `Ngaysinh` | `{{ngaysinh}}` | ❌ Bắt buộc | ✅ Có (≤ hôm nay) |
| `dateLaphs` | `Ngaylaphs` | `{{ngaylaphs}}` | ❌ Bắt buộc | ✅ Có (≤ hôm nay) |
| `datendhcccd` | `Thoihancccd` | `{{thoihancccd}}` | ❌ Bắt buộc | ❌ Không |
| `dateDH` | `Ngaydenhan` | `{{ngaydenhan}}` | ✅ Optional | ❌ Không |
| `datentk1` | `Namsinh1` | `{{namsinh1}}` | ✅ Optional | - |
| `datentk2` | `Namsinh2` | `{{namsinh2}}` | ✅ Optional | - |
| `datentk3` | `Namsinh3` | `{{namsinh3}}` | ✅ Optional | - |

---

## Chi tiết từng placeholder

### 1. {{ngaydenhan}} (dateDH)
**Mục đích**: Ngày đến hạn vay  
**Logic**:
- Nếu `dateDH.Checked = false` → `Ngaydenhan = DateTime.MinValue` → `{{ngaydenhan}} = ""`
- Nếu `dateDH.Checked = true` → `Ngaydenhan = dateDH.Value` → `{{ngaydenhan}} = "dd/MM/yyyy"`
- Cho phép chọn ngày tương lai (không validate)

**Sử dụng trong mẫu**:
- Mẫu 01/TD: Có thể có
- Mẫu 03: Có thể có
- GUQ: Không dùng

### 2. {{ngaylaphs}} (dateLaphs)
**Mục đích**: Ngày lập hồ sơ  
**Logic**:
- Bắt buộc nhập, luôn có giá trị
- Validate: nếu chọn tương lai → tự động set = `DateTime.Today`
- `{{ngaylaphs}} = "dd/MM/yyyy"`

**Sử dụng trong mẫu**:
- Mẫu 01/TD: Có
- Mẫu 03: Có
- GUQ: Có thể có

---

## Code tham khảo

### ReadForm (đọc dữ liệu từ form)
```csharp
// Ngày lập hồ sơ - BẮT BUỘC
DateTime ngaylaphs = dateLaphs.Value.Date;
if (ngaylaphs > DateTime.Today) ngaylaphs = DateTime.Today;

// Ngày đến hạn - OPTIONAL
DateTime ngaydenhan = DateTime.MinValue;
if (dateDH != null && dateDH.Checked && dateDH.Format != DateTimePickerFormat.Custom)
{
    ngaydenhan = dateDH.Value.Date;
}

// Lưu vào Customer
return new Customer
{
    // ...
    Ngaylaphs = ngaylaphs,
    Ngaydenhan = ngaydenhan,
    // ...
};
```

### ReplacePlaceholdersInWord
```csharp
var replacements = new Dictionary<string, string>
{
    // ...
    { "{{ngaylaphs}}", c.Ngaylaphs == DateTime.MinValue ? "" : c.Ngaylaphs.ToString("dd/MM/yyyy") },
    { "{{ngaydenhan}}", c.Ngaydenhan == DateTime.MinValue ? "" : c.Ngaydenhan.ToString("dd/MM/yyyy") },
    // ...
};
```

---

## Kết luận

✅ **dateDH** (Ngày đến hạn):
- ✅ Đã link đến placeholder `{{ngaydenhan}}`
- ✅ Format `dd/MM/yyyy`
- ✅ Optional (có checkbox)
- ✅ Cho phép ngày tương lai

✅ **dateLaphs** (Ngày lập hồ sơ):
- ✅ Đã link đến placeholder `{{ngaylaphs}}`
- ✅ Format `dd/MM/yyyy`
- ✅ Bắt buộc (không có checkbox)
- ✅ Validate ngày tương lai

**Cả 2 DateTimePicker đều đã được cấu hình đầy đủ và hoạt động chính xác!**
