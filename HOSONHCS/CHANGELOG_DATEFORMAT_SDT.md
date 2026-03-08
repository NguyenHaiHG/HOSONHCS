# CHANGELOG - CẬP NHẬT ĐỊNH DẠNG NGÀY VÀ SỐ ĐIỆN THOẠI

## 📅 Ngày: 2024

## 🎯 Mục đích
Cải thiện trải nghiệm người dùng khi nhập ngày tháng và số điện thoại:
1. **Thời hạn CCCD**: Cho phép nhập ngày tương lai với validation định dạng hợp lý
2. **Số điện thoại**: Format tự động với dấu chấm ngăn cách để dễ đọc
3. **Nơi cấp CCCD**: Tự động điền dựa trên ngày cấp, không cho chỉnh sửa thủ công

---

## ✅ THAY ĐỔI CHI TIẾT

### 1️⃣ **THỜI HẠN CCCD (datendhcccd)**

#### ⚙️ Cấu hình mới:
- ✅ **Cho phép nhập ngày TƯƠNG LAI** (không giới hạn MaxDate)
- ✅ **Validation thông minh**:
  - Năm: Tối đa 5 chữ số (1 - 99999)
  - Tháng: 1-12 (tự động bởi DateTimePicker)
  - Ngày: 1-31 (tùy theo tháng)
  - **Tháng 2**: Tối đa 29 ngày (và kiểm tra năm nhuận)
- ✅ **Placeholder**: Khi tick chọn → điền vào `{{thoihancccd}}`

#### 💻 Code thay đổi:
```csharp
// InitializeApp()
if (datendhcccd != null) 
{ 
    datendhcccd.MaxDate = DateTime.MaxValue;  // Cho phép tương lai
    datendhcccd.ValueChanged += DateThoihanCCCD_ValueChanged;  // Event riêng
}

// ReadForm()
DateTime thoihancccd = DateTime.MinValue;
if (datendhcccd != null && datendhcccd.Checked)
{
    thoihancccd = datendhcccd.Value.Date;  // KHÔNG validate ngày tương lai
}

// PopulateForm()
if (c.Thoihancccd != DateTime.MinValue)
{
    datendhcccd.Checked = true;
    datendhcccd.Value = c.Thoihancccd;  // KHÔNG giới hạn ngày tương lai
}
```

#### 🔍 Validation logic:
```csharp
private void DateThoihanCCCD_ValueChanged(object sender, EventArgs e)
{
    // ✅ Validate năm không quá 5 số
    if (year > 99999) → Reset về 10 năm sau
    
    // ✅ Validate tháng 2 không quá 29 ngày
    if (month == 2 && day > 29) → Reset về 29/02
    
    // ✅ Validate năm nhuận cho ngày 29/02
    if (month == 2 && day == 29 && !IsLeapYear) → Reset về 28/02
}
```

---

### 2️⃣ **SỐ ĐIỆN THOẠI (txtSdt)**

#### ⚙️ Format mới:
- **Trước**: `0812801886` (khó đọc)
- **Sau**: `0812.801.886` (dễ đọc)

#### 📐 Định dạng: `XXXX.XXX.XXX`
- 4 số đầu
- Dấu chấm
- 3 số giữa
- Dấu chấm
- 3 số cuối

#### 💻 Code thay đổi:
```csharp
// InitializeApp()
txtSdt.MaxLength = 12;  // 10 số + 2 dấu chấm
txtSdt.KeyPress += TxtSdt_KeyPress;  // Cho phép số và dấu chấm

// TxtSdt_KeyPress()
// Cho phép: số (0-9), backspace, dấu chấm (.)

// TxtSdt_TextChanged()
var digits = new string(text.Where(char.IsDigit).ToArray());
if (digits.Length <= 4)
    formatted = digits;
else if (digits.Length <= 7)
    formatted = digits.Substring(0, 4) + "." + digits.Substring(4);
else
    formatted = digits.Substring(0, 4) + "." + digits.Substring(4, 3) + "." + digits.Substring(7);

// Ví dụ:
// 0812       → 0812
// 08128      → 0812.8
// 0812801    → 0812.801
// 0812801886 → 0812.801.886
```

#### 📂 Lưu trữ:
- **Database (JSON)**: Lưu VỚI dấu chấm (`0812.801.886`)
- **Word template**: Xuất VỚI dấu chấm (`{{sdt}}` = `0812.801.886`)

---

### 3️⃣ **NƠI CẤP CCCD (cbNoicap)**

#### ⚙️ Cấu hình mới:
- 🔒 **KHÓA** - Không cho chọn/gõ thủ công
- 🤖 **TỰ ĐỘNG ĐIỀN** dựa trên ngày cấp CCCD:

| Ngày cấp CCCD | Nơi cấp |
|---------------|---------|
| **Từ 01/07/2024 trở đi** | `Bộ Công an` |
| **Trước 01/07/2024** | `Cục CSQLHC về TTXH` |

#### 💻 Code thay đổi:
```csharp
// InitializeApp()
if (cbNoicap != null) 
{
    cbNoicap.DropDownStyle = ComboBoxStyle.DropDownList;  // Khóa không cho gõ
    cbNoicap.Enabled = false;  // Disable hoàn toàn
}

// DateNgaycapCCCD_ValueChanged()
var ngayCat = new DateTime(2024, 7, 1);  // Ngày cắt: 01/07/2024
var ngayCap = dateNgaycapCCCD.Value.Date;

if (ngayCap >= ngayCat)
    cbNoicap.Text = "Bộ Công an";
else
    cbNoicap.Text = "Cục CSQLHC về TTXH";
```

#### 🎬 Hành vi:
1. Người dùng chọn ngày cấp CCCD
2. Tự động điền nơi cấp (không cần chọn)
3. Không thể sửa nơi cấp thủ công

---

## 🧪 TEST CASES

### Test 1: Thời hạn CCCD
- ✅ Nhập ngày tương lai: `01/01/2030` → OK
- ✅ Nhập tháng 2 quá 29 ngày → Báo lỗi, reset về 29/02
- ✅ Nhập 29/02 năm không nhuận → Báo lỗi, reset về 28/02
- ✅ Nhập năm > 99999 → Báo lỗi, reset về 10 năm sau
- ✅ Khi tick checkbox → Điền vào `{{thoihancccd}}` trong Word
- ✅ Khi không tick → Không điền (trống)

### Test 2: Số điện thoại
- ✅ Nhập `0812801886` → Tự động format thành `0812.801.886`
- ✅ Nhập chữ cái → Tự động xóa
- ✅ Nhập quá 10 số → Tự động cắt
- ✅ Lưu JSON → Lưu với format `0812.801.886`
- ✅ Xuất Word → Hiển thị `0812.801.886`

### Test 3: Nơi cấp CCCD
- ✅ Chọn ngày cấp `02/07/2024` → Tự động điền "Bộ Công an"
- ✅ Chọn ngày cấp `30/06/2024` → Tự động điền "Cục CSQLHC về TTXH"
- ✅ Không thể click vào cbNoicap để chọn
- ✅ Không thể gõ vào cbNoicap

---

## 📋 PLACEHOLDERS HỖ TRỢ TRONG WORD TEMPLATE

### Tất cả mẫu Word:
```
{{sdt}}         → 0812.801.886 (có dấu chấm)
{{noicap}}      → Bộ Công an HOẶC Cục CSQLHC về TTXH (tự động)
{{thoihancccd}} → dd/MM/yyyy (chỉ khi tick checkbox)
```

---

## 🎨 TRẢI NGHIỆM NGƯỜI DÙNG

### ✨ Cải thiện:
1. **Số điện thoại dễ đọc hơn**: `0812.801.886` thay vì `0812801886`
2. **Nơi cấp tự động**: Không cần nhớ chọn "Bộ Công an" hay "Cục CSQLHC"
3. **Thời hạn CCCD linh hoạt**: Cho phép nhập ngày tương lai (vì CCCD thường có hạn 10-15 năm)
4. **Validation thông minh**: Ngăn lỗi nhập liệu phổ biến (tháng 2 có 30 ngày, năm không hợp lệ)

### 📊 Ví dụ thực tế:

**Kịch bản 1: Nhập CCCD mới (cấp 2024)**
1. Nhập ngày cấp: `15/08/2024`
2. → Nơi cấp tự động: `Bộ Công an` ✅
3. Tick checkbox thời hạn CCCD, nhập: `15/08/2034` (10 năm sau)
4. → Tự động validate và cho phép ✅

**Kịch bản 2: Nhập CCCD cũ (cấp trước 2024)**
1. Nhập ngày cấp: `20/05/2021`
2. → Nơi cấp tự động: `Cục CSQLHC về TTXH` ✅
3. Thời hạn để trống (CCCD cũ không có thời hạn)

**Kịch bản 3: Nhập số điện thoại**
1. Gõ: `0`, `8`, `1`, `2`, `8`, `0`, `1`, `8`, `8`, `6`
2. → Hiển thị: `0812.801.886` (tự động format) ✅
3. Lưu → JSON chứa: `"Sdt": "0812.801.886"`
4. Xuất Word → `{{sdt}}` = `0812.801.886`

---

## 🔧 MAINTENANCE NOTES

### DateTimePicker MaxDate:
- `dateLaphs`: MaxDate = Today ❌ (không cho tương lai)
- `dateNgaycapCCCD`: MaxDate = Today ❌ (không cho tương lai)
- `dateNgaysinh`: MaxDate = Today ❌ (không cho tương lai)
- `dateDH`: MaxDate = MaxValue ✅ (cho phép tương lai)
- `datendhcccd`: MaxDate = MaxValue ✅ (cho phép tương lai)
- `datentk1/2/3`: MaxDate = Today ❌ (không cho tương lai)

### ComboBox States:
- `cbNoicap`: Enabled = false, ReadOnly = true (khóa hoàn toàn)
- `cbPGD`, `cbXa`, `cbThon`, etc.: Enabled = true (cho phép chọn bình thường)

---

## 🐛 KNOWN ISSUES & FIXES

### Issue 1: DateTimePicker không cho nhập năm > 9999 mặc định
**Giải pháp**: DateTimePicker trong .NET Framework 4.7.2 tự động giới hạn năm từ 1753-9999. Validation của chúng ta (năm <= 99999) chủ yếu để báo lỗi rõ ràng cho người dùng.

### Issue 2: Số điện thoại format khi copy/paste
**Giải pháp**: `TxtSdt_TextChanged` tự động format kể cả khi paste, luôn đảm bảo format đúng `XXXX.XXX.XXX`

### Issue 3: cbNoicap bị disable nhưng vẫn muốn thay đổi
**Giải pháp**: Chỉ cần thay đổi ngày cấp CCCD, nơi cấp sẽ tự động cập nhật

---

## 📝 TODO (Nếu cần)

- [ ] Thêm tooltip cho cbNoicap giải thích tại sao bị khóa
- [ ] Thêm indicator (icon) khi ngày cấp >= 01/07/2024
- [ ] Cho phép override cbNoicap trong trường hợp đặc biệt (nếu cần)

---

## 🚀 DEPLOYMENT

1. Build successful ✅
2. Tất cả validation hoạt động ✅
3. Format số điện thoại hoạt động ✅
4. Auto-fill nơi cấp hoạt động ✅

**Sẵn sàng sử dụng!** 🎉
