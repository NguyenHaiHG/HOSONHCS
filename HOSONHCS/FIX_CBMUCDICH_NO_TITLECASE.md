# FIX: cbMucdich1/2 - GIỮ NGUYÊN ĐỊNH DẠNG (KHÔNG TỰ ĐỘNG VIẾT HOA)

## 📅 Ngày: 2024

---

## ❌ VẤN ĐỀ TRƯỚC KHI SỬA:

Khi người dùng nhập vào `cbMucdich1` và `cbMucdich2` trong **Form1**, text bị **tự động chuyển sang Title Case** (viết hoa chữ cái đầu mỗi từ):

### Ví dụ:
```
Người dùng nhập: "nuôi lợn đen"
Sau khi lưu: "Nuôi Lợn Đen"   ❌ KHÔNG MONG MUỐN

Người dùng nhập: "TRỒNG NGỌC LINH"
Sau khi lưu: "Trồng Ngọc Linh"   ❌ KHÔNG MONG MUỐN
```

---

## 🔍 NGUYÊN NHÂN:

Trong hàm `BtnSave_Click()` (dòng 2085-2086), code đang gọi `ToTitleCase()` cho `cbmucdich1` và `cbmucdich2`:

```csharp
Mucdich1 = ToTitleCase(cbmucdich1 != null ? cbmucdich1.Text : ""),
Mucdich2 = ToTitleCase(cbmucdich2 != null ? cbmucdich2.Text : ""),
```

Hàm `ToTitleCase()` tự động viết hoa chữ cái đầu mỗi từ, gây ra thay đổi không mong muốn.

---

## ✅ GIẢI PHÁP:

Xóa `ToTitleCase()` và **GIỮ NGUYÊN** text như người dùng nhập vào.

### **Trước (dòng 2085-2086):**
```csharp
Mucdich1 = ToTitleCase(cbmucdich1 != null ? cbmucdich1.Text : ""),
Mucdich2 = ToTitleCase(cbmucdich2 != null ? cbmucdich2.Text : ""),
```

### **Sau (dòng 2085-2086):**
```csharp
Mucdich1 = (cbmucdich1 != null ? cbmucdich1.Text : ""),
Mucdich2 = (cbmucdich2 != null ? cbmucdich2.Text : ""),
```

---

## 🎯 KẾT QUẢ SAU KHI SỬA:

### ✅ **GIỮ NGUYÊN định dạng:**

| Người dùng nhập | Sau khi lưu | Kết quả |
|-----------------|-------------|---------|
| `nuôi lợn đen` | `nuôi lợn đen` | ✅ Đúng |
| `TRỒNG NGỌC LINH` | `TRỒNG NGỌC LINH` | ✅ Đúng |
| `Chăn Nuôi Bò` | `Chăn Nuôi Bò` | ✅ Đúng |
| `ABC 123` | `ABC 123` | ✅ Đúng |

### ✅ **Xuất Word cũng GIỮ NGUYÊN:**

File Word xuất ra sẽ hiển thị **CHÍNH XÁC** như người dùng đã nhập:

- `{{mucdich1}}` → `nuôi lợn đen` (nếu người dùng nhập như vậy)
- `{{mucdich2}}` → `TRỒNG NGỌC LINH` (nếu người dùng nhập như vậy)

---

## 📁 FILE THAY ĐỔI:

| File | Dòng | Thay đổi |
|------|------|----------|
| `Form1.cs` | 2085 | Xóa `ToTitleCase()` cho `Mucdich1` |
| `Form1.cs` | 2086 | Xóa `ToTitleCase()` cho `Mucdich2` |

---

## 🔍 CÁC TRƯỜNG HỢP KHÁC:

### ✅ **Các trường VẪN tự động Title Case (giữ nguyên):**

- `txtHoten` (Họ tên khách hàng)
- `txtntk1/2/3` (Người thừa kế)
- `cbPhuongan` (Phương án)

→ Vì đây là **TÊN NGƯỜI**, nên cần viết hoa chữ cái đầu.

### ✅ **Các trường KHÔNG tự động format (giữ nguyên):**

- `cbMucdich1/2` (Form1) - **ĐÃ SỬA** ✅
- `cbmd1/2/3/4/5` (Form2) - **KHÔNG BỊ ẢNH HƯỞNG** ✅
- `cbSotien`, `cbSotien1/2` - Chỉ format số tiền (thêm dấu chấm)
- `txtSdt` - Chỉ format số điện thoại (thêm dấu chấm)

---

## 🧪 CÁCH KIỂM TRA:

1. Mở Form1
2. Nhập vào `cbMucdich1`: `"nuôi lợn đen"`
3. Nhập vào `cbMucdich2`: `"TRỒNG NGỌC LINH"`
4. Nhấn **Lưu**
5. Kiểm tra file JSON trong `Customers\*.json`:
   ```json
   {
     "Mucdich1": "nuôi lợn đen",
     "Mucdich2": "TRỒNG NGỌC LINH"
   }
   ```
6. Xuất Word → Kiểm tra placeholder `{{mucdich1}}`, `{{mucdich2}}` giữ nguyên định dạng

---

## ✅ KẾT LUẬN:

- ✅ `cbMucdich1/2` giờ đây **GIỮ NGUYÊN** định dạng text như người dùng nhập
- ✅ Không còn tự động chuyển sang Title Case
- ✅ Xuất Word hiển thị **CHÍNH XÁC** như input
- ✅ Build thành công

---

**Hoàn tất!** 🎉
