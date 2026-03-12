# HƯỚNG DẪN SỬ DỤNG CHỨC NĂNG CHỈNH SỬA VÀ TẠO TỔ MỚI TRONG FORM2

## Tổng quan
Đã cấu hình Form2 với chức năng:
- **Chỉnh sửa tổ**: Click vào dòng trong dgv2 để load TOÀN BỘ thông tin tổ lên form
- **Cập nhật tổ (ghi đè)**: Bấm nút "btn03to" để ghi đè lên tổ đã chọn
- **Tạo tổ mới**: Bấm nút "btntaoto" để lưu thành tổ mới (không xóa dữ liệu đang có)

## Các thay đổi đã thực hiện

### 1. Mở rộng class ExportHistory
Đã thêm các trường để lưu đầy đủ thông tin 5 tổ viên:
- **Thông tin chung**: Pgd, Chuongtrinh
- **Họ tên**: Kh1, Kh2, Kh3, Kh4, Kh5
- **Số tiền**: Tien1, Tien2, Tien3, Tien4, Tien5
- **Phương án**: Md1, Md2, Md3, Md4, Md5
- **Thời hạn vay**: Time1, Time2, Time3, Time4, Time5
- **Đối tượng**: Dt1, Dt2, Dt3, Dt4, Dt5

### 2. Chức năng mới

#### A. Click vào dgv2 để load đầy đủ thông tin tổ
- **Event**: `Dgv2_CellClick`
- **Chức năng**: 
  - Khi click vào một dòng trong dgv2, TOÀN BỘ thông tin tổ sẽ được load lên form
  - Chuyển sang chế độ Edit (isEditMode = true)
  - Text của btn03to thay đổi thành "Cập nhật tổ (ghi đè)"
  - **Thông tin được load**:
    - Thông tin chung: PGD, Xã, Thôn, Tổ trưởng, Chương trình
    - Họ tên 5 tổ viên: txtkh1-5
    - Số tiền 5 tổ viên: cbtien1-5
    - Phương án: cbmd1-5
    - Thời hạn vay: cbtime1-5
    - Đối tượng: cbdt1-5
  - Hiển thị thông báo hướng dẫn

#### B. Nút "Cập nhật tổ (ghi đè)" - btn03to
- **Chế độ Edit** (sau khi click vào dgv2):
  - Cập nhật (ghi đè) TOÀN BỘ thông tin tổ đã chọn
  - Không tạo record mới trong dgv2
  - Xuất file Word với thông tin đã cập nhật
  - Hiển thị thông báo "Đã cập nhật thông tin tổ (ghi đè)"

- **Chế độ Add** (mặc định):
  - Tạo tổ mới và thêm vào dgv2
  - Xuất file Word
  - Hiển thị thông báo "Đã tạo tổ mới và xuất file Word"

#### C. Nút "Tạo tổ mới" - btntaoto
- **Event**: `BtnTaoTo_Click`
- **Chức năng**:
  - Lấy dữ liệu hiện tại trên form
  - Kiểm tra dữ liệu (tối thiểu 2 người + tên tổ trưởng)
  - Hiển thị hộp thoại xác nhận với thông tin tổ
  - Tạo ExportHistory mới với dữ liệu hiện tại
  - Thêm vào dgv2 (không ghi đè)
  - Chuyển về chế độ Add
  - Text của btn03to đổi về "Xuất tổ"
  - **KHÔNG xóa dữ liệu trên form**

### 3. Phương thức hỗ trợ
- `ClearAllFields()`: Xóa tất cả dữ liệu trên form (không còn được dùng trong btntaoto)
- `ClearToVienFields()`: Xóa thông tin 5 tổ viên

## Hướng dẫn sử dụng chi tiết

### Kịch bản 1: Chỉnh sửa tổ đã có (GHI ĐÈ)
1. Click vào dòng tổ trong dgv2 muốn chỉnh sửa
2. **Toàn bộ thông tin tổ sẽ hiện lại trên form** (tên, tiền, phương án, thời hạn, đối tượng)
3. Chỉnh sửa thông tin cần thiết
4. Bấm nút **"Cập nhật tổ (ghi đè)"** (btn03to)
5. → Thông tin sẽ được cập nhật, file Word sẽ được xuất lại với thông tin mới

### Kịch bản 2: Tạo tổ mới từ tổ đã có
1. Click vào dòng tổ trong dgv2
2. Thông tin tổ hiện lại trên form
3. Chỉnh sửa thông tin (ví dụ: đổi tên tổ trưởng, thay đổi thành viên)
4. Bấm nút **"Tạo tổ mới"** (btntaoto)
5. Xác nhận trong hộp thoại
6. → Tổ mới sẽ được tạo, tổ cũ vẫn còn trong dgv2

### Kịch bản 3: Tạo tổ hoàn toàn mới
1. Nhập thông tin tổ mới trên form
2. Bấm nút **"Xuất tổ"** (btn03to)
3. → Tổ mới sẽ được tạo và file Word sẽ được xuất

### Luồng hoạt động:
```
[Mở Form2] → [Chế độ Add - Mặc định]
    ↓
    ├─ [Nhập dữ liệu] → [Bấm "Xuất tổ"] → [Tạo tổ mới + Xuất Word]
    │
    └─ [Click vào dgv2] → [Chế độ Edit - Load đầy đủ thông tin]
            ↓
            ├─ [Chỉnh sửa] → [Bấm "Cập nhật tổ (ghi đè)"] → [Cập nhật tổ + Xuất Word]
            │
            └─ [Chỉnh sửa] → [Bấm "Tạo tổ mới"] → [Tạo tổ mới] → [Chế độ Add]
```

## So sánh 2 nút

| Tính năng | btn03to (Xuất tổ / Cập nhật tổ) | btntaoto (Tạo tổ mới) |
|-----------|----------------------------------|------------------------|
| **Khi nào dùng** | Xuất Word hoặc cập nhật tổ đã chọn | Tạo tổ mới từ dữ liệu hiện tại |
| **Chế độ Add** | Tạo tổ mới + Xuất Word | Tạo tổ mới KHÔNG xuất Word |
| **Chế độ Edit** | Ghi đè lên tổ đã chọn + Xuất Word | Tạo tổ mới từ dữ liệu đã load |
| **Xuất file Word** | ✓ Có | ✗ Không |
| **Xóa form** | ✗ Không | ✗ Không |
| **Hộp thoại xác nhận** | ✗ Không | ✓ Có |

## Lưu ý quan trọng
- ✅ Click vào dgv2 → Load **TOÀN BỘ** thông tin (txtkh1-5, cbtien1-5, cbmd1-5, cbtime1-5, cbdt1-5)
- ✅ Bấm **btn03to** sau khi click dgv2 → **GHI ĐÈ** lên tổ cũ
- ✅ Bấm **btntaoto** sau khi click dgv2 → **TẠO TỔ MỚI** (không ghi đè)
- ⚠️ Nút "btnxoa" vẫn giữ nguyên chức năng xóa tổ trong dgv2
- ⚠️ File Word được xuất khi bấm btn03to (cả Edit và Add mode)
- ⚠️ Btntaoto yêu cầu tối thiểu 2 người và tên tổ trưởng

## File được thay đổi
- `HOSONHCS\Form2.cs`:
  - Mở rộng class **ExportHistory** với 25+ trường mới
  - Cập nhật **Dgv2_CellClick**: Load đầy đủ thông tin tổ viên
  - Cập nhật **BtnTaoTo_Click**: Tạo tổ mới từ dữ liệu hiện tại (không xóa form)
  - Cập nhật **Btn03to_Click**: Lưu đầy đủ thông tin khi tạo/cập nhật

## Kiểm tra
Build thành công ✓
Không có lỗi compilation ✓
Logic đã được cập nhật theo yêu cầu ✓
