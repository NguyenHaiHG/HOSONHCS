# Hướng dẫn sử dụng chức năng tách 12 số CCCD

## Tính năng

Khi bấm nút **btn01TGTV** (Tạo 01/TGTV), hệ thống sẽ tự động:
1. Lấy số CCCD từ ô `txtSocccd` (12 số)
2. Tách thành 12 ký tự riêng biệt
3. Điền vào mẫu Word `01TGTV.docx` tại placeholder `{{cccd12}}`

## Cách sử dụng trong template Word

### Bước 1: Tạo bảng trong Word
Trong file `01TGTV.docx`, tạo một bảng với **12 cột** liền nhau:

| Ô 1 | Ô 2 | Ô 3 | Ô 4 | Ô 5 | Ô 6 | Ô 7 | Ô 8 | Ô 9 | Ô 10 | Ô 11 | Ô 12 |
|------|------|------|------|------|------|------|------|------|------|------|------|
|      |      |      |      |      |      |      |      |      |      |      |      |

### Bước 2: Đặt placeholder
- Trong **Ô đầu tiên** (Ô 1), gõ: `{{cccd12}}`
- Để các ô còn lại trống

### Bước 3: Định dạng bảng
- Canh giữa văn bản trong các ô
- Đặt kích thước ô phù hợp (ví dụ: rộng 1cm mỗi ô)
- Có thể thêm border để tạo ô vuông rõ ràng

## Cách hoạt động

Khi xuất file:
- Hệ thống sẽ tìm placeholder `{{cccd12}}` trong bảng
- Lấy 12 số từ CCCD (ví dụ: `001234567890`)
- Điền mỗi số vào một ô:
  - Ô 1: `0`
  - Ô 2: `0`
  - Ô 3: `1`
  - Ô 4: `2`
  - ...
  - Ô 12: `0`

## Ví dụ

**Input CCCD:** `001234567890`

**Kết quả trong Word:**

| 0 | 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 0 |
|---|---|---|---|---|---|---|---|---|---|---|---|

## Lưu ý

1. **CCCD phải có 12 số**: Nếu ít hơn, hệ thống sẽ thêm số `0` vào cuối. Nếu nhiều hơn, chỉ lấy 12 số đầu.

2. **Bảng phải có đủ 12 ô**: Placeholder `{{cccd12}}` cần được đặt ở ô đầu tiên của dòng có ít nhất 12 ô liền nhau.

3. **Chỉ lấy số**: Hệ thống tự động lọc chỉ lấy các ký tự số từ ô `txtSocccd`.

## Xử lý lỗi

- Nếu không tìm thấy `{{cccd12}}`, hệ thống sẽ bỏ qua (không báo lỗi)
- Nếu bảng có ít hơn 12 ô, chỉ điền vào các ô có sẵn
- Placeholder `{{cccd}}` vẫn hoạt động bình thường (hiển thị CCCD đầy đủ không tách)

## Code implementation

### Hàm chính
- `ProcessCCCD12Placeholder()`: Tìm và điền 12 số CCCD vào bảng Word
- `SplitCCCDInto12Boxes()`: Tách CCCD thành chuỗi có khoảng trắng (dự phòng)

### Placeholder hỗ trợ
- `{{cccd}}`: Hiển thị CCCD đầy đủ (ví dụ: `001234567890`)
- `{{socccd}}`: Giống `{{cccd}}`
- `{{cccd12}}`: Tách thành 12 ô riêng biệt (ví dụ: `0 0 1 2 3 4 5 6 7 8 9 0`)
