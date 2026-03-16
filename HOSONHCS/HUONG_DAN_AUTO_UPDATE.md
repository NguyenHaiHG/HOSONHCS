# HƯỚNG DẪN HỆ THỐNG AUTO-UPDATE

## ✅ ĐÃ TẠO FILE

- `HOSONHCS\AutoUpdater.cs` - Class quản lý auto-update
- Đã tích hợp vào Form1.cs

## 🚀 WORKFLOW HỆ THỐNG

```
┌──────────────┐         ┌─────────────────┐         ┌──────────────┐
│ Tác giả      │────────→│ GitHub Releases │←────────│ User's App   │
│ (Developer)  │ Upload  │                 │ Check   │              │
└──────────────┘         └─────────────────┘         └──────────────┘
```

## 📦 BƯỚC 1: TÁC GIẢ ĐẨY BẢN MỚI LÊN GITHUB

### 1.1. Cập nhật Assembly Version

**File**: `HOSONHCS\Properties\AssemblyInfo.cs`

```csharp
[assembly: AssemblyVersion("1.0.0")]
[assembly: AssemblyFileVersion("1.0.0")]
```

Thay `1.0.0` thành version mới (ví dụ: `1.1.0`)

### 1.2. Build Release

1. Visual Studio → Build → Configuration Manager
2. Chọn **Release** (không phải Debug)
3. Build Solution (Ctrl+Shift+B)
4. File sẽ ở: `D:\C#\HOSONHCS\bin\Release\`

### 1.3. Tạo file .zip

**Tạo folder tạm**:
```
HOSONHCS_v1.1.0\
├── HOSONHCS.exe
├── *.dll (DocumentFormat.OpenXml.dll, Newtonsoft.Json.dll, ...)
└── KHÔNG bao gồm:
    ├── Customers\   ← Dữ liệu người dùng
    ├── To\          ← Dữ liệu người dùng
    ├── BangKe\      ← Dữ liệu người dùng
    └── *.json       ← Dữ liệu người dùng
```

**Nén thành**: `HOSONHCS_v1.1.0.zip`

### 1.4. Tạo GitHub Release

1. Mở https://github.com/NguyenHaiHG/HOSONHCS/releases
2. Click **"Draft a new release"**
3. Điền thông tin:
   - **Tag**: `v1.1.0` (phải có chữ 'v')
   - **Title**: `HOSONHCS v1.1.0 - Thêm tính năng X`
   - **Description**:
     ```
     ## Thay đổi
     - Thêm bảng kê tiền
     - Sửa lỗi format ngày tháng
     - Cải thiện hiệu suất

     ## Yêu cầu
     - .NET Framework 4.7.2 trở lên
     ```
4. **Upload file**: Kéo thả `HOSONHCS_v1.1.0.zip`
5. Click **"Publish release"**

### 1.5. Xác nhận

URL sẽ có dạng:
```
https://github.com/NguyenHaiHG/HOSONHCS/releases/download/v1.1.0/HOSONHCS_v1.1.0.zip
```

## 💻 BƯỚC 2: USER SỬ DỤNG

### 2.1. Tự động kiểm tra khi khởi động

Khi mở app:
```
Form1_Load() → CheckForUpdateOnStartup()
              ↓
Gọi GitHub API (silent mode)
              ↓
    Có update?  ┌─ Không → Im lặng, app chạy bình thường
              └─ Có → Hiện dialog hỏi user
```

### 2.2. Thủ công kiểm tra

User bấm **btnUpdate**:
```
BtnUpdate_Click()
    ↓
Gọi GitHub API
    ↓
Hiện thông báo (có update hoặc đã mới nhất)
```

### 2.3. Dialog cập nhật

```
╔════════════════════════════════════════╗
║ Có phiên bản mới: 1.1.0               ║
║ Phiên bản hiện tại: 1.0.0             ║
║ Ngày phát hành: 20/12/2024            ║
║                                        ║
║ Thay đổi:                              ║
║ - Thêm bảng kê tiền                   ║
║ - Sửa lỗi format ngày tháng           ║
║                                        ║
║ Bạn có muốn cập nhật ngay bây giờ?   ║
║                                        ║
║     [ Có ]           [ Không ]        ║
╚════════════════════════════════════════╝
```

### 2.4. Tiến trình cập nhật

```
1. Tải file .zip về temp folder
   └─ Progress: 0-50%

2. Giải nén
   └─ Progress: 50-75%

3. Tạo file update.bat
   └─ Progress: 75-100%

4. Chạy batch file và thoát app
   ├─ Backup HOSONHCS.exe → HOSONHCS.exe.old
   ├─ Copy file mới (chỉ .exe và .dll)
   ├─ KHÔNG copy: Customers\, To\, BangKe\
   └─ Khởi động lại app
```

## 🔧 BƯỚC 3: CÀI ĐẶT (CHO LẦN ĐẦU)

### 3.1. Đóng Visual Studio

### 3.2. Thêm Reference

Mở file `HOSONHCS\HOSONHCS.csproj` bằng Notepad, tìm dòng:
```xml
<Reference Include="System.Core" />
```

Thêm 2 dòng sau:
```xml
<Reference Include="System.IO.Compression" />
<Reference Include="System.IO.Compression.FileSystem" />
```

Kết quả:
```xml
<Reference Include="System" />
<Reference Include="System.Core" />
<Reference Include="System.IO.Compression" />
<Reference Include="System.IO.Compression.FileSystem" />
<Reference Include="System.Xml.Linq" />
```

Lưu file.

### 3.3. Mở lại Visual Studio

Build → Thành công!

## 📡 API GITHUB

### Endpoint

```
GET https://api.github.com/repos/NguyenHaiHG/HOSONHCS/releases/latest
```

### Response mẫu

```json
{
  "tag_name": "v1.1.0",
  "name": "HOSONHCS v1.1.0",
  "body": "## Thay đổi\n- Thêm bảng kê tiền",
  "published_at": "2024-12-20T14:30:00Z",
  "assets": [
    {
      "name": "HOSONHCS_v1.1.0.zip",
      "browser_download_url": "https://github.com/.../HOSONHCS_v1.1.0.zip"
    }
  ]
}
```

## 🛡️ BẢO MẬT & AN TOÀN

### Dữ liệu người dùng KHÔNG bị mất

File `.bat` chỉ copy:
```batch
xcopy /Y /E /I "%extractPath%\*.exe" "%currentFolder%\"
xcopy /Y /E /I "%extractPath%\*.dll" "%currentFolder%\"
```

**KHÔNG** copy folder:
- ❌ `Customers\` (dữ liệu khách hàng)
- ❌ `To\` (dữ liệu tổ)
- ❌ `BangKe\` (bảng kê tiền)
- ❌ `*.json` (xinman.json, config.json)

### Rollback

Nếu update lỗi, file cũ vẫn còn:
```
HOSONHCS.exe.old  ← Bản backup
```

Đổi tên lại thành `HOSONHCS.exe` để khôi phục.

## 🧪 KIỂM TRA

### Test Release

1. Tạo release test với tag `v0.0.1-test`
2. Build app (version 0.0.0)
3. Chạy app → Sẽ thấy thông báo có update

### Test Download

```csharp
// Trong Form1, thêm button test
private async void btnTestUpdate_Click(object sender, EventArgs e)
{
    var updateInfo = await AutoUpdater.CheckForUpdateAsync(silent: false);
    if (updateInfo != null)
    {
        MessageBox.Show($"Tìm thấy: {updateInfo.Version}\nURL: {updateInfo.DownloadUrl}");
    }
}
```

## 🎯 USE CASE

### Kịch bản 1: Update bình thường

```
1. Tác giả release v1.1.0
2. User mở app (đang dùng v1.0.0)
3. App tự động kiểm tra → Có update
4. Hiện dialog → User chọn "Có"
5. Tải về (5s) → Cài đặt (2s)
6. App khởi động lại → Dùng v1.1.0
```

### Kịch bản 2: Không có internet

```
1. User mở app (không có internet)
2. CheckForUpdateAsync → WebException
3. Bỏ qua lỗi (silent)
4. App chạy bình thường
```

### Kịch bản 3: User từ chối update

```
1. Hiện dialog update
2. User chọn "Không"
3. App tiếp tục dùng bản cũ
4. Lần sau mở app → Lại hỏi
```

## 📊 THỐNG KÊ

GitHub cung cấp:
- Số lượt download
- Phiên bản nào được dùng nhiều
- Xem tại: Releases → Assets

## ❓ FAQ

**Q: Có tính phí không?**
A: Không! GitHub Releases miễn phí 100%.

**Q: Bandwidth giới hạn?**
A: Không giới hạn cho public repo.

**Q: Nếu GitHub down?**
A: App vẫn chạy bình thường, chỉ không update được.

**Q: Có thể dùng server riêng?**
A: Có, đổi `VERSION_CHECK_URL` sang URL server của bạn.

**Q: Bắt buộc update?**
A: Không, user có thể từ chối. Nếu muốn bắt buộc, thêm field `mandatory: true` trong JSON.

## 🔗 LINKS

- **Repo**: https://github.com/NguyenHaiHG/HOSONHCS
- **Releases**: https://github.com/NguyenHaiHG/HOSONHCS/releases
- **API Docs**: https://docs.github.com/en/rest/releases
