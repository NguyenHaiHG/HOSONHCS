# HÀM VÀ THUẬT TOÁN — HOSONHCS

> Tài liệu mô tả toàn bộ các hàm, phương thức và thuật toán được sử dụng trong ứng dụng.  
> Cập nhật lần cuối: 2025

---

## MỤC LỤC

1. [Khởi tạo & Vòng đời ứng dụng](#1-khởi-tạo--vòng-đời-ứng-dụng)
2. [Quản lý dữ liệu khách hàng (CRUD)](#2-quản-lý-dữ-liệu-khách-hàng-crud)
3. [Validation & Nghiệp vụ](#3-validation--nghiệp-vụ)
4. [Thuật toán tính toán nghiệp vụ](#4-thuật-toán-tính-toán-nghiệp-vụ)
5. [Xử lý template Word (OpenXML)](#5-xử-lý-template-word-openxml)
6. [Chuyển đổi PDF (Word Interop)](#6-chuyển-đổi-pdf-word-interop)
7. [Dữ liệu địa lý — TinhHelper & XinManEditor](#7-dữ-liệu-địa-lý--tinhhelper--xinmaneditor)
8. [Mã hóa AES-256 — ToanQuocEncryptor](#8-mã-hóa-aes-256--toanquocencryptor)
9. [Định dạng & UI Input](#9-định-dạng--ui-input)
10. [Bảng kê tiền — BangKeTien](#10-bảng-kê-tiền--bangketien)
11. [Ghi chú — Form1.GhiChu](#11-ghi-chú--form1ghichu)
12. [Chatbot — Form1.Chatbot](#12-chatbot--form1chatbot)
13. [Cập nhật tự động — AutoUpdater](#13-cập-nhật-tự-động--autoupdater)
14. [Thông báo ngày lễ — HolidayNoticeChecker](#14-thông-báo-ngày-lễ--holidaynoticechecker)
15. [Giao diện — UIStyler & Form1.UIStyle](#15-giao-diện--uistyler--form1uistyle)
16. [Form2 — Xuất hồ sơ nhóm](#16-form2--xuất-hồ-sơ-nhóm)
17. [Marquee & Hiệu ứng UI](#17-marquee--hiệu-ứng-ui)

---

## 1. Khởi tạo & Vòng đời ứng dụng

### `InitializeApp()` — Form1.cs
**Mục đích:** Điểm khởi tạo trung tâm, được gọi từ constructor Form1.

**Trình tự thực thi:**
```
1. Gắn event handlers cho tất cả buttons (btn01, btnDelete, btn03, btnGUQ, …)
2. Gắn TextChanged/KeyPress cho ô tên (Title Case auto-format)
3. Gắn KeyPress/TextChanged/Leave cho CCCD, SDT, nhân khẩu
4. Gắn event DateTimePicker (Leave validate, Enter show date, ValueChanged)
5. Gắn event ComboBox địa bàn (cascading: Tỉnh → PGD → Xã → Hội → Thôn → Tổ)
6. Gắn event tiền tệ (KeyPress + TextChanged format số tiền)
7. Cấu hình DataGridView (chế độ chọn, cột ẩn, readonly)
8. LoadXinManData() — nạp dữ liệu địa lý
9. LoadCustomersFromFiles() — nạp danh sách khách
10. InitializeDgvTotruong() + LoadBangKeFromFiles()
11. InitializeGhiChuTab() / InitializeChatbotTab() / InitializeTab6() / InitializeTab7()
12. BindGrid()
13. ApplyModernStyle()
14. CheckForUpdateOnStartup() [async]
15. AttachControls XinManEditor
16. InitializeMarquee()
17. Shown += CheckAndShowAsync (thông báo ngày lễ)
```

### `Form1_Load()` — Form1.cs
- Gọi `TinhHelper.PopulateComboBox(cbTinh)` và `cbTinhfix`
- Đăng ký cascading event cho cbTinhfix → cbpgdfix

---

## 2. Quản lý dữ liệu khách hàng (CRUD)

### `LoadCustomersFromFiles()` — Form1.cs
- Duyệt `Customers/*.json`
- Deserialize từng file thành `Customer`
- Nạp vào `BindingList<Customer> customers`

### `SaveCustomerToFile(Customer c)` — Form1.cs
**Thuật toán tên file:**
```
baseName = MakeFileSystemSafe(c.Hoten)
path = Customers/{baseName}.json
Nếu file tồn tại và thuộc cùng khách → ghi đè
Nếu trùng tên khác khách → thêm hậu tố _1, _2, …
```
Gọi `UpdateComputedFields(c)` trước khi serialize.

### `DeleteCustomerFiles(Customer c)` — Form1.cs
- Xóa file JSON trong `Customers/`
- Xóa folder hồ sơ trên Desktop `Hồ sơ NHCS/{HoTen}_{MM-yyyy}/`

### `ReadForm()` — Form1.cs
Đọc toàn bộ form → tạo `Customer` object.  
Thực hiện **validation nghiệp vụ inline:**
- Ngày tương lai → reset về hôm nay
- Gọi `CalcThoiHanCCCD(ngaysinh, ngaycap)`
- Throw Exception nếu CCCD đã hết hạn

### `PopulateForm(Customer c)` — Form1.cs
Điền ngược từ `Customer` → form controls.  
Dùng cờ `suppressComboChanged = true` để chặn cascading event khi load.

### `ClearForm()` — Form1.cs
Reset toàn bộ controls về trạng thái ban đầu.  
Đặt DateTimePicker về `CustomFormat = " "` (ẩn ngày).

### `UpsertCustomerInList(Customer customer)` — Form1.cs
- `editingIndex >= 0` → cập nhật phần tử hiện có
- `editingIndex < 0` → thêm mới vào cuối list

### `BindGrid()` — Form1.cs
Gán `BindingList<Customer>` làm DataSource cho DataGridView.  
Ẩn cột `_fileName`, đổi header cột `Hoten`.

---

## 3. Validation & Nghiệp vụ

### `ValidateRequiredFields()` — Form1.Validation.cs
Kiểm tra theo thứ tự:
1. Các ComboBox/TextBox bắt buộc → thêm vào `missingFields`
2. DateTimePicker bắt buộc (`CustomFormat == " "` = chưa nhập)
3. **Kiểm tra tổng tiền:** `Sotien1 + Sotien2 == Sotien`

```csharp
// Thuật toán kiểm tra tổng tiền
long total = ParseMoneyStringToLong(cbSotien.Text);
long s1    = ParseMoneyStringToLong(cbSotien1.Text);
long s2    = ParseMoneyStringToLong(cbSotien2.Text);
if (total > 0 && (s1 + s2) != total) → báo lỗi
```

Hiển thị 1 MessageBox duy nhất liệt kê tất cả trường thiếu.

### `ValidateDuplicateCccdSdt()` — Form1.Validation.cs
Kiểm tra:
1. CCCD chính không trùng CCCD người thừa kế trong cùng hồ sơ
2. Các CCCD người thừa kế không trùng nhau
3. CCCD chính không trùng với bất kỳ khách nào khác trong DB (bỏ qua bản ghi đang sửa)
4. SDT không trùng với bất kỳ khách nào khác trong DB

**Chuẩn hóa SDT:** Dùng `DigitsOnly()` → bỏ dấu chấm trước khi so sánh.

---

## 4. Thuật toán tính toán nghiệp vụ

### `CalcThoiHanCCCD(DateTime ngaysinh, DateTime ngaycap)` — Form1.cs
**Thuật toán tính thời hạn CCCD theo Luật CCCD VN:**

```
tuổi_hành_chính = năm_cấp − năm_sinh   ← KHÔNG điều chỉnh theo tháng/ngày

Milestones = {25, 40, 60}

foreach milestone m:
    if tuổi_hành_chính < m:
        hết_hạn = ngày/tháng sinh, năm = năm_sinh + m
        xử lý 29/02 trên năm không nhuận → đổi thành 28/02
        return hết_hạn

return DateTime.MinValue  → "không thời hạn" (≥ 60 tuổi)
```

**Bảng kết quả:**

| Tuổi lúc cấp | Hết hạn tại |
|---|---|
| < 25 | Sinh nhật 25 tuổi |
| 25 – 39 | Sinh nhật 40 tuổi |
| 40 – 59 | Sinh nhật 60 tuổi |
| ≥ 60 | Không thời hạn |

> ⚠️ **Lưu ý quan trọng:** Dùng `năm_cấp − năm_sinh` (không trừ 1 theo birthday thông thường).  
> Ví dụ: sinh 10/3/1986, cấp 03/03/2026 → tuổi = 2026−1986 = **40** → nhóm 40-59 → hết hạn 10/3/2046.

### `NumberToVietnameseWords(long number)` — Form1.cs
**Thuật toán chuyển số → chữ tiếng Việt:**

```
1. Tách số thành các nhóm 3 chữ số (đơn vị: nghìn, triệu, tỷ)
2. Duyệt từ nhóm cao nhất → thấp nhất
3. Mỗi nhóm 3 chữ số:
   - hàng_trăm: "X trăm"
   - hàng_chục: "mười" hoặc "X mươi"
   - hàng_đơn_vị: đặc biệt "mốt" (1 sau chục ≥ 1), "lăm" (5 sau chục ≥ 1)
   - Nếu chục = 0, đơn vị > 0: thêm "lẻ"
4. Nếu nhóm = 0 nhưng có nhóm thấp hơn ≠ 0: thêm tên đơn vị để giữ vị trí
5. Kết hợp với tên đơn vị: "nghìn", "triệu", "tỷ"
```

**Ví dụ:** 50.000.000 → "năm mươi triệu"

### `UpdateComputedFields(Customer c)` — Form1.cs
```
Sotientong = Vtc + Sotien  (nếu Sotientong chưa có)
Sotienchu  = NumberToVietnameseWords(Sotientong) + " đồng"
```

### `ParseMoneyStringToLong(string s)` — Form1.cs
- Lấy tất cả ký tự số từ chuỗi
- `long.TryParse(digits)` → trả về 0 nếu lỗi

### `CalculateNgayDenHan()` — Form1.cs
```
if dateGn.Checked:
    months = ParseThoiHanThang(cbThoihanvay.Text)
    dateDH.Value = dateGn.Value.AddMonths(months)
```

### `ParseThoiHanThang(string)` — Form1.cs
Regex `\d+` → trích số tháng từ chuỗi như "60 tháng" → 60.

---

## 5. Xử lý template Word (OpenXML)

### `CreateProfileFromTemplate(Customer c, bool include03)` — Form1.cs
```
1. Xác định danh sách template cần xuất (GetTemplateNamesForCustomer)
2. foreach template:
   a. ResolveTemplatePath()  → tìm file .docx
   b. IsDocxFile()           → kiểm tra hợp lệ
   c. File.Copy → destDoc    → copy vào thư mục output
   d. ReplacePlaceholdersInWord(destDoc, c)
3. ConvertDocxListToPdf(createdFiles)
```

### `GetTemplateNamesForCustomer(Customer c, bool include03)` — Form1.cs
| Chương trình | Template chọn |
|---|---|
| SXKD | `01 SXKD.docx` |
| GQVL | `01 GQVL.docx` |
| Mặc định | `01 HN.docx` |
| include03=true | + `03 DS.docx` |

### `ResolveTemplatePath(string templateFileName)` — Form1.cs
**Thứ tự ưu tiên tìm template:**
```
1. {BaseDir}/Templates/{name}
2. {BaseDir}/{name}
3. Directory.EnumerateFiles(BaseDir, name, AllDirectories)
4. Assembly.GetManifestResourceStream (Embedded Resource)
```
Kết quả được cache trong `Dictionary<string, string> templatePathCache`.

### `ReplacePlaceholdersInWord(string docPath, Customer c)` — Form1.cs
```
1. EnsureSotienchuFromNumeric()
2. Tạo Dictionary<string, string> replacements (100+ placeholder)
3. Gọi ReplacePlaceholdersUsingOpenXml(docPath, replacements, c)
```

**Placeholders chính:**  
`{{hoten}}`, `{{socccd}}`, `{{cccd12}}`, `{{ngaysinh}}`, `{{ngaycap}}`, `{{noicap}}`,  
`{{xa}}`, `{{thon}}`, `{{hoi}}`, `{{totruong}}`, `{{chuongtrinh}}`, `{{sotien}}`,  
`{{sotienchu}}`, `{{thoihanvay}}`, `{{ngaylaphs}}`, `{{thoihancccd}}`, `{{khau}}`, ...

### `ReplacePlaceholdersUsingOpenXml(string docPath, ...)` — Form1.cs
**Thuật toán thay thế 3 lớp:**
```
1. ProcessCCCD12Placeholder() — xử lý đặc biệt ô bảng 12 ô số CCCD
2. Mở WordprocessingDocument (chế độ edit)
3. Foreach part (Main, Headers, Footers, Footnotes, Comments):
   a. ReplaceInTableCells()  — thay trong từng Text node của ô bảng
   b. ReplaceInParagraphs()  — thay trong từng Text node của đoạn văn
   c. TryReplacePlaceholdersAcrossRuns() — thay placeholder bị tách qua nhiều Run
   d. Quét Text trực tiếp (lần cuối, backup)
```

### `TryReplacePlaceholdersAcrossRuns(OpenXmlPart part, ...)` — Form1.cs
**Vấn đề:** Word thường tách `{{hoten}}` thành nhiều `<w:r><w:t>` riêng biệt.  
**Giải pháp:**
```
Foreach placeholder key:
  1. Xây regex: \{\{\s*{token}\s*\}\}
  2. Duyệt sliding window qua danh sách Text nodes
  3. Nối text tích lũy cho đến khi tìm thấy match
  4. Tính startNode/endOffset và endNode/endOffset
  5. Đặt text node đầu = prefix + replacement + suffix
  6. Xóa trắng các text node còn lại trong khoảng
```

### `ProcessCCCD12Placeholder(MainDocumentPart, string cccd)` — Form1.cs
- Tìm ô bảng chứa `{{cccd12}}`
- Điền 12 chữ số CCCD vào 12 ô liền kề (mỗi ô 1 chữ số)

### `ReplaceIgnoreCase(string input, string oldValue, string newValue)` — Form1.cs
- Nếu oldValue là `{{token}}`: dùng regex `\{\{\s*{token}\s*\}\}` (cho phép whitespace)
- Else: `Regex.Replace(input, Regex.Escape(oldValue), newValue, IgnoreCase)`

---

## 6. Chuyển đổi PDF (Word Interop)

### `ConvertDocxListToPdf(List<string> docxPaths)` — Form1.cs
```
1. Tạo 1 instance Word.Application (Visible=false) dùng chung
2. foreach docxPath:
   ConvertDocxToPdf(wordApp, docxPath)
3. wordApp.Quit(false)
4. Marshal.ReleaseComObject(wordApp)
5. GC.Collect() + GC.WaitForPendingFinalizers()
6. Xóa tất cả file .docx nguồn
7. Trả về danh sách .pdf
```

### `ConvertDocxToPdf(Word.Application, string docxPath)` — Form1.cs
```
doc = wordApp.Documents.Open(docxPath)
doc.ExportAsFixedFormat(pdfPath, wdExportFormatPDF)
doc.Close(false)
return pdfPath
```

---

## 7. Dữ liệu địa lý — TinhHelper & XinManEditor

### `TinhHelper.BuildCache()` — TinhHelper.cs
```
1. ToanQuocEncryptor.GetJson() → giải mã toanquoc.enc
2. JArray.Parse(json) → duyệt từng tỉnh
3. Tạo TinhModel { tinh, pgds[] }
4. Lọc "nan" khỏi groups (giữ assoc "nan" cho UI)
5. Nạp vào Dictionary<string, TinhModel> _cache
```

### `TinhHelper.PopulateComboBox(ComboBox cb)` — TinhHelper.cs
- Lấy danh sách tên tỉnh từ cache
- Thêm vào ComboBox

### `CbPGD_SelectedIndexChanged()` / `CbXa_SelectedIndexChanged()` / ... — Form1.cs
**Cascading load địa lý:**
```
Tỉnh →(load pgds)→ PGD →(load communes)→ Xã
   →(load associations)→ Hội →(load villages)→ Thôn
   →(load groups)→ Tổ (tổ trưởng)
```
Dùng cờ `suppressComboChanged` để tránh đệ quy.

### `XinManEditor` — XinManEditor.cs
- Login bằng `haihg` / `Haihg23`
- Hiển thị dữ liệu PGD trong DataGridView dạng bảng phẳng
- Tìm kiếm: `TxtSearch_TextChanged` → highlight row match
- Lưu: ghi lại vào `toanquoc.enc` qua `ToanQuocEncryptor`

---

## 8. Mã hóa AES-256 — ToanQuocEncryptor

### `ToanQuocEncryptor.GetJson()` — ToanQuocEncryptor.cs
```
1. Nếu cache (_decryptedJson != null) → trả về ngay
2. Nếu tồn tại toanquoc.enc → Decrypt()
3. Nếu tồn tại toanquoc.json → Encrypt() → lưu .enc → xóa .json
4. Cache kết quả trong _decryptedJson (RAM)
```

### `Encrypt(string json)` / `Decrypt(byte[] data)` — ToanQuocEncryptor.cs
- Thuật toán: **AES-256-CBC**
- Key: `AppSecrets.AesKey` (32 bytes)
- IV: `AppSecrets.AesIV` (16 bytes)
- Input: UTF-8 string → `AesCryptoServiceProvider.CreateEncryptor()`
- Output: byte array (không đọc được nếu không có Key+IV)

---

## 9. Định dạng & UI Input

### `CbMoney_TextChanged(object sender, EventArgs e)` — Form1.cs
**Áp dụng cho:** `cbSotien`, `cbSotien1`, `cbSotien2`
```
1. Lấy digits = chỉ các ký tự số
2. value = ParseMoneyStringToLong(digits)
3. formatted = String.Format("{0:N0}", value, InvariantCulture).Replace(",", ".")
4. Nếu formatted != text hiện tại → cập nhật (dùng suppressMoneyChange)
```
**Kết quả:** `50000000` → `50.000.000`

### `TxtSdt_TextChanged()` — Form1.cs
**Format số điện thoại theo pattern 4-3-3:**
```
digits = chỉ số (tối đa 10)
≤4 → digits
≤7 → XXXX.XXX
=10 → XXXX.XXX.XXX
```

### `TxtName_TextChanged()` / `TxtName_Leave()` — Form1.cs
- `TextChanged`: `CapitalizeWords()` — viết hoa chữ đầu mỗi từ realtime
- `Leave`: `ToTitleCase()` — dùng `CultureInfo("vi-VN").TextInfo.ToTitleCase()` để dọn dẹp cuối

### `TxtCccd_TextChanged()` / `TxtCccd_Leave()` — Form1.cs
- `TextChanged`: chỉ giữ ký tự số
- `Leave`: kiểm tra phải đúng 12 chữ số

### `DatePicker_Enter()` — Form1.cs
- Khi click vào DateTimePicker đang ở trạng thái `CustomFormat = " "` → hiện ngày hôm nay

### `DateNgaycapCCCD_ValueChanged()` — Form1.cs
```
if ngaycap ≥ 01/07/2024 → cbNoicap = "Bộ Công an"
else                     → cbNoicap = "Cục CSQLHC về TTXH"
+ Gọi DateNgaysinh_ValueChanged() để tính lại thời hạn CCCD
```

### `DateNgaysinh_ValueChanged()` — Form1.cs
```
ngaycap = dateNgaycapCCCD.Value (nếu đã nhập) else DateTime.Today
thoihan = CalcThoiHanCCCD(ngaysinh, ngaycap)
if thoihan == MinValue → datendhcccd hiện "'không thời hạn'"
else                   → datendhcccd hiện "dd/MM/yyyy"
```

### `ApplyPhuonganState(string phuongan)` — Form1.cs
Tự động điều chỉnh trạng thái `cbmucdich1` / `cbmucdich2` theo `cbPhuongan`:
- "Nâng cấp, sửa chữa CTNS, CTVS" → điền + khóa cả 2
- "Xây mới CTNS, CTVS" → điền + khóa cả 2
- Item hợp lệ khác → điền cbmucdich1 = phuongan, khóa cbmucdich2
- Nhập tay tự do → mở cả 2

### `CbChuongtrinh_SelectedIndexChanged()` — Form1.cs
Tự động điền + khóa / mở `cbDoituong` dựa theo chương trình:
| Chương trình | Đối tượng |
|---|---|
| Hộ nghèo | "Hộ nghèo" (khóa) |
| Hộ cận nghèo | "Hộ cận nghèo" (khóa) |
| Hộ mới thoát nghèo | "Hộ mới thoát nghèo" (khóa) |
| SXKD | "Hộ GĐ SXKD VKK" (khóa) |
| Nước sạch VSMTNT | "HGĐ cư trú tại VNT" (khóa) |
| GQVL | ["Người lao động", "NLĐ là người DTTS"] (mở dropdown) |

---

## 10. Bảng kê tiền — BangKeTien

### `InitializeDataGridView(DataGridView dgv)` — BangKeTien.cs
Tạo cấu trúc 3 cột: Mệnh giá (readonly), Số lượng (editable), Thành tiền (computed).  
Thêm 9 dòng tương ứng 9 mệnh giá VNĐ.

### `GetTongThanhTien(DataGridView)` — BangKeTien.cs
```
Σ (MenhGia[i] × SoLuong[i]) với i = 0..8
```

### `GetSoTienSoSach(DataGridView)` / `GetChenhLech(DataGridView)` — BangKeTien.cs
- Đọc dòng cuối (dòng sổ sách nhập tay)
- `ChenhLech = TongTien - SoTienSoSach`

### `SaveToFile(BangKeData)` / `LoadAllFromFiles()` — BangKeTien.cs
- Serialize/Deserialize JSON trong `BangKe/*.json`

---

## 11. Ghi chú — Form1.GhiChu

### `InitializeGhiChuTab()` — Form1.GhiChu.cs
Xây dựng UI động (không dùng Designer): Toolbar + SplitContainer + ListView + RichTextBox.

### `SaveNote()` / `DeleteNote()` — Form1.GhiChu.cs
- Serialize `NoteItem` → `Notes/{Id}.json`
- Sắp xếp: Pinned trước, sau đó theo `UpdatedAt` desc

### `SearchNotes(string query)` — Form1.GhiChu.cs
- So sánh `query` với `Title` + `Content` (case-insensitive)
- Cập nhật ListView realtime theo từng ký tự gõ

---

## 12. Chatbot — Form1.Chatbot

### `InitializeChatbotTab()` — Form1.Chatbot.cs
Xây dựng UI 2 panel (chat trái + training phải) hoàn toàn bằng code.

### `ProcessUserMessage(string userText)` — Form1.Chatbot.cs
**Thuật toán tìm câu trả lời:**
```
1. Chuẩn hóa input: lowercase, bỏ dấu
2. foreach KnowledgeItem (sắp xếp Priority desc):
   score = 0
   foreach keyword in item.Keywords:
     if normalizedInput.Contains(normalizedKeyword): score++
   if score > bestScore: bestAnswer = item
3. Nếu bestScore > 0: trả lời từ KnowledgeItem
4. Else: trả lời mặc định "Xin lỗi, tôi chưa có câu trả lời…"
```

### `SaveKnowledgeItem()` / `LoadKnowledge()` — Form1.Chatbot.cs
- Serialize/Deserialize JSON trong `ChatBot/knowledge/*.json`

### `SaveSession()` / `LoadHistory()` — Form1.Chatbot.cs
- Serialize/Deserialize `ChatSession` JSON trong `ChatBot/history/*.json`

---

## 13. Cập nhật tự động — AutoUpdater

### `CheckForUpdateAsync(bool silent)` — AutoUpdater.cs
```
1. HttpWebRequest → GET https://api.github.com/repos/NguyenHaiHG/HOSONHCS/releases/latest
2. Parse JSON response → lấy tag_name, body, assets[0].browser_download_url
3. So sánh version: latestVersion > CurrentVersion (Assembly version)
4. Nếu có update → trả về UpdateInfo
5. Nếu không → trả về null
```
**Protocol:** TLS 1.2 (`ServicePointManager.SecurityProtocol`)

### `ShowUpdateDialogAsync(UpdateInfo)` — AutoUpdater.cs
```
1. Hiện MessageBox thông tin phiên bản mới
2. Nếu user đồng ý:
   a. Download file mới về temp
   b. Chạy installer / copy exe
   c. Application.Exit()
3. Nếu từ chối → return false → App tắt (bắt buộc update)
```

---

## 14. Thông báo ngày lễ — HolidayNoticeChecker

### `CheckAndShowAsync(Form parentForm)` — HolidayNoticeChecker.cs
```
1. GET https://raw.githubusercontent.com/NguyenHaiHG/HOSONHCS/master/holiday_notice.json
2. Deserialize → List<HolidayNoticeMessage>
3. foreach message:
   if message.IsActiveToday() && !AlreadyShownToday(message.Id):
     Form.Invoke → HolidayNoticeForm.ShowDialog()
     MarkAsShown(message.Id)
```

### `HolidayNoticeMessage.IsActiveToday()` — HolidayNotice.cs
```
today >= DateTime.Parse(StartDate) && today <= DateTime.Parse(EndDate) && Enabled
```

### Cơ chế chống hiển thị lặp
File `.notice_shown` lưu danh sách ID đã hiện hôm nay (1 dòng/ID).  
Reset khi sang ngày mới.

---

## 15. Giao diện — UIStyler & Form1.UIStyle

### `UIStyler` (static class) — UIStyler.cs
Cung cấp bộ màu **Teal Theme:**

| Hằng số | Màu RGB | Dùng cho |
|---|---|---|
| `BgMain` | (13, 75, 75) | Nền chính |
| `BgPanel` | (18, 95, 92) | Panel/Sidebar |
| `BgCard` | (24, 112, 108) | Card/Nội dung |
| `Primary` | (100, 235, 228) | Accent sáng |
| `TextMain` | (232, 250, 248) | Chữ chính |
| `BtnGreen` | (45, 195, 85) | Nút OK/Lưu |
| `BtnRed` | (235, 70, 85) | Nút Xóa |
| `BtnBlue` | (50, 130, 255) | Nút xuất |

### `ApplyModernStyle()` — Form1.UIStyle.cs
Duyệt đệ quy tất cả controls → áp dụng màu, font, border radius theo UIStyler.

---

## 16. Form2 — Xuất hồ sơ nhóm

### `Form2(List<Customer> selected)` — Form2.cs
- Nhận danh sách khách đã chọn từ Form1 (cùng PGD, Chương trình, Tổ)
- Hiển thị trong DataGridView để chỉnh sửa nhóm

### `Btn03Group_Click()` → Form1.cs → `new Form2(selected).ShowDialog()`
Lọc: chỉ cho phép chọn khách cùng PGD + Chương trình + Tổ trưởng + Xã + Thôn.

### `ExportGroup()` — Form2.cs
```
1. Đọc dữ liệu từ DataGridView (tối đa 5 tổ viên)
2. Serialize thành ToData → Form2State.json (lưu trạng thái)
3. ReplacePlaceholdersInWordForGroup(docPath, c, entriesText)
4. Tạo file 03 DS nhóm với danh sách nhiều tổ viên
5. ConvertDocxListToPdf()
```

### `CalcThoiHanCCCD()` — Form2.cs
Cùng thuật toán như Form1 (duplicate để Form2 độc lập).

---

## 17. Marquee & Hiệu ứng UI

### `InitializeMarquee()` — Form1.cs
```
marqueeText = "     PHẦN MỀM TẠO HỒ SƠ VAY VỐN     "
label14MarqueeText = "     {label14.Text}     "
marqueeTimer.Interval = 100ms → Start()
```

### `MarqueeTimer_Tick()` — Form1.cs
```
Mỗi 100ms:
  marqueePosition = (marqueePosition + 1) % marqueeText.Length
  this.Text  = marqueeText.Substring(position) + marqueeText.Substring(0, position)
  label14.Text = (tương tự cho label14MarqueeText)
```
**Hiệu ứng:** Text xoay vòng từ phải sang trái trên title bar và label14.

---

## PHỤ LỤC — Danh sách hàm theo file

| File | Hàm chính |
|---|---|
| `Form1.cs` | `InitializeApp`, `ReadForm`, `PopulateForm`, `ClearForm`, `BtnSave_Click`, `CalcThoiHanCCCD`, `NumberToVietnameseWords`, `ReplacePlaceholdersInWord`, `ConvertDocxListToPdf`, `MarqueeTimer_Tick` |
| `Form1.Validation.cs` | `ValidateRequiredFields`, `ValidateDuplicateCccdSdt`, `DigitsOnly` |
| `Form1.Chatbot.cs` | `InitializeChatbotTab`, `ProcessUserMessage`, `SaveKnowledgeItem`, `LoadKnowledge`, `SaveSession` |
| `Form1.GhiChu.cs` | `InitializeGhiChuTab`, `SaveNote`, `DeleteNote`, `SearchNotes` |
| `Form1.UIStyle.cs` | `ApplyModernStyle` |
| `Form1.DocViewer.cs` | `InitializeTab7` |
| `Form1.ComboBoxLogic.cs` | *(placeholder — logic trong Form1.cs)* |
| `BangKeTien.cs` | `InitializeDataGridView`, `GetTongThanhTien`, `GetChenhLech`, `SaveToFile`, `LoadAllFromFiles` |
| `TinhHelper.cs` | `BuildCache`, `PopulateComboBox`, `LoadTinhModel`, `GetProvinceNames` |
| `XinManEditor.cs` | `AttachControls`, `LoadFromPgdEntry`, `BtnLogin_Click`, `TxtSearch_TextChanged`, `Logout` |
| `ToanQuocEncryptor.cs` | `GetJson`, `Encrypt`, `Decrypt` |
| `AutoUpdater.cs` | `CheckForUpdateAsync`, `ShowUpdateDialogAsync` |
| `HolidayNoticeChecker.cs` | `CheckAndShowAsync`, `AlreadyShownToday`, `MarkAsShown` |
| `Form2.cs` | `ExportGroup`, `CalcThoiHanCCCD`, `ReplacePlaceholdersInWordForGroup` |
| `UIStyler.cs` | *(static color constants)* |

---

*File được tạo tự động bởi GitHub Copilot — HOSONHCS v1.x*
