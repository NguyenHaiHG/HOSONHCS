# CẤU TRÚC DỮ LIỆU — HOSONHCS

> Tài liệu mô tả toàn bộ các model, class, enum và cấu trúc lưu trữ dữ liệu của ứng dụng.
> Cập nhật lần cuối: 2025

---

## MỤC LỤC

1. [Model chính — Customer](#1-model-chính--customer)
2. [Dữ liệu địa lý — XinMan / Tinh](#2-dữ-liệu-địa-lý--xinman--tinh)
3. [Bảng kê tiền — BangKeData](#3-bảng-kê-tiền--bangkedata)
4. [Dữ liệu tổ nhóm — ToData](#4-dữ-liệu-tổ-nhóm--todata)
5. [Ghi chú — NoteItem](#5-ghi-chú--noteitem)
6. [Chatbot — KnowledgeItem / ChatMessage / ChatSession](#6-chatbot--knowledgeitem--chatmessage--chatsession)
7. [Thông báo ngày lễ — HolidayNoticeMessage](#7-thông-báo-ngày-lễ--holidaynoticemessage)
8. [Cập nhật tự động — UpdateInfo](#8-cập-nhật-tự-động--updateinfo)
9. [Lịch sử xuất hồ sơ — ExportHistory](#9-lịch-sử-xuất-hồ-sơ--exporthistory)
10. [Cấu trúc file & thư mục lưu trữ](#10-cấu-trúc-file--thư-mục-lưu-trữ)
11. [Luồng dữ liệu tổng thể](#11-luồng-dữ-liệu-tổng-thể)

---

## 1. Model chính — `Customer`

**File:** `Form1.cs` (inner class trong namespace HOSONHCS)  
**Lưu trữ:** `Customers/*.json` (mỗi khách 1 file)  
**Serializer:** `Newtonsoft.Json`

```
Customer
├── Thông tin cá nhân
│   ├── Hoten          : string   — Họ và tên (Title Case)
│   ├── Socccd         : string   — Số CCCD (12 chữ số)
│   ├── GioiTinh       : string   — Nam / Nữ
│   ├── Nhandang       : string   — Nhân dạng (CCCD / CMND)
│   ├── Ngaysinh       : DateTime — Ngày sinh
│   ├── Dantoc         : string   — Dân tộc
│   ├── Sdt            : string   — Số điện thoại (format 0xxx.xxx.xxx)
│   └── Nhankhau       : string   — Số nhân khẩu (tối đa 2 chữ số)
│
├── Thông tin CCCD
│   ├── Ngaycap        : DateTime — Ngày cấp CCCD
│   ├── Noicap         : string   — Nơi cấp (Bộ Công an / Cục CSQLHC)
│   ├── Thoihancccd    : DateTime — Ngày hết hạn CCCD (tính từ ngày sinh + ngày cấp)
│   └── ThoihancccdText: string   — "không thời hạn" hoặc "dd/MM/yyyy"
│
├── Thông tin địa bàn
│   ├── Tinh           : string   — Tỉnh
│   ├── PGD            : string   — Phòng giao dịch
│   ├── Xa             : string   — Xã
│   ├── Thon           : string   — Thôn
│   ├── Hoi            : string   — Hội quản lý
│   ├── Totruong       : string   — Tên tổ trưởng
│   └── To             : string   — (dự phòng, thường rỗng)
│
├── Thông tin khoản vay
│   ├── Chuongtrinh    : string   — Chương trình vay
│   ├── Vtc            : string   — Vốn tự có (format tiền VN)
│   ├── Phuongan       : string   — Phương án vay
│   ├── Thoihanvay     : string   — Thời hạn vay (vd: "60 tháng")
│   ├── Phanky         : string   — Phân kỳ trả nợ
│   ├── Sotien         : string   — Tổng số tiền vay (format 50.000.000)
│   ├── Sotien1        : string   — Số tiền mục đích 1
│   ├── Sotien2        : string   — Số tiền mục đích 2
│   ├── Sotientong     : string   — Tổng (Vtc + Sotien), tính tự động
│   ├── Sotienchu      : string   — Số tiền bằng chữ tiếng Việt
│   ├── Soluong1       : string   — Số lượng mục đích 1
│   ├── Soluong2       : string   — Số lượng mục đích 2
│   ├── Mucdich1       : string   — Mục đích vay vốn 1
│   ├── Mucdich2       : string   — Mục đích vay vốn 2
│   ├── Doituong1      : string   — Đối tượng thụ hưởng
│   └── Doituong2      : string   — Đối tượng thụ hưởng 2
│
├── Thông tin thời gian hồ sơ
│   ├── Ngaylaphs      : DateTime — Ngày lập hồ sơ
│   ├── Ngaygiaingaan  : DateTime — Ngày giải ngân
│   └── Ngaydenhan     : DateTime — Ngày đến hạn (tính tự động)
│
├── Người thừa kế (NTK) — tối đa 3 người
│   ├── Ntk1/2/3       : string   — Họ tên người thừa kế
│   ├── CccdNtk1/2/3   : string   — Số CCCD người thừa kế
│   ├── Namsinh1/2/3   : string   — Ngày sinh NTK (dd/MM/yyyy)
│   └── Qh1/2/3        : string   — Quan hệ với khách hàng
│
└── Metadata
    └── _fileName      : string   — Tên file JSON lưu trữ [JsonIgnore]
```

**Ràng buộc nghiệp vụ:**
- `Sotien1 + Sotien2 == Sotien` (bắt buộc khi cả hai có giá trị)
- `Socccd` phải đúng 12 chữ số
- `Sdt` phải đúng 10 chữ số
- `ThoihancccdText`: tính theo tuổi hành chính VN tại ngày cấp

---

## 2. Dữ liệu địa lý — XinMan / Tinh

**File:** `TinhHelper.cs`, `Form1.cs`  
**Nguồn:** `toanquoc.enc` (AES-256) → giải mã thành JSON trong RAM

### 2a. Cấu trúc phân cấp địa lý

```
TinhModel                          ← 1 tỉnh
├── tinh    : string               — Tên tỉnh
└── pgds    : List<TinhPgdEntry>   — Danh sách phòng giao dịch
    └── TinhPgdEntry
        ├── pgd      : string      — Tên PGD
        └── communes : List<Commune>

Commune (Xã)
├── name         : string
├── associations : List<Association>   — Danh sách Hội
│   └── Association
│       ├── name           : string   — Tên Hội ("nan" = không có hội)
│       ├── code           : string
│       ├── villages       : List<Village>
│       └── managedVillages: List<string>
└── villages     : List<Village>       — Thôn trực thuộc xã

Village (Thôn)
├── name   : string
└── groups : List<string>             — Danh sách tên tổ trưởng
```

### 2b. XinManModel (legacy, PGD đơn lẻ)

```
XinManModel
├── pgd      : string
└── communes : List<Commune>
```

### 2c. Cache & Mã hóa

| Thành phần | Mô tả |
|---|---|
| `toanquoc.enc` | File AES-256 chứa toàn bộ dữ liệu địa lý |
| `AppSecrets.AesKey` | Key 32 bytes (256-bit) |
| `AppSecrets.AesIV` | IV 16 bytes (128-bit) |
| `TinhHelper._cache` | Dictionary<string, TinhModel> — cache trong RAM |

---

## 3. Bảng kê tiền — `BangKeData`

**File:** `BangKeData.cs`, `BangKeTien.cs`  
**Lưu trữ:** `BangKe/*.json`

```
BangKeData
├── Totruong    : string                    — Tên tổ trưởng
├── ChiTiet     : Dictionary<long, long>    — Mệnh giá → Số lượng
│                  Key  : long (500000, 200000, 100000, 50000,
│                               20000, 10000, 5000, 2000, 1000)
│                  Value: long (số tờ)
├── TongTien    : long    — Tổng tiền mặt = Σ(mệnh giá × số lượng)
├── SoTienSoSach: long    — Số tiền theo sổ sách
├── ChenhLech   : long    — TongTien - SoTienSoSach
├── NgayTao     : DateTime
└── _fileName   : string
```

**Mệnh giá hỗ trợ (VNĐ):**  
`500.000 | 200.000 | 100.000 | 50.000 | 20.000 | 10.000 | 5.000 | 2.000 | 1.000`

---

## 4. Dữ liệu tổ nhóm — `ToData`

**File:** `ToData.cs`  
**Dùng bởi:** `Form2.cs` (xuất hồ sơ nhóm)  
**Lưu trữ:** `Form2State.json`

```
ToData
├── Thông tin chung
│   ├── Pgd, Xa, Thon, Totruong, Chuongtrinh : string
│   ├── NgayXuat    : DateTime
│   └── SoThanhVien : int (1-5)
│
└── Dữ liệu 5 tổ viên (suffix 1→5)
    ├── Kh1..5      : string  — Họ tên
    ├── Tien1..5    : string  — Số tiền vay
    ├── Md1..5      : string  — Mục đích / Phương án
    ├── Time1..5    : string  — Thời hạn vay
    └── Dt1..5      : string  — Đối tượng thụ hưởng
```

---

## 5. Ghi chú — `NoteItem`

**File:** `NoteItem.cs`, `Form1.GhiChu.cs`  
**Lưu trữ:** `Notes/*.json` (mỗi ghi chú 1 file)

```
NoteItem
├── Id        : string    — GUID (Guid.NewGuid().ToString("N"))
├── Title     : string    — Tiêu đề
├── Content   : string    — Nội dung (hỗ trợ xuống dòng)
├── Category  : string    — Danh mục (mặc định "Chung")
├── IsPinned  : bool      — Ghim lên đầu
├── CreatedAt : DateTime
└── UpdatedAt : DateTime
```

---

## 6. Chatbot — `KnowledgeItem` / `ChatMessage` / `ChatSession`

**File:** `KnowledgeItem.cs`, `ChatMessage.cs`, `Form1.Chatbot.cs`  
**Lưu trữ:**  
- Kiến thức: `ChatBot/knowledge/*.json`  
- Lịch sử: `ChatBot/history/*.json`

### KnowledgeItem (cơ sở kiến thức chatbot)

```
KnowledgeItem
├── Id        : string    — GUID
├── Category  : string    — Danh mục câu hỏi
├── Question  : string    — Câu hỏi mẫu
├── Keywords  : string[]  — Từ khóa để matching
├── Answer    : string    — Câu trả lời
├── Priority  : int       — Độ ưu tiên (1-10, cao hơn = ưu tiên hơn)
├── IsActive  : bool      — Bật/tắt
└── CreatedAt : DateTime
```

### ChatMessage

```
ChatMessage
├── Role      : string   — "user" | "bot"
├── Content   : string   — Nội dung tin nhắn
└── Timestamp : DateTime
```

### ChatSession (1 phiên chat)

```
ChatSession
├── Id       : string              — GUID
├── Date     : DateTime
└── Messages : List<ChatMessage>
```

---

## 7. Thông báo ngày lễ — `HolidayNoticeMessage`

**File:** `HolidayNotice.cs`, `HolidayNoticeChecker.cs`, `HolidayNoticeForm.cs`  
**Nguồn:** `holiday_notice.json` (fetch từ GitHub raw)

```
HolidayNoticeMessage
├── Id        : string  — ID duy nhất (tránh hiện lại trong ngày)
├── Title     : string  — Tiêu đề popup
├── Content   : string  — Nội dung (hỗ trợ \n)
├── Emoji     : string  — Emoji hiển thị header
├── StartDate : string  — "yyyy-MM-dd" — ngày bắt đầu hiệu lực
├── EndDate   : string  — "yyyy-MM-dd" — ngày kết thúc
└── Enabled   : bool    — Bật/tắt thông báo

HolidayNoticeData (root JSON)
└── messages  : List<HolidayNoticeMessage>
```

**File `.notice_shown`:** Lưu ID các thông báo đã hiển thị hôm nay (tránh popup lặp lại).

---

## 8. Cập nhật tự động — `UpdateInfo`

**File:** `AutoUpdater.cs`  
**Nguồn:** GitHub Releases API  
`https://api.github.com/repos/NguyenHaiHG/HOSONHCS/releases/latest`

```
UpdateInfo
├── Version     : string  — Phiên bản mới (vd: "1.2.3")
├── DownloadUrl : string  — URL tải file .exe / .zip
├── ReleaseNotes: string  — Ghi chú release
└── PublishedAt : DateTime
```

**So sánh phiên bản:** `Assembly.GetExecutingAssembly().GetName().Version`  
Format: `{Major}.{Minor}.{Build}`

---

## 9. Lịch sử xuất hồ sơ — `ExportHistory`

**File:** `Form2.cs`

```
ExportHistory
├── Xa          : string
├── Thon        : string
├── Totruong    : string
├── Chuongtrinh : string
├── NgayXuat    : DateTime
└── SoKhach     : int
```

---

## 10. Cấu trúc File & Thư mục lưu trữ

```
{BaseDirectory}/                      ← Thư mục chứa file .exe
│
├── Customers/                        ← Hồ sơ khách hàng
│   ├── NguyenVanA.json              ← 1 file = 1 Customer
│   └── TranThiB.json
│
├── BangKe/                           ← Bảng kê tiền
│   └── TotruongX.json               ← 1 file = 1 BangKeData
│
├── Notes/                            ← Ghi chú
│   └── {GUID}.json                  ← 1 file = 1 NoteItem
│
├── ChatBot/
│   ├── knowledge/                    ← Cơ sở kiến thức chatbot
│   │   └── {GUID}.json             ← 1 file = 1 KnowledgeItem
│   └── history/                     ← Lịch sử hội thoại
│       └── {GUID}.json             ← 1 file = 1 ChatSession
│
├── Templates/                        ← File Word template (.docx)
│   ├── 01 HN.docx
│   ├── 01 SXKD.docx
│   ├── 01 GQVL.docx
│   ├── 03 DS.docx
│   ├── GUQ.docx
│   ├── 01TGTV.docx
│   └── BIA.docx
│
├── toanquoc.enc                      ← Dữ liệu địa lý mã hóa AES-256
├── Form2State.json                   ← Trạng thái Form2 (ToData)
├── mau10c.rtf                        ← Văn bản xem (Tab 7)
├── .notice_shown                     ← ID thông báo đã hiển thị hôm nay
│
└── Desktop/Hồ sơ NHCS/              ← Output hồ sơ xuất ra
    └── {HoTen}_{MM-yyyy}/
        ├── {HoTen}_01_HN.pdf
        ├── {HoTen}_03_DS.pdf
        └── ...
```

---

## 11. Luồng dữ liệu tổng thể

```
[User nhập form]
        │
        ▼
[ValidateRequiredFields()]  ──FAIL──► MessageBox cảnh báo
        │ PASS
        ▼
[ReadForm()] → Customer object
        │
        ├──► SaveCustomerToFile()  → Customers/{name}.json
        │
        ├──► CalcThoiHanCCCD()    → Tính hạn CCCD (từ ngày sinh + ngày cấp)
        │
        ├──► UpdateComputedFields()
        │       ├── Sotientong = Vtc + Sotien
        │       └── Sotienchu  = NumberToVietnameseWords(Sotientong)
        │
        └──► CreateProfileFromTemplate()
                ├── ResolveTemplatePath()    → Templates/*.docx
                ├── ReplacePlaceholdersInWord()
                │       └── OpenXML replace {{placeholder}} → giá trị thực
                └── ConvertDocxListToPdf()   → Word Interop → *.pdf → Desktop
```

---

*File được tạo tự động bởi GitHub Copilot — HOSONHCS v1.x*
