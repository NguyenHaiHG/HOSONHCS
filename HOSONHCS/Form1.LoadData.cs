using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace HOSONHCS
{
    public partial class Form1
    {
        // ===== CẤU TRÚC TẠM ĐỂ BUILD JSON =====
        private class ExcelRow
        {
            public string Tinh { get; set; }
            public string Pgd { get; set; }
            public string Xa { get; set; }
            public string HoiDoanThe { get; set; }
            public string Thon { get; set; }
            public string TenTotruong { get; set; }
        }

        private void btnLoaddata_Click(object sender, EventArgs e)
        {
            if (xinManEditor == null || !xinManEditor.IsLoggedIn)
            {
                MessageBox.Show("Bạn phải đăng nhập trước khi sử dụng chức năng này.",
                    "Chưa đăng nhập", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Chọn file Excel dữ liệu tổ chức";
                dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
                dlg.FilterIndex = 1;
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    var rows = ReadExcelRows(dlg.FileName);
                    if (rows == null || rows.Count == 0)
                    {
                        MessageBox.Show("File Excel không có dữ liệu hoặc sai cấu trúc.\n\nYêu cầu:\n  Cột A: Tên Tỉnh\n  Cột B: Phòng giao dịch\n  Cột C: Xã\n  Cột D: Hội đoàn thể\n  Cột E: Thôn\n  Cột F: Tên tổ trưởng",
                            "Lỗi dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var jsonArray = BuildJsonFromRows(rows);
                    var jsonText = JsonConvert.SerializeObject(jsonArray, Formatting.Indented);

                    var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "toanquoc.json");
                    File.WriteAllText(outputPath, jsonText, Encoding.UTF8);

                    // Làm mới cache
                    TinhHelper.RefreshCache();

                    // Reload combobox tỉnh nếu có
                    if (cbTinhfix != null) TinhHelper.PopulateComboBox(cbTinhfix);
                    if (cbTinh != null) TinhHelper.PopulateComboBox(cbTinh);

                    MessageBox.Show(
                        $"Nhập dữ liệu thành công!\n\n" +
                        $"  Số dòng đọc được: {rows.Count}\n" +
                        $"  Số tỉnh: {jsonArray.Count}\n" +
                        $"  File lưu tại: {outputPath}",
                        "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xử lý file Excel:\n" + ex.Message,
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private List<ExcelRow> ReadExcelRows(string filePath)
        {
            var result = new List<ExcelRow>();
            using (var doc = SpreadsheetDocument.Open(filePath, false))
            {
                var wbPart = doc.WorkbookPart;
                var sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null) return result;

                var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
                var sharedStrings = wbPart.SharedStringTablePart?.SharedStringTable;

                string GetCellValue(Cell cell)
                {
                    if (cell == null) return string.Empty;
                    var val = cell.InnerText ?? string.Empty;
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        if (sharedStrings != null && int.TryParse(val, out int idx))
                            return sharedStrings.ElementAt(idx).InnerText?.Trim() ?? string.Empty;
                    }
                    return val.Trim();
                }

                bool isFirstRow = true;
                foreach (var row in wsPart.Worksheet.Descendants<Row>())
                {
                    // Bỏ qua dòng tiêu đề (dòng 1)
                    if (isFirstRow) { isFirstRow = false; continue; }

                    var cells = row.Elements<Cell>().ToList();
                    string GetColValue(string colLetter)
                    {
                        var c = cells.FirstOrDefault(x => GetColumnLetter(x.CellReference) == colLetter);
                        return GetCellValue(c);
                    }

                    var tinh = GetColValue("A");
                    var pgd = GetColValue("B");
                    var xa = GetColValue("C");
                    var hoi = GetColValue("D");
                    var thon = GetColValue("E");
                    var totruong = GetColValue("F");

                    if (string.IsNullOrWhiteSpace(tinh) && string.IsNullOrWhiteSpace(pgd) &&
                        string.IsNullOrWhiteSpace(xa) && string.IsNullOrWhiteSpace(hoi) &&
                        string.IsNullOrWhiteSpace(thon) && string.IsNullOrWhiteSpace(totruong))
                        continue;

                    result.Add(new ExcelRow
                    {
                        Tinh = tinh,
                        Pgd = pgd,
                        Xa = xa,
                        HoiDoanThe = hoi,
                        Thon = thon,
                        TenTotruong = totruong
                    });
                }
            }
            return result;
        }

        private static string GetColumnLetter(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return string.Empty;
            var sb = new StringBuilder();
            foreach (char c in cellRef)
            {
                if (char.IsLetter(c)) sb.Append(c);
                else break;
            }
            return sb.ToString().ToUpper();
        }

        private static string BuildAssociationCode(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return string.Empty;
            var words = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();
            foreach (var w in words)
                if (w.Length > 0) sb.Append(char.ToUpper(w[0]));
            return sb.ToString();
        }

        private List<TinhModel> BuildJsonFromRows(List<ExcelRow> rows)
        {
            // Nhóm theo Tỉnh -> PGD -> Xã -> Hội đoàn thể -> Thôn -> Tổ trưởng
            var tinhDict = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>>>(StringComparer.OrdinalIgnoreCase);

            // Dùng để giữ thứ tự xuất hiện
            var tinhOrder = new List<string>();
            var pgdOrder = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var xaOrder = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var hoiOrder = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var thonOrder = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var r in rows)
            {
                var tinh = r.Tinh?.Trim() ?? string.Empty;
                var pgd = r.Pgd?.Trim() ?? string.Empty;
                var xa = r.Xa?.Trim() ?? string.Empty;
                var hoi = r.HoiDoanThe?.Trim() ?? string.Empty;
                var thon = r.Thon?.Trim() ?? string.Empty;
                var totruong = r.TenTotruong?.Trim() ?? string.Empty;

                if (!tinhDict.ContainsKey(tinh))
                {
                    tinhDict[tinh] = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>>(StringComparer.OrdinalIgnoreCase);
                    tinhOrder.Add(tinh);
                }
                var pgdDict = tinhDict[tinh];

                var tinhPgdKey = tinh + "\0" + pgd;
                if (!pgdDict.ContainsKey(pgd))
                {
                    pgdDict[pgd] = new Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>(StringComparer.OrdinalIgnoreCase);
                    if (!pgdOrder.ContainsKey(tinh)) pgdOrder[tinh] = new List<string>();
                    pgdOrder[tinh].Add(pgd);
                }
                var xaDict = pgdDict[pgd];

                var pgdXaKey = tinhPgdKey + "\0" + xa;
                if (!xaDict.ContainsKey(xa))
                {
                    xaDict[xa] = new Dictionary<string, Dictionary<string, List<string>>>(StringComparer.OrdinalIgnoreCase);
                    if (!xaOrder.ContainsKey(tinhPgdKey)) xaOrder[tinhPgdKey] = new List<string>();
                    xaOrder[tinhPgdKey].Add(xa);
                }
                var hoiDict = xaDict[xa];

                var xaHoiKey = pgdXaKey + "\0" + hoi;
                if (!hoiDict.ContainsKey(hoi))
                {
                    hoiDict[hoi] = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                    if (!hoiOrder.ContainsKey(pgdXaKey)) hoiOrder[pgdXaKey] = new List<string>();
                    hoiOrder[pgdXaKey].Add(hoi);
                }
                var thonDict = hoiDict[hoi];

                if (!thonDict.ContainsKey(thon))
                {
                    thonDict[thon] = new List<string>();
                    if (!thonOrder.ContainsKey(xaHoiKey)) thonOrder[xaHoiKey] = new List<string>();
                    thonOrder[xaHoiKey].Add(thon);
                }
                if (!string.IsNullOrEmpty(totruong) && !thonDict[thon].Contains(totruong))
                    thonDict[thon].Add(totruong);
            }

            // Build TinhModel list
            var result = new List<TinhModel>();
            foreach (var tinhName in tinhOrder)
            {
                var pgdList = new List<TinhPgdEntry>();
                var pgds = pgdOrder.ContainsKey(tinhName) ? pgdOrder[tinhName] : new List<string>();
                foreach (var pgdName in pgds)
                {
                    var tinhPgdKey = tinhName + "\0" + pgdName;
                    var communes = new List<Commune>();
                    var xas = xaOrder.ContainsKey(tinhPgdKey) ? xaOrder[tinhPgdKey] : new List<string>();
                    foreach (var xaName in xas)
                    {
                        var pgdXaKey = tinhPgdKey + "\0" + xaName;
                        var associations = new List<Association>();
                        var hois = hoiOrder.ContainsKey(pgdXaKey) ? hoiOrder[pgdXaKey] : new List<string>();
                        foreach (var hoiName in hois)
                        {
                            var xaHoiKey = pgdXaKey + "\0" + hoiName;
                            var villages = new List<Village>();
                            var thons = thonOrder.ContainsKey(xaHoiKey) ? thonOrder[xaHoiKey] : new List<string>();
                            foreach (var thonName in thons)
                            {
                                villages.Add(new Village
                                {
                                    name = thonName,
                                    groups = tinhDict[tinhName][pgdName][xaName][hoiName][thonName]
                                });
                            }
                            associations.Add(new Association
                            {
                                name = hoiName,
                                code = BuildAssociationCode(hoiName),
                                villages = villages
                            });
                        }
                        communes.Add(new Commune { name = xaName, associations = associations });
                    }
                    pgdList.Add(new TinhPgdEntry { pgd = pgdName, communes = communes });
                }
                result.Add(new TinhModel { tinh = tinhName, pgds = pgdList });
            }
            return result;
        }
    }
}
