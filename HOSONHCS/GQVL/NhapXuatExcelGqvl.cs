using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace HOSONHCS
{
    public partial class Form1
    {
        private static readonly string[] GqvlExcelHeaders = new[]
        {
            "Sohd", "Sotienvay", "Mucdich", "Ngayhopdong", "Laisuat", "Ngaygiaingan",
            "Stk", "ChuyenKhoan", "DcPhuongAn", "Tenkh", "Ngaysinh", "SdtKh", "Cccd",
            "NgaycapCccd", "DiachiKh", "Pgd", "TenLanhDao", "SdtPgd", "DiachiPgd",
            "Ong", "GiamDoc", "SoUyQuyen", "NgayUyQuyen", "Thoihanvay", "Phanky"
        };

        private void XuatExcelMauGqvl()
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = "Mau_GQVL.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                CreateGqvlExcel(sfd.FileName, new List<KhachHangGqvl>());
                MessageBox.Show("Đã tạo file Excel mẫu GQVL.", "GQVL", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void NhapExcelGqvl()
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                if (ofd.ShowDialog() != DialogResult.OK) return;

                var items = ReadGqvlExcel(ofd.FileName);
                int imported = 0;
                foreach (var item in items)
                {
                    string duplicateMessage;
                    if (!KiemTraTrungGqvl.KiemTraKhachMoi(item, gqvlCustomers, -1, out duplicateMessage))
                    {
                        MessageBox.Show(
                            "Dòng Excel bị bỏ qua do CCCD hoặc SĐT đã tồn tại:\n" + duplicateMessage,
                            "GQVL",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        continue;
                    }

                    gqvlCustomers.Add(item);
                    imported++;
                }

                SaveGqvlCustomers();
                BindGqvlGrid();
                MessageBox.Show("Đã upload " + imported + " khách hàng GQVL từ Excel.", "GQVL", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CreateGqvlExcel(string path, IList<KhachHangGqvl> rows)
        {
            using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                sheetData.Append(CreateRow(GqvlExcelHeaders));

                if (rows != null)
                {
                    foreach (var item in rows)
                    {
                        sheetData.Append(CreateRow(new[]
                        {
                            item.Sohd, item.Sotienvay, item.Mucdich, FormatExcelDate(item.Ngayhopdong), item.Laisuat, FormatExcelDate(item.Ngaygiaingan),
                            item.Stk, item.ChuyenKhoan ? "1" : "0", item.DcPhuongAn, item.Tenkh, FormatExcelDate(item.Ngaysinh), item.SdtKh, item.Cccd,
                            FormatExcelDate(item.NgaycapCccd), item.DiachiKh, item.Pgd, item.TenLanhDao, item.SdtPgd, item.DiachiPgd,
                            item.Ong ? "1" : "0", item.GiamDoc ? "1" : "0", item.SoUyQuyen, FormatExcelDate(item.NgayUyQuyen), item.Thoihanvay, item.Phanky
                        }));
                    }
                }

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "GQVL" });
                workbookPart.Workbook.Save();
            }
        }

        private List<KhachHangGqvl> ReadGqvlExcel(string path)
        {
            var result = new List<KhachHangGqvl>();
            using (var doc = SpreadsheetDocument.Open(path, false))
            {
                var workbookPart = doc.WorkbookPart;
                var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var rows = worksheetPart.Worksheet.Descendants<Row>().Skip(1);

                foreach (var row in rows)
                {
                    var values = row.Elements<Cell>().Select(c => ReadCell(workbookPart, c)).ToList();
                    if (values.Count == 0 || values.All(string.IsNullOrWhiteSpace)) continue;

                    result.Add(new KhachHangGqvl
                    {
                        Sohd = Get(values, 0),
                        Sotienvay = Get(values, 1),
                        Mucdich = Get(values, 2),
                        Ngayhopdong = ParseDate(Get(values, 3)),
                        Laisuat = Get(values, 4),
                        Ngaygiaingan = ParseDate(Get(values, 5)),
                        Stk = Get(values, 6),
                        ChuyenKhoan = Get(values, 7) == "1",
                        DcPhuongAn = Get(values, 8),
                        Tenkh = Get(values, 9),
                        Ngaysinh = ParseDate(Get(values, 10)),
                        SdtKh = Get(values, 11),
                        Cccd = Get(values, 12),
                        NgaycapCccd = ParseDate(Get(values, 13)),
                        DiachiKh = Get(values, 14),
                        Pgd = Get(values, 15),
                        TenLanhDao = Get(values, 16),
                        SdtPgd = Get(values, 17),
                        DiachiPgd = Get(values, 18),
                        Ong = Get(values, 19) != "0",
                        GiamDoc = Get(values, 20) != "0",
                        SoUyQuyen = Get(values, 21),
                        NgayUyQuyen = ParseDate(Get(values, 22)),
                        Thoihanvay = Get(values, 23),
                        Phanky = Get(values, 24)
                    });
                }
            }
            return result;
        }

        private static Row CreateRow(IEnumerable<string> values)
        {
            var row = new Row();
            foreach (var value in values)
                row.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(value ?? string.Empty) });
            return row;
        }

        private static string ReadCell(WorkbookPart workbookPart, Cell cell)
        {
            string value = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }

        private static string Get(IList<string> values, int index)
        {
            return index >= 0 && index < values.Count ? values[index] : string.Empty;
        }

        private static string FormatExcelDate(DateTime value)
        {
            return value == DateTime.MinValue ? string.Empty : value.ToString("dd/MM/yyyy");
        }

        private static DateTime ParseDate(string value)
        {
            DateTime date;
            return DateTime.TryParse(value, out date) ? date : DateTime.MinValue;
        }
    }
}
