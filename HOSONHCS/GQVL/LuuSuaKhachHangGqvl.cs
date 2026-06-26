using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using Newtonsoft.Json;
namespace HOSONHCS
{
    public partial class Form1
    {
        private BindingList<KhachHangGqvl> gqvlCustomers = new BindingList<KhachHangGqvl>();
        private int gqvlEditingIndex = -1;

        private string GetGqvlDataFolder()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GQVL_Data");
        }

        private string GetGqvlDataPath()
        {
            return Path.Combine(GetGqvlDataFolder(), "khachhang_gqvl.json");
        }

        private void LoadGqvlCustomers()
        {
            try
            {
                Directory.CreateDirectory(GetGqvlDataFolder());
                string path = GetGqvlDataPath();
                if (File.Exists(path))
                {
                    var list = JsonConvert.DeserializeObject<BindingList<KhachHangGqvl>>(File.ReadAllText(path));
                    gqvlCustomers = list ?? new BindingList<KhachHangGqvl>();
                }
            }
            catch
            {
                gqvlCustomers = new BindingList<KhachHangGqvl>();
            }
        }

        private void SaveGqvlCustomers()
        {
            Directory.CreateDirectory(GetGqvlDataFolder());
            File.WriteAllText(GetGqvlDataPath(), JsonConvert.SerializeObject(gqvlCustomers, Formatting.Indented));
        }

        private void UpsertGqvlCustomer(KhachHangGqvl item)
        {
            if (gqvlEditingIndex >= 0 && gqvlEditingIndex < gqvlCustomers.Count)
            {
                gqvlCustomers[gqvlEditingIndex] = item;
            }
            else
            {
                string duplicateMessage;
                if (!KiemTraTrungGqvl.KiemTraKhachMoi(item, gqvlCustomers, -1, out duplicateMessage))
                    throw new InvalidOperationException(duplicateMessage);

                gqvlCustomers.Add(item);
                gqvlEditingIndex = gqvlCustomers.Count - 1;
            }

            SaveGqvlCustomers();
            BindGqvlGrid();
        }

        private void BindGqvlGrid()
        {
            if (dgvGQVL == null) return;
            dgvGQVL.AutoGenerateColumns = true;
            dgvGQVL.DataSource = gqvlCustomers;
            ConfigureGqvlGridColumns();
        }

        private void ConfigureGqvlGridColumns()
        {
            if (dgvGQVL == null || dgvGQVL.Columns.Count == 0) return;

            SetGqvlColumn("Chon", "Chọn", 0);
            SetGqvlColumn("Sohd", "Số HĐ", 1);
            SetGqvlColumn("Tenkh", "Tên khách hàng", 2);
            SetGqvlColumn("Sotienvay", "Số tiền vay", 3);
            SetGqvlColumn("Mucdich", "Mục đích", 4);
            SetGqvlColumn("Ngayhopdong", "Ngày HĐ", 5);
            SetGqvlColumn("Laisuat", "Lãi suất", 6);
            SetGqvlColumn("Ngaygiaingan", "Ngày giải ngân", 7);
            SetGqvlColumn("Stk", "Số TK", 8);
            SetGqvlColumn("ChuyenKhoan", "Chuyển khoản", 9);
            SetGqvlColumn("DcPhuongAn", "Đ/c phương án", 10);
            SetGqvlColumn("Ngaysinh", "Ngày sinh", 11);
            SetGqvlColumn("SdtKh", "SĐT KH", 12);
            SetGqvlColumn("Cccd", "CCCD", 13);
            SetGqvlColumn("NgaycapCccd", "Ngày cấp CCCD", 14);
            SetGqvlColumn("ThoihanCccd", "Hạn CCCD", 15);
            SetGqvlColumn("ThoihanCccdText", "Hạn CCCD (text)", 16);
            SetGqvlColumn("NoicapCccd", "Nơi cấp CCCD", 17);
            SetGqvlColumn("DiachiKh", "Địa chỉ KH", 18);
            SetGqvlColumn("Pgd", "PGD", 19);
            SetGqvlColumn("TenLanhDao", "Tên lãnh đạo", 20);
            SetGqvlColumn("SdtPgd", "SĐT PGD", 21);
            SetGqvlColumn("DiachiPgd", "Đ/c PGD", 22);
            SetGqvlColumn("Ong", "Ông", 23);
            SetGqvlColumn("GiamDoc", "Giám đốc", 24);
            SetGqvlColumn("SoUyQuyen", "Số ủy quyền", 25);
            SetGqvlColumn("NgayUyQuyen", "Ngày ủy quyền", 26);
            SetGqvlColumn("Thoihanvay", "Thời hạn vay", 27);
            SetGqvlColumn("Phanky", "Phân kỳ", 28);
            SetGqvlColumn("_fileName", "Tên file", 29);
        }

        private void SetGqvlColumn(string columnName, string headerText, int displayIndex)
        {
            if (!dgvGQVL.Columns.Contains(columnName)) return;

            var column = dgvGQVL.Columns[columnName];
            column.HeaderText = headerText;
            column.DisplayIndex = displayIndex;
        }

        private static bool IsSameGqvlCustomer(KhachHangGqvl left, KhachHangGqvl right)
        {
            if (left == null || right == null) return false;

            return string.Equals(NormalizeGqvlText(left.Tenkh), NormalizeGqvlText(right.Tenkh), StringComparison.OrdinalIgnoreCase)
                && string.Equals(NormalizeGqvlText(left.Cccd), NormalizeGqvlText(right.Cccd), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeGqvlText(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private string FindGqvlOutputFolderForCustomer(KhachHangGqvl item)
        {
            foreach (var customer in gqvlCustomers)
            {
                if (!IsSameGqvlCustomer(customer, item)) continue;
                if (string.IsNullOrWhiteSpace(customer._fileName)) continue;
                if (Directory.Exists(customer._fileName))
                    return customer._fileName;
            }

            return null;
        }

        private string GetGqvlOutputRootFolder()
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            return Path.Combine(desktop, "Hồ sơ NHCS", "GQVL");
        }

        private string BuildGqvlCustomerFolderName(KhachHangGqvl item)
        {
            string folderName = MakeFileSystemSafe(item.Tenkh);
            if (!string.IsNullOrWhiteSpace(item.Cccd))
                folderName += "_" + MakeFileSystemSafe(item.Cccd);

            return folderName;
        }

        private string GetOrCreateGqvlOutputFolder(KhachHangGqvl item)
        {
            if (!string.IsNullOrWhiteSpace(item._fileName) && Directory.Exists(item._fileName))
                return item._fileName;

            string existingFolder = FindGqvlOutputFolderForCustomer(item);
            if (!string.IsNullOrWhiteSpace(existingFolder))
            {
                item._fileName = existingFolder;
                return existingFolder;
            }

            string folder = Path.Combine(GetGqvlOutputRootFolder(), BuildGqvlCustomerFolderName(item));
            item._fileName = folder;
            return folder;
        }

        private void AddGqvlCustomerAsNew(KhachHangGqvl item)
        {
            string duplicateMessage;
            if (!KiemTraTrungGqvl.KiemTraKhachMoi(item, gqvlCustomers, -1, out duplicateMessage))
                throw new InvalidOperationException(duplicateMessage);

            item.Chon = false;

            string existingFolder = FindGqvlOutputFolderForCustomer(item);
            item._fileName = existingFolder;

            gqvlEditingIndex = -1;
            gqvlCustomers.Add(item);
            gqvlEditingIndex = gqvlCustomers.Count - 1;

            SaveGqvlCustomers();
            BindGqvlGrid();

            if (dgvGQVL != null && gqvlEditingIndex >= 0 && gqvlEditingIndex < dgvGQVL.Rows.Count)
            {
                dgvGQVL.ClearSelection();
                dgvGQVL.Rows[gqvlEditingIndex].Selected = true;
                try { dgvGQVL.FirstDisplayedScrollingRowIndex = gqvlEditingIndex; } catch { }
            }
        }
    }
}
