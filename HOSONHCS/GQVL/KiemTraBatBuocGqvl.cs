using System.Collections.Generic;
using System.Text;

namespace HOSONHCS
{
    internal static class KiemTraBatBuocGqvl
    {
        public static bool KiemTra(KhachHangGqvl item, out string message)
        {
            var missing = new List<string>();

            AddIfEmpty(missing, item.Sohd, "Số hợp đồng");
            AddIfEmpty(missing, item.Sotienvay, "Số tiền vay");
            AddIfEmpty(missing, item.Mucdich, "Mục đích vay");
            AddIfEmpty(missing, item.Laisuat, "Lãi suất");
            AddIfEmpty(missing, item.DcPhuongAn, "Địa chỉ phương án");
            AddIfEmpty(missing, item.Tenkh, "Tên khách hàng");
            AddIfEmpty(missing, item.SdtKh, "SĐT khách hàng");
            AddIfEmpty(missing, item.Cccd, "CCCD khách hàng");
            AddIfEmpty(missing, item.NoicapCccd, "Nơi cấp CCCD");
            AddIfEmpty(missing, item.DiachiKh, "Địa chỉ khách hàng");
            AddIfEmpty(missing, item.Pgd, "Tên PGD");
            AddIfEmpty(missing, item.TenLanhDao, "Tên lãnh đạo");
            AddIfEmpty(missing, item.SdtPgd, "SĐT PGD");
            AddIfEmpty(missing, item.DiachiPgd, "Địa chỉ PGD");
            AddIfEmpty(missing, item.Thoihanvay, "Thời hạn vay");
            AddIfEmpty(missing, item.Phanky, "Phân kỳ");

            if (!item.ChuyenKhoan && string.IsNullOrWhiteSpace(item.Stk) == false)
            {
                // Cho phép bỏ qua, vì khi tiền mặt textbox sẽ bị khóa/xóa ở UI.
            }

            if (item.ChuyenKhoan)
                AddIfEmpty(missing, item.Stk, "Số tài khoản khi chọn chuyển khoản");

            if (!item.GiamDoc)
                AddIfEmpty(missing, item.SoUyQuyen, "Số ủy quyền khi chọn Phó giám đốc");

            string cccdError;
            if (!KiemTraCccdGqvl.KiemTraCccd(item.Cccd, item.Ngaysinh, out cccdError))
                missing.Add(cccdError);

            if (missing.Count == 0)
            {
                message = null;
                return true;
            }

            var sb = new StringBuilder();
            sb.AppendLine("Vui lòng kiểm tra các thông tin GQVL sau:");
            foreach (var field in missing)
                sb.AppendLine("- " + field);

            message = sb.ToString();
            return false;
        }

        private static void AddIfEmpty(List<string> missing, string value, string label)
        {
            if (string.IsNullOrWhiteSpace(value))
                missing.Add(label);
        }
    }
}
