using System;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1
    {
        private void XoaKhachHangGqvlDangChon()
        {
            if (dgvGQVL == null || dgvGQVL.CurrentRow == null)
            {
                MessageBox.Show("Vui lòng chọn khách hàng GQVL cần xóa.", "GQVL", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int index = dgvGQVL.CurrentRow.Index;
            if (index < 0 || index >= gqvlCustomers.Count) return;

            var item = gqvlCustomers[index];
            var result = MessageBox.Show("Xóa khách hàng GQVL \"" + item.Tenkh + "\"?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes) return;

            gqvlCustomers.RemoveAt(index);
            gqvlEditingIndex = -1;
            SaveGqvlCustomers();
            BindGqvlGrid();
            ClearGqvlForm();
        }
    }
}
