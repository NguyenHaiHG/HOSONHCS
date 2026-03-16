using System;
using System.Linq;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1 : Form
    {
        // ============================================
        // BẢNG KÊ TIỀN (TAB 4)
        // ============================================

        private void InitializeDgvTotruong()
        {
            try
            {
                if (dgvTotruong == null) return;

                dgvTotruong.AutoGenerateColumns = false;
                dgvTotruong.AllowUserToAddRows = false;
                dgvTotruong.AllowUserToDeleteRows = false;
                dgvTotruong.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgvTotruong.MultiSelect = false;
                dgvTotruong.ReadOnly = true;

                dgvTotruong.Columns.Clear();

                dgvTotruong.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "Totruong",
                    HeaderText = "Tổ trưởng",
                    DataPropertyName = "Totruong",
                    Width = 150
                });

                dgvTotruong.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "TongTien",
                    HeaderText = "Tổng tiền (VNĐ)",
                    DataPropertyName = "TongTien",
                    Width = 150,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                });

                dgvTotruong.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "NgayTao",
                    HeaderText = "Ngày tạo",
                    DataPropertyName = "NgayTao",
                    Width = 120,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Format = "dd/MM/yyyy HH:mm"
                    }
                });

                dgvTotruong.DataSource = bangKeList;
                dgvTotruong.CellClick += DgvTotruong_CellClick;
            }
            catch { }
        }

        private void DgvTotruong_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgvTotruong == null || e.RowIndex < 0) return;

                var bangKe = dgvTotruong.Rows[e.RowIndex].DataBoundItem as BangKeData;
                if (bangKe == null) return;

                if (txtTotruong != null)
                {
                    txtTotruong.Text = bangKe.Totruong;
                }

                if (dgvbangke1 != null)
                {
                    BangKeTien.LoadDataFromBangKe(dgvbangke1, bangKe);
                }

                MessageBox.Show(
                    $"Đã load bảng kê của: {bangKe.Totruong}\n" +
                    $"Tổng tiền mặt: {bangKe.TongTien:N0} VNĐ\n" +
                    $"Sổ sách: {bangKe.SoTienSoSach:N0} VNĐ\n" +
                    $"Chênh lệch: {bangKe.ChenhLech:N0} VNĐ",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch { }
        }

        private void LoadBangKeFromFiles()
        {
            try
            {
                bangKeList.Clear();

                var loadedList = BangKeTien.LoadAllFromFiles();
                foreach (var bangKe in loadedList)
                {
                    bangKeList.Add(bangKe);
                }
            }
            catch { }
        }

        private void BtnLuubangke_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtTotruong == null || string.IsNullOrWhiteSpace(txtTotruong.Text))
                {
                    MessageBox.Show("Vui lòng nhập tên Tổ trưởng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string totruong = txtTotruong.Text.Trim();

                if (dgvbangke1 == null)
                {
                    MessageBox.Show("Không tìm thấy bảng kê tiền!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var chiTiet = BangKeTien.GetChiTiet(dgvbangke1);
                long tongTien = BangKeTien.GetTongThanhTien(dgvbangke1);
                long soTienSoSach = BangKeTien.GetSoTienSoSach(dgvbangke1);
                long chenhLech = BangKeTien.GetChenhLech(dgvbangke1);

                if (tongTien == 0)
                {
                    MessageBox.Show("Bảng kê chưa có dữ liệu (Tổng = 0)!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var existing = bangKeList.FirstOrDefault(b => string.Equals(b.Totruong, totruong, StringComparison.OrdinalIgnoreCase));

                if (existing != null)
                {
                    var result = MessageBox.Show($"Tổ trưởng '{totruong}' đã có bảng kê.\n\nBạn có muốn cập nhật?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        existing.ChiTiet = chiTiet;
                        existing.TongTien = tongTien;
                        existing.SoTienSoSach = soTienSoSach;
                        existing.ChenhLech = chenhLech;
                        existing.NgayTao = DateTime.Now;

                        BangKeTien.SaveToFile(existing);

                        if (dgvTotruong != null)
                        {
                            dgvTotruong.Refresh();
                        }

                        MessageBox.Show($"Đã cập nhật bảng kê cho: {totruong}\nTổng tiền mặt: {tongTien:N0} VNĐ\nSổ sách: {soTienSoSach:N0} VNĐ\nChênh lệch: {chenhLech:N0} VNĐ", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    var bangKe = new BangKeData
                    {
                        Totruong = totruong,
                        ChiTiet = chiTiet,
                        TongTien = tongTien,
                        SoTienSoSach = soTienSoSach,
                        ChenhLech = chenhLech,
                        NgayTao = DateTime.Now
                    };

                    BangKeTien.SaveToFile(bangKe);
                    bangKeList.Add(bangKe);

                    MessageBox.Show($"Đã lưu bảng kê cho: {totruong}\nTổng tiền mặt: {tongTien:N0} VNĐ\nSổ sách: {soTienSoSach:N0} VNĐ\nChênh lệch: {chenhLech:N0} VNĐ", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu bảng kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnXoabangke_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvTotruong == null || dgvTotruong.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn tổ trưởng cần xóa trong danh sách!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var bangKe = dgvTotruong.SelectedRows[0].DataBoundItem as BangKeData;
                if (bangKe == null) return;

                var result = MessageBox.Show($"Bạn có chắc muốn xóa bảng kê của '{bangKe.Totruong}'?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    BangKeTien.DeleteFile(bangKe);
                    bangKeList.Remove(bangKe);
                    MessageBox.Show($"Đã xóa bảng kê của: {bangKe.Totruong}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xóa bảng kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTaobangke_Click(object sender, EventArgs e)
        {
            try
            {
                var result = MessageBox.Show("Bạn có chắc muốn tạo bảng kê mới?\n\nDữ liệu hiện tại sẽ bị xóa (chưa lưu).", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    if (dgvbangke1 != null)
                    {
                        BangKeTien.ResetAll(dgvbangke1);
                    }

                    if (txtTotruong != null)
                    {
                        txtTotruong.Text = "";
                    }

                    MessageBox.Show("Đã tạo bảng kê mới!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tạo bảng kê mới: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
