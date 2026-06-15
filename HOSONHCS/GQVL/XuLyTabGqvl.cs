using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1
    {
        private void InitializeGqvlModule()
        {
            try
            {
                LoadGqvlCustomers();
                BindGqvlGrid();
                ConfigureGqvlControls();
                ConfigureGqvlInputRestrictions();

                btnHdGQVL.Click += BtnHdGQVL_Click;
                btnCapnhatkh.Click += BtnCapnhatkh_Click;
                btxXoaGQVL.Click += BtxXoaGQVL_Click;
                btnExcelmau.Click += BtnExcelmau_Click;
                btnUpexcel.Click += BtnUpexcel_Click;

                rdT.CheckedChanged += GqvlPaymentRadio_CheckedChanged;
                rdC.CheckedChanged += GqvlPaymentRadio_CheckedChanged;
                rdGd.Click += GqvlChucVu_Click;
                rdPgd.Click += GqvlChucVu_Click;
                rdO.Click += GqvlDanhXung_Click;
                rdB.Click += GqvlDanhXung_Click;

                dtNamsinhkhGQVL.ValueChanged += GqvlCccdDateChanged;
                dtNgaycapGQVL.ValueChanged += GqvlCccdDateChanged;
                dgvGQVL.CellClick += DgvGQVL_CellClick;
                dgvGQVL.CellValueChanged += DgvGQVL_CellValueChanged;
                dgvGQVL.CurrentCellDirtyStateChanged += DgvGQVL_CurrentCellDirtyStateChanged;

                UpdateGqvlPaymentState();
                UpdateGqvlAuthorizationState();
                UpdateGqvlCccdInfo();
            }
            catch { }
        }

        private void ConfigureGqvlControls()
        {
            txtCccdkhGQVL.MaxLength = 12;
            txtCccdkhGQVL.KeyPress += TxtDigitsOnly_KeyPress;
            txtStkGQVL.KeyPress += TxtDigitsOnly_KeyPress;

            rdO.AutoCheck = false;
            rdB.AutoCheck = false;
            rdGd.AutoCheck = false;
            rdPgd.AutoCheck = false;

            AddDefaultItems(cbLsGQVL, "6,6", "7,92", "9,0");
            AddDefaultItems(cbStGQVL, "50.000.000", "100.000.000", "150.000.000", "200.000.000");

            if (!rdT.Checked && !rdC.Checked) rdT.Checked = true;
            if (!rdO.Checked && !rdB.Checked) rdO.Checked = true;
            if (!rdGd.Checked && !rdPgd.Checked) rdGd.Checked = true;
        }

        private static void AddDefaultItems(ComboBox comboBox, params string[] values)
        {
            if (comboBox == null || comboBox.Items.Count > 0) return;
            comboBox.Items.AddRange(values);
        }

        private KhachHangGqvl ReadGqvlForm()
        {
            string thoiHanCccdText;
            DateTime thoiHanCccd = KiemTraCccdGqvl.TinhNgayHetHan(dtNamsinhkhGQVL.Value.Date, dtNgaycapGQVL.Value.Date, out thoiHanCccdText);

            return new KhachHangGqvl
            {
                Sohd = txtSohd.Text.Trim(),
                Sotienvay = cbStGQVL.Text.Trim(),
                Mucdich = txtMdGQVL.Text.Trim(),
                Ngayhopdong = dtNgayhdGQVL.Value.Date,
                Laisuat = cbLsGQVL.Text.Trim(),
                Ngaygiaingan = dtNgaygnGQVL.Value.Date,
                Stk = txtStkGQVL.Text.Trim(),
                ChuyenKhoan = rdC.Checked,
                DcPhuongAn = txtDcGQVL.Text.Trim(),
                Tenkh = txtTenkhGQVL.Text.Trim(),
                Ngaysinh = dtNamsinhkhGQVL.Value.Date,
                SdtKh = txtSdtkhGQVL.Text.Trim(),
                Cccd = txtCccdkhGQVL.Text.Trim(),
                NgaycapCccd = dtNgaycapGQVL.Value.Date,
                ThoihanCccd = thoiHanCccd,
                ThoihanCccdText = thoiHanCccdText,
                NoicapCccd = txtNoicapGQVL.Text.Trim(),
                DiachiKh = txtDckhGQVL.Text.Trim(),
                Pgd = txtTenPGD.Text.Trim(),
                TenLanhDao = txtTenld.Text.Trim(),
                SdtPgd = txtSdtPGD.Text.Trim(),
                DiachiPgd = txtDcPGD.Text.Trim(),
                Ong = rdO.Checked,
                GiamDoc = rdGd.Checked,
                SoUyQuyen = txtUq.Text.Trim(),
                NgayUyQuyen = dtNgayuq.Value.Date,
                Thoihanvay = cbThoihanvayGQVL.Text.Trim(),
                Phanky = cbPhankyGQVL.Text.Trim()
            };
        }

        private void PopulateGqvlForm(KhachHangGqvl item)
        {
            if (item == null) return;

            txtSohd.Text = item.Sohd ?? "";
            cbStGQVL.Text = item.Sotienvay ?? "";
            txtMdGQVL.Text = item.Mucdich ?? "";
            if (item.Ngayhopdong != DateTime.MinValue) dtNgayhdGQVL.Value = item.Ngayhopdong;
            cbLsGQVL.Text = item.Laisuat ?? "";
            if (item.Ngaygiaingan != DateTime.MinValue) dtNgaygnGQVL.Value = item.Ngaygiaingan;
            txtStkGQVL.Text = item.Stk ?? "";
            rdC.Checked = item.ChuyenKhoan;
            rdT.Checked = !item.ChuyenKhoan;
            txtDcGQVL.Text = item.DcPhuongAn ?? "";
            txtTenkhGQVL.Text = item.Tenkh ?? "";
            if (item.Ngaysinh != DateTime.MinValue) dtNamsinhkhGQVL.Value = item.Ngaysinh;
            txtSdtkhGQVL.Text = item.SdtKh ?? "";
            txtCccdkhGQVL.Text = item.Cccd ?? "";
            if (item.NgaycapCccd != DateTime.MinValue) dtNgaycapGQVL.Value = item.NgaycapCccd;
            if (item.ThoihanCccd != DateTime.MinValue && item.ThoihanCccd != DateTime.MaxValue.Date) dtNgaydhccGQVL.Value = item.ThoihanCccd;
            txtNoicapGQVL.Text = item.NoicapCccd ?? "";
            txtDckhGQVL.Text = item.DiachiKh ?? "";
            txtTenPGD.Text = item.Pgd ?? "";
            txtTenld.Text = item.TenLanhDao ?? "";
            txtSdtPGD.Text = item.SdtPgd ?? "";
            txtDcPGD.Text = item.DiachiPgd ?? "";
            rdO.Checked = item.Ong;
            rdB.Checked = !item.Ong;
            rdGd.Checked = item.GiamDoc;
            rdPgd.Checked = !item.GiamDoc;
            txtUq.Text = item.SoUyQuyen ?? "";
            if (item.NgayUyQuyen != DateTime.MinValue) dtNgayuq.Value = item.NgayUyQuyen;
            cbThoihanvayGQVL.Text = item.Thoihanvay ?? "";
            cbPhankyGQVL.Text = item.Phanky ?? "";

            UpdateGqvlPaymentState();
            UpdateGqvlAuthorizationState();
        }

        private void ClearGqvlForm()
        {
            txtSohd.Clear();
            cbStGQVL.Text = "";
            txtMdGQVL.Clear();
            cbLsGQVL.Text = "";
            txtStkGQVL.Clear();
            txtDcGQVL.Clear();
            txtTenkhGQVL.Clear();
            txtSdtkhGQVL.Clear();
            txtCccdkhGQVL.Clear();
            txtDckhGQVL.Clear();
            txtUq.Clear();
            cbThoihanvayGQVL.Text = "";
            cbPhankyGQVL.Text = "";
            gqvlEditingIndex = -1;
            UpdateGqvlPaymentState();
            UpdateGqvlAuthorizationState();
        }

        private void BtnCapnhatkh_Click(object sender, EventArgs e)
        {
            var item = ReadGqvlForm();
            string message;
            if (!KiemTraBatBuocGqvl.KiemTra(item, out message))
            {
                MessageBox.Show(message, "Thiếu thông tin GQVL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            UpsertGqvlCustomer(item);
            MessageBox.Show("Đã cập nhật khách hàng GQVL.", "GQVL", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnHdGQVL_Click(object sender, EventArgs e)
        {
            try
            {
                var selected = GetSelectedGqvlCustomers();
                if (selected.Count == 0)
                {
                    var current = ReadGqvlForm();
                    string message;
                    if (!KiemTraBatBuocGqvl.KiemTra(current, out message))
                    {
                        MessageBox.Show(message, "Thiếu thông tin GQVL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    UpsertGqvlCustomer(current);
                    selected.Add(current);
                }

                var files = new List<string>();
                foreach (var item in selected)
                {
                    string message;
                    if (!KiemTraBatBuocGqvl.KiemTra(item, out message))
                    {
                        MessageBox.Show("Khách hàng " + item.Tenkh + " chưa đủ thông tin:\n" + message, "GQVL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    files.Add(XuatHopDongGqvl(item));
                }

                SaveGqvlCustomers();
                BindGqvlGrid();
                OpenGqvlFiles(files);
                MessageBox.Show("Đã tạo " + files.Count + " hợp đồng GQVL.", "GQVL", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tạo hợp đồng GQVL: " + ex.Message, "GQVL", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<KhachHangGqvl> GetSelectedGqvlCustomers()
        {
            return gqvlCustomers.Where(x => x.Chon).ToList();
        }

        private void BtxXoaGQVL_Click(object sender, EventArgs e)
        {
            XoaKhachHangGqvlDangChon();
        }

        private void BtnExcelmau_Click(object sender, EventArgs e)
        {
            XuatExcelMauGqvl();
        }

        private void BtnUpexcel_Click(object sender, EventArgs e)
        {
            NhapExcelGqvl();
        }

        private void DgvGQVL_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= gqvlCustomers.Count) return;
            gqvlEditingIndex = e.RowIndex;
            PopulateGqvlForm(gqvlCustomers[e.RowIndex]);
        }

        private void DgvGQVL_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvGQVL.IsCurrentCellDirty) dgvGQVL.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvGQVL_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            SaveGqvlCustomers();
        }

        private void GqvlPaymentRadio_CheckedChanged(object sender, EventArgs e)
        {
            UpdateGqvlPaymentState();
        }

        private void UpdateGqvlPaymentState()
        {
            if (txtStkGQVL == null) return;
            txtStkGQVL.Enabled = rdC.Checked;
            if (rdT.Checked)
                txtStkGQVL.Text = "";
        }

        private void GqvlChucVu_Click(object sender, EventArgs e)
        {
            rdGd.Checked = sender == rdGd;
            rdPgd.Checked = sender == rdPgd;
            UpdateGqvlAuthorizationState();
        }

        private void GqvlDanhXung_Click(object sender, EventArgs e)
        {
            rdO.Checked = sender == rdO;
            rdB.Checked = sender == rdB;
        }

        private void UpdateGqvlAuthorizationState()
        {
            bool isDirector = rdGd.Checked;
            txtUq.Enabled = !isDirector;
            dtNgayuq.Enabled = !isDirector;
            if (isDirector)
                txtUq.Text = "";
        }

        private void GqvlCccdDateChanged(object sender, EventArgs e)
        {
            UpdateGqvlCccdInfo();
        }

        private void UpdateGqvlCccdInfo()
        {
            try
            {
                txtNoicapGQVL.Text = KiemTraCccdGqvl.TinhNoiCap(dtNgaycapGQVL.Value.Date);
                string display;
                DateTime expiry = KiemTraCccdGqvl.TinhNgayHetHan(dtNamsinhkhGQVL.Value.Date, dtNgaycapGQVL.Value.Date, out display);
                if (expiry != DateTime.MaxValue.Date && expiry >= dtNgaydhccGQVL.MinDate && expiry <= dtNgaydhccGQVL.MaxDate)
                    dtNgaydhccGQVL.Value = expiry;
            }
            catch { }
        }
    }
}
