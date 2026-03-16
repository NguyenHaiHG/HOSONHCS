using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    /// <summary>
    /// Class quản lý Bảng kê tiền (TabPage4)
    /// Tính toán số lượng và thành tiền theo mệnh giá
    /// </summary>
    public class BangKeTien
    {
        // ============================================
        // DANH SÁCH MỆNH GIÁ TIỀN VIỆT NAM
        // ============================================

        /// <summary>
        /// Danh sách mệnh giá tiền Việt Nam (từ cao xuống thấp)
        /// </summary>
        private static readonly long[] MenhGiaList = new long[]
        {
            500000,   // 500 nghìn
            200000,   // 200 nghìn
            100000,   // 100 nghìn
            50000,    // 50 nghìn
            20000,    // 20 nghìn
            10000,    // 10 nghìn
            5000,     // 5 nghìn
            2000,     // 2 nghìn
            1000      // 1 nghìn
        };

        // ============================================
        // KHỞI TẠO DATAGRIDVIEW
        // ============================================

        /// <summary>
        /// Khởi tạo cấu trúc cho dgvbangke1
        /// </summary>
        /// <param name="dgv">DataGridView cần cấu hình</param>
        public static void InitializeDataGridView(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return;

                // Xóa tất cả cột và dòng cũ
                dgv.Columns.Clear();
                dgv.Rows.Clear();

                // Cấu hình chung
                dgv.AutoGenerateColumns = false;
                dgv.AllowUserToAddRows = false;
                dgv.AllowUserToDeleteRows = false;
                dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;
                dgv.MultiSelect = false;
                dgv.RowHeadersVisible = true;
                dgv.ColumnHeadersVisible = true;
                dgv.EditMode = DataGridViewEditMode.EditOnEnter;

                // Thêm 3 cột: Mệnh giá, Số lượng, Thành tiền

                // Cột 1: Mệnh giá (ReadOnly, hiển thị)
                var colMenhGia = new DataGridViewTextBoxColumn
                {
                    Name = "MenhGia",
                    HeaderText = "Mệnh giá (VNĐ)",
                    Width = 120,
                    ReadOnly = true,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0",
                        BackColor = System.Drawing.Color.LightGray
                    }
                };
                dgv.Columns.Add(colMenhGia);

                // Cột 2: Số lượng (Editable, nhập số)
                var colSoLuong = new DataGridViewTextBoxColumn
                {
                    Name = "SoLuong",
                    HeaderText = "Số lượng",
                    Width = 80,
                    ReadOnly = false,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                };
                dgv.Columns.Add(colSoLuong);

                // Cột 3: Thành tiền (ReadOnly, tự động tính)
                var colThanhTien = new DataGridViewTextBoxColumn
                {
                    Name = "ThanhTien",
                    HeaderText = "Thành tiền (VNĐ)",
                    Width = 140,
                    ReadOnly = true,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0",
                        BackColor = System.Drawing.Color.LightYellow
                    }
                };
                dgv.Columns.Add(colThanhTien);

                // Thêm các dòng mệnh giá
                foreach (var menhGia in MenhGiaList)
                {
                    int rowIndex = dgv.Rows.Add();
                    dgv.Rows[rowIndex].Cells["MenhGia"].Value = menhGia;
                    dgv.Rows[rowIndex].Cells["SoLuong"].Value = 0;
                    dgv.Rows[rowIndex].Cells["ThanhTien"].Value = 0;
                }

                // Thêm dòng tổng cộng
                int totalRowIndex = dgv.Rows.Add();
                var totalRow = dgv.Rows[totalRowIndex];
                totalRow.Cells["MenhGia"].Value = "TỔNG TIỀN MẶT";
                totalRow.Cells["SoLuong"].Value = 0;
                totalRow.Cells["ThanhTien"].Value = 0;

                // Style cho dòng tổng
                totalRow.DefaultCellStyle.Font = new System.Drawing.Font(dgv.Font, System.Drawing.FontStyle.Bold);
                totalRow.DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                totalRow.ReadOnly = true;

                // Thêm dòng "Số tiền sổ sách" (có thể nhập)
                int soSachRowIndex = dgv.Rows.Add();
                var soSachRow = dgv.Rows[soSachRowIndex];
                soSachRow.Cells["MenhGia"].Value = "SỔ SÁCH";
                soSachRow.Cells["SoLuong"].Value = "";
                soSachRow.Cells["ThanhTien"].Value = 0;
                soSachRow.DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                soSachRow.Cells["MenhGia"].ReadOnly = true;
                soSachRow.Cells["SoLuong"].ReadOnly = true;
                soSachRow.Cells["ThanhTien"].ReadOnly = false; // Cho phép nhập số tiền sổ sách

                // Thêm dòng "Chênh lệch" (tự động tính)
                int chenhLechRowIndex = dgv.Rows.Add();
                var chenhLechRow = dgv.Rows[chenhLechRowIndex];
                chenhLechRow.Cells["MenhGia"].Value = "CHÊNH LỆCH";
                chenhLechRow.Cells["SoLuong"].Value = "";
                chenhLechRow.Cells["ThanhTien"].Value = 0;
                chenhLechRow.DefaultCellStyle.Font = new System.Drawing.Font(dgv.Font, System.Drawing.FontStyle.Bold);
                chenhLechRow.DefaultCellStyle.BackColor = System.Drawing.Color.LightCoral;
                chenhLechRow.ReadOnly = true;

                // Thêm dòng "Bằng chữ"
                int bangChuRowIndex = dgv.Rows.Add();
                var bangChuRow = dgv.Rows[bangChuRowIndex];
                bangChuRow.Cells["MenhGia"].Value = "BẰNG CHỮ";
                bangChuRow.Cells["SoLuong"].Value = "";
                bangChuRow.Cells["ThanhTien"].Value = "";
                bangChuRow.DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                bangChuRow.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                bangChuRow.ReadOnly = true;
                bangChuRow.Height = 40; // Tăng chiều cao để hiện đủ chữ

                // Đăng ký sự kiện
                dgv.CellValueChanged += Dgvbangke1_CellValueChanged;
                dgv.CellEndEdit += Dgvbangke1_CellEndEdit;
                dgv.EditingControlShowing += Dgvbangke1_EditingControlShowing;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khởi tạo Bảng kê tiền: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================
        // XỬ LÝ SỰ KIỆN
        // ============================================

        /// <summary>
        /// Xử lý sự kiện khi giá trị cell thay đổi
        /// Tính toán lại Thành tiền và Tổng cộng
        /// </summary>
        private static void Dgvbangke1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var dgv = sender as DataGridView;
                if (dgv == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;

                // Xử lý khi sửa cột "SoLuong" trong các dòng mệnh giá
                if (dgv.Columns[e.ColumnIndex].Name == "SoLuong" && e.RowIndex < MenhGiaList.Length)
                {
                    RecalculateRow(dgv, e.RowIndex);
                    RecalculateTotal(dgv);
                }

                // Xử lý khi sửa "Số tiền sổ sách" (dòng thứ 10)
                if (dgv.Columns[e.ColumnIndex].Name == "ThanhTien" && e.RowIndex == MenhGiaList.Length + 1)
                {
                    RecalculateChenhLech(dgv);
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện khi kết thúc edit cell
        /// </summary>
        private static void Dgvbangke1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var dgv = sender as DataGridView;
                if (dgv == null || e.RowIndex < 0) return;

                // Xử lý các dòng mệnh giá
                if (e.RowIndex < MenhGiaList.Length)
                {
                    RecalculateRow(dgv, e.RowIndex);
                    RecalculateTotal(dgv);
                }

                // Xử lý dòng sổ sách
                if (e.RowIndex == MenhGiaList.Length + 1)
                {
                    RecalculateChenhLech(dgv);
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện khi hiển thị control chỉnh sửa
        /// Chỉ cho phép nhập số
        /// </summary>
        private static void Dgvbangke1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            try
            {
                var dgv = sender as DataGridView;
                if (dgv == null) return;

                // Chỉ xử lý cột "SoLuong"
                if (dgv.CurrentCell.OwningColumn.Name != "SoLuong") return;

                var textBox = e.Control as TextBox;
                if (textBox != null)
                {
                    // Bỏ event cũ (tránh đăng ký nhiều lần)
                    textBox.KeyPress -= TextBox_KeyPress_NumbersOnly;

                    // Đăng ký event mới: chỉ cho phép nhập số
                    textBox.KeyPress += TextBox_KeyPress_NumbersOnly;
                }
            }
            catch { }
        }

        /// <summary>
        /// Event handler: chỉ cho phép nhập số
        /// </summary>
        private static void TextBox_KeyPress_NumbersOnly(object sender, KeyPressEventArgs e)
        {
            // Cho phép: số (0-9), Backspace, Delete
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                e.Handled = true;
            }
        }

        // ============================================
        // TÍNH TOÁN
        // ============================================

        /// <summary>
        /// Tính toán lại Thành tiền cho 1 dòng
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <param name="rowIndex">Index của dòng cần tính</param>
        private static void RecalculateRow(DataGridView dgv, int rowIndex)
        {
            try
            {
                if (dgv == null || rowIndex < 0 || rowIndex >= MenhGiaList.Length) return;

                var row = dgv.Rows[rowIndex];

                // Lấy mệnh giá
                var menhGiaCell = row.Cells["MenhGia"].Value;
                if (menhGiaCell == null) return;
                long menhGia = Convert.ToInt64(menhGiaCell);

                // Lấy số lượng
                var soLuongCell = row.Cells["SoLuong"].Value;
                long soLuong = 0;
                if (soLuongCell != null)
                {
                    if (!long.TryParse(soLuongCell.ToString(), out soLuong))
                    {
                        soLuong = 0;
                    }
                }

                // Tính thành tiền
                long thanhTien = menhGia * soLuong;

                // Cập nhật cell
                row.Cells["SoLuong"].Value = soLuong;
                row.Cells["ThanhTien"].Value = thanhTien;
            }
            catch { }
        }

        /// <summary>
        /// Tính toán lại dòng Tổng cộng
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        private static void RecalculateTotal(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return;

                long tongSoLuong = 0;
                long tongThanhTien = 0;

                // Tính tổng (bỏ qua dòng cuối - dòng tổng)
                for (int i = 0; i < MenhGiaList.Length; i++)
                {
                    var row = dgv.Rows[i];

                    var soLuongCell = row.Cells["SoLuong"].Value;
                    if (soLuongCell != null && long.TryParse(soLuongCell.ToString(), out long soLuong))
                    {
                        tongSoLuong += soLuong;
                    }

                    var thanhTienCell = row.Cells["ThanhTien"].Value;
                    if (thanhTienCell != null && long.TryParse(thanhTienCell.ToString(), out long thanhTien))
                    {
                        tongThanhTien += thanhTien;
                    }
                }

                // Cập nhật dòng tổng tiền mặt (dòng thứ 9)
                int totalRowIndex = MenhGiaList.Length;
                dgv.Rows[totalRowIndex].Cells["SoLuong"].Value = tongSoLuong;
                dgv.Rows[totalRowIndex].Cells["ThanhTien"].Value = tongThanhTien;

                // Tính chênh lệch
                RecalculateChenhLech(dgv);
            }
            catch { }
        }

        /// <summary>
        /// Tính toán chênh lệch và bằng chữ
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        private static void RecalculateChenhLech(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return;

                // Lấy tổng tiền mặt (dòng 9)
                int tongTienMatRowIndex = MenhGiaList.Length;
                var tongTienMatCell = dgv.Rows[tongTienMatRowIndex].Cells["ThanhTien"].Value;
                long tongTienMat = 0;
                if (tongTienMatCell != null && long.TryParse(tongTienMatCell.ToString(), out long tienMat))
                {
                    tongTienMat = tienMat;
                }

                // Lấy số tiền sổ sách (dòng 10)
                int soSachRowIndex = MenhGiaList.Length + 1;
                var soSachCell = dgv.Rows[soSachRowIndex].Cells["ThanhTien"].Value;
                long soTienSoSach = 0;
                if (soSachCell != null && long.TryParse(soSachCell.ToString().Replace(".", "").Replace(",", ""), out long tienSach))
                {
                    soTienSoSach = tienSach;
                }

                // Tính chênh lệch = Tiền mặt - Sổ sách
                long chenhLech = tongTienMat - soTienSoSach;

                // Cập nhật dòng chênh lệch (dòng 11)
                int chenhLechRowIndex = MenhGiaList.Length + 2;
                dgv.Rows[chenhLechRowIndex].Cells["ThanhTien"].Value = chenhLech;

                // Cập nhật dòng bằng chữ (dòng 12)
                int bangChuRowIndex = MenhGiaList.Length + 3;
                string bangChu = NumberToVietnameseWords(Math.Abs(chenhLech));
                if (chenhLech < 0)
                {
                    bangChu = "Âm " + bangChu.ToLower();
                }
                dgv.Rows[bangChuRowIndex].Cells["ThanhTien"].Value = bangChu + " đồng";
            }
            catch { }
        }

        /// <summary>
        /// Chuyển số thành chữ tiếng Việt
        /// </summary>
        private static string NumberToVietnameseWords(long number)
        {
            if (number == 0) return "Không";

            string[] ones = { "", "Một", "Hai", "Ba", "Bốn", "Năm", "Sáu", "Bảy", "Tám", "Chín" };
            string[] teens = { "Mười", "Mười một", "Mười hai", "Mười ba", "Mười bốn", "Mười lăm", "Mười sáu", "Mười bảy", "Mười tám", "Mười chín" };
            string[] tens = { "", "", "Hai mươi", "Ba mươi", "Bốn mươi", "Năm mươi", "Sáu mươi", "Bảy mươi", "Tám mươi", "Chín mươi" };
            string[] thousands = { "", "Nghìn", "Triệu", "Tỷ", "Nghìn tỷ", "Triệu tỷ" };

            if (number < 0) return "Âm " + NumberToVietnameseWords(-number);
            if (number < 10) return ones[number];
            if (number < 20) return teens[number - 10];
            if (number < 100)
            {
                int ten = (int)(number / 10);
                int one = (int)(number % 10);
                if (one == 0) return tens[ten];
                if (one == 1 && ten > 1) return tens[ten] + " mốt";
                if (one == 5 && ten > 1) return tens[ten] + " lăm";
                return tens[ten] + " " + ones[one].ToLower();
            }

            if (number < 1000)
            {
                int hundred = (int)(number / 100);
                int rest = (int)(number % 100);
                string result = ones[hundred] + " trăm";
                if (rest == 0) return result;
                // Dùng "lẻ" cho số < 10, "linh" cho số >= 10
                if (rest < 10)
                    result += " lẻ " + ones[rest].ToLower();
                else
                    result += " " + NumberToVietnameseWords(rest).ToLower();
                return result;
            }

            for (int i = thousands.Length - 1; i >= 0; i--)
            {
                long divisor = (long)Math.Pow(1000, i + 1);
                if (number >= divisor)
                {
                    long high = number / divisor;
                    long low = number % divisor;
                    string result = NumberToVietnameseWords(high) + " " + thousands[i + 1];
                    if (low == 0) return result;

                    // Xử lý phần thấp hơn
                    if (low < 10)
                        result += " không trăm lẻ " + ones[low].ToLower();
                    else if (low < 100)
                        result += " không trăm " + NumberToVietnameseWords(low).ToLower();
                    else
                        result += " " + NumberToVietnameseWords(low).ToLower();

                    return result;
                }
            }

            return "";
        }

        // ============================================
        // TIỆN ÍCH
        // ============================================

        /// <summary>
        /// Reset tất cả về 0
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        public static void ResetAll(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return;

                // Reset các dòng mệnh giá
                for (int i = 0; i < MenhGiaList.Length; i++)
                {
                    dgv.Rows[i].Cells["SoLuong"].Value = 0;
                    dgv.Rows[i].Cells["ThanhTien"].Value = 0;
                }

                // Reset số tiền sổ sách
                int soSachRowIndex = MenhGiaList.Length + 1;
                dgv.Rows[soSachRowIndex].Cells["ThanhTien"].Value = 0;

                RecalculateTotal(dgv);
            }
            catch { }
        }

        /// <summary>
        /// Lấy tổng tiền mặt
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <returns>Tổng số tiền mặt</returns>
        public static long GetTongThanhTien(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return 0;

                int totalRowIndex = MenhGiaList.Length;
                var cell = dgv.Rows[totalRowIndex].Cells["ThanhTien"].Value;

                if (cell != null && long.TryParse(cell.ToString(), out long total))
                {
                    return total;
                }
            }
            catch { }

            return 0;
        }

        /// <summary>
        /// Lấy số tiền sổ sách
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <returns>Số tiền sổ sách</returns>
        public static long GetSoTienSoSach(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return 0;

                int soSachRowIndex = MenhGiaList.Length + 1;
                var cell = dgv.Rows[soSachRowIndex].Cells["ThanhTien"].Value;

                if (cell != null && long.TryParse(cell.ToString().Replace(".", "").Replace(",", ""), out long soSach))
                {
                    return soSach;
                }
            }
            catch { }

            return 0;
        }

        /// <summary>
        /// Lấy chênh lệch
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <returns>Chênh lệch tiền mặt - sổ sách</returns>
        public static long GetChenhLech(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return 0;

                int chenhLechRowIndex = MenhGiaList.Length + 2;
                var cell = dgv.Rows[chenhLechRowIndex].Cells["ThanhTien"].Value;

                if (cell != null && long.TryParse(cell.ToString(), out long chenhLech))
                {
                    return chenhLech;
                }
            }
            catch { }

            return 0;
        }

        /// <summary>
        /// Lấy chi tiết bảng kê (danh sách mệnh giá và số lượng)
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <returns>Dictionary với key = mệnh giá, value = số lượng</returns>
        public static Dictionary<long, long> GetChiTiet(DataGridView dgv)
        {
            var result = new Dictionary<long, long>();

            try
            {
                if (dgv == null) return result;

                for (int i = 0; i < MenhGiaList.Length; i++)
                {
                    var row = dgv.Rows[i];

                    var menhGiaCell = row.Cells["MenhGia"].Value;
                    var soLuongCell = row.Cells["SoLuong"].Value;

                    if (menhGiaCell != null && soLuongCell != null)
                    {
                        long menhGia = Convert.ToInt64(menhGiaCell);
                        long soLuong = Convert.ToInt64(soLuongCell);

                        if (soLuong > 0)
                        {
                            result[menhGia] = soLuong;
                        }
                    }
                }
            }
            catch { }

            return result;
        }

        /// <summary>
        /// Load dữ liệu vào bảng kê từ Dictionary
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <param name="data">Dictionary với key = mệnh giá, value = số lượng</param>
        public static void LoadData(DataGridView dgv, Dictionary<long, long> data)
        {
            try
            {
                if (dgv == null || data == null) return;

                // Reset tất cả về 0
                ResetAll(dgv);

                // Load dữ liệu
                for (int i = 0; i < MenhGiaList.Length; i++)
                {
                    long menhGia = MenhGiaList[i];

                    if (data.ContainsKey(menhGia))
                    {
                        dgv.Rows[i].Cells["SoLuong"].Value = data[menhGia];
                        RecalculateRow(dgv, i);
                    }
                }

                RecalculateTotal(dgv);
            }
            catch { }
        }

        /// <summary>
        /// Load dữ liệu từ BangKeData vào bảng kê (bao gồm cả số tiền sổ sách)
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <param name="bangKe">BangKeData</param>
        public static void LoadDataFromBangKe(DataGridView dgv, BangKeData bangKe)
        {
            try
            {
                if (dgv == null || bangKe == null) return;

                // Load chi tiết mệnh giá
                LoadData(dgv, bangKe.ChiTiet);

                // Load số tiền sổ sách
                int soSachRowIndex = MenhGiaList.Length + 1;
                dgv.Rows[soSachRowIndex].Cells["ThanhTien"].Value = bangKe.SoTienSoSach;

                // Tính lại chênh lệch
                RecalculateChenhLech(dgv);
            }
            catch { }
        }

        /// <summary>
        /// Export dữ liệu bảng kê thành chuỗi text (để hiển thị hoặc in)
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <returns>Chuỗi text mô tả bảng kê</returns>
        public static string ExportToText(DataGridView dgv)
        {
            try
            {
                if (dgv == null) return "";

                var lines = new List<string>();
                lines.Add("========== BẢNG KÊ TIỀN MẶT ==========");
                lines.Add("");
                lines.Add(string.Format("{0,-20} {1,15} {2,20}", "Mệnh giá", "Số lượng", "Thành tiền"));
                lines.Add(new string('-', 60));

                for (int i = 0; i < MenhGiaList.Length; i++)
                {
                    var row = dgv.Rows[i];
                    var menhGia = Convert.ToInt64(row.Cells["MenhGia"].Value);
                    var soLuong = Convert.ToInt64(row.Cells["SoLuong"].Value);
                    var thanhTien = Convert.ToInt64(row.Cells["ThanhTien"].Value);

                    if (soLuong > 0)
                    {
                        lines.Add(string.Format("{0,-20:N0} {1,15:N0} {2,20:N0}", menhGia, soLuong, thanhTien));
                    }
                }

                lines.Add(new string('-', 60));

                int totalRowIndex = dgv.Rows.Count - 1;
                var tongSoLuong = Convert.ToInt64(dgv.Rows[totalRowIndex].Cells["SoLuong"].Value);
                var tongThanhTien = Convert.ToInt64(dgv.Rows[totalRowIndex].Cells["ThanhTien"].Value);

                lines.Add(string.Format("{0,-20} {1,15:N0} {2,20:N0}", "TỔNG CỘNG", tongSoLuong, tongThanhTien));
                lines.Add("");
                lines.Add("======================================");

                return string.Join(Environment.NewLine, lines);
            }
            catch
            {
                return "";
            }
        }

        // ============================================
        // QUẢN LÝ FILE JSON (LƯU/TẢI BẢNG KÊ)
        // ============================================

        /// <summary>
        /// Đường dẫn folder lưu trữ file JSON bảng kê
        /// </summary>
        private static string GetBangKeFolderPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BangKe");
        }

        /// <summary>
        /// Đảm bảo folder "BangKe" tồn tại
        /// </summary>
        private static void EnsureBangKeFolder()
        {
            var folder = GetBangKeFolderPath();
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
        }

        /// <summary>
        /// Tạo tên file an toàn từ tên tổ trưởng
        /// </summary>
        private static string MakeFileSystemSafe(string input)
        {
            if (string.IsNullOrEmpty(input)) input = "unknown";
            var invalid = Path.GetInvalidFileNameChars();
            foreach (var ch in invalid)
                input = input.Replace(ch.ToString(), "_");
            return input.Trim();
        }

        /// <summary>
        /// Lưu BangKeData vào file JSON
        /// </summary>
        /// <param name="bangKe">Dữ liệu bảng kê cần lưu</param>
        public static void SaveToFile(BangKeData bangKe)
        {
            try
            {
                if (bangKe == null) return;

                EnsureBangKeFolder();

                string fileName;
                if (!string.IsNullOrEmpty(bangKe._fileName))
                {
                    // Cập nhật file hiện có
                    fileName = bangKe._fileName;
                }
                else
                {
                    // Tạo file mới
                    var safeName = MakeFileSystemSafe(bangKe.Totruong);
                    fileName = $"{safeName}.json";

                    // Tránh trùng tên
                    var folder = GetBangKeFolderPath();
                    var fullPath = Path.Combine(folder, fileName);
                    int i = 1;
                    while (File.Exists(fullPath))
                    {
                        fileName = $"{safeName}_{i}.json";
                        fullPath = Path.Combine(folder, fileName);
                        i++;
                    }

                    bangKe._fileName = fileName;
                }

                var path = Path.Combine(GetBangKeFolderPath(), fileName);
                var json = JsonConvert.SerializeObject(bangKe, Formatting.Indented);
                File.WriteAllText(path, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi lưu bảng kê: {ex.Message}");
            }
        }

        /// <summary>
        /// Load tất cả bảng kê từ folder "BangKe"
        /// </summary>
        /// <returns>Danh sách BangKeData</returns>
        public static List<BangKeData> LoadAllFromFiles()
        {
            var result = new List<BangKeData>();

            try
            {
                EnsureBangKeFolder();
                var folder = GetBangKeFolderPath();

                foreach (var file in Directory.EnumerateFiles(folder, "*.json"))
                {
                    try
                    {
                        var json = File.ReadAllText(file, Encoding.UTF8);
                        var bangKe = JsonConvert.DeserializeObject<BangKeData>(json);

                        if (bangKe != null)
                        {
                            bangKe._fileName = Path.GetFileName(file);
                            result.Add(bangKe);
                        }
                    }
                    catch
                    {
                        // Bỏ qua file lỗi
                    }
                }
            }
            catch
            {
                // Folder không tồn tại hoặc lỗi khác
            }

            return result;
        }

        /// <summary>
        /// Xóa file JSON của bảng kê
        /// </summary>
        /// <param name="bangKe">Bảng kê cần xóa</param>
        public static void DeleteFile(BangKeData bangKe)
        {
            try
            {
                if (bangKe == null || string.IsNullOrEmpty(bangKe._fileName)) return;

                var path = Path.Combine(GetBangKeFolderPath(), bangKe._fileName);
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi xóa file bảng kê: {ex.Message}");
            }
        }

        /// <summary>
        /// Load bảng kê từ file cụ thể
        /// </summary>
        /// <param name="fileName">Tên file JSON</param>
        /// <returns>BangKeData hoặc null nếu không load được</returns>
        public static BangKeData LoadFromFile(string fileName)
        {
            try
            {
                var path = Path.Combine(GetBangKeFolderPath(), fileName);
                if (!File.Exists(path)) return null;

                var json = File.ReadAllText(path, Encoding.UTF8);
                var bangKe = JsonConvert.DeserializeObject<BangKeData>(json);

                if (bangKe != null)
                {
                    bangKe._fileName = fileName;
                }

                return bangKe;
            }
            catch
            {
                return null;
            }
        }
    }
}
