using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1 : Form
    {
        // ============================================
        // BUTTON CLICK & DATAGRIDVIEW HANDLERS
        // ============================================

         private void Dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
         {
         }

         private bool IsSelectionCompatibleWithChecked(int rowIndex)
         {
             try
             {
                 var candidate = dgv.Rows[rowIndex].DataBoundItem as Customer;
                 if (candidate == null) return true;

                 var checkedCustomers = new List<Customer>();
                 foreach (DataGridViewRow row in dgv.Rows)
                 {
                     if (row.Index == rowIndex) continue;
                     try
                     {
                         var ccell = row.Cells["colSelect"];
                         bool checkedVal = false;
                         if (ccell != null && ccell.Value != null)
                         {
                             if (ccell.Value is bool bb) checkedVal = bb;
                             else bool.TryParse(ccell.Value.ToString(), out checkedVal);
                         }
                         if (checkedVal) { var item = row.DataBoundItem as Customer; if (item != null) checkedCustomers.Add(item); }
                     }
                     catch { }
                 }

                 if (checkedCustomers.Count == 0) return true;

                 foreach (var other in checkedCustomers)
                 {
                     if (!string.Equals((candidate.PGD ?? "").Trim(), (other.PGD ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer does not match PGD of already selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     if (!string.Equals((candidate.Chuongtrinh ?? "").Trim(), (other.Chuongtrinh ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer does not match Chương trình of already selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     var candTo = (candidate.Totruong ?? candidate.To ?? "").Trim();
                     var otherTo = (other.Totruong ?? other.To ?? "").Trim();
                     if (!string.Equals(candTo, otherTo, StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(candTo))
                     {
                         MessageBox.Show("Selected customer must have the same Tổ / Tổ trưởng as other selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     if (!string.Equals((candidate.Xa ?? "").Trim(), (other.Xa ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer must be in the same Xã as other selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     if (!string.Equals((candidate.Thon ?? "").Trim(), (other.Thon ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer must be in the same Thôn as other selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }
                 }

                 return true;
             }
             catch { return true; }
         }

         private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
         {
             if (e.RowIndex >= 0 && e.RowIndex < customers.Count)
             {
                 if (e.ColumnIndex >= 0 && dgv.Columns[e.ColumnIndex].Name == "colSelect") return;
                 editingIndex = e.RowIndex;
                 PopulateForm(customers[e.RowIndex]);
             }
         }

         private void Dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
         {
             if (e.RowIndex >= 0 && e.RowIndex < customers.Count)
             {
                 if (e.ColumnIndex >= 0 && dgv.Columns[e.ColumnIndex].Name == "colSelect") return;
                 editingIndex = e.RowIndex;
                 PopulateForm(customers[e.RowIndex]);
             }
         }

         private async void BtnSave_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;
                 else customer._fileName = customer._fileName;

                 SaveCustomerToFile(customer);
                 var createdFiles = await Task.Run(() => CreateProfileFromTemplate(customer, include03: false));

                 UpsertCustomerInList(customer);

                 BindGrid();
                 ClearForm();

                 MessageBox.Show(
                     "✅ Đã xuất mẫu 01/TD thành công!\n\n" +
                     $"📄 Khách hàng: {customer.Hoten}\n" +
                     $"💼 Chương trình: {customer.Chuongtrinh}\n" +
                     $"💵 Số tiền: {customer.Sotien}",
                     "✅ Xuất mẫu 01/TD",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Information
                 );

                 OpenCreatedFiles(createdFiles);
             }
             catch (Exception ex)
             {
                 MessageBox.Show(
                     $"❌ Lỗi khi xuất mẫu 01/TD:\n\n{ex.Message}",
                     "❌ Lỗi",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Error
                 );
             }
         }

         private async void Btn03_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     return ExportSpecificTemplate(customer, "03 DS.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("Export (03 DS) thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi export 03: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void BtnGUQ_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     return ExportSpecificTemplate(customer, "GUQ.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("Export GUQ thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi export GUQ: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void Btn01tgtv_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     return ExportSpecificTemplate(customer, "01TGTV.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("✅ Xuất mẫu 01TGTV thành công.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("❌ Lỗi khi xuất mẫu 01TGTV: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void BtnBia_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     return ExportSpecificTemplate(customer, "BIA.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("✅ Xuất bìa hồ sơ thành công.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("❌ Lỗi khi xuất bìa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private void BtnDelete_Click(object sender, EventArgs e)
         {
             if (dgv == null || dgv.SelectedRows.Count == 0)
             {
                 MessageBox.Show("Chọn khách để xoá.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 return;
             }

             var idx = dgv.SelectedRows[0].Index;
             if (idx < 0 || customers == null || idx >= customers.Count) return;

             var c = customers[idx];
             var r = MessageBox.Show($"Bạn có muốn xóa khách hàng \"{c.Hoten}\" không?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
             if (r != DialogResult.Yes) return;

             try
             {
                 DeleteCustomerFiles(c);
                 customers.RemoveAt(idx);
                 BindGrid();
                 ClearForm();
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

        private List<Customer> GetSelectedCustomers() { var list = new List<Customer>(); try { foreach (DataGridViewRow row in dgv.SelectedRows) try { var item = row.DataBoundItem as Customer; if (item != null) list.Add(item); } catch { } } catch { } return list; }

        private void Btn03Group_Click(object sender, EventArgs e) { try { var selected = GetSelectedCustomers(); var f2 = new Form2(selected); f2.ShowDialog(); } catch (Exception ex) { MessageBox.Show("Lỗi khi mở Form nhóm: " + ex.Message); } }

        private void BtnTaokh_Click(object sender, EventArgs e)
        {
            try
            {
                var customer = ReadForm();

                if (string.IsNullOrWhiteSpace(customer.Hoten))
                {
                    MessageBox.Show("Vui lòng nhập Họ và tên để tạo khách hàng mới.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                customer._fileName = null;

                try
                {
                    SaveCustomerToFile(customer);

                    if (customers != null)
                    {
                        customers.Add(customer);
                    }

                    BindGrid();

                    MessageBox.Show($"Đã tạo khách hàng mới '{customer.Hoten}' thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    editingIndex = -1;
                    ClearForm();

                    try
                    {
                        if (dgv != null && dgv.SelectedRows.Count > 0)
                        {
                            dgv.ClearSelection();
                        }
                    }
                    catch { }

                    try { txtHoten.Focus(); } catch { }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi tạo khách hàng mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tạo mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            try
            {
                try { if (txtUsername != null) txtUsername.Text = ""; } catch { }
                try { if (txtPassword != null) txtPassword.Text = ""; } catch { }

                try
                {
                    if (xinManEditor != null)
                    {
                        xinManEditor.Logout();
                    }
                }
                catch { }

                MessageBox.Show("Đã đăng xuất thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đăng xuất: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void Btnall_Click(object sender, EventArgs e)
        {
            if (!ValidateRequiredFields()) return;

            try
            {
                var customer = ReadForm();
                string existingFile = null;
                if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                {
                    try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                }

                if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                SaveCustomerToFile(customer);
                UpsertCustomerInList(customer);

                var allCreatedFiles = new List<string>();

                await Task.Run(() =>
                {
                    try
                    {
                        var files01 = CreateProfileFromTemplate(customer, include03: false);
                        if (files01 != null) allCreatedFiles.AddRange(files01);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu 01: {ex.Message}");
                    }

                    try
                    {
                        var file03 = ExportSpecificTemplate(customer, "03 DS.docx");
                        if (!string.IsNullOrEmpty(file03)) allCreatedFiles.Add(file03);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu 03: {ex.Message}");
                    }

                    try
                    {
                        var fileGUQ = ExportSpecificTemplate(customer, "GUQ.docx");
                        if (!string.IsNullOrEmpty(fileGUQ)) allCreatedFiles.Add(fileGUQ);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu GUQ: {ex.Message}");
                    }

                    try
                    {
                        var file01TGTV = ExportSpecificTemplate(customer, "01TGTV.docx");
                        if (!string.IsNullOrEmpty(file01TGTV)) allCreatedFiles.Add(file01TGTV);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu 01TGTV: {ex.Message}");
                    }

                    try
                    {
                        var fileBIA = ExportSpecificTemplate(customer, "BIA.docx");
                        if (!string.IsNullOrEmpty(fileBIA)) allCreatedFiles.Add(fileBIA);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Lỗi tạo bìa hồ sơ: {ex.Message}");
                    }
                });

                BindGrid();
                ClearForm();

                MessageBox.Show(
                    $"✅ Đã tạo toàn bộ hồ sơ thành công!\n\n" +
                    $"📄 Khách hàng: {customer.Hoten}\n" +
                    $"📁 Số file tạo: {allCreatedFiles.Count}\n\n" +
                    $"Bao gồm:\n" +
                    $"- Bìa hồ sơ (BIA)\n" +
                    $"- Mẫu 01 (TD/SXKD/GQVL)\n" +
                    $"- Mẫu 03 DS\n" +
                    $"- Mẫu GUQ\n" +
                    $"- Mẫu 01TGTV",
                    "✅ Tạo toàn bộ hồ sơ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                OpenCreatedFiles(allCreatedFiles);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Lỗi khi tạo toàn bộ hồ sơ:\n\n{ex.Message}",
                    "❌ Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}
