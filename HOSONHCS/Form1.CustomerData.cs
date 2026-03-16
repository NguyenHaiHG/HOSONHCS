using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    public partial class Form1 : Form
    {
        // ============================================
        // CUSTOMER DATA (CRUD, FORM READ/POPULATE)
        // ============================================

        private string GetCustomersFolderPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customers");
        }

        private void EnsureCustomersFolder()
        {
            var folder = GetCustomersFolderPath();
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
        }

        private void LoadCustomersFromFiles()
        {
            customers = new BindingList<Customer>();
            try
            {
                EnsureCustomersFolder();
                var folder = GetCustomersFolderPath();
                foreach (var file in Directory.EnumerateFiles(folder, "*.json"))
                {
                    try
                    {
                        var json = File.ReadAllText(file, Encoding.UTF8);
                        var c = JsonConvert.DeserializeObject<Customer>(json);
                        if (c != null)
                        {
                            c._fileName = Path.GetFileName(file);
                            customers.Add(c);
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load customer files: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            customers.ListChanged += Customers_ListChanged;
        }

        private void Customers_ListChanged(object sender, ListChangedEventArgs e)
        {
        }

        private void BindGrid()
        {
            try
            {
                dgv.DataSource = null;
                dgv.DataSource = customers;

                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    col.ReadOnly = true;
                }

                if (dgv.Columns["Hoten"] != null)
                {
                    dgv.Columns["Hoten"].DisplayIndex = 1;
                    dgv.Columns["Hoten"].HeaderText = "Họ và tên";
                }
                if (dgv.Columns["_fileName"] != null) dgv.Columns["_fileName"].Visible = false;
            }
            catch { }
        }

        private string GetCustomerJsonPath(Customer c)
        {
            if (!string.IsNullOrEmpty(c._fileName))
                return Path.Combine(GetCustomersFolderPath(), c._fileName);

            var baseName = MakeFileSystemSafe(c.Hoten);
            var file = baseName + ".json";
            var folder = GetCustomersFolderPath();
            var path = Path.Combine(folder, file);
            int i = 1;
            while (File.Exists(path))
            {
                file = $"{baseName}_{i}.json";
                path = Path.Combine(folder, file);
                i++;
            }
            return path;
        }

        private string GetCustomerJsonPathByName(string hoten)
        {
            var baseName = MakeFileSystemSafe(hoten);
            var folder = GetCustomersFolderPath();
            var candidate = Path.Combine(folder, baseName + ".json");
            if (File.Exists(candidate)) return candidate;
            var found = Directory.EnumerateFiles(folder, baseName + "*.json").FirstOrDefault();
            return found;
        }

        private void SaveCustomerToFile(Customer c)
        {
            EnsureCustomersFolder();
            string path;
            if (!string.IsNullOrEmpty(c._fileName))
                path = Path.Combine(GetCustomersFolderPath(), c._fileName);
            else
            {
                var baseName = MakeFileSystemSafe(c.Hoten);
                var file = baseName + ".json";
                var folder = GetCustomersFolderPath();
                path = Path.Combine(folder, file);
                int i = 1;
                while (File.Exists(path))
                {
                    var existing = File.ReadAllText(path, Encoding.UTF8);
                    try
                    {
                        var ec = JsonConvert.DeserializeObject<Customer>(existing);
                        if (ec != null && string.Equals(MakeFileSystemSafe(ec.Hoten), MakeFileSystemSafe(c.Hoten), StringComparison.OrdinalIgnoreCase))
                            break;
                    }
                    catch { }
                    file = $"{baseName}_{i}.json";
                    path = Path.Combine(folder, file);
                    i++;
                }
                c._fileName = Path.GetFileName(path);
            }

            UpdateComputedFields(c);

            var json = JsonConvert.SerializeObject(c, Formatting.Indented);
            File.WriteAllText(path, json, Encoding.UTF8);
        }

        private void DeleteCustomerFiles(Customer c)
        {
            try
            {
                if (!string.IsNullOrEmpty(c._fileName))
                {
                    var jsonPath = Path.Combine(GetCustomersFolderPath(), c._fileName);
                    if (File.Exists(jsonPath)) File.Delete(jsonPath);
                }
                else
                {
                    var found = Directory.EnumerateFiles(GetCustomersFolderPath(), MakeFileSystemSafe(c.Hoten) + "*.json").FirstOrDefault();
                    if (!string.IsNullOrEmpty(found) && File.Exists(found)) File.Delete(found);
                }

                var folder = GetProfileFolderPath(c);
                if (Directory.Exists(folder)) Directory.Delete(folder, true);
            }
            catch { }
        }

        private void UpsertCustomerInList(Customer customer)
        {
            if (customer == null) return;

            try
            {
                if (editingIndex >= 0 && editingIndex < customers.Count)
                {
                    customer._fileName = customers[editingIndex]._fileName;
                    customers[editingIndex] = customer;
                }
                else
                {
                    customers.Add(customer);
                }

                editingIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi cập nhật danh sách: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Customer ReadForm()
        {
             string ntk1 = "", ntk2 = "", ntk3 = "";
             string cccdntk1 = "", cccdntk2 = "", cccdntk3 = "";
             string namsinh1 = "", namsinh2 = "", namsinh3 = "";
             string qh1 = "", qh2 = "", qh3 = "";

             try { if (txtntk1 != null) ntk1 = txtntk1.Text.Trim(); } catch { }
             try { if (txtntk2 != null) ntk2 = txtntk2.Text.Trim(); } catch { }
             try { if (txtntk3 != null) ntk3 = txtntk3.Text.Trim(); } catch { }

             try { if (txtcccd1 != null) cccdntk1 = txtcccd1.Text.Trim(); } catch { }
             try { if (txtcccd2 != null) cccdntk2 = txtcccd2.Text.Trim(); } catch { }
             try { if (txtcccd3 != null) cccdntk3 = txtcccd3.Text.Trim(); } catch { }

             try 
             { 
                 if (datentk1 != null) 
                 {
                     if (datentk1.Checked)
                         namsinh1 = datentk1.Value.ToString("dd/MM/yyyy");
                     else
                         namsinh1 = "";
                 }
             } catch { }
             try 
             { 
                 if (datentk2 != null) 
                 {
                     if (datentk2.Checked)
                         namsinh2 = datentk2.Value.ToString("dd/MM/yyyy");
                     else
                         namsinh2 = "";
                 }
             } catch { }
             try 
             { 
                 if (datentk3 != null) 
                 {
                     if (datentk3.Checked)
                         namsinh3 = datentk3.Value.ToString("dd/MM/yyyy");
                     else
                         namsinh3 = "";
                 }
             } catch { }

             try { if (cbqh1 != null) qh1 = cbqh1.Text.Trim(); } catch { }
             try { if (cbqh2 != null) qh2 = cbqh2.Text.Trim(); } catch { }
             try { if (cbqh3 != null) qh3 = cbqh3.Text.Trim(); } catch { }

             DateTime ngaycap = dateNgaycapCCCD.Value.Date;
             if (ngaycap > DateTime.Today) ngaycap = DateTime.Today;

             DateTime ngaysinh = DateTime.MinValue;
             if (dateNgaysinh != null)
             {
                 ngaysinh = dateNgaysinh.Value.Date;
                 if (ngaysinh > DateTime.Today) ngaysinh = DateTime.Today;
             }

             DateTime ngaylaphs = dateLaphs.Value.Date;
             if (ngaylaphs > DateTime.Today) ngaylaphs = DateTime.Today;

             DateTime ngaydenhan = DateTime.MinValue;
             if (dateDH != null && dateDH.Checked)
             {
                 ngaydenhan = dateDH.Value.Date;
             }

             DateTime thoihancccd = DateTime.MinValue;
             if (datendhcccd != null)
             {
                 thoihancccd = datendhcccd.Value.Date;

                 if (thoihancccd < DateTime.Today)
                 {
                     throw new Exception($"CCCD đã hết hạn ngày {thoihancccd:dd/MM/yyyy}.\n\nKhông thể tạo hồ sơ với CCCD hết hạn.\n\nVui lòng cập nhật CCCD mới.");
                 }
             }

             return new Customer
             {
                 Hoten = ToTitleCase(txtHoten.Text.Trim()),
                 Socccd = txtSocccd.Text,
                 GioiTinh = (cbGioitinh != null ? cbGioitinh.Text : ""),
                 Nhandang = cbNhandang.Text,
                 Ngaycap = ngaycap,
                 Ngaysinh = ngaysinh,
                 Noicap = cbNoicap.Text,
                 Xa = cbXa.Text,
                 Thon = cbThon.Text,
                 Hoi = cbHoi.Text,
                 Totruong = cbTo.Text,
                 To = "",
                 PGD = cbPGD.Text,
                 Chuongtrinh = cbChuongtrinh.Text,
                 Vtc = (cbVtc != null ? cbVtc.Text : ""),
                 Phuongan = (cbPhuongan != null ? cbPhuongan.Text : ""),
                 Thoihanvay = cbThoihanvay.Text,
                 Phanky = (cbPhanky != null ? cbPhanky.Text : ""),
                 Sotien = cbSotien.Text,
                 Sotien1 = cbSotien1.Text,
                 Sotien2 = cbSotien2.Text,
                 Soluong1 = (cbDoituong1 != null ? cbDoituong1.Text : ""),
                 Soluong2 = (cbDoituong2 != null ? cbDoituong2.Text : ""),
                 Sotientong = "",
                 Sotienchu = "",
                 Mucdich1 = (cbmucdich1 != null ? cbmucdich1.Text : ""),
                 Mucdich2 = (cbmucdich2 != null ? cbmucdich2.Text : ""),
                 Doituong1 = (cbDoituong != null ? cbDoituong.Text : ""),
                 Doituong2 = "",
                 Ngaylaphs = ngaylaphs,
                 Ngaydenhan = ngaydenhan,
                 Thoihancccd = thoihancccd,
                 Dantoc = (cbDantoc != null ? cbDantoc.Text : ""),
                 Sdt = (txtSdt != null ? txtSdt.Text : ""),
                 Nhankhau = (txtNhankhau != null ? txtNhankhau.Text.Trim() : ""),
                 Ntk1 = ToTitleCase(ntk1), Ntk2 = ToTitleCase(ntk2), Ntk3 = ToTitleCase(ntk3),
                 CccdNtk1 = cccdntk1, CccdNtk2 = cccdntk2, CccdNtk3 = cccdntk3,
                 Namsinh1 = namsinh1, Namsinh2 = namsinh2, Namsinh3 = namsinh3,
                 Qh1 = qh1, Qh2 = qh2, Qh3 = qh3
             };
        }

        private void PopulateForm(Customer c)
        {
             if (c == null) return;

             txtHoten.Text = c.Hoten ?? "";
             txtSocccd.Text = c.Socccd ?? "";
             cbNhandang.Text = c.Nhandang ?? "";
             var ngaycap = c.Ngaycap == DateTime.MinValue ? DateTime.Today : c.Ngaycap;
             if (ngaycap > DateTime.Today) ngaycap = DateTime.Today;
             dateNgaycapCCCD.Value = ngaycap;

             try 
             { 
                 if (dateNgaysinh != null) 
                 {
                     var ngaysinh = c.Ngaysinh == DateTime.MinValue ? DateTime.Today : c.Ngaysinh;
                     if (ngaysinh > DateTime.Today) ngaysinh = DateTime.Today;
                     dateNgaysinh.Value = ngaysinh;
                 }
             } catch { }
             cbNoicap.Text = c.Noicap ?? "";

             try { if (cbGioitinh != null) cbGioitinh.Text = c.GioiTinh ?? ""; } catch { }
             try { if (cbDantoc != null) cbDantoc.Text = c.Dantoc ?? ""; } catch { }
             try { if (txtSdt != null) txtSdt.Text = c.Sdt ?? ""; } catch { }
             try { if (txtNhankhau != null) txtNhankhau.Text = c.Nhankhau ?? ""; } catch { }

             suppressComboChanged = true;
             try
             {
                 cbPGD.Text = c.PGD ?? ""; 

                 string jsonFileName = GetJsonFileNameFromPGD(c.PGD ?? "");

                 LoadXinManData(jsonFileName);
                 if (cbXa.Items.Count == 0 && xinmanModel != null) 
                     foreach (var com in xinmanModel.communes) 
                         if (!string.IsNullOrWhiteSpace(com.name) && !cbXa.Items.Contains(com.name)) 
                             cbXa.Items.Add(com.name);

                 try { if (!string.IsNullOrEmpty(c.Xa)) { if (!cbXa.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Xa, StringComparison.OrdinalIgnoreCase))) cbXa.Items.Add(c.Xa); cbXa.Text = c.Xa; } else cbXa.Text = ""; } catch { }
                 try { if (!string.IsNullOrEmpty(c.Thon)) { if (!cbThon.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Thon, StringComparison.OrdinalIgnoreCase))) cbThon.Items.Add(c.Thon); cbThon.Text = c.Thon; } else cbThon.Text = ""; } catch { }
                 try { if (!string.IsNullOrEmpty(c.Hoi)) { if (!cbHoi.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Hoi, StringComparison.OrdinalIgnoreCase))) cbHoi.Items.Add(c.Hoi); cbHoi.Text = c.Hoi; } else cbHoi.Text = ""; } catch { }
                 try { var toVal = !string.IsNullOrEmpty(c.Totruong) ? c.Totruong : (!string.IsNullOrEmpty(c.To) ? c.To : ""); if (!string.IsNullOrEmpty(toVal)) { if (!cbTo.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), toVal, StringComparison.OrdinalIgnoreCase))) cbTo.Items.Add(toVal); cbTo.Text = toVal; } else cbTo.Text = ""; } catch { }
             }
             finally { suppressComboChanged = false; }

             cbChuongtrinh.Text = c.Chuongtrinh ?? "";
             try { if (cbVtc != null) cbVtc.Text = c.Vtc ?? ""; } catch { }
             try { if (cbPhuongan != null) cbPhuongan.Text = c.Phuongan ?? ""; } catch { }
             cbThoihanvay.Text = c.Thoihanvay ?? "";
             try { if (cbPhanky != null) cbPhanky.Text = c.Phanky ?? ""; } catch { }

             cbSotien.Text = c.Sotien ?? "";
             cbSotien1.Text = c.Sotien1 ?? "";
             cbSotien2.Text = c.Sotien2 ?? "";

             try { if (cbmucdich1 != null) cbmucdich1.Text = c.Mucdich1 ?? ""; } catch { }
             try { if (cbmucdich2 != null) cbmucdich2.Text = c.Mucdich2 ?? ""; } catch { }
             try { if (cbDoituong != null) cbDoituong.Text = c.Doituong1 ?? ""; } catch { }
             try { if (cbDoituong1 != null) cbDoituong1.Text = c.Soluong1 ?? ""; } catch { }
             try { if (cbDoituong2 != null) cbDoituong2.Text = c.Soluong2 ?? ""; } catch { }

             var ngaylaphs = c.Ngaylaphs == DateTime.MinValue ? DateTime.Today : c.Ngaylaphs;
             if (ngaylaphs > DateTime.Today) ngaylaphs = DateTime.Today;
             dateLaphs.Value = ngaylaphs;
             try 
             { 
                 if (dateDH != null) 
                 {
                     dateDH.ShowCheckBox = true;
                     if (c.Ngaydenhan == DateTime.MinValue)
                     {
                         dateDH.Checked = false;
                     }
                     else
                     {
                         var ngaydenhan = c.Ngaydenhan;
                         if (ngaydenhan > DateTime.Today) ngaydenhan = DateTime.Today;
                         dateDH.Format = DateTimePickerFormat.Custom;
                         dateDH.CustomFormat = "dd/MM/yyyy";
                         dateDH.Checked = true;
                         dateDH.Value = ngaydenhan;
                     }
                 }
             } catch { }
             try 
             { 
                 if (datendhcccd != null) 
                 {
                     if (c.Thoihancccd == DateTime.MinValue)
                     {
                         datendhcccd.Format = DateTimePickerFormat.Custom;
                         datendhcccd.CustomFormat = "dd/MM/yyyy";
                         datendhcccd.Value = DateTime.Now;
                     }
                     else
                     {
                         var thoihancccd = c.Thoihancccd;
                         datendhcccd.Format = DateTimePickerFormat.Custom;
                         datendhcccd.CustomFormat = "dd/MM/yyyy";
                         datendhcccd.Value = thoihancccd;
                     }
                 }
             } catch { }

             try { if (txtntk1 != null) txtntk1.Text = c.Ntk1 ?? ""; } catch { }
             try { if (txtntk2 != null) txtntk2.Text = c.Ntk2 ?? ""; } catch { }
             try { if (txtntk3 != null) txtntk3.Text = c.Ntk3 ?? ""; } catch { }

             try { if (txtcccd1 != null) txtcccd1.Text = c.CccdNtk1 ?? ""; } catch { }
             try { if (txtcccd2 != null) txtcccd2.Text = c.CccdNtk2 ?? ""; } catch { }
             try { if (txtcccd3 != null) txtcccd3.Text = c.CccdNtk3 ?? ""; } catch { }

             try 
             { 
                 if (datentk1 != null) 
                 {
                     datentk1.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh1))
                     {
                         datentk1.Checked = false;
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh1);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk1.Checked = true;
                             datentk1.Value = parsedDate;
                         }
                         else
                         {
                             datentk1.Checked = false;
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datentk2 != null) 
                 {
                     datentk2.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh2))
                     {
                         datentk2.Checked = false;
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh2);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk2.Checked = true;
                             datentk2.Value = parsedDate;
                         }
                         else
                         {
                             datentk2.Checked = false;
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datentk3 != null) 
                 {
                     datentk3.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh3))
                     {
                         datentk3.Checked = false;
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh3);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk3.Checked = true;
                             datentk3.Value = parsedDate;
                         }
                         else
                         {
                             datentk3.Checked = false;
                         }
                     }
                 }
             } catch { }

             try { if (cbqh1 != null) cbqh1.Text = c.Qh1 ?? ""; } catch { }
             try { if (cbqh2 != null) cbqh2.Text = c.Qh2 ?? ""; } catch { }
             try { if (cbqh3 != null) cbqh3.Text = c.Qh3 ?? ""; } catch { }
        }

        private void ClearForm()
        {
             try { txtHoten.Clear(); } catch { } try { txtSocccd.Text = ""; } catch { } try { cbNhandang.Text = ""; } catch { }
             try { dateNgaycapCCCD.Value = DateTime.Today; } catch { }
             try { if (dateNgaysinh != null) dateNgaysinh.Value = DateTime.Today; } catch { }
             try { cbNoicap.Text = ""; } catch { } try { cbXa.Items.Clear(); cbThon.Items.Clear(); cbHoi.Items.Clear(); cbTo.Items.Clear(); } catch { }
             try { cbXa.Text = ""; cbThon.Text = ""; cbHoi.Text = ""; cbTo.Text = ""; } catch { }
             try { cbChuongtrinh.Text = ""; cbThoihanvay.Text = ""; cbSotien.Text = ""; cbSotien1.Text = ""; cbSotien2.Text = ""; } catch { }
             try { if (cbmucdich1 != null) cbmucdich1.Text = ""; if (cbmucdich2 != null) cbmucdich2.Text = ""; } catch { }
             try { dateLaphs.Value = DateTime.Today; cbPGD.Text = ""; editingIndex = -1; ResetVisibilityToDefault(); } catch { }

             try { if (cbDoituong != null) cbDoituong.Enabled = true; cbDoituong.Text = ""; } catch { }

             try { if (dateDH != null) dateDH.Checked = false; } catch { }
             try { if (datendhcccd != null) datendhcccd.Checked = false; } catch { }
             try { if (datentk1 != null) datentk1.Checked = false; } catch { }
             try { if (datentk2 != null) datentk2.Checked = false; } catch { }
             try { if (datentk3 != null) datentk3.Checked = false; } catch { }

             try { if (txtntk1 != null) txtntk1.Text = ""; } catch { }
             try { if (txtntk2 != null) txtntk2.Text = ""; } catch { }
             try { if (txtntk3 != null) txtntk3.Text = ""; } catch { }
             try { if (txtcccd1 != null) txtcccd1.Text = ""; } catch { }
             try { if (txtcccd2 != null) txtcccd2.Text = ""; } catch { }
             try { if (txtcccd3 != null) txtcccd3.Text = ""; } catch { }
             try { if (cbqh1 != null) cbqh1.Text = ""; } catch { }
             try { if (cbqh2 != null) cbqh2.Text = ""; } catch { }
             try { if (cbqh3 != null) cbqh3.Text = ""; } catch { }

             try { if (txtSdt != null) txtSdt.Text = ""; } catch { }
             try { if (txtNhankhau != null) txtNhankhau.Text = ""; } catch { }
        }
    }
}
