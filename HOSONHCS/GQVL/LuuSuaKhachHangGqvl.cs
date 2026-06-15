using System;
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
        }
    }
}
