using System;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    internal class CauHinhCoDinhGqvl
    {
        public string TenPgd { get; set; }
        public string TenLanhDao { get; set; }
        public string SdtPgd { get; set; }
        public string DiaChiPgd { get; set; }
        public bool Ong { get; set; }
        public bool GiamDoc { get; set; }
        public string SoUyQuyen { get; set; }
        public DateTime NgayUyQuyen { get; set; }
    }

    public partial class Form1
    {
        private bool suppressGqvlPersistentSettings;

        private string GetGqvlPersistentSettingsPath()
        {
            return Path.Combine(GetGqvlDataFolder(), "cauhinh_codinh_gqvl.json");
        }

        private void LoadGqvlPersistentSettings()
        {
            try
            {
                string path = GetGqvlPersistentSettingsPath();
                if (!File.Exists(path)) return;

                var config = JsonConvert.DeserializeObject<CauHinhCoDinhGqvl>(File.ReadAllText(path));
                if (config == null) return;

                suppressGqvlPersistentSettings = true;

                txtTenPGD.Text = config.TenPgd ?? "";
                txtTenld.Text = config.TenLanhDao ?? "";
                txtSdtPGD.Text = config.SdtPgd ?? "";
                txtDcPGD.Text = config.DiaChiPgd ?? "";
                rdO.Checked = config.Ong;
                rdB.Checked = !config.Ong;
                rdGd.Checked = config.GiamDoc;
                rdPgd.Checked = !config.GiamDoc;
                txtUq.Text = config.SoUyQuyen ?? "";
                if (config.NgayUyQuyen != DateTime.MinValue)
                    dtNgayuq.Value = config.NgayUyQuyen;
            }
            catch { }
            finally
            {
                suppressGqvlPersistentSettings = false;
            }
        }

        private void SaveGqvlPersistentSettings()
        {
            if (suppressGqvlPersistentSettings) return;

            try
            {
                Directory.CreateDirectory(GetGqvlDataFolder());
                var config = new CauHinhCoDinhGqvl
                {
                    TenPgd = txtTenPGD.Text,
                    TenLanhDao = txtTenld.Text,
                    SdtPgd = txtSdtPGD.Text,
                    DiaChiPgd = txtDcPGD.Text,
                    Ong = rdO.Checked,
                    GiamDoc = rdGd.Checked,
                    SoUyQuyen = txtUq.Text,
                    NgayUyQuyen = dtNgayuq.Value.Date
                };

                File.WriteAllText(GetGqvlPersistentSettingsPath(), JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            catch { }
        }

        private void AttachGqvlPersistentSettingsEvents()
        {
            txtTenPGD.TextChanged += GqvlPersistentSetting_Changed;
            txtTenld.TextChanged += GqvlPersistentSetting_Changed;
            txtSdtPGD.TextChanged += GqvlPersistentSetting_Changed;
            txtDcPGD.TextChanged += GqvlPersistentSetting_Changed;
            txtUq.TextChanged += GqvlPersistentSetting_Changed;
            dtNgayuq.ValueChanged += GqvlPersistentSetting_Changed;
            rdO.CheckedChanged += GqvlPersistentSetting_Changed;
            rdB.CheckedChanged += GqvlPersistentSetting_Changed;
            rdGd.CheckedChanged += GqvlPersistentSetting_Changed;
            rdPgd.CheckedChanged += GqvlPersistentSetting_Changed;
        }

        private void GqvlPersistentSetting_Changed(object sender, EventArgs e)
        {
            SaveGqvlPersistentSettings();
        }
    }
}
