using System;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1
    {
        private void ConfigureGqvlInputRestrictions()
        {
            ConfigureNumberOnly(txtSohd);
            ConfigureNumberOnly(cbStGQVL);
            ConfigureNumberOnly(cbThoihanvayGQVL);
            ConfigureNumberOnly(cbPhankyGQVL);
            ConfigureNumberOnly(cbLsGQVL);
            ConfigureNumberOnly(txtStkGQVL);
            ConfigureNumberOnly(txtSdtkhGQVL, 12);
            ConfigureNumberOnly(txtCccdkhGQVL, 12);
            ConfigureNumberOnly(txtSdtPGD);
            cbStGQVL.TextChanged += CbStGQVL_TextChanged;
            txtSdtkhGQVL.TextChanged += TxtSdtkhGQVL_TextChanged;

            ConfigureTextOnly(txtMdGQVL);
            ConfigureTextOnly(txtTenkhGQVL);
            ConfigureTextOnly(txtNoicapGQVL);
            ConfigureTextOnly(txtTenPGD);
            ConfigureTextOnly(txtTenld);
        }

        private void ConfigureNumberOnly(TextBox textBox, int maxLength = 0)
        {
            if (textBox == null) return;
            if (maxLength > 0) textBox.MaxLength = maxLength;
            textBox.KeyPress += GqvlNumberOnly_KeyPress;
        }

        private void ConfigureNumberOnly(ComboBox comboBox)
        {
            if (comboBox == null) return;
            comboBox.KeyPress += GqvlNumberOnly_KeyPress;
        }

        private void ConfigureTextOnly(TextBox textBox)
        {
            if (textBox == null) return;
            textBox.KeyPress += GqvlTextOnly_KeyPress;
        }

        private void GqvlNumberOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar)) return;

            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void GqvlTextOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar) || char.IsWhiteSpace(e.KeyChar)) return;

            if (!char.IsLetter(e.KeyChar))
                e.Handled = true;
        }

        private bool suppressGqvlMoneyFormat;
        private bool suppressGqvlPhoneFormat;

        private void CbStGQVL_TextChanged(object sender, EventArgs e)
        {
            if (suppressGqvlMoneyFormat) return;
            try
            {
                var comboBox = sender as ComboBox;
                if (comboBox == null) return;

                string digits = GetDigits(comboBox.Text);
                if (string.IsNullOrEmpty(digits)) return;

                long value;
                if (!long.TryParse(digits, out value)) return;

                string formatted = string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:N0}", value).Replace(",", ".");
                if (formatted == comboBox.Text) return;

                suppressGqvlMoneyFormat = true;
                comboBox.Text = formatted;
                comboBox.SelectionStart = comboBox.Text.Length;
            }
            finally
            {
                suppressGqvlMoneyFormat = false;
            }
        }

        private void TxtSdtkhGQVL_TextChanged(object sender, EventArgs e)
        {
            if (suppressGqvlPhoneFormat) return;
            try
            {
                var textBox = sender as TextBox;
                if (textBox == null) return;

                string digits = GetDigits(textBox.Text);
                if (digits.Length > 10) digits = digits.Substring(0, 10);

                string formatted = FormatPhoneGqvl(digits);
                if (formatted == textBox.Text) return;

                suppressGqvlPhoneFormat = true;
                textBox.Text = formatted;
                textBox.SelectionStart = textBox.Text.Length;
            }
            finally
            {
                suppressGqvlPhoneFormat = false;
            }
        }

        private static string FormatPhoneGqvl(string digits)
        {
            if (string.IsNullOrEmpty(digits)) return string.Empty;
            if (digits.Length <= 4) return digits;
            if (digits.Length <= 7) return digits.Substring(0, 4) + "." + digits.Substring(4);
            return digits.Substring(0, 4) + "." + digits.Substring(4, 3) + "." + digits.Substring(7);
        }

        private static string GetDigits(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var chars = new System.Text.StringBuilder();
            foreach (char c in value)
            {
                if (char.IsDigit(c)) chars.Append(c);
            }
            return chars.ToString();
        }
    }
}
