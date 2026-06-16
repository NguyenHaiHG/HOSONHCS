using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace HOSONHCS
{
    public class Customer
    {
        public string Hoten { get; set; }
        public string Socccd { get; set; }
        public string GioiTinh { get; set; }
        public string Dantoc { get; set; }
        public string Sdt { get; set; } = "";
        public string Stk { get; set; } = "";
        public string Nganhang { get; set; } = "";
        public string Nhankhau { get; set; } = "";
        public string Nhandang { get; set; }
        public DateTime Ngaycap { get; set; }
        public DateTime Ngaysinh { get; set; }
        public string Noicap { get; set; }
        public string Xa { get; set; }
        public string Thon { get; set; }
        public string Hoi { get; set; }
        public string Totruong { get; set; }
        public string To { get; set; }
        public string Tinh { get; set; }
        public string PGD { get; set; }
        public string Chuongtrinh { get; set; }
        public string Vtc { get; set; }
        public string Phuongan { get; set; }
        public DateTime Ngaydenhan { get; set; }
        public DateTime Ngaygiaingaan { get; set; }
        public DateTime Thoihancccd { get; set; }
        public string ThoihancccdText { get; set; }
        public string Thoihanvay { get; set; }
        public string Sotien { get; set; }
        public string Sotien1 { get; set; }
        public string Sotien2 { get; set; }
        public string Sotien3 { get; set; }
        public string Sotientong { get; set; }
        public string Sotienchu { get; set; }
        public string Soluong1 { get; set; }
        public string Soluong2 { get; set; }
        public string Soluong3 { get; set; }
        public string Mucdich1 { get; set; }
        public string Mucdich2 { get; set; }
        public string Mucdich3 { get; set; }
        public string Doituong1 { get; set; }
        public string Doituong2 { get; set; }
        public DateTime Ngaylaphs { get; set; }
        public string Phanky { get; set; }

        // GUQ fields
        public string Ntk1 { get; set; } = "";
        public string Ntk2 { get; set; } = "";
        public string Ntk3 { get; set; } = "";
        public string Ntk4 { get; set; } = "";
        public string CccdNtk1 { get; set; } = "";
        public string CccdNtk2 { get; set; } = "";
        public string CccdNtk3 { get; set; } = "";
        public string CccdNtk4 { get; set; } = "";
        public string Namsinh1 { get; set; } = "";
        public string Namsinh2 { get; set; } = "";
        public string Namsinh3 { get; set; } = "";
        public string Namsinh4 { get; set; } = "";
        public string Qh1 { get; set; } = "";
        public string Qh2 { get; set; } = "";
        public string Qh3 { get; set; } = "";
        public string Qh4 { get; set; } = "";

        // internal: filename used to persist
        [JsonIgnore]
        public string _fileName { get; set; }

        public Customer()
        {
            Ngaycap = DateTime.MinValue;
            Ngaysinh = DateTime.MinValue;
            Ngaylaphs = DateTime.MinValue;
            Ngaydenhan = DateTime.MinValue;
            Ngaygiaingaan = DateTime.MinValue;
            Thoihancccd = DateTime.MinValue;
        }
    }

    public class XinManModel
    {
        public string pgd { get; set; }
        public List<Commune> communes { get; set; } = new List<Commune>();
    }

    public class Commune
    {
        public string name { get; set; }
        public List<Association> associations { get; set; } = new List<Association>();
        public List<Village> villages { get; set; } = new List<Village>();
    }

    public class Association
    {
        public string name { get; set; }
        public string code { get; set; }
        public List<Village> villages { get; set; } = new List<Village>();
        public List<string> managedVillages { get; set; } = new List<string>();
    }

    public class Village
    {
        public string name { get; set; }
        public List<string> groups { get; set; } = new List<string>();
    }

    public class TinhPgdEntry
    {
        public string pgd { get; set; }
        public List<Commune> communes { get; set; } = new List<Commune>();
    }

    public class TinhModel
    {
        public string tinh { get; set; }
        public List<TinhPgdEntry> pgds { get; set; } = new List<TinhPgdEntry>();
    }
}
