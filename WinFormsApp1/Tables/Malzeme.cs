using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class Malzeme
    {
        [Description("Stok kodu")]
        public string StokKodu { get; set; }

        [Description("Malzeme Adı")]
        public string MalAdi { get; set; }

        [Description("Miktar")]
        public string İhtiyacMiktar { get; set; }
    }
}
