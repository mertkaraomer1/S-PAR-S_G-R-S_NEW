using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class Proforma
    {
        [Description("SatırNo")]
        public string SatırNo { get; set; }

        [Description("Kodu")]
        public string Kodu { get; set; }

        [Description("İsmi")]
        public string Ismi { get; set; }

        [Description("Proje kodu")]
        public string ProjeKodu { get; set; }

        [Description("Miktar")]
        public string Miktar { get; set; }

        [Description("Br")]
        public string Br { get; set; }

        [Description("Birim fiyat")]
        public string BirimFiyat{get; set;}
        [Description("Tutarı")]
        public string Tutari { get; set; }

        [Description("Açıklama")]
        public string? Acıklama { get; set; }

        [Description("Açıklama2")]
        public string? Acıklama2 { get; set; }

    }
}
