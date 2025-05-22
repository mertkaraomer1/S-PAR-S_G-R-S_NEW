
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class STOKLAR
    {
        public Guid sto_Guid { get; set; }
        public string sto_kod { get; set; }
        public string sto_isim { get; set; }
        public string sto_marka_kodu {  get; set; }
        public double sto_birim1_agirlik { get; set; }
        public string sto_birim2_ad { get; set; }
        public double sto_birim2_katsayi { get; set; }
        public short sto_siparis_sure { get; set; }
        public string sto_anagrup_kod { get; set; }
        public string sto_altgrup_kod { get; set; }
    }
}
