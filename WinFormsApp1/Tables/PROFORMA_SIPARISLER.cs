using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class PROFORMA_SIPARISLER
    {
        public Guid pro_Guid { get; set; }
        public int pro_evrakno_sira { get; set; }
        public int pro_satirno { get; set; }
        public string pro_stokkodu { get; set; }
        public double pro_bfiyati { get; set; }
        public double pro_miktar { get; set; }
        public double pro_tutari { get; set; }
        public string pro_aciklama { get; set; }
        public string pro_aciklama2 { get; set; }
        public string pro_projekodu { get; set; }

    }
}
