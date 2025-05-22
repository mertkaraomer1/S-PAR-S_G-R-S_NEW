using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1;

namespace WinFormsApp1.Tables
{
    public class ISEMIRLERI
    {

        public Guid is_Guid { get; set; }
        public string is_ProjeKodu { get; set; }
        public string is_BagliOlduguIsemri { get; set; }
        public string is_SiparisNo_Seri { get; set; }
        public int is_SiparisNo_Sira { get; set; }
        public byte is_EmriDurumu { get; set; }
        public string is_Kod { get; set; }
        public DateTime is_create_date { get; set; }
        public DateTime is_lastup_date { get; set; }
        public DateTime is_Emri_PlanBaslamaTarihi { get; set; }


    }
}
