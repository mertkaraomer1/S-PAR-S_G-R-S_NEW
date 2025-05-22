using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SİPARİŞ_GİRİŞ.Tables
{
    public class CARI_HESAP_HAREKETLERI
    {
        public Guid cha_Guid { get; set; }
        public byte cha_evrak_tip { get; set; }
        public double cha_aratoplam { get; set; }
        public DateTime cha_belge_tarih { get; set; }
        public double cha_d_kur { get; set; }
        public string cha_srmrkkodu { get; set; }
        public string cha_projekodu { get; set; }
    }
}
