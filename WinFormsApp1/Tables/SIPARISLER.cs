using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class SIPARISLER
    {
        public Guid sip_Guid { get; set; }
        public string sip_evrakno_seri { get; set; }
        public byte sip_tip { get; set; }
        public int sip_evrakno_sira { get; set; }
        public int sip_satirno { get; set; }
        public string sip_stok_kod { get; set; }
        public double sip_miktar { get; set; }
        public double sip_b_fiyat {  get; set; }
        public double sip_doviz_kuru {  get; set; }
        public DateTime sip_create_date { get; set; }
        public DateTime sip_tarih {  get; set; }
        public double sip_tutar {  get; set; }
        public string sip_aciklama {  get; set; }
        public double sip_teslim_miktar { get; set; }
        public string sip_cari_sormerk {  get; set; }
        public string sip_stok_sormerk { get; set; }
        public string sip_musteri_kod { get; set; }
        public byte sip_birim_pntr {  get; set; }
        public string sip_projekodu { get; set; }
        public DateTime sip_teslim_tarih { get; set; }
        public bool? sip_kapat_fl { get; set; }

    }
}
