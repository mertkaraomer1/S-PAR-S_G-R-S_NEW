using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class URETIM_OPERASYON_HAREKETLERI
    {
        public Guid OpT_Guid { get; set; }
        public DateTime OpT_baslamatarihi { get; set; }
        public DateTime OpT_bitis_tarihi { get; set; }
        public string OpT_IsEmriKodu { get; set; }
        public string OpT_UrunKodu { get; set; }
        public string OpT_ismerkezi { get; set; }
        public string OpT_PersonelKodu { get; set; }
        public double OpT_TamamlananMiktar { get; set; }
        public int OpT_TamamlananSure { get; set; }
        public int Opt_SetupSuresi { get; set; }
        public double Opt_BozukMiktar { get; set; }
        public string OpT_OperasyonKodu { get; set; }
        public short OpT_OperasyonSafhaNo { get; set; }
        public int OpT_EvrakSatirNo {  get; set; }
    }
}
