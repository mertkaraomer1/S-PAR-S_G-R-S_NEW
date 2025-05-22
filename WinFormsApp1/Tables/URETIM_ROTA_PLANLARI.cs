using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class URETIM_ROTA_PLANLARI
    {
        public Guid RtP_Guid { get; set; }
        public string RtP_IsEmriKodu { get; set; }
        public short RtP_OperasyonSafhaNo { get; set; }
        public int RtP_PlanlananSure { get; set; }
        public int RtP_TamamlananSure { get; set; }
        public string RtP_OperasyonKodu { get; set; }
        public string RtP_UrunKodu { get; set; }
        public double RtP_PlanlananMiktar { get; set; }
        public double RtP_TamamlananMiktar { get; set; }
        public string RtP_PlanlananIsMerkezi { get; set; }
        public int RtP_PlanlananSetupSuresi { get; set; }
        public DateTime Rtp_PlanlananBaslamaTarihi { get; set; }
        public DateTime Rtp_PlanlananBitisTarihi { get; set; }
        public DateTime RtP_create_date {  get; set; }
    }
}
