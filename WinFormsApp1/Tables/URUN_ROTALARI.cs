namespace SİPARİŞ_GİRİŞ.Tables
{
    public class URUN_ROTALARI
    {
        public Guid URt_Guid { get; set; }
        public DateTime URt_create_date { get; set; }
        public string URt_RotaUrunKodu { get; set; }
        public string URt_IsmerkeziveyaGrupKod { get; set; }
        public int? URt_SabitHazirlikSuresi { get; set; }
        public int? URt_DegiskenOperasyonSuresi { get; set; }
        public string URt_OpKod { get; set; }

    }
}
