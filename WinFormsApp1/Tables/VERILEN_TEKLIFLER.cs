using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SİPARİŞ_GİRİŞ.Tables
{
    public class VERILEN_TEKLIFLER
    {
        public Guid tkl_Guid { get; set; }
        public short tkl_DBCno { get; set; }
        public int? tkl_SpecRECno { get; set; }
        public int? tkl_iptal { get; set; }
        public int? tkl_fileid { get; set; }
        public int? tkl_hidden { get; set; }
        public int? tkl_kilitli { get; set; }
        public int? tkl_degisti { get; set; }
        public int? tkl_checksum { get; set; }
        public short? tkl_create_user { get; set; }
        public DateTime tkl_create_date { get; set; }
        public short? tkl_lastup_user { get; set; }
        public DateTime? tkl_lastup_date { get; set; }
        public string tkl_special1 { get; set; }
        public string tkl_special2 { get; set; }
        public string tkl_special3 { get; set; }
        public int? tkl_firmano { get; set; }
        public int? tkl_subeno { get; set; }
        public string tkl_stok_kod { get; set; }
        public string tkl_cari_kod { get; set; }
        public string tkl_evrakno_seri { get; set; } // Özel tür için farklı bir yapı gerekebilir
        public int tkl_evrakno_sira { get; set; }
        public DateTime? tkl_evrak_tarihi { get; set; }
        public int tkl_satirno { get; set; }
        public string tkl_belge_no { get; set; } // Özel tür için farklı bir yapı gerekebilir
        public DateTime? tkl_belge_tarih { get; set; }
        public double? tkl_asgari_miktar { get; set; }
        public short? tkl_teslimat_suresi { get; set; }
        public DateTime? tkl_baslangic_tarihi { get; set; }
        public DateTime? tkl_Gecerlilik_Sures { get; set; }
        public double? tkl_Brut_fiyat { get; set; }
        public int? tkl_Odeme_Plani { get; set; }
        public double? tkl_Alisfiyati { get; set; }
        public double? tkl_karorani { get; set; }
        public double? tkl_miktar { get; set; }
        public string tkl_Aciklama { get; set; }
        public byte? tkl_doviz_cins { get; set; }
        public double? tkl_doviz_kur { get; set; }
        public double? tkl_alt_doviz_kur { get; set; }
        public double? tkl_iskonto1 { get; set; }
        public double? tkl_iskonto2 { get; set; }
        public double? tkl_iskonto3 { get; set; }
        public double? tkl_iskonto4 { get; set; }
        public double? tkl_iskonto5 { get; set; }
        public double? tkl_iskonto6 { get; set; }
        public double? tkl_masraf1 { get; set; }
        public double? tkl_masraf2 { get; set; }
        public double? tkl_masraf3 { get; set; }
        public double? tkl_masraf4 { get; set; }
        public byte? tkl_vergi_pntr { get; set; }
        public double? tkl_vergi { get; set; }
        public byte? tkl_masraf_vergi_pnt { get; set; }
        public double? tkl_masraf_vergi { get; set; }
        public byte? tkl_isk_mas1 { get; set; }
        public byte? TKL_ISK_MAS2 { get; set; }
        public byte? TKL_ISK_MAS3 { get; set; }
        public byte? TKL_ISK_MAS4 { get; set; }
        public byte? TKL_ISK_MAS5 { get; set; }
        public byte? TKL_ISK_MAS6 { get; set; }
        public byte? TKL_ISK_MAS7 { get; set; }
        public byte? TKL_ISK_MAS8 { get; set; }
        public byte? TKL_ISK_MAS9 { get; set; }
        public byte? TKL_ISK_MAS10 { get; set; }
        public int? TKL_SAT_ISKMAS1 { get; set; }
        public int? TKL_SAT_ISKMAS2 { get; set; }
        public int? TKL_SAT_ISKMAS3 { get; set; }
        public int? TKL_SAT_ISKMAS4 { get; set; }
        public int? TKL_SAT_ISKMAS5 { get; set; }
        public int? TKL_SAT_ISKMAS6 { get; set; }
        public int? TKL_SAT_ISKMAS7 { get; set; }
        public int? TKL_SAT_ISKMAS8 { get; set; }
        public int? TKL_SAT_ISKMAS9 { get; set; }
        public int? TKL_SAT_ISKMAS10 { get; set; }
        public int? TKL_VERGISIZ_FL { get; set; }
        public int? TKL_KAPAT_FL { get; set; }
        public string TKL_TESLIMTURU { get; set; }
        public string tkl_ProjeKodu { get; set; }
        public string tkl_Sorumlu_Kod { get; set; }
        public int? tkl_adres_no { get; set; }
        public Guid tkl_yetkili_uid { get; set; }
        public byte? tkl_durumu { get; set; }
        public string tkl_TedarikEdilecekCari { get; set; }
        public int? tkl_fiyat_liste_no { get; set; }
        public double? tkl_Birimfiyati { get; set; }
        public string tkl_paket_kod { get; set; }
        public double? tkl_teslim_miktar { get; set; }
        public short? tkl_OnaylayanKulNo { get; set; }
        public bool? tkl_cagrilabilir_fl { get; set; }
        public byte? tkl_harekettipi { get; set; }
        public string tkl_cari_sormerk { get; set; }
        public string tkl_stok_sormerk { get; set; }
        public string tkl_kapatmanedenkod { get; set; }
        public string tkl_servisisemrikodu { get; set; }
        public byte? tkl_birim_pntr { get; set; }
        public byte? tkl_cari_tipi { get; set; }
        public string tkl_HareketGrupKodu1 { get; set; }
        public string tkl_HareketGrupKodu2 { get; set; }
        public string tkl_HareketGrupKodu3 { get; set; }
        public double? tkl_Olcu1 { get; set; }
        public double? tkl_Olcu2 { get; set; }
        public double? tkl_Olcu3 { get; set; }
        public double? tkl_Olcu4 { get; set; }
        public double? tkl_Olcu5 { get; set; }
        public byte? tkl_FormulMiktarNo { get; set; }
        public double? tkl_FormulMiktar { get; set; }
        public byte? tkl_Tevkifat_turu { get; set; }
        public bool? tkl_tevkifat_sifirlandi_fl { get; set; }
    }
}
