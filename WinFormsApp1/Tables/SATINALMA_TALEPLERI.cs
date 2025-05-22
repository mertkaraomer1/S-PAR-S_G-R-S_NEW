using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class SATINALMA_TALEPLERI
    {
        [Key]
        [Column("stl_Guid")]
        public Guid stl_Guid { get; set; }

        [Column("stl_DBCno")]
        public int stl_DBCno { get; set; }

        [Column("stl_SpecRECno")]
        public int stl_SpecRECno { get; set; }

        [Column("stl_iptal")]
        public int stl_iptal { get; set; }

        [Column("stl_fileid")]
        public int stl_fileid { get; set; }

        [Column("stl_hidden")]
        public int stl_hidden { get; set; }

        [Column("stl_kilitli")]
        public int stl_kilitli { get; set; }

        [Column("stl_degisti")]
        public int stl_degisti { get; set; }

        [Column("stl_checksum")]
        public int stl_checksum { get; set; }

        [Column("stl_create_user")]
        public string stl_create_user { get; set; }

        [Column("stl_create_date")]
        public DateTime stl_create_date { get; set; }

        [Column("stl_lastup_user")]
        public string stl_lastup_user { get; set; }

        [Column("stl_lastup_date")]
        public DateTime stl_lastup_date { get; set; }

        [Column("stl_special1")]
        public string stl_special1 { get; set; }

        [Column("stl_special2")]
        public string stl_special2 { get; set; }

        [Column("stl_special3")]
        public string stl_special3 { get; set; }

        [Column("stl_firmano")]
        public int stl_firmano { get; set; }

        [Column("stl_subeno")]
        public int stl_subeno { get; set; }

        [Column("stl_tarihi")]
        public DateTime stl_tarihi { get; set; }

        [Column("stl_teslim_tarihi")]
        public DateTime stl_teslim_tarihi { get; set; }

        [Column("stl_evrak_seri")]
        public string stl_evrak_seri { get; set; }

        [Column("stl_evrak_sira")]
        public int stl_evrak_sira { get; set; }

        [Column("stl_satir_no")]
        public int stl_satir_no { get; set; }

        [Column("stl_belge_no")]
        public string stl_belge_no { get; set; }

        [Column("stl_belge_tarihi")]
        public DateTime stl_belge_tarihi { get; set; }

        [Column("stl_Sor_Merk")]
        public string stl_Sor_Merk { get; set; }

        [Column("stl_Stok_kodu")]
        public string stl_Stok_kodu { get; set; }

        [Column("stl_Satici_Kodu")]
        public string stl_Satici_Kodu { get; set; }

        [Column("stl_miktari")]
        public double stl_miktari { get; set; }

        [Column("stl_teslim_miktari")]
        public int stl_teslim_miktari { get; set; }

        [Column("stl_aciklama")]
        public string stl_aciklama { get; set; }

        [Column("stl_aciklama2")]
        public string stl_aciklama2 { get; set; }

        [Column("stl_depo_no")]
        public int stl_depo_no { get; set; }

        [Column("stl_kapat_fl")]
        public int stl_kapat_fl { get; set; }

        [Column("stl_projekodu")]
        public string stl_projekodu { get; set; }

        [Column("stl_parti_kodu")]
        public string stl_parti_kodu { get; set; }

        [Column("stl_lot_no")]
        public string stl_lot_no { get; set; }

        [Column("stl_OnaylayanKulNo")]
        public string stl_OnaylayanKulNo { get; set; }

        [Column("stl_cagrilabilir_fl")]
        public int stl_cagrilabilir_fl { get; set; }

        [Column("stl_harekettipi")]
        public int stl_harekettipi { get; set; }

        [Column("stl_talep_eden")]
        public string stl_talep_eden { get; set; }

        [Column("stl_kapatmanedenkod")]
        public string stl_kapatmanedenkod { get; set; }

        [Column("stl_KaynakSip_uid")]
        public Guid stl_KaynakSip_uid { get; set; }

        [Column("stl_HareketGrupKodu1")]
        public string stl_HareketGrupKodu1 { get; set; }

        [Column("stl_HareketGrupKodu2")]
        public string stl_HareketGrupKodu2 { get; set; }

        [Column("stl_HareketGrupKodu3")]
        public string stl_HareketGrupKodu3 { get; set; }

        [Column("stl_birim_pntr")]
        public int stl_birim_pntr { get; set; }

        [Column("stl_Franchise_Fiyati")]
        public int stl_Franchise_Fiyati { get; set; }
    }
}
