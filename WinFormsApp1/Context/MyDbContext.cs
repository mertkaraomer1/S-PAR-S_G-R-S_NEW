using Microsoft.EntityFrameworkCore;
using SİPARİŞ_GİRİŞ.Tables;
using WinFormsApp1.Tables;

namespace WinFormsApp1.Context
{
    public class MyDbContext : DbContext
    {
        public DbSet<ISEMIRLERI> ISEMIRLERI { get; set; }
        public DbSet<URETIM_MALZEME_PLANLAMA> URETIM_MALZEME_PLANLAMA { get; set; }
        public DbSet<ISEMIRLERI_USER> ISEMIRLERI_USER { get; set; }
        public DbSet<STOKLAR> STOKLAR { get; set; }
        public DbSet<SATINALMA_TALEPLERI> SATINALMA_TALEPLERI { get; set; }
        public DbSet<PROFORMA_SIPARISLER> PROFORMA_SIPARISLER { get; set; }
        public DbSet<SIPARISLER> SIPARISLER { get; set; }
        public DbSet<URETIM_ROTA_PLANLARI> URETIM_ROTA_PLANLARI { get; set; }
        public DbSet<URETIM_OPERASYON_HAREKETLERI> URETIM_OPERASYON_HAREKETLERI { get; set; }
        public DbSet<URUN_RECETELERI> URUN_RECETELERI { get; set; }
        public DbSet<URUN_ROTALARI> URUN_ROTALARI {  get; set; }
        public DbSet<VERILEN_TEKLIFLER> VERILEN_TEKLIFLER {  get; set; }
        public DbSet<IS_MERKEZLERI> IS_MERKEZLERI { get; set; }
        public DbSet<URUNLER> URUNLER { get; set; }
        public DbSet<zz_bom>zz_bom { get; set; }
        public DbSet<SAYIM_SONUCLARI> SAYIM_SONUCLARI { get; set; }
        public DbSet<CARI_HESAPLAR> CARI_HESAPLAR { get; set; }
        public DbSet<CARI_HESAP_HAREKETLERI> CARI_HESAP_HAREKETLERI { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Burada veritabanı bağlantı bilgilerini tanımlayın.
            // Örnek olarak SQL Server kullanalım:
            string connectionString = "Data Source=SRV-MIKRO;Initial Catalog=MikroDB_V16_ICM;Integrated Security=True;Connect Timeout=10;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";
            optionsBuilder.UseSqlServer(connectionString);
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // Tablo adını ve sütun adlarını özelleştirme
            modelBuilder.Entity<ISEMIRLERI>().ToTable("ISEMIRLERI").HasKey(x => x.is_Guid);
            modelBuilder.Entity<ISEMIRLERI>().Property(e => e.is_ProjeKodu).HasColumnName("is_ProjeKodu");
            modelBuilder.Entity<URETIM_MALZEME_PLANLAMA>().ToTable("URETIM_MALZEME_PLANLAMA").HasKey(x => x.upl_Guid);
            modelBuilder.Entity<ISEMIRLERI_USER>().ToTable("ISEMIRLERI_USER").HasKey(x => x.Record_uid);
            modelBuilder.Entity<STOKLAR>().ToTable("STOKLAR").HasKey(x => x.sto_Guid);
            modelBuilder.Entity<SATINALMA_TALEPLERI>().ToTable("SATINALMA_TALEPLERI").HasKey(x => x.stl_Guid);
            modelBuilder.Entity<PROFORMA_SIPARISLER>().ToTable("PROFORMA_SIPARISLER").HasKey(x => x.pro_Guid);
            modelBuilder.Entity<SIPARISLER>().ToTable("SIPARISLER").HasKey(x => x.sip_Guid);
            modelBuilder.Entity<URETIM_ROTA_PLANLARI>().ToTable("URETIM_ROTA_PLANLARI").HasKey(x => x.RtP_Guid);
            modelBuilder.Entity<URETIM_OPERASYON_HAREKETLERI>().ToTable("URETIM_OPERASYON_HAREKETLERI").HasKey(x => x.OpT_Guid);
            modelBuilder.Entity<URUN_RECETELERI>().ToTable("URUN_RECETELERI").HasKey(x => x.rec_Guid);
            modelBuilder.Entity<VERILEN_TEKLIFLER>().ToTable("VERILEN_TEKLIFLER").HasKey(x=>x.tkl_Guid);
            modelBuilder.Entity<URUN_ROTALARI>().ToTable("URUN_ROTALARI").HasKey(x=>x.URt_Guid);
            modelBuilder.Entity<IS_MERKEZLERI>().ToTable("IS_MERKEZLERI").HasKey(x=>x.IsM_Guid);
            modelBuilder.Entity<URUNLER>().ToTable("URUNLER").HasKey(x => x.uru_Guid);
            modelBuilder.Entity<zz_bom>().ToTable("zz_bom").HasKey(x=>x.lrf);
            modelBuilder.Entity<SAYIM_SONUCLARI>().ToTable("SAYIM_SONUCLARI").HasKey(x => x.sym_Guid);
            modelBuilder.Entity<CARI_HESAPLAR>().ToTable("CARI_HESAPLAR").HasKey(x => x.cari_Guid);
            modelBuilder.Entity<CARI_HESAP_HAREKETLERI>().ToTable("CARI_HESAP_HAREKETLERI").HasKey(x => x.cha_Guid);
        }

    }
}
