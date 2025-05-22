using Microsoft.EntityFrameworkCore;
using SİPARİŞ_GİRİŞ.Tables;
using WinFormsApp1.Tables;

namespace WinFormsApp1.Context
{
    public class SRFDbContext : DbContext
    {
        public DbSet<SARF_MALZEME_KOD> SARF_MALZEME_KOD { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Burada veritabanı bağlantı bilgilerini tanımlayın.
            // Örnek olarak SQL Server kullanalım:
            string connectionString = "Data Source=SRV-MIKRO;Initial Catalog=Muh_Plan_Prog1;Integrated Security=True;Connect Timeout=10;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";
            optionsBuilder.UseSqlServer(connectionString);
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // Tablo adını ve sütun adlarını özelleştirme
            modelBuilder.Entity<SARF_MALZEME_KOD>().ToTable("SARF_MALZEME_KOD").HasKey(x => x.ID);

        }

    }
}
