using Microsoft.EntityFrameworkCore;
using SİPARİŞ_GİRİŞ.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Tables;

namespace SİPARİŞ_GİRİŞ.Context
{
    public class TLBContext:DbContext
    {
        public DbSet<SATINALMA_TALEPLERI1> SATINALMA_TALEPLERI1 { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Burada veritabanı bağlantı bilgilerini tanımlayın.
            // Örnek olarak SQL Server kullanalım:
            string connectionString = "Data Source=SRV-MIKRO;Initial Catalog=MikroDB_V16_ICM;Integrated Security=True;Connect Timeout=10;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";
            optionsBuilder.UseSqlServer(connectionString);
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<SATINALMA_TALEPLERI1>()
                .ToTable("SATINALMA_TALEPLERI") // Tablo adı burada belirtiliyor
                .HasKey(x => x.stl_Guid); // Anahtar alanı
        }
    }
}
