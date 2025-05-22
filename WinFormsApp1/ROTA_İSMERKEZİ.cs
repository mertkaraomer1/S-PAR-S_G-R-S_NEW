using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinFormsApp1.Context;

namespace SİPARİŞ_GİRİŞ
{
    public partial class ROTA_İSMERKEZİ : Form
    {
        private MyDbContext dbContext;
        public ROTA_İSMERKEZİ()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }

        private void ROTA_İSMERKEZİ_Load(object sender, EventArgs e)
        {
            var kodlar = new List<string> { "001", "010", "025", "022", "044", "041", "042", "024", "026", "004", "023", "002", "075", "078" };

            // Veritabanından iş merkezlerini al
            var ismerkezleri = dbContext.IS_MERKEZLERI
                .Where(x => kodlar.Contains(x.IsM_Kodu))
                .Select(x => x.IsM_Kodu)
                .Distinct()
                .ToList();

            // Başlangıca "İş merkezi seçiniz..." ekle
            ismerkezleri.Insert(0, "İş merkezi seçiniz...");

            // ComboBox'a ata
            comboBox1.DataSource = ismerkezleri;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string isMerkeziKodu = comboBox1.SelectedItem.ToString();

            // İş merkezi seçilen iş emirleri ve ürün kodlarını al
            var ismerkezisemir = dbContext.URETIM_ROTA_PLANLARI
                .Where(x => x.RtP_PlanlananIsMerkezi == isMerkeziKodu)
                .Select(x => new { isemri = x.RtP_IsEmriKodu, urunkodu = x.RtP_UrunKodu })
                .ToList();

            var uplUrStokKodAnd0 = dbContext.URETIM_ROTA_PLANLARI
                .Join(dbContext.ISEMIRLERI,
                    uretim => uretim.RtP_IsEmriKodu,
                    isEmir => isEmir.is_Kod,
                    (uretim, isEmir) => new { Uretim = uretim, IsEmir = isEmir })
                .Where(joined =>
                    joined.IsEmir.is_EmriDurumu == 1 &&
                    joined.IsEmir.is_create_date.Date > DateTime.Parse("2025-01-01") &&
                    joined.Uretim.RtP_UrunKodu.Contains(".") &&
                    joined.Uretim.RtP_IsEmriKodu == joined.IsEmir.is_Kod &&
                    joined.Uretim.RtP_PlanlananIsMerkezi == isMerkeziKodu) // <- Bunu eklemek bazen gereksiz görünse de join sonrası filtreye yardımcı olur
                .Select(joined => new
                {
                    UplUrStokKod = joined.Uretim.RtP_UrunKodu,
                    isemrikodu = joined.IsEmir.is_Kod,
                    isemriduruu = joined.IsEmir.is_EmriDurumu,
                    projekodu = joined.IsEmir.is_ProjeKodu,
                    işmerkezi = joined.Uretim.RtP_PlanlananIsMerkezi,
                    süre = joined.Uretim.RtP_PlanlananSure, 
                    planlananbaşlamatarihi = joined.IsEmir.is_Emri_PlanBaslamaTarihi,
                })
                .Distinct() // <- Eğer tekrar varsa temizler
                .ToList();
            // isemrikodu listesi
            var isemriKodlari = uplUrStokKodAnd0.Select(x => x.isemrikodu).ToList();

            // Eşleşen iş emirleri (bağlı olduğu iş emri = isemrikodu)
            var bagliIsEmirleri = dbContext.ISEMIRLERI
                .Where(x => isemriKodlari.Contains(x.is_BagliOlduguIsemri) && x.is_EmriDurumu == 2)
                .Select(x => x.is_BagliOlduguIsemri)
                .ToHashSet();

            // AdvancedDataGridView’e ata
            advancedDataGridView1.DataSource = uplUrStokKodAnd0;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                var isemrikodu = row.Cells["isemrikodu"].Value?.ToString();
                if (isemrikodu != null && bagliIsEmirleri.Contains(isemrikodu))
                {
                    row.DefaultCellStyle.BackColor = Color.LightGreen;
                }
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                // Eğer ValueType null ise, varsayılan bir veri türü kullanabilirsiniz.
                Type columnType = column.ValueType ?? typeof(string);
                dt.Columns.Add(column.HeaderText, columnType);
            }

            // Satırları ekle
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataRow dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasını başlatın
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalışma kitabı oluşturun
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // DataTable'ı Excel çalışma sayfasına aktarın (tablo başlıklarını da dahil etmek için)
            int rowIndex = 1;

            // Başlıkları yaz
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                worksheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                worksheet.Cells[1, j + 1].Font.Bold = true;
            }

            // Verileri yaz
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowIndex++;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }
        }
    }
}
