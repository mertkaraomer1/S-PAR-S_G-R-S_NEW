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
    public partial class MRP_ONCESİ_TAHMİNİ_SURE : Form
    {
        private MyDbContext dbContext;
        public MRP_ONCESİ_TAHMİNİ_SURE()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var ismerkezleri = dbContext.IS_MERKEZLERI
            .Where(x => x.IsM_Kodu == "001"
            || x.IsM_Kodu == "010"
            || x.IsM_Kodu == "025"
            || x.IsM_Kodu == "022"
            || x.IsM_Kodu == "044"
            || x.IsM_Kodu == "041"
            || x.IsM_Kodu == "042"
            || x.IsM_Kodu == "024"
            || x.IsM_Kodu == "026"
            || x.IsM_Kodu == "004"
            || x.IsM_Kodu == "023"
            || x.IsM_Kodu == "002"
            || x.IsM_Kodu == "075"
            || x.IsM_Kodu == "078")
            .Select(x => x.IsM_Kodu)
            .ToList();

            // DataGridView sütunlarını temizle
            advancedDataGridView1.Columns.Clear();


            // Sütunları ekle (ismerkezleri listesinden)
            foreach (var kod in ismerkezleri)
            {
                advancedDataGridView1.Columns.Add(kod, kod);
            }

            // Tarihi al
            DateTime olusmatarih = dateTimePicker2.Value.Date;

            // Verileri çek (toplam süreleri al)
            var reçete = dbContext.URUN_ROTALARI
                .Where(x => x.URt_create_date.Date == olusmatarih)
                .GroupBy(x => x.URt_IsmerkeziveyaGrupKod)
                .Select(g => new
                {
                    IsMerkezKod = g.Key,
                    ToplamSure = Math.Round(g.Sum(x => (double)x.URt_SabitHazirlikSuresi) / 60, 2)
                })
                .ToList();

            // Yeni bir satır oluştur
            var newRow = new DataGridViewRow();
            newRow.CreateCells(advancedDataGridView1);

            // Her sütun için ilgili IsMerkezKod varsa ToplamSure yaz
            for (int i = 0; i < ismerkezleri.Count; i++)
            {
                var kod = ismerkezleri[i];
                var veri = reçete.FirstOrDefault(r => r.IsMerkezKod == kod);
                newRow.Cells[i].Value = veri != null ? veri.ToplamSure.ToString() : "";
            }

            // Satırı ekle
            advancedDataGridView1.Rows.Add(newRow);

            int sifirOlanToplamAdet = dbContext.URUN_ROTALARI
            .Where(x => x.URt_create_date.Date == olusmatarih && x.URt_SabitHazirlikSuresi == 0 && x.URt_OpKod != "009")
            .Count();

            textBox1.Text = sifirOlanToplamAdet.ToString();

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
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
