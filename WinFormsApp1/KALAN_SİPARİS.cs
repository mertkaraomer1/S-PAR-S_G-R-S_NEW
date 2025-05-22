using Microsoft.Office.Interop.Excel;
using SİPARİŞ_GİRİŞ.Context;
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
    public partial class KALAN_SİPARİS : Form
    {
        private MyDbContext dbContext;
        private TLBContext TLBdbContext;

        public KALAN_SİPARİS()
        {
            TLBdbContext = new TLBContext();
            dbContext = new MyDbContext();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Sonuçları tutacak tablo
            System.Data.DataTable sonucTablosu = new System.Data.DataTable();
            sonucTablosu.Columns.Add("Ay", typeof(string));
            sonucTablosu.Columns.Add("Toplam Tutar", typeof(string));

            var siparisler = dbContext.SIPARISLER
                .Where(x => x.sip_evrakno_seri == ""
                    && x.sip_teslim_tarih.Date >= new DateTime(2025, 1, 1)
                    && x.sip_teslim_miktar < x.sip_miktar)
                .Select(x => new
                {
                    x.sip_teslim_tarih,
                    x.sip_miktar,
                    x.sip_teslim_miktar,
                    x.sip_b_fiyat,
                    x.sip_doviz_kuru
                })
                .OrderByDescending(x => x.sip_teslim_tarih)
                .ToList();




            // Teslim tarihine göre gruplama: Ay ve Yıl bazında
            var aylikGruplar = siparisler
                .GroupBy(s => new { Ay = s.sip_teslim_tarih.Month, Yil = s.sip_teslim_tarih.Year })
                .OrderBy(g => g.Key.Yil).ThenBy(g => g.Key.Ay).ToList();

            foreach (var grup in aylikGruplar)
            {
                double toplamTutar = 0;

                foreach (var siparis in grup)
                {
                    double miktarFarki = Convert.ToDouble(siparis.sip_miktar) - Convert.ToDouble(siparis.sip_teslim_miktar);


                    double birimFiyat = Convert.ToDouble(siparis.sip_b_fiyat);
                    double doviz = Convert.ToDouble(siparis.sip_doviz_kuru);
                    double tutar = Math.Round(miktarFarki * birimFiyat * doviz, 2);
                    toplamTutar += tutar;
                }

                DataRow yeniSatir = sonucTablosu.NewRow();
                string ayAdi = new DateTime(grup.Key.Yil, grup.Key.Ay, 1).ToString("MMMM yyyy", new System.Globalization.CultureInfo("tr-TR"));
                yeniSatir["Ay"] = ayAdi;

                // TL ibaresiyle ve para birimi formatında yaz
                yeniSatir["Toplam Tutar"] = toplamTutar.ToString("C2", new System.Globalization.CultureInfo("tr-TR")); // örn: 12.345,67 ₺

                sonucTablosu.Rows.Add(yeniSatir);
            }


            // AdvancedDataGridView'e ata
            advancedDataGridView1.DataSource = sonucTablosu;

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            System.Data.DataTable dt = new System.Data.DataTable();

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
                    if (cell.Value != null) // Hücre değeri null değilse
                    {
                        dataRow[cell.ColumnIndex] = cell.Value;
                    }
                    else if (cell is DataGridViewCheckBoxCell) // CheckBox hücresi ise
                    {
                        dataRow[cell.ColumnIndex] = (cell.Value != null && (bool)cell.Value) ? "True" : "False"; // CheckBox değeri null değilse ve true ise "True" olarak ayarla, değilse "False" olarak ayarla
                    }
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasını başlatın
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalışma kitabı oluşturun
            Workbook workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            _Worksheet worksheet = (_Worksheet)workbook.Worksheets[1];

            // DataTable'ı Excel çalışma sayfasına aktarın
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
                    // Metin olarak aktar
                    worksheet.Cells[rowIndex, j + 1].NumberFormat = "@";
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            // 2. DataGridView'ı temizle
            advancedDataGridView1.DataSource = null;
            advancedDataGridView1.Rows.Clear();
            advancedDataGridView1.Refresh();
        }
    }
}
