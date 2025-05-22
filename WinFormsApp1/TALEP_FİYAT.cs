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
using WinFormsApp1.Tables;

namespace SİPARİŞ_GİRİŞ
{
    public partial class TALEP_FİYAT : Form
    {
        private MyDbContext dbContext;
        private TLBContext TLBdbContext;

        public TALEP_FİYAT()
        {
            TLBdbContext = new TLBContext();
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        List<SATINALMA_TALEPLERI1> satınalmaList = new List<SATINALMA_TALEPLERI1>();
        private void button3_Click(object sender, EventArgs e)
        {
            // Kullanıcıdan alınan sıra numarasını tamsayıya dönüştürün
            int sirano = Convert.ToInt32(textBox3.Text);

            // Veritabanından satın alma taleplerini alın
            var satınalma = TLBdbContext.SATINALMA_TALEPLERI1.Where(x => x.stl_evrak_sira == sirano)
                .ToList();

            // Satın alma taleplerini satınalmaList'e ekleyin
            satınalmaList.AddRange(satınalma);

            MessageBox.Show("Eklendi...");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Sonuçları tutacak tablo
            System.Data.DataTable sonucTablosu = new System.Data.DataTable();
            sonucTablosu.Columns.Add("Ay", typeof(string)); // Nisan 2025, Mayıs 2025 gibi
            sonucTablosu.Columns.Add("Toplam Fiyat", typeof(string));

            // Teslim tarihine göre gruplama: Ay bazında
            var aylikGruplar = satınalmaList
                .GroupBy(s => new { Ay = s.stl_teslim_tarihi.Month, Yil = s.stl_teslim_tarihi.Year })
                .OrderBy(g => g.Key.Yil).ThenBy(g => g.Key.Ay);

            foreach (var grup in aylikGruplar)
            {
                double toplamFiyat = 0;

                foreach (var satinalma in grup)
                {
                    string stokKodu = satinalma.stl_Stok_kodu;
                    double miktar = Convert.ToDouble(satinalma.stl_miktari);
                    string projeKodu = satinalma.stl_Sor_Merk;
                    textBox1.Text = projeKodu;
                    // SIPARISLER tablosundan bu stok koduna ait en son sipariş
                    var eslesenSiparis = dbContext.SIPARISLER
                        .Where(s => s.sip_stok_kod == stokKodu)
                        .OrderByDescending(s => s.sip_tarih)
                        .Select(x => new
                        {
                            x.sip_b_fiyat,
                            x.sip_doviz_kuru
                        })
                        .FirstOrDefault();

                    if (eslesenSiparis != null)
                    {
                        double birimFiyat = Convert.ToDouble(eslesenSiparis.sip_b_fiyat);
                        double doviz = Convert.ToDouble(eslesenSiparis.sip_doviz_kuru);
                        double fiyat = Math.Round((birimFiyat * doviz) * miktar, 2);
                        toplamFiyat += fiyat;
                    }
                }

                // Yeni satır ekle
                DataRow yeniSatir = sonucTablosu.NewRow();
                string ayAdi = new DateTime(grup.Key.Yil, grup.Key.Ay, 1).ToString("MMMM yyyy", new System.Globalization.CultureInfo("tr-TR"));
                yeniSatir["Ay"] = ayAdi;
                // TL ibaresiyle ve para birimi formatında yaz
                yeniSatir["Toplam Fiyat"] = toplamFiyat.ToString("C2", new System.Globalization.CultureInfo("tr-TR")); // örn: 12.345,67 ₺
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
            // 1. Listeyi temizle
            satınalmaList.Clear();
            textBox1.Clear();
            // 2. DataGridView'ı temizle
            advancedDataGridView1.DataSource = null;
            advancedDataGridView1.Rows.Clear();
            advancedDataGridView1.Refresh();
        }
    }
}
