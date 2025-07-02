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
    public partial class YÖNETİM_EKRANI : Form
    {
        private MyDbContext dbContext;
        private TLBContext TLBdbContext;
        public YÖNETİM_EKRANI()
        {
            TLBdbContext = new TLBContext();
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        // DataTable oluştur
        System.Data.DataTable dt = new System.Data.DataTable();
        private void YÖNETİM_EKRANI_Load(object sender, EventArgs e)
        {
            Getir();
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            advancedDataGridView1.DataSource = null; // Bu yeterli
            advancedDataGridView1.Columns.Clear();   // İsteğe bağlı (gerekirse)
            Getir();
        }
        public void Getir()
        {

            dt.Columns.Add("Ay - Yıl", typeof(string)); // İlk sütun

            DateTime baslangicTarihi = new DateTime(2024, 1, 1);

            for (int i = 0; i < 24; i++)
            {
                DateTime tarih = baslangicTarihi.AddMonths(i);
                string ayYil = tarih.ToString("MMMM yyyy", new System.Globalization.CultureInfo("tr-TR"));
                dt.Rows.Add(ayYil);

                if (tarih.Month == 12)
                {
                    DataRow toplamYilSatir = dt.NewRow();
                    toplamYilSatir["Ay - Yıl"] = $"Toplam {tarih.Year}";
                    dt.Rows.Add(toplamYilSatir);
                }
            }

            advancedDataGridView1.DataSource = dt;

            advancedDataGridView1.Columns.Add("Satınalma Talep Toplam Tutar", "Satınalma Talep Toplam Tutar");
            advancedDataGridView1.Columns.Add("Toplam Açık Sipariş Tutar", "Toplam Açık Sipariş Tutar");
            advancedDataGridView1.Columns.Add("Toplam Sipariş Tutar", "Toplam Sipariş Tutar");
            advancedDataGridView1.Columns.Add("Satınalma Siparişlerine Kesilen Toplam Fatura Tutarı", "Satınalma Siparişlerine Kesilen Toplam Fatura Tutarı");
            advancedDataGridView1.Columns.Add("Satış Siparişlerine Kesilen Toplam Fatura Tutarı", "Satış Siparişlerine Kesilen Toplam Fatura Tutarıı");
            advancedDataGridView1.Columns.Add("Makine Satış Açık Sipariş Toplam Tutar", "Makine Satış Açık Sipariş Toplam Tutar");
            advancedDataGridView1.Columns.Add("Satış Sonrası Açık Sipariş Toplam Tutar", "Satış Sonrası Açık Sipariş Toplam Tutar");



            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                string olusmaTarihi = row.Cells["Ay - Yıl"].Value?.ToString();

                if (!string.IsNullOrEmpty(olusmaTarihi))
                {
                    if (!DateTime.TryParseExact(olusmaTarihi, "MMMM yyyy", new System.Globalization.CultureInfo("tr-TR"), System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                        continue;

                    int ay = parsedDate.Month;
                    int yil = parsedDate.Year;

                    var siparisler = dbContext.SIPARISLER
                        .Where(x => x.sip_evrakno_seri == "" && x.sip_teslim_tarih.Year == yil && x.sip_teslim_tarih.Month == ay && x.sip_teslim_miktar < x.sip_miktar && x.sip_kapat_fl != true)
                        .Select(x => new { x.sip_miktar, x.sip_teslim_miktar, x.sip_b_fiyat, x.sip_doviz_kuru })
                        .ToList();

                    var Makinesiparisler = dbContext.SIPARISLER
                        .Where(x => (x.sip_evrakno_seri == "MX" || x.sip_evrakno_seri == "MS") && x.sip_teslim_tarih.Year == yil && x.sip_teslim_tarih.Month == ay && x.sip_teslim_miktar < x.sip_miktar)
                        .Select(x => new { x.sip_miktar, x.sip_teslim_miktar, x.sip_b_fiyat, x.sip_doviz_kuru })
                        .ToList();

                    var SatısSonrasısiparisler = dbContext.SIPARISLER
                        .Where(x => (x.sip_evrakno_seri == "XS" || x.sip_evrakno_seri == "3S" || x.sip_evrakno_seri == "2S" || x.sip_evrakno_seri == "OX") && x.sip_teslim_tarih.Year == yil && x.sip_teslim_tarih.Month == ay && x.sip_teslim_miktar < x.sip_miktar)
                        .Select(x => new { x.sip_miktar, x.sip_teslim_miktar, x.sip_b_fiyat, x.sip_doviz_kuru })
                        .ToList();

                    var ToplamSatınalmaSiparisler = dbContext.SIPARISLER
                        .Where(x => x.sip_evrakno_seri == "" && x.sip_teslim_tarih.Year == yil && x.sip_teslim_tarih.Month == ay && x.sip_teslim_tarih.Date >= new DateTime(2025, 1, 1))
                        .Select(x => new { x.sip_miktar, x.sip_teslim_miktar, x.sip_b_fiyat, x.sip_doviz_kuru })
                        .ToList();

                    var satinAlmalar = TLBdbContext.SATINALMA_TALEPLERI1
                        .Where(x => x.stl_teslim_tarihi.Year == yil && x.stl_teslim_tarihi.Month == ay && x.stl_miktari != x.stl_teslim_miktari && x.stl_teslim_miktari == 0 && x.stl_kapat_fl != true)
                        .ToList();

                    var stokKodlari = satinAlmalar.Select(x => x.stl_Stok_kodu).Distinct().ToList();

                    var siparisler1 = dbContext.SIPARISLER
                        .Where(s => stokKodlari.Contains(s.sip_stok_kod))
                        .OrderByDescending(s => s.sip_tarih)
                        .ToList()
                        .GroupBy(s => s.sip_stok_kod)
                        .Select(g => new
                        {
                            StokKodu = g.Key,
                            Fiyat = g.First().sip_b_fiyat,
                            Kur = g.First().sip_doviz_kuru
                        })
                        .ToList();

                    double toplamFiyat = satinAlmalar
                        .Join(siparisler1,
                              stl => stl.stl_Stok_kodu,
                              sip => sip.StokKodu,
                              (stl, sip) => Math.Round(Convert.ToDouble(sip.Fiyat) * Convert.ToDouble(sip.Kur) * stl.stl_miktari, 2))
                        .Sum();

                    var SatısFatura = dbContext.CARI_HESAP_HAREKETLERI
                        .Where(x => x.cha_belge_tarih.Year == yil && x.cha_belge_tarih.Month == ay && x.cha_evrak_tip == 63 && x.cha_projekodu != "")
                        .Select(x => new { x.cha_aratoplam, x.cha_d_kur })
                        .ToList();

                    var AlısFatura = dbContext.CARI_HESAP_HAREKETLERI
                        .Where(x => x.cha_belge_tarih.Year == yil && x.cha_belge_tarih.Month == ay && x.cha_evrak_tip == 0 && x.cha_projekodu != "")
                        .Select(x => new { x.cha_aratoplam, x.cha_d_kur })
                        .ToList();

                    double toplamTutar = siparisler.Sum(x => (x.sip_miktar - x.sip_teslim_miktar) * x.sip_b_fiyat * x.sip_doviz_kuru);
                    double MakinatoplamTutar = Makinesiparisler.Sum(x => (x.sip_miktar - x.sip_teslim_miktar) * x.sip_b_fiyat * x.sip_doviz_kuru);
                    double SAtisSonrasıtoplamTutar = SatısSonrasısiparisler.Sum(x => (x.sip_miktar - x.sip_teslim_miktar) * x.sip_b_fiyat * x.sip_doviz_kuru);
                    double MaliYıltoplamTutar = ToplamSatınalmaSiparisler.Sum(x => x.sip_miktar * x.sip_b_fiyat * x.sip_doviz_kuru);
                    double SatısFaturaTutar = SatısFatura.Sum(x => x.cha_aratoplam * x.cha_d_kur);
                    double AlısFaturaTutar = AlısFatura.Sum(x => x.cha_aratoplam * x.cha_d_kur);

                    row.Cells["Satınalma Talep Toplam Tutar"].Value = toplamFiyat.ToString("C2");
                    row.Cells["Toplam Açık Sipariş Tutar"].Value = toplamTutar.ToString("C2");
                    row.Cells["Toplam Sipariş Tutar"].Value = MaliYıltoplamTutar.ToString("C2");
                    row.Cells["Satış Siparişlerine Kesilen Toplam Fatura Tutarı"].Value = SatısFaturaTutar.ToString("C2");
                    row.Cells["Satınalma Siparişlerine Kesilen Toplam Fatura Tutarı"].Value = AlısFaturaTutar.ToString("C2");
                    row.Cells["Makine Satış Açık Sipariş Toplam Tutar"].Value = MakinatoplamTutar.ToString("C2");
                    row.Cells["Satış Sonrası Açık Sipariş Toplam Tutar"].Value = SAtisSonrasıtoplamTutar.ToString("C2");

                }
            }

            var hesaplanacakSutunlar = new[]
       {
    "Satınalma Talep Toplam Tutar",
    "Toplam Açık Sipariş Tutar",
    "Toplam Sipariş Tutar",
    "Satınalma Siparişlerine Kesilen Toplam Fatura Tutarı",
    "Satış Siparişlerine Kesilen Toplam Fatura Tutarı",
    "Makine Satış Açık Sipariş Toplam Tutar",
    "Satış Sonrası Açık Sipariş Toplam Tutar"
};

            var yillar = new[] { 2024, 2025 };

            foreach (int yil in yillar)
            {
                var toplamRow = advancedDataGridView1.Rows
                    .Cast<DataGridViewRow>()
                    .FirstOrDefault(r => r.Cells["Ay - Yıl"].Value?.ToString() == $"Toplam {yil}");

                if (toplamRow == null) continue;

                // 🟡 Sarı renge boyama
                toplamRow.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;

                foreach (string columnName in hesaplanacakSutunlar)
                {
                    double toplam = 0;

                    foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                    {
                        string ayYil = row.Cells["Ay - Yıl"].Value?.ToString();
                        if (ayYil == null || ayYil.StartsWith("Toplam")) continue;

                        if (DateTime.TryParseExact(ayYil, "MMMM yyyy", new System.Globalization.CultureInfo("tr-TR"),
                            System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                        {
                            if (parsedDate.Year == yil)
                            {
                                string degerStr = row.Cells[columnName].Value?.ToString()?.Replace("₺", "").Replace(".", "").Replace(",", ".") ?? "0";

                                if (double.TryParse(degerStr, System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out double value))
                                {
                                    toplam += value;
                                }
                            }
                        }
                    }

                    toplamRow.Cells[columnName].Value = toplam.ToString("C2", new System.Globalization.CultureInfo("tr-TR"));
                }
            }

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
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
    }
}
