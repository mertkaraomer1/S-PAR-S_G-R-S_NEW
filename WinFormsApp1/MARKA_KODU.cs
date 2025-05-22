using ExcelDataReader;
using SİPARİŞ_GİRİŞ.Tables;
using System.Data;
using System.Data.OleDb;
using System.IO;
using WinFormsApp1.Context;

namespace SİPARİŞ_GİRİŞ
{
    public partial class MARKA_KODU : Form
    {
        private MyDbContext dbContext;
        private SRFDbContext SRFDb;
        public MARKA_KODU()
        {
            dbContext = new MyDbContext();
            SRFDb = new SRFDbContext();
            InitializeComponent();
        }
        System.Data.DataTable data = new System.Data.DataTable();
        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx;*.xls";
                if (file.ShowDialog() == DialogResult.OK)
                {
                    string tamYol = file.FileName;

                    // Excel dosyasını okuyun
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance); // Encoding kaydet

                    using (var stream = File.Open(tamYol, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true // İlk satırı başlık olarak kullan
                                }
                            });

                            if (result.Tables.Count > 0)
                            {
                                DataTable data = result.Tables[0]; // İlk sayfadaki verileri al

                                foreach (DataRow row in data.Rows)
                                {
                                    Dictionary<string, object> item = new Dictionary<string, object>();
                                    foreach (DataColumn column in data.Columns)
                                    {
                                        item[column.ColumnName] = row[column]; // Sütun adını anahtar, hücre değerini değer olarak ekle
                                    }
                                    advancedDataGridView1.DataSource = data;
                                }

                                MessageBox.Show("BOM Listesi Yüklendi");
                            }
                            else
                            {
                                MessageBox.Show("Excel dosyasında veri bulunamadı.");
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Dosya seçilmedi.");
                }
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("Dosya hatası: " + ioEx.Message);
            }
            catch (FormatException formatEx)
            {
                MessageBox.Show("Format hatası: " + formatEx.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Bir hata oluştu: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var transaction = dbContext.Database.BeginTransaction())
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Cells["KODU"].Value != null)
                    {
                        string Kodu = row.Cells["KODU"].Value.ToString();
                        string ismi = row.Cells["İSMİ"].Value?.ToString();
                        string MARKAKODU = row.Cells["MARKA KODU"].Value?.ToString();

                        var match = dbContext.STOKLAR.Where(x => x.sto_kod == Kodu && x.sto_isim == ismi).FirstOrDefault();

                        if (match != null)
                        {
                            // Eşleşen kayıtları güncelle
                            match.sto_marka_kodu = MARKAKODU;

                            // Veritabanı güncellemesi
                            dbContext.STOKLAR.Update(match);
                        }
                    }
                }
                // Veritabanındaki değişiklikleri kaydet
                dbContext.SaveChanges();

                // Transaction'ı onayla
                transaction.Commit();
                MessageBox.Show("Güncelleme tamamlandı");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Cells["Part Number"].Value != null)
                    {
                        string kodu = row.Cells["Part Number"].Value.ToString();
                        string ismi = row.Cells["Title"].Value?.ToString();

                        // Yeni bir SARF_MALZEME_KOD nesnesi oluştur
                        var malzeme = new SARF_MALZEME_KOD
                        {
                            KODU = kodu,
                            ADİ = ismi
                            // Diğer alanlar varsa onları da burada doldurun
                        };

                        // Yeni malzemeyi veritabanına ekle
                        SRFDb.SARF_MALZEME_KOD.Add(malzeme);
                    }
                }

                // Veritabanındaki değişiklikleri kaydet
                int changes = SRFDb.SaveChanges(); // Değişiklik sayısını al

                // Eğer değişiklik yapılmadıysa, kullanıcıya bilgi ver
                if (changes > 0)
                {
                    MessageBox.Show("Güncelleme tamamlandı");
                }
                else
                {
                    MessageBox.Show("Hiçbir kayıt güncellenmedi.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata oluştu: {ex.Message}");
            }


        }


    }
}
