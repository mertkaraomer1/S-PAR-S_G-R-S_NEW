using Microsoft.EntityFrameworkCore;
using System.Data;
using System.Data.OleDb;
using System.Drawing.Drawing2D;
using WinFormsApp1.Context;
using WinFormsApp1.Tables;

namespace WinFormsApp1
{
    public partial class ProformaSiparis : Form
    {
        private MyDbContext dbContext;
        public ProformaSiparis()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        List<Proforma> ProformaList = new List<Proforma>();
        private void button2_Click(object sender, EventArgs e)
        {
            //dataGridView1.Rows.Clear();
            try
            {
                // Dosya seçme penceresi açmak için
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx";
                file.ShowDialog();

                // seçtiğimiz excel'in tam yolu
                string tamYol = file.FileName;

                //Excel bağlantı adresi
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                //bağlantı 
                OleDbConnection baglanti = new(baglantiAdresi);

                //tüm verileri seçmek için select sorgumuz. Sayfa1 olan kısmı Excel'de hangi sayfayı açmak istiyosanız orayı yazabilirsiniz.
                OleDbCommand komut = new OleDbCommand("Select * From [" + textBox2.Text + "$]", baglanti);

                //bağlantıyı açıyoruz.
                baglanti.Open();

                //Gelen verileri DataAdapter'e atıyoruz.
                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                //Grid'imiz için bir DataTable oluşturuyoruz.
                System.Data.DataTable data = new System.Data.DataTable();

                //DataAdapter'da ki verileri data adındaki DataTable'a dolduruyoruz.
                da.Fill(data);

                // Tutari ve bfiyat alanlarını virgülden sonra 2 karakter olacak şekilde düzenle
                foreach (DataRow row in data.Rows)
                {
                    if (row["Tutarı"] != DBNull.Value)
                    {
                        row["Tutarı"] = decimal.Parse(row["Tutarı"].ToString()).ToString("F2");
                    }
                    if (row["Birim fiyat"] != DBNull.Value)
                    {
                        row["Birim fiyat"] = decimal.Parse(row["Birim fiyat"].ToString()).ToString("F2");
                    }
                }

                //DataGrid'imizin kaynağını oluşturduğumuz DataTable ile dolduruyoruz.
                ProformaList = DataTableConverter.ConvertTo<Proforma>(data);
                dataGridView1.DataSource = ProformaList;
                MessageBox.Show("Fiyat Listesi Yüklendi");
            }
            catch (Exception ex)
            {
                // Hata alırsak ekrana bastırıyoruz.
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (var context = new MyDbContext())
                {
                    // Disable triggers
                    context.Database.ExecuteSqlRaw("DISABLE TRIGGER ALL ON [dbo].[PROFORMA_SIPARISLER]");

                    int evrakNoSira = Convert.ToInt32(textBox1.Text);

                    // Retrieve all entities to be updated
                    var recordsToUpdate = context.PROFORMA_SIPARISLER
                        .Where(p => p.pro_evrakno_sira == evrakNoSira)
                        .ToList();

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        string stokKodu = Convert.ToString(row.Cells[1].Value);
                        string projeKodu = Convert.ToString(row.Cells[3].Value);
                        int satirno = Convert.ToInt32(row.Cells[0].Value);

                        // Find the matching entity in the recordsToUpdate list
                        var existingRecord = recordsToUpdate.FirstOrDefault(p =>
                            p.pro_stokkodu == stokKodu && p.pro_projekodu == projeKodu && p.pro_satirno == satirno);

                        if (existingRecord != null)
                        {
                            double yeniBirimFiyat = Convert.ToDouble(row.Cells[6].Value.ToString());
                            string yeniAciklama = row.Cells[8].Value != null ? row.Cells[8].Value.ToString() : string.Empty;

                            // Ensure yeniAciklama is not longer than 50 characters
                            if (yeniAciklama.Length > 50)
                            {
                                yeniAciklama = yeniAciklama.Substring(0, 50);
                            }

                            double yeniTutari = Math.Round(Convert.ToDouble(row.Cells[7].Value.ToString()), 2);

                            // Update the entity properties
                            existingRecord.pro_bfiyati = yeniBirimFiyat;
                            existingRecord.pro_tutari = yeniTutari;
                            existingRecord.pro_aciklama = yeniAciklama;
                        }
                    }


                    // Save changes to the database
                    context.SaveChanges();

                    // Enable triggers
                    context.Database.ExecuteSqlRaw("ENABLE TRIGGER ALL ON [dbo].[PROFORMA_SIPARISLER]");

                    MessageBox.Show("TÜM KAYITLAR BAŞARI İLE GÜNCELLENDİ...");
                    // Optionally, update the DataGridView's data source here
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
                MessageBox.Show("Inner Exception: " + ex.InnerException?.Message);
            }



        }

        private void ProformaSiparis_Load(object sender, EventArgs e)
        {
            toolStripLabel1.ToolTipText =

                "1.Excel Sayfa İsmini Gir\n" +

                "2.Listele\n" +

                "3.Evrak Sıra No Gir\n" +
                
                "4.Güncelle";
        }
    }
}
