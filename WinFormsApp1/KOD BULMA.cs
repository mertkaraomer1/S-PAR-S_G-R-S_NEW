using SİPARİŞ_GİRİŞ.Tables;
using System.Data;
using System.Data.OleDb;
using WinFormsApp1.Context;
using WinFormsApp1.Tables;

namespace SİPARİŞ_GİRİŞ
{
    public partial class KOD_BULMA : Form
    {
        private MyDbContext dbContext;
        public KOD_BULMA()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        System.Data.DataTable data = new System.Data.DataTable();
        List<kodlar> ProformaList = new List<kodlar>();
        private void button1_Click(object sender, EventArgs e)
        {
            data.Columns.Clear();
            data.Rows.Clear();
            data.Clear();
            if (checkBox2.Checked == true)
            {
                try
                {
                    OpenFileDialog file = new OpenFileDialog();
                    file.Filter = "Excel Dosyası |*.xlsx";
                    file.ShowDialog();

                    string tamYol = file.FileName;

                    string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                    OleDbConnection baglanti = new OleDbConnection(baglantiAdresi);

                    OleDbCommand komut = new OleDbCommand("Select * From [" + textBox2.Text + "$]", baglanti);

                    baglanti.Open();

                    OleDbDataAdapter da = new OleDbDataAdapter(komut);

                    da.Fill(data);

                    // Veri kaynağını DataGridView'e ata
                    advancedDataGridView1.DataSource = data;

                    MessageBox.Show("Bom Listesi Yüklendi");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else if (checkBox3.Checked == true)
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

                    //DataGrid'imizin kaynağını oluşturduğumuz DataTable ile dolduruyoruz.
                    //dataGridView1.DataSource = data;

                    ProformaList = DataTableConverter.ConvertTo<kodlar>(data);
                    MessageBox.Show("Fiyat Listesi Yüklendi");
                }
                catch (Exception ex)
                {
                    // Hata alırsak ekrana bastırıyoruz.
                    MessageBox.Show(ex.Message);
                }
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            // DataGridView'e yeni bir sütun ekleyin
            DataGridViewTextBoxColumn combinedColumn = new DataGridViewTextBoxColumn();
            combinedColumn.Name = "CombinedColumn";
            combinedColumn.HeaderText = "Combined";
            advancedDataGridView1.Columns.Add(combinedColumn);
            // Tüm satırları döngüyle gezerek 2. ve 4. sütunları birleştirin ve yeni sütuna atayın
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (!row.IsNewRow) // Yeni satır değilse işleme devam et
                {
                    string value2 = row.Cells[1].Value?.ToString();
                    string value4 = row.Cells[3].Value?.ToString();
                    row.Cells["CombinedColumn"].Value = $"{value4}-{value2}";
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // DataGridView'e yeni bir sütun ekleyin
            DataGridViewTextBoxColumn combinedColumn = new DataGridViewTextBoxColumn();
            combinedColumn.Name = "CombinedColumn1";
            combinedColumn.HeaderText = "KOD";
            advancedDataGridView1.Columns.Add(combinedColumn);


            // Tüm satırları döngüyle gezerek 4. sütundan gelen veriyi sto_isim ile eşleştir ve sto_kod'u KOD sütununa yaz
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (!row.IsNewRow) // Yeni satır değilse işleme devam et
                {
                    string value4 = row.Cells[4].Value?.ToString();

                    // sto_isim ile eşleşen sto_kod'u bul
                    var sto_kod = dbContext.STOKLAR.Where(x => x.sto_isim == value4).Select(x => x.sto_kod).FirstOrDefault();

                    // sto_kod'u KOD sütununa yaz
                    row.Cells["CombinedColumn1"].Value = sto_kod ?? "Eşleşme Yok"; // Eşleşme yoksa "Eşleşme Yok" yaz
                }
            }
        }
        System.Data.DataTable table = new System.Data.DataTable();

        private void button4_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("SATIR_NO");
            table.Columns.Add("STOK KODU");
            table.Columns.Add("TARİH");
            table.Columns.Add("TÜKETİM KODU");
            table.Columns.Add("TALEP MİKTARI");
            table.Columns.Add("HAMMADE MALİYETİ");
            table.Columns.Add("PLANLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TAMAMLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("SATINALMA FİYATI");
            table.Columns.Add("FASON MALİYETİ");
            table.Columns.Add("PLANLANAN MALİYETİ");
            table.Columns.Add("TAMAMLANAN MALİYETİ");
            table.Columns.Add("HAMMADDE-FASON-İŞÇİLİK MALİYETİ");

            int DKIsciMaliyet = Convert.ToInt32(textBox1.Text);
            int sayac = 1;

            foreach (var item in ProformaList)
            {
                var tuketimkod = dbContext.URUN_RECETELERI.Where(x => x.rec_anakod.Substring(0, 13) == item.PartNumber && !x.rec_tuketim_kod.StartsWith("9"))
                 .Select(x => new
                 {

                     item.PartNumber,
                     item.ItemQTY,
                     x.rec_tuketim_kod,
                     gereklimiktar = x.rec_tuketim_miktar * int.Parse(item.ItemQTY)
                 }
             ).ToList();

                var tuketimkod1 = dbContext.URUN_RECETELERI.Where(x => x.rec_anakod.Substring(0, 13) == item.PartNumber && 
                (x.rec_tuketim_kod.StartsWith("9")||x.rec_tuketim_kod.Substring(14,2)=="01"))
                     .Select(x => new
                     {

                         item.PartNumber,
                         item.ItemQTY,
                         x.rec_tuketim_kod,
                         gereklimiktar = x.rec_tuketim_miktar * int.Parse(item.ItemQTY)
                     }
                    ).ToList();
                double fiyat = 0;
                foreach (var item3 in tuketimkod1)
                {
                    var MalzemeMaliyeti = dbContext.SIPARISLER
                .Where(x => x.sip_stok_kod == item3.rec_tuketim_kod) // rec_tuketim_kod kullanıldı
                .OrderByDescending(x => x.sip_create_date)
                .Select(x => new
                {
                    TopFiyat = x.sip_b_fiyat * x.sip_doviz_kuru * item3.gereklimiktar,
                })
                .FirstOrDefault();
                    // Her bir TopFiyat değerini fiyat değişkenine ekle
                    fiyat += MalzemeMaliyeti?.TopFiyat ?? 0;
                }

                double planlananMaliyet = 0;
                double TamamlananMaliyet = 0;
                double planlananiscimaliyet = 0;
                double tamamlananiscimaliyet = 0;
                double malzememaliyet = 0;
                double fabrikamaliyet = 0;
                foreach (var item2 in tuketimkod)
                {
                    var İşçiMaliyetiPlanlanan = dbContext.URETIM_ROTA_PLANLARI
                   .Where(x => x.RtP_UrunKodu == item2.rec_tuketim_kod)
                   .OrderByDescending(x => x.RtP_create_date)
                   .GroupBy(x => x.RtP_IsEmriKodu) // RtP_IsEmriKodu'na göre grupla
                   .Select(group => new
                   {
                       RtP_IsEmriKodu = group.Key, // Gruplanmış anahtar (IsEmriKodu)
                       TopFiyatIsci = group.Sum(x => x.RtP_PlanlananSure+x.RtP_PlanlananSetupSuresi) / 60 * DKIsciMaliyet // Toplam işçi maliyeti hesapla
                   })
                   .FirstOrDefault(); // İlk grup için öğeyi seç

                    planlananiscimaliyet += İşçiMaliyetiPlanlanan?.TopFiyatIsci ?? 0;

                    var İşçiMaliyetiTamamlanan = dbContext.URETIM_OPERASYON_HAREKETLERI
                        .Where(x => x.OpT_UrunKodu == item2.rec_tuketim_kod)
                        .OrderByDescending(x => x.OpT_baslamatarihi)
                        .GroupBy(x => x.OpT_IsEmriKodu)
                        .Select(x => new
                        {
                            OpT_IsEmriKodu = x.Key,
                            TopFiyatIsciTam = x.Sum(x => (x.OpT_TamamlananSure/x.OpT_TamamlananMiktar)*int.Parse(item.ItemQTY)+x.Opt_SetupSuresi) / 60 * DKIsciMaliyet,

                        })
                        .FirstOrDefault();
                    tamamlananiscimaliyet += İşçiMaliyetiTamamlanan?.TopFiyatIsciTam ?? 0;



                    planlananMaliyet = fiyat + (İşçiMaliyetiPlanlanan?.TopFiyatIsci ?? 0);
                    TamamlananMaliyet = fiyat + (İşçiMaliyetiTamamlanan?.TopFiyatIsciTam ?? 0);
                }

                var MalzemeSatisFiyati1 = dbContext.SIPARISLER
                  .Where(x =>x.sip_stok_kod.Length>13 && x.sip_stok_kod.Substring(0, 13) == item.PartNumber)
                  .OrderByDescending(x => x.sip_create_date)
                  .Select(x => new
                  {
                      x.sip_stok_kod,
                      x.sip_tarih,
                      TopFiyatSatis = x.sip_b_fiyat * x.sip_doviz_kuru * int.Parse(item.ItemQTY),
                  })
                  .FirstOrDefault();

                var MalzemeSatisFiyati = dbContext.SIPARISLER
                  .Where(x =>x.sip_stok_kod == item.PartNumber)
                  .OrderByDescending(x => x.sip_create_date)
                  .Select(x => new
                  {
                      x.sip_stok_kod,
                      x.sip_tarih,
                      TopFiyatSatis = x.sip_b_fiyat *x.sip_doviz_kuru * int.Parse(item.ItemQTY),
                  })
                  .FirstOrDefault();
                fabrikamaliyet = TamamlananMaliyet + MalzemeSatisFiyati1?.TopFiyatSatis ?? 0;
                table.Rows.Add(
                sayac++,
                item.PartNumber,
                MalzemeSatisFiyati?.sip_tarih.Date.ToString("yyyy/MM/dd"),
                string.Join(" ; ", tuketimkod1.Select(tk => tk.rec_tuketim_kod)),
                item.ItemQTY,
                Math.Round(fiyat, 2),
                Math.Round(planlananiscimaliyet, 2),
                Math.Round(tamamlananiscimaliyet, 2),
                Math.Round(MalzemeSatisFiyati?.TopFiyatSatis ?? 0, 2),
                Math.Round(MalzemeSatisFiyati1?.TopFiyatSatis??0,2),
                Math.Round(planlananMaliyet, 2),
                Math.Round(TamamlananMaliyet, 2),
                Math.Round(fabrikamaliyet,2));
            }
            advancedDataGridView1.DataSource = table;
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
