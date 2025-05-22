using ExcelDataReader;
using Microsoft.EntityFrameworkCore;
using SİPARİŞ_GİRİŞ.Tables;
using System.Data;
using System.IO;
using System.Linq;
using WinFormsApp1.Context;
using Zuby.ADGV;

namespace SİPARİŞ_GİRİŞ
{
    public partial class SURE_TRANSFERI : Form
    {
        private MyDbContext dbContext;
        public SURE_TRANSFERI()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        DataTable table = new DataTable();
        private void SURE_TRANSFERI_Load(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("URUN KODU");
            table.Columns.Add("IS MERKEZI");
            table.Columns.Add("HAZIRLIK SURESI");
            table.Columns.Add("TAMAMLANAN SURE");
            table.Columns.Add("TAMAMLANAN MİKTAR");

            var süre5dk = dbContext.URETIM_OPERASYON_HAREKETLERI
                .Where(x => !x.OpT_UrunKodu.Contains("-")
                            && x.OpT_ismerkezi != null
                            && (x.OpT_UrunKodu.EndsWith(".001")
                                || x.OpT_UrunKodu.EndsWith(".002")
                                || x.OpT_UrunKodu.EndsWith(".123")
                                || x.OpT_UrunKodu.EndsWith(".010")))
                .ToList() // **Veriyi belleğe al**
                .GroupBy(x => x.OpT_UrunKodu)
                .Select(g => g
                    .OrderByDescending(x => x.OpT_TamamlananSure) // En yüksek süreye göre sırala
                    .FirstOrDefault() // En yüksek tamamlanan süreye sahip kaydı al
                )
                .Select(x => new
                {
                    x.OpT_UrunKodu,
                    x.OpT_ismerkezi,
                    planlanasüre = Math.Round(x.OpT_TamamlananMiktar != 0 ? x.Opt_SetupSuresi / x.OpT_TamamlananMiktar : 0, 0),
                    tamamlanasüre = Math.Round(x.OpT_TamamlananMiktar != 0 ? x.OpT_TamamlananSure / x.OpT_TamamlananMiktar : 0, 0),
                    x.OpT_bitis_tarihi,
                    x.OpT_TamamlananMiktar
                })
                .ToList();



            foreach (var item in süre5dk)
            {
                table.Rows.Add(item.OpT_UrunKodu, item.OpT_ismerkezi, item.planlanasüre, item.tamamlanasüre, item.OpT_TamamlananMiktar);
            }
            advancedDataGridView1.DataSource = table;
        }

        private void button1_Click(object sender, EventArgs e)
        {



            using (var transaction = dbContext.Database.BeginTransaction())
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Cells["URUN KODU"].Value != null)
                    {
                        string urunkodu = row.Cells["URUN KODU"].Value.ToString();
                        string ismerkezi = row.Cells["IS MERKEZI"].Value?.ToString();
                        int hazirliksuresi;
                        int tamamlanansure;

                        // int dönüşümü güvenli şekilde yap
                        int.TryParse(row.Cells["HAZIRLIK SURESI"].Value?.ToString(), out hazirliksuresi);
                        int.TryParse(row.Cells["TAMAMLANAN SURE"].Value?.ToString(), out tamamlanansure);

                        // veritabanından eşleşen kayıtları alın
                        var matchingdata = dbContext.URUN_ROTALARI
                            .Where(ur => !ur.URt_RotaUrunKodu.Contains("-")
                                         && ur.URt_RotaUrunKodu == urunkodu && ur.URt_DegiskenOperasyonSuresi == 0).FirstOrDefault();

                        if (matchingdata != null)
                        {
                            // eşleşen kayıtları güncelle
                            matchingdata.URt_IsmerkeziveyaGrupKod = ismerkezi;
                            matchingdata.URt_SabitHazirlikSuresi = tamamlanansure;


                            // veritabanı güncellemesi
                            dbContext.URUN_ROTALARI.Update(matchingdata);
                        }
                    }
                }
                            var güncellenecekler = dbContext.URUN_ROTALARI
                .Where(x => !x.URt_RotaUrunKodu.Contains("-")
                            && (x.URt_RotaUrunKodu.EndsWith(".014") || x.URt_RotaUrunKodu.EndsWith(".139")))
                .ToList();

                foreach (var item in güncellenecekler)
                {
                    item.URt_SabitHazirlikSuresi = 604800;
                    if (item.URt_RotaUrunKodu.EndsWith(".014"))
                    {
                        item.URt_IsmerkeziveyaGrupKod = "075";
                    }
                    else if (item.URt_RotaUrunKodu.EndsWith(".139"))
                    {
                        item.URt_IsmerkeziveyaGrupKod = "078";
                    }
                }


                dbContext.SaveChanges();

                // transaction'ı onayla
                transaction.Commit();
                MessageBox.Show("GÜNCELLEME TAMAMLANDI.");
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("sıra no");
            table.Columns.Add("ürün kodu");
            table.Columns.Add("işemri kodu");
            table.Columns.Add("miktar");
            table.Columns.Add("is emri durumu");
            DateTime targetDate = new DateTime(2024, 1, 1); // Hedef tarihi belirle

            // İlk sonuçlar
            var results = dbContext.ISEMIRLERI
                .Where(result => result.is_EmriDurumu == 2 && result.is_create_date > targetDate)
                .Join(dbContext.URETIM_ROTA_PLANLARI,
                      isem => isem.is_Kod,
                      ump => ump.RtP_IsEmriKodu,
                      (isem, ump) => new
                      {
                          ump.RtP_IsEmriKodu,
                          isem.is_create_date,
                          isem.is_Kod,
                          isem.is_EmriDurumu,
                          ump.RtP_UrunKodu,
                          isem.is_BagliOlduguIsemri,
                          ump.RtP_PlanlananMiktar
                      })
                .Where(ump => ump.RtP_UrunKodu.Contains(".")
                              && ump.RtP_UrunKodu.Substring(14, 2) == "01"
                              && (ump.RtP_UrunKodu.EndsWith(".000")
                                  || ump.RtP_UrunKodu.EndsWith(".120")
                                  || ump.RtP_UrunKodu.EndsWith(".121")
                                  || ump.RtP_UrunKodu.EndsWith(".122")
                                  || ump.RtP_UrunKodu.EndsWith(".127")))
                .ToList();


            // İlk sonuçlar
            var results1 = dbContext.ISEMIRLERI
                .Where(result => result.is_EmriDurumu == 1 && result.is_create_date > targetDate)
                .Join(dbContext.URETIM_ROTA_PLANLARI,
                      isem => isem.is_Kod,
                      ump => ump.RtP_IsEmriKodu,
                      (isem, ump) => new
                      {
                          ump.RtP_IsEmriKodu,
                          isem.is_create_date,
                          isem.is_Kod,
                          isem.is_EmriDurumu,
                          ump.RtP_UrunKodu,
                          isem.is_BagliOlduguIsemri,
                          ump.RtP_PlanlananMiktar
                      })
                .Where(ump => ump.RtP_UrunKodu.Contains(".")
                              && ump.RtP_UrunKodu.Substring(14, 2) == "02")
                .ToList();


            // Eşleştirilmiş sonuçlar
            var matchedResults = results
                .Join(results1,
                      res => res.is_BagliOlduguIsemri, // results listesindeki eşleştirilecek alan
                      res1 => res1.is_Kod,            // results1 listesindeki eşleştirilecek alan
                      (res, res1) => new
                      {
                          res.is_create_date,
                          res.is_Kod,
                          res.is_EmriDurumu,
                          res.RtP_UrunKodu,
                          res.is_BagliOlduguIsemri,
                          res.RtP_PlanlananMiktar,
                          Matched_create_date = res1.is_create_date,
                          Matched_EmriDurumu = res1.is_EmriDurumu
                      })
                .ToList();

            // Gruplama ve toplama işlemi
            var groupedResults = matchedResults
                .GroupBy(mr => mr.RtP_UrunKodu) // upl_kodu'na göre gruplama
                .Select(group => new
                {
                    UplKodu = group.Key, // Grup anahtarı (upl_kodu)
                    TotalUplMiktar = group.Sum(mr => mr.RtP_PlanlananMiktar), // upl_miktar toplamı
                    IsKodList = group.Select(x => x.is_Kod).Distinct().ToList(), // Eşsiz is_Kod listesi
                    IsEmriDurumuList = group.Select(x => x.is_EmriDurumu).Distinct().ToList() // Eşsiz is_EmriDurumu listesi
                })
                .ToList();

            // Tabloyu doldurun
            int sayac = 1;
            foreach (var item1 in groupedResults)
            {
                table.Rows.Add(
                    sayac++,
                    item1.UplKodu,
                    string.Join(", ", item1.IsKodList), // is_Kod listesini virgülle birleştir
                    item1.TotalUplMiktar,
                    string.Join(", ", item1.IsEmriDurumuList) // is_EmriDurumu listesini virgülle birleştir
                );
            }

            // AdvancedDataGridView'e tabloyu bağlayın
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

        private void button3_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("isemrikodu");
            table.Columns.Add("ilk13");
            table.Columns.Add("orta14-16");
            table.Columns.Add("son");
            table.Columns.Add("isemri durumu");
            table.Columns.Add("miktar");
            DateTime targetDate = new DateTime(2024, 1, 1); // Hedef tarihi belirle
                                                            // İlk sonuçlar

            // İlk sonuçlar
            var results = dbContext.URETIM_ROTA_PLANLARI
                .Where(ump => ump.RtP_UrunKodu.Contains(".")
                              && ump.RtP_UrunKodu.Length == 13)
                .Join(dbContext.ISEMIRLERI,
                      ump => ump.RtP_IsEmriKodu,
                      isem => isem.is_Kod,
                      (ump, isem) => new
                      {
                          ump.RtP_IsEmriKodu,
                          isem.is_create_date,
                          isem.is_Kod,
                          isem.is_EmriDurumu,
                          ump.RtP_UrunKodu,
                          isem.is_BagliOlduguIsemri,
                          ump.RtP_PlanlananMiktar
                      })
                .Where(x => x.is_EmriDurumu == 1 && x.is_create_date > targetDate).Distinct()
                .ToList(); // Belleğe al

            // TempResults için sorgu
            var tempResults = dbContext.URETIM_ROTA_PLANLARI
                .Where(ump => ump.RtP_UrunKodu.Contains(".")
                              && ump.RtP_UrunKodu.Length > 13
                              && !ump.RtP_UrunKodu.EndsWith(".014")
                                                            && !ump.RtP_UrunKodu.EndsWith(".128")
                              && ump.RtP_UrunKodu.Substring(14, 2) != "02" && ump.RtP_UrunKodu.Substring(14, 2) != "01")
                .Join(dbContext.ISEMIRLERI,
                      ump => ump.RtP_IsEmriKodu,
                      isem => isem.is_Kod,
                      (ump, isem) => new
                      {
                          ump.RtP_IsEmriKodu,
                          isem.is_create_date,
                          isem.is_Kod,
                          isem.is_EmriDurumu,
                          ump.RtP_UrunKodu,
                          isem.is_BagliOlduguIsemri,
                          ump.RtP_PlanlananMiktar
                      })
                .Where(x => x.is_EmriDurumu == 2 && x.is_create_date > targetDate).Distinct()
                .ToList(); // Belleğe al

            // Sonuçları karşılaştır ve yeni bir liste oluştur
            var matchedResults = tempResults
                .Where(temp => results.Any(result => result.RtP_UrunKodu == temp.RtP_UrunKodu.Substring(0, 13))
                ).Distinct().ToList();




            foreach (var item in matchedResults)
            {
                // RtP_UrunKodu'nu parçala
                string urunKodu = item.RtP_UrunKodu;
                string isemri = item.RtP_IsEmriKodu;
                string ilk13 = urunKodu.Substring(0, 13); // İlk 13 karakter
                string on4_17 = urunKodu.Substring(14, 2); // 14-16. karakterler (3 karakter)
                string kalan = urunKodu.Substring(17); // 17. karakterden sonrasını al
                string isemrikdou = isemri;
                // Yeni bir satır ekle
                table.Rows.Add(isemrikdou, ilk13, on4_17, kalan, item.is_EmriDurumu, item.RtP_PlanlananMiktar);
            }

            // AdvancedDataGridView'e tabloyu bağlayın
            advancedDataGridView1.DataSource = table;





        }
        List<Dictionary<string, object>> customerList = new List<Dictionary<string, object>>();
        private void button4_Click(object sender, EventArgs e)
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

                                // customerList'i temizleyin
                                customerList.Clear();

                                foreach (DataRow row in data.Rows)
                                {
                                    Dictionary<string, object> item = new Dictionary<string, object>();
                                    foreach (DataColumn column in data.Columns)
                                    {
                                        item[column.ColumnName] = row[column]; // Sütun adını anahtar, hücre değerini değer olarak ekle
                                    }
                                    customerList.Add(item);
                                }

                                // AdvancedDataGridView'e verileri yükle
                                advancedDataGridView1.DataSource = data;

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

        private void button5_Click(object sender, EventArgs e)
        {
            List<SAYIM_SONUCLARI> sayımsonucları = new List<SAYIM_SONUCLARI>();

            int sayac = 0;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    // DataGridView'den verileri al
                    int DEPO_NO = Convert.ToInt32(row.Cells["DEPO NO"].Value);
                    string STOK_KODU = row.Cells["STOK KODU"].Value.ToString();
                    double MİKTAR = Convert.ToDouble(row.Cells["MİKTAR"].Value);

                    int BİRİM = 0; // Varsayılan değer

                    var birimValue = row.Cells["BİRİM"].Value;

                    if (birimValue != null && !string.IsNullOrWhiteSpace(birimValue.ToString()))
                    {
                        // Dönüşüm yaparken hata almayı önlemek için TryParse kullan
                        if (!int.TryParse(birimValue.ToString(), out BİRİM))
                        {
                            Console.WriteLine("BİRİM değeri geçersiz bir sayı formatında.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("BİRİM değeri null veya boş.");
                    }


                    // Benzersiz değerleri ayarlayın
                    int evrakSeri = 20248; // Örnek olarak "T" olarak ayarlandı, siz iş mantığınıza göre değiştirmelisiniz
                    SAYIM_SONUCLARI talep1 = new SAYIM_SONUCLARI
                    {
                        sym_Guid = Guid.NewGuid(),
                        sym_DBCno = 0,
                        sym_SpecRECno = 0,
                        sym_iptal = 0,
                        sym_fileid = 28,
                        sym_hidden = 0,
                        sym_kilitli = 0,
                        sym_degisti = 0,
                        sym_checksum = 0,
                        sym_create_user = 31,
                        sym_create_date = new DateTime(2024, 12, 31),
                        sym_lastup_user = 31,
                        sym_lastup_date = new DateTime(2024, 12, 31),
                        sym_special1 = string.Empty,
                        sym_special2 = string.Empty,
                        sym_special3 = string.Empty,
                        sym_tarihi = new DateTime(2024, 12, 31),
                        sym_depono = DEPO_NO,
                        sym_evrakno = evrakSeri,
                        sym_satirno = sayac++,
                        sym_Stokkodu = STOK_KODU,
                        sym_reyonkodu = 0,
                        sym_koridorkodu = 0,
                        sym_rafkodu = 0,
                        sym_miktar1 = MİKTAR,
                        sym_miktar2 = 0,
                        sym_miktar3 = 0,
                        sym_miktar4 = 0,
                        sym_miktar5 = 0,
                        sym_birim_pntr = BİRİM,
                        sym_barkod = STOK_KODU,
                        sym_renkno = 0,
                        sym_bedenno = 0,
                        sym_parti_kodu = string.Empty,
                        sym_lot_no = 0,
                        sym_serino = string.Empty,


                    };

                    sayımsonucları.Add(talep1);
                }
            }

            // Veritabanına ekle
            dbContext.SAYIM_SONUCLARI.AddRange(sayımsonucları);
            dbContext.SaveChanges(); // Değişiklikleri kaydet

            MessageBox.Show("KAYDEDİLDİ...");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("sıra no");
            table.Columns.Add("ürün kodu");
            table.Columns.Add("miktar");

            DateTime targetDate = new DateTime(2024, 1, 1); // Hedef tarihi belirle

            // İlk sonuçlar
            var results = dbContext.ISEMIRLERI
                .Where(result => result.is_EmriDurumu == 1 && result.is_create_date > targetDate)
                .Join(dbContext.URETIM_ROTA_PLANLARI,
                      isem => isem.is_Kod,
                      ump => ump.RtP_IsEmriKodu,
                      (isem, ump) => new
                      {
                          ump.RtP_IsEmriKodu,
                          isem.is_create_date,
                          isem.is_Kod,
                          isem.is_EmriDurumu,
                          ump.RtP_UrunKodu,
                          isem.is_BagliOlduguIsemri,
                          ump.RtP_PlanlananMiktar
                      })
                .Where(ump => ump.RtP_UrunKodu.Contains(".")
                              && ump.RtP_UrunKodu.Substring(14, 2) == "01"
                              && (ump.RtP_UrunKodu.EndsWith(".000")
                                  || ump.RtP_UrunKodu.EndsWith(".120")
                                  || ump.RtP_UrunKodu.EndsWith(".121")
                                  || ump.RtP_UrunKodu.EndsWith(".122")
                                  || ump.RtP_UrunKodu.EndsWith(".127")))
                .ToList();
            // Tabloyu doldurun
            int sayac = 1;
            foreach (var item1 in results)
            {
                var hammaddekod = dbContext.URUN_RECETELERI.Where(x => x.rec_anakod == item1.RtP_UrunKodu).Select
                    (x => new
                    {
                        x.rec_tuketim_kod,
                        miktar = x.rec_tuketim_miktar * item1.RtP_PlanlananMiktar,
                    }
                    ).FirstOrDefault();

                table.Rows.Add(
                    sayac++,
                    hammaddekod.rec_tuketim_kod,
                    Math.Round(hammaddekod.miktar, 2)

                );
            }

            // AdvancedDataGridView'e tabloyu bağlayın
            advancedDataGridView1.DataSource = table;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("sıra no");
            table.Columns.Add("ürün kodu");
            table.Columns.Add("miktar");
            table.Columns.Add("teslim miktarı");
            DateTime targetDate = new DateTime(2024, 1, 1); // Hedef tarihi belirle

            var results = dbContext.ISEMIRLERI
                .Where(result => result.is_EmriDurumu == 1 && result.is_create_date > targetDate)
                .Join(dbContext.URETIM_MALZEME_PLANLAMA,
                      isem => isem.is_Kod,
                      ump => ump.upl_isemri,
                      (isem, ump) => ump.upl_kodu) // burada ump.upl_kodu'nun string olduğundan emin ol
                .Where(ump => ump.Contains(".") && ump.EndsWith(".014")) // ump burada string olmalı
                .ToList();


            var components = dbContext.SIPARISLER
                .Where(x => results.Contains(x.sip_aciklama)
                            && x.sip_miktar == x.sip_teslim_miktar
                            && (x.sip_stok_kod.StartsWith("02.") ||
                                x.sip_stok_kod.StartsWith("03.") ||
                                x.sip_stok_kod.StartsWith("04.") ||
                                x.sip_stok_kod.StartsWith("05.")))
                .GroupBy(x => x.sip_stok_kod) // sip_stok_kod'a göre grupla
                .Select(g => new
                {
                    sip_stok_kod = g.Key, // grup anahtarı
                    toplam_miktar = g.Sum(x => x.sip_miktar), // sip_miktarları topla
                    birim = g.First().sip_birim_pntr

                })
                .ToList();
            // Tabloyu doldurun
            int sayac = 1;
            foreach (var item1 in components)
            {

                table.Rows.Add(
                    sayac++,
                    item1.sip_stok_kod,
                    item1.toplam_miktar,
                    item1.birim

                );
            }

            // AdvancedDataGridView'e tabloyu bağlayın
            advancedDataGridView1.DataSource = table;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Cells["ROTALAR"].Value != null)
                {
                    string name = row.Cells["ROTALAR"].Value.ToString();
                    string markakod = row.Cells["MARKA KODU"].Value.ToString();
                    string anagrupkod = row.Cells["ANA GRUP KODU"].Value.ToString();
                    string altgrupkod = row.Cells["ALT GRUP KODU"].Value.ToString();
                    // STOKLAR tablosunda eşleşen tüm stok kayıtlarını getir

                    var stokListesi = dbContext.STOKLAR
                        .Where(ch => ch.sto_kod.StartsWith(name))
                        .ToList();

                    if (stokListesi.Any()) // Eğer kayıt varsa
                    {
                        // 'GÜN' değerini short (Int16) olarak dönüştür
                        if (short.TryParse(row.Cells["GÜN"].Value?.ToString(), out short agirlik))
                        {
                            foreach (var stok in stokListesi)
                            {
                                // sto_siparis_sure değerini güncelle
                                stok.sto_siparis_sure = agirlik;
                                stok.sto_marka_kodu = markakod.Length > 25 ? markakod.Substring(0, 25) : markakod;
                                stok.sto_anagrup_kod = anagrupkod.Length > 25 ? anagrupkod.Substring(0, 25) : anagrupkod;
                                stok.sto_altgrup_kod = altgrupkod.Length > 25 ? altgrupkod.Substring(0, 25) : altgrupkod;

                            }

                            // Değişiklikleri kaydet
                            dbContext.SaveChanges();
                        }
                        else
                        {
                            MessageBox.Show("Hata: 'GÜN' değeri geçerli bir sayı değil!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            MessageBox.Show("GÜNCELLEME TAMAMLANDI.");

        }

        private void button9_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Cells["ROTALAR"].Value != null)
                {
                    string name = row.Cells["ROTALAR"].Value.ToString();
                    string markakod = row.Cells["SURE"].Value.ToString();


                    var stokListesi = dbContext.STOKLAR
                        .Where(ch => ch.sto_kod.EndsWith(name) && ch.sto_siparis_sure == 0)
                        .ToList();

                    if (stokListesi.Any()) // Eğer kayıt varsa
                    {
                        // 'GÜN' değerini short (Int16) olarak dönüştür
                        if (short.TryParse(row.Cells["SURE"].Value?.ToString(), out short agirlik))
                        {
                            foreach (var stok in stokListesi)
                            {
                                // sto_siparis_sure değerini güncelle
                                stok.sto_siparis_sure = agirlik;
                            }

                            // Değişiklikleri kaydet
                            dbContext.SaveChanges();
                        }
                        else
                        {
                            MessageBox.Show("Hata: 'GÜN' değeri geçerli bir sayı değil!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            MessageBox.Show("GÜNCELLEME TAMAMLANDI.");

        }

        private void button10_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Cells["ROTALAR"].Value != null)
                {
                    string name = row.Cells["ROTALAR"].Value.ToString();
                    string markakod = row.Cells["SURE"].Value.ToString();


                    var stokListesi = dbContext.STOKLAR
                        .Where(ch => ch.sto_kod==name && ch.sto_siparis_sure == 0)
                        .ToList();

                    if (stokListesi.Any()) // Eğer kayıt varsa
                    {
                        // 'GÜN' değerini short (Int16) olarak dönüştür
                        if (short.TryParse(row.Cells["SURE"].Value?.ToString(), out short agirlik))
                        {
                            foreach (var stok in stokListesi)
                            {
                                // sto_siparis_sure değerini güncelle
                                stok.sto_siparis_sure = agirlik;
                            }

                            // Değişiklikleri kaydet
                            dbContext.SaveChanges();
                        }
                        else
                        {
                            MessageBox.Show("Hata: 'GÜN' değeri geçerli bir sayı değil!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            MessageBox.Show("GÜNCELLEME TAMAMLANDI.");
        }
    }
}
