using ExcelDataReader;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using SİPARİŞ_GİRİŞ.Context;
using SİPARİŞ_GİRİŞ.Tables;
using System.Data;
using System.IO;
using System.Linq;
using WinFormsApp1.Context;
using WinFormsApp1.Tables;
using Zuby.ADGV;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace SİPARİŞ_GİRİŞ
{
    public partial class BOM_KONTROL : Form
    {
        private MyDbContext dbContext;
        private TLBContext TLBdbContext;
        public BOM_KONTROL()
        {
            dbContext = new MyDbContext();
            TLBdbContext = new TLBContext();
            InitializeComponent();
        }
        List<Dictionary<string, object>> customerList = new List<Dictionary<string, object>>();
        private void button3_Click(object sender, EventArgs e)
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
                                    customerList.Add(item);
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
        List<Dictionary<string, object>> lrfList = new List<Dictionary<string, object>>();
        private void button2_Click(object sender, EventArgs e)
        {

            // Eşleşmeyenleri bul
            var eslesmeyenler = customerList
                .Where(customer =>
                    !lrfList.Any(e =>
                        e.ContainsKey("item_sirano") && // Anahtarın varlığını kontrol et
                        e["item_sirano"].ToString() == customer["Item"].ToString()))
                .ToList();

            // Eşleşmeyenleri AdvancedDataGridView'e ekle
            advancedDataGridView1.DataSource = null; // Mevcut verileri temizle
            advancedDataGridView1.DataSource = eslesmeyenler.Select(x => new
            {
                Item = x["Item"],
                PartNumber = x["Part Number"],
                Title = x["Title"],
                PartNumberERP = lrfList.FirstOrDefault(e =>
                    e.ContainsKey("item_sirano") && // Anahtarın varlığını kontrol et
                    e["item_sirano"].ToString() == x["Item"].ToString())?["item_partnumber"] // Eşleştir listesinden item_partnumber değerini al
            }).ToList();

            MessageBox.Show($"{eslesmeyenler.Count} eşleşmeyen kayıt bulundu.");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int lrf = Convert.ToInt32(textBox1.Text);
            string ismrikodu = textBox2.Text;
            var selectedDate1 = dateTimePicker2.Value.Date;
            var selectedDate = dateTimePicker1.Value.Date;
            int Durum = Convert.ToInt32(comboBox1.Text.ToString());

            // Eşleştirilecek verileri al
            var eşleştir = dbContext.zz_bom
                .Where(x => x.dsref == lrf && (x.item_partnumber.StartsWith("02.") || x.item_partnumber.StartsWith("01.")))
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                })
                .ToList();

            var query = dbContext.URETIM_MALZEME_PLANLAMA
                .Where(ump => string.IsNullOrEmpty(ump.upl_urstokkod)) // upl_urstokkodu boş olanları filtrele
                .Join(dbContext.ISEMIRLERI
                    .Where(x => x.is_Kod.Substring(0, 8) == ismrikodu &&
                                 x.is_create_date.Date >= selectedDate &&
                                 x.is_create_date.Date <= selectedDate1 &&
                                 x.is_EmriDurumu == Durum),
                    ump => ump.upl_isemri,
                    isem => isem.is_Kod,
                    (ump, isem) => new { UretimMalzeme = ump, IsEmri = isem })
                .Select(x => new
                {
                    x.UretimMalzeme.upl_kodu,
                    x.IsEmri.is_Guid,
                })
                .Join(dbContext.ISEMIRLERI_USER
                    .Where(x => x.is_emri_tipi == "KK_IE"),
                    combined => combined.is_Guid,
                    usr => usr.Record_uid,
                    (combined, usr) => new { combin = combined, user = usr })
                .Select(x => new
                {
                    x.combin.upl_kodu
                })
                .ToList();




            // Eşleştir ve query listesini karşılaştır
            var olmayanlar = eşleştir
                .Where(e => !query.Any(q => q.upl_kodu == e.item_partnumber)) // item_partnumber ile eşleştir
                .ToList();

            // AdvancedDataGridView'ye verileri ekle
            advancedDataGridView1.DataSource = olmayanlar.Select(o => new
            {
                o.item_partnumber,

            }).ToList();
        }

        // Satın alma talepleri için bir liste tanımlayın
        List<SATINALMA_TALEPLERI1> satınalmaList = new List<SATINALMA_TALEPLERI1>();
        private void button4_Click(object sender, EventArgs e)
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
        // ERPList'i tanımla ve boş bir liste olarak başlat
        List<string> ERPList = new List<string>();
        private void button5_Click(object sender, EventArgs e)
        {
            // ERPList'i yeni bir liste yapısına dönüştürüyoruz
            List<ERPItem> ERPList = new List<ERPItem>();

            foreach (var item in lrfList)
            {
                // Boş veya null kontrolü
                if (string.IsNullOrEmpty(item["item_partnumber"]?.ToString())) continue;

                // Yeni ERPItem nesnesini oluştur
                var erpItem = new ERPItem
                {
                    EksikOlan = item["fason"]?.ToString() == "EVET" ? item["item_partnumber"]?.ToString() : item["hammadde_erp_kodu"]?.ToString(),
                    BagliParca = item["item_partnumber"]?.ToString(),
                    Title = item["title"]?.ToString(),
                    fason = item["fason"]?.ToString(),
                    item_qty = item.ContainsKey("Adet") ? Convert.ToInt32(item["Adet"]) : 0, // Adet'e erişim
                    MalzemeTanim = item["malzemetanım"]?.ToString(),
                    ItemSirano = item["item_sirano"]?.ToString(),
                };

                // Koşullara göre ERPList'e ekleme
                if (item["item_partnumber"].ToString().StartsWith("02."))
                {
                    if (item["fason"]?.ToString() == "HAYIR" && !item["rota"].ToString().StartsWith("KAYNAK"))
                    {
                        ERPList.Add(erpItem);
                    }
                    else if (item["fason"]?.ToString() == "EVET")
                    {
                        ERPList.Add(erpItem);
                    }
                }
                else
                {
                    ERPList.Add(erpItem);
                }
            }



            // Eşleştir ve query listesini karşılaştır
            var olmayanlar = ERPList
                .Where(e => !satınalmaList.Any(q => q.stl_Stok_kodu == e.EksikOlan)) // Eksik olan ile eşleştir
                .ToList();

            // AdvancedDataGridView'ye verileri ekle
            advancedDataGridView1.DataSource = olmayanlar.Where(x => x.EksikOlan != "").Select(o => new
            {
                ItemSirano = o.ItemSirano,
                EksikOlan = o.EksikOlan,
                BagliParca = o.BagliParca,
                Title = o.Title,
                fason = o.fason,
                item_qty = o.item_qty,
                MalzemeTanım = o.MalzemeTanim,
            }).Distinct().ToList();

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

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ERPList.Clear();
            customerList.Clear();
            textBox1.Clear();
            textBox2.Clear();
            // AdvancedDataGridView'deki verileri temizlemek için
            advancedDataGridView1.DataSource = null;
            advancedDataGridView1.Rows.Clear(); // Satırları da temizlemek için


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                button6.Visible = true;
            }
            else if (checkBox1.Checked == false)
            {
                button6.Visible = false;
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            int lrf = Convert.ToInt32(textBox1.Text);

            // Eşleştirilecek verileri al
            var eşleştir = dbContext.zz_bom
                .Where(x => x.dsref == lrf)
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                })
                .ToList();

            // item_partnumber'ları aynı olan ama title'ları birebir eşleşmeyenleri bul
            var eslesmeyenler = eşleştir
                .GroupBy(x => x.item_partnumber)
                .Where(g => g.Select(x => x.title).Distinct().Count() > 1)
                .SelectMany(g => g.Where(x =>
                    g.Any(y => x.title != y.title)) // Title'lar birebir eşleşmiyor
                )
                .ToList();

            // Eşleşmeyenleri AdvancedDataGridView'e ekle
            advancedDataGridView1.DataSource = null; // Mevcut verileri temizle
            advancedDataGridView1.DataSource = eslesmeyenler.Select(x => new
            {
                ItemPartNumber = x.item_partnumber,
                ItemSirano = x.item_sirano,
                Title = x.title
            }).ToList();




        }


        private void button7_Click(object sender, EventArgs e)
        {
            int lrf = Convert.ToInt32(textBox1.Text);

            // Eşleştirilecek verileri al
            var eşleştir = dbContext.zz_bom
                .Where(x => x.dsref == lrf)
                .Select(x => new
                {
                    item_partnumber = x.item_partnumber ?? "Varsayılan Değer",
                    item_sirano = x.item_sirano ?? "Varsayılan Değer",
                    title = x.title ?? "Varsayılan Değer",
                    hammadde_erp_kodu = x.hammadde_erp_kodu ?? "Varsayılan Değer",
                    fason = x.fason ?? "Varsayılan Değer",
                    rota1 = x.rota1 ?? "Varsayılan Değer",
                    item_qty = x.item_qty, // Eğer item_qty null olabiliyorsa varsayılan olarak 0 atandı
                    malzemetanim = x.malzemetanim ?? "Varsayılan Değer"
                })
                .ToList();


            foreach (var item in eşleştir)
            {
                var dictionary = new Dictionary<string, object>
                {
                    { "item_partnumber", item.item_partnumber ?? "Varsayılan Değer" },
                    { "item_sirano", item.item_sirano ?? "Varsayılan Değer" },
                    { "title", item.title ?? "Varsayılan Değer" },
                    { "hammadde_erp_kodu", item.hammadde_erp_kodu ?? "Varsayılan Değer" },
                    { "fason", item.fason ?? "Varsayılan Değer" },
                    { "rota", item.rota1 ?? "Varsayılan Değer" },
                    { "Adet", item.item_qty },
                    { "malzemetanım", item.malzemetanim ?? "Varsayılan Değer" }
                };
                lrfList.Add(dictionary);
            }
            MessageBox.Show("Kaydedildi...");
            // Artık lrfList kullanıma hazır

        }

        private void button8_Click(object sender, EventArgs e)
        {
            var siranoList = customerList
              .Select(s => Convert.ToInt32(((Dictionary<string, object>)s)["sirano"]))
              .ToList();

            // Satın alma taleplerini tutacak liste
            var satınalmaListesi = new List<SATINALMA_TALEPLERI1>();

            // Her bir sirano için veritabanından satın alma taleplerini al
            foreach (var sirano in siranoList)
            {
                var satınalma = TLBdbContext.SATINALMA_TALEPLERI1
                    .Where(x => x.stl_evrak_sira == sirano)
                    .ToList();

                // Satın alma taleplerini satınalmaList'e ekleyin
                satınalmaListesi.AddRange(satınalma);
            }

            // Satın alma taleplerini satınalmaList'e ekleyin
            satınalmaList.AddRange(satınalmaListesi);

            // Mesaj kutusunu göster
            MessageBox.Show("Eklendi...");
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                groupBox1.Visible = true;
                button2.Visible = true;
            }
            else if (checkBox2.Checked == false)
            {
                groupBox1.Visible = false;
                button2.Visible = false;
            }

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                groupBox1.Visible = true;
                button2.Visible = true;
                groupBox2.Visible = true;
                button1.Visible = true;
            }
            else if (checkBox3.Checked == false)
            {
                groupBox1.Visible = false;
                button2.Visible = false;
                groupBox2.Visible = false;
                button1.Visible = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                groupBox1.Visible = true;
                button2.Visible = true;
                groupBox3.Visible = true;
                button5.Visible = true;
            }
            else if (checkBox4.Checked == false)
            {
                groupBox1.Visible = false;
                button2.Visible = false;
                groupBox3.Visible = false;
                button5.Visible = false;
            }
        }
    }
}
